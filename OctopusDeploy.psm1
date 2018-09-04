#----------------------------------------------------------------------------------
# Module OctopusDeploy.PSM1
#
# Cmdlets geared towards Octopus Deploy
#
# Author: Jeff Buenting
#----------------------------------------------------------------------------------

#----------------------------------------------------------------------------------
# Client Cmdlets
#----------------------------------------------------------------------------------

Function Install-OctoClientTentacle {

<#
    .Synopsis
        Deploy Octopus Client Tentacle

    .Description
        Installs and Configures an Octopus Deploy Client Tentacle.

    .Parameter ComputerName
        Name of the computer where the client will be installed

    .Parameter Path
        Source path for the Octopus Tentacle install.

    .Parameter OctopusServerThumbprint
        Thumbprint from the Octopus Deploy Server

    .Parameter OctopusUR
        URI Web Address of the Octopus Deploy Server

    .Parameter OctopusAPIKey
        Octopus Deploy API Key

    .Parameter OctopusEnvironemtn
        Environment Name this client should be added to.

    .Parameter OctopusRoles
        Roles this client will be assigned.

    .Example
        Install client

        install-OctoClientTentacle -ComputerName $ComputerName -Path "c:\temp\Octopus.Tentacle.3.22.0-x64.msi" -OctopusServerThumbprint $octoThumb -octopusURI $OctopusURI -octopusApiKey $OctoAPIKey -OctopusEnvironment 'JB04' -OctopusRoles 'Web' -verbose

    .Link
        https://octopus.com/docs/infrastructure/windows-targets/automating-tentacle-installation

    .Notes
        Author : Jeff Buenting
        Date : 2018 Aug 23
        
#>

    [CmdletBinding()]
    Param (
        [Parameter( ValueFromPipeline = $True )]
        [String[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter( Mandatory = $True )]
        [String]$Path,

        [Parameter( Mandatory = $True )]
        [String]$OctopusServerThumbprint,

        [Parameter( Mandatory = $True )]
        [String]$octopusURI, 
        
        [Parameter( Mandatory = $True )]
        [String]$octopusApiKey,

        [Parameter( Mandatory = $True )]
        [String]$OctopusEnvironment,

        [Parameter( Mandatory = $True )]
        [String[]]$OctopusRoles
    )

    Process {
        Foreach ( $C in $ComputerName ) {
            Write-Verbose "Installing Octopus Deploy Client Tentacle on $C"

            Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                $VerbosePreference = $Using:VerbosePreference

                Try {
                    Write-Verbose "Installing Octopus Tentacle Client"
                    Start-Process -FilePath msiexec -argumentlist "/i $Using:Path /quiet" -wait -ErrorAction Stop
                }
                Catch {
                    $EXceptionMessage = $_.Exception.Message
                    $ExceptionType = $_.exception.GetType().fullname
                    Throw "Install-OctopusClientTentacle: Error installing Tentacle.`n`n     $ExceptionMessage`n`n     Exception : $ExceptionType"  
                }

                Try {
                    Write-Verbose "Configuring Octopus Tentacle Client"
                    Start-Process -FilePath "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" -ArgumentList "create-instance --instance ""Tentacle"" --config ""C:\Octopus\Tentacle.config""" -wait -ErrorAction Stop
                    Start-Process -FilePath "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" -ArgumentList "new-certificate --instance ""Tentacle"" --if-blank" -Wait -ErrorAction Stop
                    Start-Process -FilePath "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" -ArgumentList "configure --instance ""Tentacle"" --reset-trust" -wait -ErrorAction Stop
                    Start-Process -FilePath "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" -ArgumentList "configure --instance ""Tentacle"" --app ""C:\Octopus\Applications"" --port ""10933"" --noListen ""False""" -wait  -ErrorAction Stop
                    Start-Process -FilePath "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" -ArgumentList "configure --instance ""Tentacle"" --trust ""$Using:OctopusServerThumbprint""" -Wait -ErrorAction Stop
    
                    Write-Verbose "Creating Firewall Rule for Octopus Tentacle"
                    if ( -Not ( Get-NetFirewallRule -Name "Octopus Deploy Tentacle" -ErrorAction SilentlyContinue ) ) {
                        New-NetFirewallRule -Name "Octopus Deploy Tentacle" -DisplayName "Octopus Deploy Tentacle" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 10933 -ErrorAction Stop 
                    }

                    Write-Verbose "Restarting Octopus Tentacle"
                    Start-Process -FilePath "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" -ArgumentList "service --instance ""Tentacle"" --install --stop --start" -Wait -ErrorAction Stop
                }
                Catch {
                    $EXceptionMessage = $_.Exception.Message
                    $ExceptionType = $_.exception.GetType().fullname
                    Throw "Install-OctopusClientTentacle: Error Configuring Tentacle.`n`n     $ExceptionMessage`n`n     Exception : $ExceptionType"  
                }

                Write-Verbose "Get the Client Thumbprint from the log"
                $OctoLog = Get-Content c:\Octopus\Logs\OctopusTentacle.txt 

                # ----- parse the log and look for certificate generation error
                $OctoError = (($octolog | Select-String -Pattern 'FATAL  No certificate has been generated for this Tentacle. Please run the new-certificate command before starting.' ) -split 'FATAL')[-1]

                if ( $OctoError ) {
                    Write-Warning "Install-OctopusClientTentacle : $OctoError"

                    Start-Process -FilePath "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" -ArgumentList "new-certificate --instance ""Tentacle"" --if-blank" -Wait -ErrorAction Stop
                    Start-Process -FilePath "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" -ArgumentList "service --instance ""Tentacle"" --install --stop --start" -Wait -ErrorAction Stop
                }
                
                # ----- Reload the log file
                $OctoLog = Get-Content c:\Octopus\Logs\OctopusTentacle.txt 

                # ----- parse the log, From the line after the one containing A new certificate has been generated, select the thumbprint after the last space.
                $ClientThumb = ($OctoLog[($OctoLog | Select-string -Pattern 'A new certificate has been generated').LineNumber].split( ' '))[-1]

                Write-verbose "Register client with Octopus Server"

                Add-Type -Path 'C:\Program Files\Octopus Deploy\Tentacle\Newtonsoft.Json.dll'
                Add-Type -Path 'C:\Program Files\Octopus Deploy\Tentacle\Octopus.Client.dll'

                $endpoint = new-object Octopus.Client.OctopusServerEndpoint $Using:octopusURI, $Using:octopusApiKey
                $repository = new-object Octopus.Client.OctopusRepository $endpoint

                $tentacle = New-Object Octopus.Client.Model.MachineResource

                $tentacle.name = "Tentacle registered from client"

                $OctoEnv = $repository.Environments.FindByName($Using:OctopusEnvironment)
                
                $tentacle.EnvironmentIds.Add($OctoEnv.ID)

                # ----- Loop and add all listed Roles
                Foreach ( $R in $Using:OctopusRoles ) {
                    $tentacle.Roles.Add($R)
                }

                $tentacleEndpoint = New-Object Octopus.Client.Model.Endpoints.ListeningTentacleEndpointResource
                $tentacle.EndPoint = $tentacleEndpoint
                $tentacle.Endpoint.Uri = "https://$($Using:ComputerName):10933"
                $tentacle.Endpoint.Thumbprint = $ClientThumb

                $repository.machines.create($tentacle)
            }
        }
    }
}


#----------------------------------------------------------------------------------
# Module OctopusDeploy.PSM1
#
# Cmdlets geared towards Octopus Deploy
#
# Author: Jeff Buenting
#----------------------------------------------------------------------------------

#----------------------------------------------------------------------------------
# Server Cmdlets
#----------------------------------------------------------------------------------

Function Connect-ODServer {

<#
    .Synopsis
        Connects to an Octopus Deploy environment

    .Description
        Initial connection to the Octopus Deploy Server so further data can be extracted.

    .Parameter OctopusDLL
        Location of the Octopus DLL assemblies

    .Parameter APIKey
        API Key unique to your environment.  Allows the server to know that you are authorized

    .Parameter URI
        Octopus Deploy URI

    .Link
        https://dalmirogranias.wordpress.com/2014/09/19/octopus-api-and-powershell-getting-the-libraries-and-connecting-to-octopus/
#>

    [CmdletBinding()]
    Param (
        [String]$OctopusDLL,

        #[String]$APIKey = 'API-N7TOWCVGG3SMVNIKUDDIMB1ITS',
        [String]$APIKey = 'API-TMTEQFVCCCI6URHQYJAAOQHJKC',

        [String]$URI = 'http://octopus2012/api'
    )

    # ----- Adding Octopus Deploy Libraries

    if ( [String]::IsNullOrEmpty($OctopusDLL) ) {
            Write-Verbose "OctopusDLL Path is empty."

            if ( Test-path -Path "C:\Program Files\Octopus Deploy\Tentacle" ) {
                Write-Verbose "Loading DLLs from Octopus Tentacle"
                $DLLPath = "C:\Program Files\Octopus Deploy\Tentacle"
            }

            if ( Test-Path -path "C:\Program Files\Octopus Deploy\Octopus" ) {
                Write-Verbose "Loading DLLs from Octopus Server"
                $DLLPath = "C:\Program Files\Octopus Deploy\Octopus"
            }
        }
        Else {
            Write-Verbose "Loading Octopus DLLs from $OctopusDLL"
            $DLLPath = $OctopusDLL
    }
    
    # ----- Load DLL Assemblies
    if ( $DLLPath -ne $Null ) {
            Write-verbose "Loading DLLs"
            Add-Type -Path "$DLLPath\Octopus.Client.dll"
            Add-Type -Path "$DLLPath\Octopus.Platform.dll"
            Add-Type -Path "$DLLPath\Newtonsoft.json.dll"
        }
        Else {
            Write-Verbose "ERROR Throwing Exception"
            Throw "Octopus DLLs do not exist at specified location.`n`nVerify path and try again."
            
    }

    # ----- Creating a connection
    $endpoint = new-object Octopus.Client.OctopusServerEndpoint $URI,$apikey
    $repository = new-object Octopus.Client.OctopusRepository $endpoint

    Write-Output $repository

}

#----------------------------------------------------------------------------------

Function Get-ODEnvironment {

<# 
    .Synopsis
        Gets Octopus Deploy Environments.

    .Description
        Gets Octopus Deploy Environments.  These are what defines the group of computers used to deploy Octopus Deploy project deployments.

    .Parameter OctopusConnection
        Connection to the Octopus Deploy Server established by executing Connect-ODServer. 
        
    .Parameter Name
        Name of the environment to return.  If not provided then all environments will be returned.  Wildcards are supported.
        
    .Example
        Returns all environments

        $Connection = Connect-ODServer -OctopusDLL 'F:\OneDrive - StratusLIVE, LLC\Scripts\OctopusDeploy\OctopusDLLs'

        Get-ODEnvironment -OctopusConnection $Connection 

    .Example
        Returns all environments beginning with S

        $Connection = Connect-ODServer -OctopusDLL 'F:\OneDrive - StratusLIVE, LLC\Scripts\OctopusDeploy\OctopusDLLs'

        Get-ODEnvironment -OctopusConnection $Connection -Name "s*" 

    .Example
        Return the enviroment named Lab

        $Connection = Connect-ODServer -OctopusDLL 'F:\OneDrive - StratusLIVE, LLC\Scripts\OctopusDeploy\OctopusDLLs' -Verbose

        Get-ODEnvironment -OctopusConnection $Connection -Name "Lab" -verbose

    .Note
        Author: Jeff Buenting
        Date: 2015 DEC 09
#>

    [CmdletBinding(DefaultParameterSetName='Default')]
    Param (
        [Parameter(ParameterSetName='Default',Mandatory=$True)]
        [Parameter(ParameterSetName='Name',Mandatory=$True)]
        [Octopus.Client.OctopusRepository]$OctopusConnection,

        [Parameter(ParameterSetName='Name')]
        [String[]]$Name
    )
    
    Switch ( $PSCmdlet.ParameterSetName ) {
        'Name' {
            Write-Verbose "ParameterSet: Name"
          
            if ( ( $Name | Measure-Object ).count -eq 1 ) {
                    
                    # ----- Because $Name is a array of Strings, the contains method does not work correctly.  Need to reference the first element of the array ( which is the only
                    # ----- element because count is 1) to use the contains method.
                    if ( $Name[0].contains('*') ) {
                            Write-Verbose "Finding Wildcards"
                            Write-Output ( $OctopusConnection.Environments.FindAll() | where Name -Like $Name )
                        }
                        else {
                            Write-verbose "Finding Environments: $Name"
                            Write-Output $OctopusConnection.Environments.FindByName( $Name )
                    }
                }
                else {
                    Write-Verbose "Finding Environment: $($Name | out-string) "
                    Write-Output $OctopusConnection.Environments.FindByNames( $Name )
            }
        }

        'Default' {
            Write-Verbose "Finding all Environments"
            $OctopusConnection.Environments.FindAll()
        }
    }
}

#----------------------------------------------------------------------------------

Function Get-ODMachine {

<#
    .Synopsis
        Returns Octopus Deploy Machine.

    .Description
        Returns Octopus Deploy Machines.  

    .Parameter OctopusConnection
         Connection to the Octopus Deploy Server established by executing Connect-ODServer. 
        
    .Parameter Name
        Name of the Machine to return.  If not provided then all Machines will be returned.  Wildcards are supported.

    .Parameter Environment
        Octopus Deploy Environment to get machines from.

    .Example Connection to the Octopus Deploy Server established by executing Connect-ODServer. 
        
    .Parameter Name
        Name of the environment to return.  If not provided then all environments will be returned.  Wildcards are supported.
        
    .Example
        Retrieves all of the machines that Octopus Deploy knows about
        
        $Connection = Connect-ODServer -OctopusDLL 'F:\OneDrive - StratusLIVE, LLC\Scripts\OctopusDeploy\OctopusDLLs' 
        Get-ODMachine -OctopusConnection $Connection 

    .Example
        Retrieves all of the machines in the Lab Environment.
        
        $Connection = Connect-ODServer -OctopusDLL 'F:\OneDrive - StratusLIVE, LLC\Scripts\OctopusDeploy\OctopusDLLs' 
        Get-ODEnvironment -OctopusConnection $Connection -Name Lab | Get-ODMachine -OctopusConnection $Connection 

    .Example
        Retrieves the machine named lab1
        
        $Connection = Connect-ODServer -OctopusDLL 'F:\OneDrive - StratusLIVE, LLC\Scripts\OctopusDeploy\OctopusDLLs' 
        Get-ODMachine -OctopusConnection $Connection -Name Lab1

    .Note
        Author: Jeff Buenting
        Date: 2015 DEC 09
   
          
#>

    [CmdletBinding(DefaultParameterSetName='Default')]
    Param (
        [Parameter(ParameterSetName='Default',Mandatory=$True)]
        [Parameter(ParameterSetName='Name',Mandatory=$True)]
        [Parameter(ParameterSetName='Environment',Mandatory=$True)]
        [Octopus.Client.OctopusRepository]$OctopusConnection,

        [Parameter(ParameterSetName='Name',Mandatory=$True,ValueFromPipeline=$True)]
        [String[]]$Name,

        [Parameter(ParameterSetName='Environment',Mandatory=$True,ValueFromPipeline=$True)]
        [Octopus.Client.Model.EnvironmentResource[]]$Environment
    )
    
    Process {
        Switch ( $PSCmdlet.ParameterSetName ) {
            'Name' {
                Write-Verbose "Returning Machines by Name"
          
                if ( ( $Name | Measure-Object ).count -eq 1 ) {
                    
                        # ----- Because $Name is a array of Strings, the contains method does not work correctly.  Need to reference the first element of the array ( which is the only
                        # ----- element because count is 1) to use the contains method.
                        if ( $Name[0].contains('*') ) {
                                Write-Verbose "Finding Wildcards"
                                Write-Output ( $OctopusConnection.Machines.FindAll() | where Name -Like $Name )
                            }
                            else {
                                # ----- Returning single machine
                                Write-verbose "Finding Machine: $Name"
                                Write-Output $OctopusConnection.Machines.FindByName( $Name )
                        }
                    }
                    else {
                        # ----- Name is an array of machine names.  Returning the names provided
                        Write-Verbose "Finding Machines: $($Name | out-string) "
                        Write-Output $OctopusConnection.Machines.FindByNames( $Name )
                }
            }

            'Environment' {
                Write-Verbose "Returning Environment Machines"
                foreach ( $E in $Environment ) {
                    Write-Verbose "     Environment: $E.Name"
                    Write-Output ($OctopusConnection.machines.FindAll() | where EnvironmentIDs -contains $E.Id)
                }
            }

            'Default' {
                Write-Verbose "Returning all Machines"
                Write-output $OctopusConnection.machines.FindAll()
            }
        }
    }
}

#----------------------------------------------------------------------------------
# Service Cmdlets
#----------------------------------------------------------------------------------

Function Set-ODService {

<#
    .Synopsis
        Change the OctopusDeploy Tentacle Service Properties.

    .Description
        Changes the OctopusDeploy Tentacle Service properties depending on what is supplied in the parameters.

    .Parameter ComputerName
        Name of the computer running the service

    .parameter Credential
        Logon credentials to change on the service

    .Example
        change the logon credentials for the OctopusDeploy Service on ServerA

        Set-ODService -Computername ServerA -Credential $Cred

    .Link
        https://4sysops.com/archives/managing-services-the-powershell-way-part-8-service-accounts/
        
#>

    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline=$True)]
        [String[]]$ComputerName = "LocalHost",

        [Alias("Cred")]
        [pscredential]$Credential
    )

    Process {
        Foreach ( $C in $ComputerName ) {
            Write-Verbose "Changeing OctopusDeploy Tentacle Service settings on $C"            
            
            # ----- Have to use Get-CIMInstance instead of Get-Service because we need to change the Logon user.  Get-Service does not allow that.  
            $Service = Get-CimInstance -ComputerName $C -ClassName Win32_Service -Filter "Name = 'OctopusDeploy Tentacle'"
            
            If ( $Credential -ne $Null ) {
                Write-Verbose "Setting Logon Credentials"
                $Service | Invoke-CIMMethod -Name Change -Arguments @{StartName=$Credential.UserName;StartPassword=($Credential.GetNetworkCredential()).Password}
                #$Service.Change($null,$null,$null,$null,$null,$null,$Credential.UserName,($Credential.GetNetworkCredential()).Password)
            }

            Write-Verbose "Restarting Service"
            $Service | Invoke-CIMMethod -Name StopService
            $Service | Invoke-CimMethod -Name StartService
        }  
    }
}
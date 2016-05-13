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

$OctopusDeploy = Connect-ODServer -OctopusDLL 'F:\OneDrive - StratusLIVE, LLC\Scripts\OctopusDeploy\OctopusDLLs' -Verbose

$OctopusDeploy.users.findall()
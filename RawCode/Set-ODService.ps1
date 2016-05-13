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

$Cred = Get-Credential

Set-ODService -ComputerName jeffb-rb03 -Credential $Cred
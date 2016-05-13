import-module 'F:\OneDrive - StratusLIVE, LLC\Scripts\Modules\OctopusDeploy\OctopusDeploy.psm1'

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

$Connection = Connect-ODServer -OctopusDLL 'F:\OneDrive - StratusLIVE, LLC\Scripts\OctopusDeploy\OctopusDLLs' -Verbose

$E = Get-ODEnvironment -OctopusConnection $Connection -Name "J*" -verbose
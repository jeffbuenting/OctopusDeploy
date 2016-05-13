import-module 'F:\OneDrive - StratusLIVE, LLC\Scripts\Modules\OctopusDeploy\OctopusDeploy.psm1'

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

$Connection = Connect-ODServer -OctopusDLL 'F:\OneDrive - StratusLIVE, LLC\Scripts\OctopusDeploy\OctopusDLLs' -Verbose

$E = Get-ODEnvironment -OctopusConnection $Connection -verbose 
Get-ODMachine -OctopusConnection $Connection -name 'jeffb-iis03','jeffb-rb03' -Verbose
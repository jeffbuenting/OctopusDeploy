<#
    .Synopsis
        Imports solutions into CRM org

    .Description
        Called from Octopus.  Rotates thru each Org in an environment and utilizing all CRM frontends (called via octopus) it will import the new solutions
#>


$SLConfigPath = '\\vaslnas\Deploys\slconfigs'


#---------------------------------------------------------------------------------------
# Publish Duplicate Detection Rules for an organization
#---------------------------------------------------------------------------------------

function Publish-CrmDuplicateRules ([Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn, [guid[]]$RuleIds) {
    # Helper function to create PublishDuplicateRuleRequests
    function Create-PublishRuleRequest {
        Process {
            $newrequest = new-object Microsoft.Crm.Sdk.Messages.PublishDuplicateRuleRequest
            $newrequest.DuplicateRuleId = $_
            $newrequest
        }
    }

    $RuleIds | Create-PublishRuleRequest | % { 
        [Microsoft.Crm.Sdk.Messages.PublishDuplicateRuleResponse]$conn.ExecuteCrmOrganizationRequest($_, $null) 
    }
}

#---------------------------------------------------------------------------------------

# Import the xRM CI Framework Dynamics CRM Cmdlets
Import-Module “C:\Program Files (x86)\Xrm CI Framework\CRM 2011\PowerShell Cmdlets\Xrm.Framework.CI.PowerShell.dll”

if ($CRMType -eq "365IFD") {
    Import-Module “$SLConfigPath\Powershell\Modules\Xrm.Framework.CI.PowerShell.Cmdlets\Xrm.Framework.CI.PowerShell.Cmdlets.dll”
}
    else {
    Import-Module “C:\Program Files (x86)\Xrm CI Framework\CRM 2011\PowerShell Cmdlets\Xrm.Framework.CI.PowerShell.dll”
}
                
# import the crm powershell module if it is not already loaded
if (-Not (Get-Module -Name Microsoft.Xrm.Data.Powershell)) {
    Import-Module $SLConfigPath\Powershell\Modules\Microsoft.Xrm.Data.Powershell
}

#Set Basic Variables

#$CRMEnvironment = $OctopusParameters['Octopus.Environment.Name']
$CRMEnvironment = 'QA4' 

# ----- Number of CRM Front end servers in the environment.  Ideally this should come from $OctopusParameters['Octopus.Environment.MachinesinRoles[CRM]']  But I can'f figure out how to get this to work
$CRMServers = 'QA4' | Sort-Object
$CRMServerCount = $CRMServers.count

# ----- Extract which server we are on by the last digit in its name.  This is used to increment the Orgs index.
$ServerNum = $CRMServers.IndexOf($env:COMPUTERNAME)



$ErrorActionPreference = "Stop"
$SLSolutions=@("StratusLiveSolution","VolunteerandEventsSolution","PlannedGivingModule")

#$srcpath = $OctopusParameters['Octopus.Tentacle.Agent.ApplicationDirectoryPath'] + "\" + $OctopusParameters['Octopus.Environment.Name'] + "\StratusLive.Crm.Installer"
$srcpath = "C:\Octopus\Applications\QA4\StratusLive.Crm.Installer"

#$SrcRoot=  ($srcpath + "\" + $OctopusParameters['Octopus.Release.Number'])
$SrcRoot=  ($srcpath + "\7.2.3.2090")

#$SLConfigFilesPath = $SLConfigPath + "\" + $OctopusParameters['Octopus.Environment.Name']
$SLConfigFilesPath = $SLConfigPath + "\QA4"

Write-Output "WHoAmI?"
Whoami


# ----- Copy the installer config file to the install directory
Write-Output "SLConfigFilesPath = $SLConfigFilesPath"
Write-Output "SrcRoot = $SrcRoot "
if ( Test-Path -Path $SLConfigFilesPath\StratusLive.Crm.Installer.exe.config ) {
        Write-Output "Copying Config files"
        Copy-item -Path $SLConfigFilesPath\StratusLive.Crm.Installer.exe.config -Destination $SrcRoot\StratusLive.Crm.Installer.exe.config -Force
    }
    else {
        Write-Warning "StratusLive.Crm.Installer.exe.config file does not exist for this Environment: $CRMEnvironment"
        Exit -1
}

# ----- Get a list of Orgs installed on this CRM instance
# Pull connection string from mscrm key in registry
$Regkey = Get-ItemProperty -path hklm:SOFTWARE\Microsoft\MSCRM  -Name configdb

#Get SQL Server from connection string by dropping everything after the first semicolon then everything before the first equal sign. 
$CRMDB = $Regkey.configdb -replace ";.*" -replace ".*="

#Connect to SQL
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
$serverInstance = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $CRMDB

#Pull List of Active Orgs From MSCRM_Config table
$db = $serverInstance.Databases["MSCRM_Config"]
$ds = $db.ExecuteWithResults("SELECT UniqueName FROM dbo.Organization where state = 1")
$Orgs = $ds.Tables.UniqueName

#Write the Active Orgs to Console
write-Output "The following Orgs are active on this server:"
write-Output $orgs

#Attach to SQL Server with SLConfig Information
$ServerInstance = $Null
$serverInstance = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $SLConfigSQLServer

# ----- Adding this check to see if a connection to the SQL server was successfull.  If the Product property is not blank then the connection was successfull.
if ( -Not ($ServerInstance.Product) ) { 
    Write-Warning "`n`nConnection to SQL $SLConfigSQLServer Failed.  Pausing and will try again." 
    Start-Sleep -Seconds 300
    
    $ServerInstance = $Null
    $serverInstance = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $SLConfigSQLServer
    
    if ( -Not ($ServerInstance.Product) ) { 
        Write-Warning "Connection to SQL $SLConfigSQLServer Failed a second time.  Exiting."
        Exit -1
    }
}



# ----- Loop through Orgs selecting certain ones depending on CRMServerCount and ServerNum
For ( $I = $ServerNum; $I -le $Orgs.Count-1; $I += $CRMServerCount ) {
    Write-Output "-----"
    Write-Output "Processing $Orgs[$I]`n"

    #Get Data for current org
    $db = $Null
    $db = $serverInstance.Databases["SLDeployConfigInfo"]

    $LoaderVars = $Null
    $targetCrmConnectionUrl = $Null
    $PWCrmConnectionUrl = $Null
    $PScrmOrg = $Null
    
    Write-Output "DB = $($DB.Name)"
    Write-Output "Org = $($Orgs[$I])"
    Write-Output "Environment = $CRMEnvironment"
    
    $LoaderVars = ($db.ExecuteWithResults("SELECT * FROM dbo.LoaderVariables WHERE (Org='$Orgs[$I]') AND (Environment='$CRMEnvironment') ")).Tables

    # ---- Continue if $DeployConfigVar contains data
    if ( $LoaderVars.Port -ne $Null ) {
        Write-Output "Org data Found for URL: $($LoaderVars.URL)"

        #Define the CRM connection
        if ($CRMType -eq "IFD") {
            Write-Output "IFD Environment"
            $ifdcrmuser = "$($LoaderVars.domain)\$($LoaderVars.crmuser)"
            $targetCrmConnectionUrl = “ServiceUri=$($LoaderVars.URL)/XRMServices/2011/Organization.svc`;Timeout=02:00:00.0; Username=$ifdcrmuser; Password=$($LoaderVars.crmpassword)"
        }
        elseif ($CRMType -eq "365IFD") {
            Write-Output "IFD Environment"
            $365ifdcrmuser = "$($LoaderVars.crmuser)@$($LoaderVars.domain).com"
            $targetCrmConnectionUrl = “Username=$365ifdcrmuser;Password=$($LoaderVars.crmpassword);Domain=$($LoaderVars.crmserver);AuthType=IFD;RequireNewInstance=True;Url=$($LoaderVars.URL)/$($LoaderVars.org)"
        }
        elseif ($CRMType -eq "AD") {
            Write-Output "AD Environment"
            $targetCrmConnectionUrl = “url=$($LoaderVars.URL)/$($LoaderVars.crmorg)/XRMServices/2011/Organization.svc`;Timeout=02:00:00.0; Username=$($LoaderVars.crmuser); Password=$($LoaderVars.crmpassword); Domain=$($LoaderVars.domain)"
        }

        #Connect to Org for Dedupe Rules
        $PSCRMuser = "$($LoaderVars.crmuser)@$($LoaderVars.domain)"
        $PWCrmConnectionUrl = “Username=$PScrmuser;Password=$($LoaderVars.crmpassword);Domain=$($LoaderVars.crmserver);AuthType=IFD;Url=$($LoaderVars.URL)/$($LoaderVars.org)"
    
        # Create connection to CRM Organization
        $PScrmOrg = Get-Crmconnection -connectionstring $PWCrmConnectionUrl

        # Retrieve active DuplicateRule records
        $duplicaterules = Get-CrmRecords -conn $PScrmOrg 'duplicaterule' -FilterAttribute statecode -FilterOperator eq -FilterValue Active
        write-host "Getting List of enabled Duplicate Detection Rules"
       
        # Get the guid for DuplicateRuleId field
        $duplicateruleIds = $duplicaterules['CrmRecords'] | select -ExpandProperty duplicateruleid
        write-host "The following Duplicate Detection Rules are enabled"
        write-host $duplicateruleIds
       
        Write-host "Determine if auditing is enabled"
        $Auditingresult = Get-CrmRecords -conn $PScrmOrg -EntityLogicalName organization -Fields organizationid, isauditenabled
        $Auditingorg = $Auditingresult['CrmRecords'][0]
        write-host $Auditingorg.isauditenabled
       
        # ----- Upgrade Solution
        Foreach ( $Solution in $SLSolutions ) {
            # ----- Set Solution Path
            Switch ( $Solution ) {                
                "StratusLiveSolution" {
                    $ImportPath = $SrcRoot + "\solutions\stratuslivesolution_release.zip"
                }

                "VolunteerandEventsSolution" {
                    $ImportPath = $SrcRoot + "\solutions\engagementsolution_release.zip"
                }
                    
                "PlannedGivingModule" {
                    $ImportPath = $SrcRoot + "\solutions\plannedgivingsolution_release.zip"
                }
            }

            # Check if Solution is installed
            $InstalledSolution = Get-XrmSolution -ConnectionString $targetCrmConnectionUrl -UniqueSolutionName $Solution
           
            #if Solution is installed, Upgrade it
            if ($InstalledSolution -ne $null) {
                    Write-Output “Current $Solution version: $($InstalledSolution.Version)"
                    Write-Output “Importing Solution from: $ImportPath”
                    
                    #Import CRM Solution
                        #Import-XrmSolution -ConnectionString $targetCrmConnectionUrl -SolutionFilePath $ImportPath -PublishWorkflows $true -OverwriteUnmanagedCustomizations $true 

                    $importJobId = [guid]::NewGuid()  
  #                  $asyncOperationId = Import-XrmSolution -ConnectionString $targetCrmConnectionUrl -SolutionFilePath $ImportPath -publishWorkflows $true -overwriteUnmanagedCustomizations $true -SkipProductUpdateDependencies $false -ImportAsync $true -WaitForCompletion $true -ImportJobId $importJobId -AsyncWaitTimeout 3600 -ConvertToManaged $true
 
                    Write-Output "Solution Import Completed. Import Job Id: $importJobId"

                    if ($logsDirectory) {
                        $importLogFile = $logsDirectory + "\" + $solution + [System.DateTime]::Now.ToString("yyyy_MM_dd__HH_mm") + ".xml"
                    }

#                    $importJob = Get-XrmSolutionImportLog -ImportJobId $importJobId -ConnectionString $targetCrmConnectionUrl -OutputFile $importLogFile

#                    $importProgress = $importJob.Progress
#                    $importResult = (Select-Xml -Content $importJob.Data -XPath "//solutionManifest/result/@result").Node.Value
#                    $importErrorText = (Select-Xml -Content $importJob.Data -XPath "//solutionManifest/result/@errortext").Node.Value


                    Write-Output "Import Progress: $importProgress"
                    Write-Output "Import Result: $importResult"
                    Write-Output "Import Error Text: $importErrorText"
                    #Write-Output $importJob.Data

   


            }
            else {
                    Write-Output "$Solution not installed, this script will only upgrade what is installed."
                    Write-Output *
            }
        }


        #import the crm powershell module if it is not already loaded
        if (-Not (Get-Module -Name Microsoft.Xrm.Data.Powershell)) {
                Import-Module $SLConfigPath\Powershell\Modules\Microsoft.Xrm.Data.Powershell
            }
            
        # Publish any duplicate rules that might have been unpublished during solution import
        write-host "Republishing Duplicate Detection Rules"
            
        if ($duplicateruleIds -ne $null){
#            Publish-CrmDuplicateRules -conn $PScrmOrg -RuleIds $duplicateruleIds
        }
            
        write-host Restoring Auditing State
        if ($Auditingorg.isauditenabled -eq 'Yes') {
                # enable auditing for org if it was enabled before solution import
#              set-crmrecord -conn $PScrmOrg -EntityLogicalName organization -Id $Auditingorg.organizationid -Fields @{'isauditenabled'=$true}
            }
    }
    else {
        Write-Warning "Data was not returned for CRM Org $($Orgs[$I]), Verify that data has been entered in SLDeployConfigInfo.LoaderVariables on $SLConfigSQLServer}"
    }
}



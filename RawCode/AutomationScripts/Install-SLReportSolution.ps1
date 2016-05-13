
Function Reset-CRMWebAndServices {

    [CmdletBinding()]
    param()

    iisreset

    if ( Get-WmiObject -Class Win32_Service -Filter "Name='MSCRMAsyncService'" ) {
        net stop MSCRMAsyncService 
        net start MSCRMAsyncService
    }
      
    if ( Get-WmiObject -Class Win32_Service -Filter "Name='MSCRMAsyncService$maintenance'" ) {
        net stop MSCRMAsyncService$maintenance 
        net start MSCRMAsyncService$maintenance
    }
      
    if ( Get-WmiObject -Class Win32_Service -Filter "Name='MSCRMSandboxService'" ) {
        net stop MSCRMSandboxService 
        net start MSCRMSandboxService
    }

    if ( Get-WmiObject -Class Win32_Service -Filter "Name='MSCRMUnzipService'" ) {
        net stop MSCRMUnzipService 
        net start MSCRMUnzipService
    }

    Start-Sleep -s 30  
}

#---------------------------------------------------------------------------------------

# Import the xRM CI Framework Dynamics CRM Cmdlets
Import-Module “C:\Program Files (x86)\Xrm CI Framework\CRM 2011\PowerShell Cmdlets\Xrm.Framework.CI.PowerShell.dll”

$CRMVersion = "2011"
$SLConfigSQLServer = 'jeffb-sql03'
$CRMType = 'IFD'

$CRMEnvironment = 'JeffsLab'



#Set Basic Variables

$SLConfigPath = '\\vaslnas\Deploys\SLConfigs'


$ErrorActionPreference = "Stop"
$SLSolutions=@("StratusLiveSolution","VolunteerandEventsSolution","PlannedGivingModule", "SLReportsSolution")
#$SLSolutions=@("SLReportsSolution")
$srcpath = "c:\temp\jeffb03\stratuslive.crm.installer"                               #$OctopusParameters['Octopus.Tentacle.Agent.ApplicationDirectoryPath'] + "\" + $OctopusParameters['Octopus.Environment.Name'] + "\StratusLive.Crm.Installer"
$SrcRoot=  "$SRCPath"                #($srcpath + "\" + $OctopusParameters['Octopus.Release.Number'])
$ReportingPath = $SLConfigPath + "\ReportsSolution"
$SLConfigFilesPath = "$SLConfigPath\jeffb03"                                                         #$SLConfigPath + "\" + $OctopusParameters['Octopus.Environment.Name']

# ----- Copy Reporting Solution to Folder
if ( $CRMVersion -eq '2011' ) {
        Write-output "Copying 2011 Report Solution"
        Copy-item -Path $ReportingPath\SLReportsSolution.zip -Destination $SrcRoot\SLReportsSolution.zip -Force
    }
    else {
        Write-output "Copying 2016 Report Solution"
        Copy-item -Path $ReportingPath\SLReports2013_6_1.zip -Destination $SrcRoot\SLReports2013_6_1.zip -Force
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

# ----- Loop through Orgs
Foreach ($org in $Orgs) {
    Write-Output "-----"
    Write-Output "Processing $Org`n"

    # ----- Set to Null to ensure there are no values still hanging around in LoadVars
    $LoaderVars = $Null

    #Attach to SQL Server with SLConfig Information
    $serverInstance = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $SLConfigSQLServer

    #Get Data for current org
    $db = $serverInstance.Databases["SLDeployConfigInfo"]
    $LoaderVars = ($db.ExecuteWithResults("SELECT * FROM dbo.LoaderVariables WHERE (Org='$Org') AND (Environment='$CRMEnvironment') ")).Tables

    # ---- Continue if $DeployConfigVar contains data
    if ( $LoaderVars.Port -ne $Null ) {
             Write-Output "Org data Found for URL: $($LoaderVars.URL)"

            #Define the CRM connection
            if ($CRMType -eq "IFD") {
                    Write-Output "IFD Environment"
                    $ifdcrmuser = "$($LoaderVars.domain)\$($LoaderVars.crmuser)"
                    $targetCrmConnectionUrl = “ServiceUri=$($LoaderVars.URL)/XRMServices/2011/Organization.svc`;Timeout=02:00:00.0; Username=$ifdcrmuser; Password=$($LoaderVars.crmpassword)"
                }
                elseif ($CRMType -eq "AD") {
                    Write-Output "AD Environment"
                    $targetCrmConnectionUrl = “url=$($LoaderVars.URL)/$($LoaderVars.crmorg)/XRMServices/2011/Organization.svc`;Timeout=02:00:00.0; Username=$($LoaderVars.crmuser); Password=$($LoaderVars.crmpassword); Domain=$($LoaderVars.domain)"
            }

            # ----- Upgrade Solution
            #Set Solution Path (only dealing with the Report Solution in this script )
            $ImportPath = $SrcRoot + "\SLReportsSolution.zip"

            # Check if Solution is installed
            $InstalledSolution = Get-XrmSolution -ConnectionString $targetCrmConnectionUrl -UniqueSolutionName "SLReportsSolution"
           
            #if Solution is installed, Upgrade it
            if ($InstalledSolution -ne $null) {
                    Write-Output “Current $Solution version: $($InstalledSolution.Version)"
                    Write-Output “Importing Solution from: $ImportPath”
                    
                    #Import CRM Solution
                    Import-XrmSolution -ConnectionString $targetCrmConnectionUrl -SolutionFilePath $ImportPath -PublishWorkflows $true -OverwriteUnmanagedCustomizations $true 
                    
                    Write-Output "Import Complete, Resetting Services"
                    Write-Output *
                    Reset-CRMWebAndServices | out-null
                }
                else {
                    Write-Output "SLReportsSolution not installed, this script will only upgrade what is installed."
                    Write-Output *
            }

            # ----- Publish Report Customizations
            Write-host "Publishing Customizations"
            Write-host "*"
            try {
                    Publish-XrmCustomizations -ConnectionString $targetCrmConnectionUrl 
                }
                Catch {
                    write-host Publishing Customizations Failed, Please Perform Manually
                    write-host "*"
            }
            #Reset Services
            Reset-CRMWebAndServices | out-null

        }
        else {
            Write-Warning "Data was not returned for CRM Org $org, Verify that data has been entered in SLDeployConfigInfo.LoaderVariables on $SLConfigSQLServer}"
    }
}



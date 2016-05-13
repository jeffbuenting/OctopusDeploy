#Set Basic Variables
$ErrorActionPreference = "Stop"
$SLSolutions=@("StratusLiveSolution","VolunteerandEventsSolution","PlannedGivingModule", "SLReportsSolution")
#$SLSolutions=@("SLReportsSolution")
$srcpath = $OctopusParameters['Octopus.Tentacle.Agent.ApplicationDirectoryPath'] + "\" + $OctopusParameters['Octopus.Environment.Name'] + "\StratusLive.Crm.Installer"
$SrcRoot= ($srcpath + "\" + $OctopusParameters['Octopus.Release.Number'])
$ReportingPath = $SLConfigPath + "\ReportsSolution"
$SLConfigFilesPath = $SLConfigPath + "\" + $OctopusParameters['Octopus.Environment.Name']

#Copy Reporting Solution to Folder
Copy-item -Path $ReportingPath\SLReportsSolution.zip -Destination $SrcRoot\SLReportsSolution.zip -Force

#Copy install.config and loader.config
Copy-item -Path $SLConfigFilesPath\StratusLive.Crm.Installer.exe.config -Destination $SrcRoot\StratusLive.Crm.Installer.exe.config -Force
Copy-item -Path $SLConfigFilesPath\StratusLive.Crm.Loader.exe.config -Destination $SrcRoot\StratusLive.Crm.Loader.exe.config -Force

#Define Reset Function to easily reset services between imports
Function reset
{iisreset

if ( Get-WmiObject -Class Win32_Service -Filter "Name='MSCRMAsyncService'" ) 
      {net stop MSCRMAsyncService 
      net start MSCRMAsyncService}
      
if ( Get-WmiObject -Class Win32_Service -Filter "Name='MSCRMAsyncService$maintenance'" ) 
      {net stop MSCRMAsyncService$maintenance 
      net start MSCRMAsyncService$maintenance}
      
if ( Get-WmiObject -Class Win32_Service -Filter "Name='MSCRMSandboxService'" ) 
      {net stop MSCRMSandboxService 
      net start MSCRMSandboxService}

if ( Get-WmiObject -Class Win32_Service -Filter "Name='MSCRMUnzipService'" ) 
      {net stop MSCRMUnzipService 
      net start MSCRMUnzipService}

Start-Sleep -s 30  
}

#Pull connection string from mscrm key in registry
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
write-host "The following Orgs are active on this server:"
foreach ($org in $Orgs){write-host "     "$org}

# Import the xRM CI Framework Dynamics CRM Cmdlets
Import-Module “C:\Program Files (x86)\Xrm CI Framework\CRM 2011\PowerShell Cmdlets\Xrm.Framework.CI.PowerShell.dll”

#Loop through Orgs
Foreach ($org in $Orgs)
{

#Attach to SQL Server with SLConfig Information
$serverInstance = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $SLConfigSQLServer

#Get Data for current org
$db = $serverInstance.Databases["SLDeployConfigInfo"]
$ds = $db.ExecuteWithResults("SELECT * FROM dbo.LoaderVariables WHERE org='$org'")

#If $ds returned data, continue

if ($ds.tables.port -ne $null){
#Assign Data to Variables
$crmurl=$ds.tables.url
$port=[int]$ds.tables.port
$crmserver=$ds.tables.crmserver
$crmorg=$ds.tables.org
$domain=$ds.tables.domain
$crmuser=$ds.tables.crmuser
$crmpassword=$ds.tables.crmpassword
$sqlserver=$ds.tables.sqlserver
$dbusername=$ds.tables.sqluser
$dbpassword=$ds.tables.sqlpassword
$externaldbname=$ds.tables.sldbname
$ifdcrmuser = $domain + "\" + $crmuser

write-host *
write-host Processing $org
write-host *

#Define the CRM connection
if ($CRMType -eq "IFD")
{
$targetCrmConnectionUrl = “ServiceUri=" + $CRMURL + "/XRMServices/2011/Organization.svc`;Timeout=02:00:00.0; Username=" + $ifdcrmuser + "; Password=" + $crmpassword
}
elseif ($CRMType -eq "AD")
{
$targetCrmConnectionUrl = “url=" + $CRMURL + "/" + $crmorg + "/XRMServices/2011/Organization.svc`;Timeout=02:00:00.0; Username=" + $crmuser + "; Password=" + $crmpassword + "; Domain=" + $domain    
}
Foreach ($Solution in $SLSolutions){

#Upgrade Solution
#Set Solution Path
if ($Solution -eq "StratusLiveSolution")
    {
    $ImportPath = $SrcRoot + "\stratuslivesolution_release.zip"
    }
elseif ($Solution -eq "VolunteerandEventsSolution")
    {
    $ImportPath = $SrcRoot + "\engagementsolution_release.zip"
    }
elseif ($Solution -eq "PlannedGivingModule")
    {
    $ImportPath = $SrcRoot + "\plannedgivingsolution_release.zip"
    }
#elseif ($Solution -eq "GrantsSolution")
    #{
    #$ImportPath = $SrcRoot + "\grantssolution_release.zip"
    #}
elseif ($Solution -eq "SLReportsSolution")
    {
    $ImportPath = $SrcRoot + "\SLReportsSolution.zip"
    }
    

#Check if Solution is installed
$InstalledSolution = Get-XrmSolution -ConnectionString $targetCrmConnectionUrl -UniqueSolutionName $Solution
#if Solution is installed, Upgrade it
if ($InstalledSolution -ne $null)
{
    Write-Host “Current $Solution version: " $InstalledSolution.Version
    Write-Host “Importing Solution from: $ImportPath”
    #Import CRM Solution
    Import-XrmSolution -ConnectionString $targetCrmConnectionUrl -SolutionFilePath $ImportPath -PublishWorkflows $true -OverwriteUnmanagedCustomizations $true 
    Write-Host "Import Complete, Resetting Services"
    write-host *
    reset | out-null
}
else
{
    Write-Host "$Solution not installed, this script will only upgrade what is installed."
    write-host *
}
}
#Publish Report Customizations
Write-host "Publishing Customizations"
Write-host "*"
try
{
Publish-XrmCustomizations -ConnectionString $targetCrmConnectionUrl 
}
Catch
{
write-host Publishing Customizations Failed, Please Perform Manually
write-host "*"
}
#Reset Services
reset | out-null

#Change directory to the current set of artifacts beign deployed
cd $SrcRoot

#Run Loader, passing the ~ into the command simulates an enter to get past the hit any key to continue prompt in loader commandline
Write-host `nRunning Loader
write-host *
"~" | cmd /c ("StratusLive.Crm.Loader.exe -port $port -server $crmserver -org $crmorg -domain $domain -crmuser $crmuser -crmpassword $crmpassword -dbserver $sqlserver -dbuser $dbusername -dbpassword $dbpassword -externaldatabase $externaldbname -logFileName $crmorg.txt")
write-host `n*

#Clear the variables to ensure the next loop has good data
clear-variable -name port
clear-variable -name crmserver
clear-variable -name crmurl
clear-variable -name domain
clear-variable -name crmuser
clear-variable -name crmpassword
clear-variable -name sqlserver
clear-variable -name dbusername
clear-variable -name dbpassword
clear-variable -name externaldbname
clear-variable -name ifdcrmuser
}
else {write-host Data was not returned for CRM Org $org, Verify that data has been entered in SLDeployConfigInfo.LoaderVariables on $SLConfigSQLServer}
}

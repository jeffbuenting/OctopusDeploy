<#
    Runs the Stratus StratusLive.Crm.Loader.exe for each org in th environment

    2016 JAN 08
        - JDB -- $SLConfigPath added to allow duplicate Org names to exist in different environments.  Without this check all of the orgs will get updated at the same time.  CustDev / Prod.  Could cause issues.

        - JDB -- Step was separated from the others to allow loader to be run separately

#>





$CRMVersion = "2011"
$SLConfigSQLServer = 'jeffb-sql03'
$CRMType = 'IFD'

$CRMEnvironment = 'JeffsLab'



#Set Basic Variables

$SLConfigPath = '\\vaslnas\Deploys\SLConfigs'


$ErrorActionPreference = "Stop"
$SLSolutions=@("StratusLiveSolution","VolunteerandEventsSolution","PlannedGivingModule")
#$SLSolutions=@("SLReportsSolution")
$srcpath = "c:\temp\jeffb03\stratuslive.crm.installer"                               #$OctopusParameters['Octopus.Tentacle.Agent.ApplicationDirectoryPath'] + "\" + $OctopusParameters['Octopus.Environment.Name'] + "\StratusLive.Crm.Installer"
$SrcRoot=  "$SRCPath"                #($srcpath + "\" + $OctopusParameters['Octopus.Release.Number'])
$ReportingPath = $SLConfigPath + "\ReportsSolution"
$SLConfigFilesPath = "$SLConfigPath\jeffb03"                                                         #$SLConfigPath + "\" + $OctopusParameters['Octopus.Environment.Name']

# ----- Copy the loader config file to the install directory
if ( Test-Path -Path $SLConfigFilesPath\StratusLive.Crm.Loader.exe.config ) {
        Copy-item -Path $SLConfigFilesPath\StratusLive.Crm.Loader.exe.config -Destination $SrcRoot\StratusLive.Crm.Loader.exe.config -Force
    }
    else {
        Write-Warning "StratusLive.Crm.Loader.exe.config file does not exist for this Environment: $CRMEnvironment"
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

# ----- Loop through Orgs
Foreach ($org in $Orgs) {
    Write-Output "-----"
    Write-Output "Processing $Org`n"

    $LoaderVars = $Null

    #Attach to SQL Server with SLConfig Information
    $serverInstance = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $SLConfigSQLServer

    #Get Data for current org
    $db = $serverInstance.Databases["SLDeployConfigInfo"]
    $LoaderVars = ($db.ExecuteWithResults("SELECT * FROM dbo.LoaderVariables WHERE (Org='$Org') AND (Environment='$CRMEnvironment') ")).Tables

    # ---- Continue if $DeployConfigVar contains data
    if ( $LoaderVars.Port -ne $Null ) {
   
    
             Write-Output "Org data Found for URL: $($LoaderVars.URL)"

            #Change directory to the current set of artifacts beign deployed
            cd $SrcRoot

            #Run Loader, passing the ~ into the command simulates an enter to get past the hit any key to continue prompt in loader commandline
            Write-Output `nRunning Loader
            Write-Output *
            # write-output "StratusLive.Crm.Loader.exe -port $($LoaderVars.port) -server $($LoaderVars.crmserver) -org $($LoaderVars.org) -domain $($LoaderVars.domain) -crmuser $($LoaderVars.crmuser) -crmpassword $($LoaderVars.crmpassword) -dbserver $($LoaderVars.sqlserver) -dbuser $($LoaderVars.SQLUser) -dbpassword $($LoaderVars.SQLPassword) -externaldatabase $($LoaderVars.SLDBName) -logFileName $($LoaderVars.crmorg).txt"
            



            $Loader = "~" | cmd /c ("StratusLive.Crm.Loader.exe -port $($LoaderVars.port) -server $($LoaderVars.crmserver) -org $($LoaderVars.org) -domain $($LoaderVars.domain) -crmuser $($LoaderVars.crmuser) -crmpassword $($LoaderVars.crmpassword) -dbserver $($LoaderVars.sqlserver) -dbuser $($LoaderVars.SQLUser) -dbpassword $($LoaderVars.SQLPassword) -externaldatabase $($LoaderVars.SLDBName) -logFileName $($LoaderVars.crmorg).txt")
      
            Write-Output $Loader

            # ----- Check Loader output for Exceptions
            If ( $Loader -match 'Exception' ) { Write-Warning "Exception running Loader.  Check log file for specific error (Note: StratusLive.Crm.Loader.exe.Config needs to be configured to run loader from the command line)" }

            Write-Output `n*


        }
        else {
            Write-Warning "Data was not returned for CRM Org $org, Verify that data has been entered in SLDeployConfigInfo.LoaderVariables on $SLConfigSQLServer"
    }
}
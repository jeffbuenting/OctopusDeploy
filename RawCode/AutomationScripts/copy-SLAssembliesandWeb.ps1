$CRMPath = 'C:\Program Files\Microsoft Dynamics CRM'
$CRMWebServices = 'C:\inetpub\wwwroot\StratusLive-WebServices'
$StratusLiveWeb = 'C:\inetpub\wwwroot\StratusLive-Web'


$srcpath = "c:\temp\jeffb03\stratuslive.crm.installer"                               #$OctopusParameters['Octopus.Tentacle.Agent.ApplicationDirectoryPath'] + "\" + $OctopusParameters['Octopus.Environment.Name'] + "\StratusLive.Crm.Installer"
$scriptPath ="$SRCPath"                # ($srcpath + "\" + $OctopusParameters['Octopus.Release.Number'])


write-host Setting Up GAC

# ----- Checking for PowerShell version as version 4 best practice is to store modules in c:\program files\windowspowershell\modules
#if ( $PSVersionTable.PSVersion.Major -lt 4 ) {
        Copy-Item -Path $scriptPath\gac -Recurse -Destination "C:\Windows\System32\WindowsPowerShell\v1.0\Modules" -force 
        Set-Location "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\gac"
 #   }
 #   else {
 #        Copy-Item -Path $scriptPath\gac -Recurse -Destination "C:\Program Files\WindowsPowerShell\Modules" -force
#}
Import-Module gac

write-host Copying Assemblies
Copy-Item -Path "$scriptPath\assemblies\*" -Recurse -Destination "$CRMPath\Server\bin\assembly" -force

write-host Removing old files
Remove-Item $CRMWebServices\* -Exclude web.config -Recurse -force
Remove-Item $StratusLiveWeb\* -Exclude *.config -Recurse -force

write-host Copying New Web Files
Copy-Item -Path "$scriptPath\CRM-WebServices\*" -Exclude web.config,*.cs,*.csproj -Recurse -Destination "$CRMWebServices" -force 
Copy-Item -Path "$scriptPath\StratusLiveWeb\*" -Exclude *.cs,*.csproj -Recurse -Destination "$StratusLiveWeb" -force 

write-host Gac Assemblies
$fileEntries = Get-ChildItem -name -Path $scriptPath\assemblies\ -Filter *.dll -exclude FileHelpers.dll, LumenWorks.Framework.IO.dll, Microsoft.IdentityModel.dll, Microsoft.Xrm.Sdk.dll, microsoft.crm.sdk.proxy.dll, microsoft.xrm.sdk.workflow.dll, EntityFramework.dll, EntityFramework.SqlServer.dll
Set-Location "$CRMPath\Server\bin\assembly"
$glist = Get-GacAssembly 
foreach($fileName in $fileEntries) 
{ 
   if($filename -eq $glist)
   {
   $rev = $fileName.Substring(0,$fileName.Length-4)
   Remove-GacAssembly $rev
   }
   Add-GacAssembly $fileName
}         
		  
iisreset
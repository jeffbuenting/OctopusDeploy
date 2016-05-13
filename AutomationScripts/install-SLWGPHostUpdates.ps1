

#Set the IISRoot (2nd Gen servers in Rackspace have only a C drive, Performance Series us a data Drive of E:)
$IISRoot = 'c:\inetpub'

$SRCServer = '\\jeffb-rb03'

# ----- Update WorkPlace Giving Hosts

Import-Module "\\sl-jeffb\F$\OneDrive - StratusLIVE, LLC\Scripts\Modules\stratuslive\Stratuslive.psd1" -Force
Import-Module "\\sl-jeffb\F$\OneDrive - StratusLIVE, LLC\Scripts\Modules\FileSystem\FileSystem.psm1" -Force

$WPGHost = Get-SLWPGHost -IISRoot $IISRoot\wwwroot -Verbose | Backup-SLWPGHost -HostBackupRoot $IISRoot\Backup -Compress -Passthru -Verbose 
$WPGHost | Update-SLWPGHost -UpdateSource "$srcserver\wpg_source\4.1.8.170)" -verbose  
$WPGHost | foreach {
    $LastBackup = (Get-Childitem $IISRoot\backup\*.zip | Sort-object creationtime | Select-Object -last 1).FullName
    Write-Output "Restoring Last Backup: $LastBackup"
    Restore-SLWPGHost -WPGHostName $_.Name -WPGHostPath $_.Path -BackupPath $LastBackup -Zip -Recipes -Translations -Verbose
}


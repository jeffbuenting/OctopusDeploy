#----------------------------------------------------------------------------------
# Configure for Octopus
#
# OctopusDeploy Client Silent Install:  http://docs.octopusdeploy.com/display/OD/Automating+Tentacle+installation
#----------------------------------------------------------------------------------

$OctopusServiceSource = 'F:\Sources\octopusdeploy'


$CRMServer = 'jeffb-crm03.stratuslivedemo.com'



# ----- Distribute Octopus Deploy client install file to servers
$OctopusInstall = Get-ChildItem -Path $OctopusServiceSource | Sort-Object LastWriteTime | Select-Object -last 1

$OctopusInstall | FL *

$OctopusInstall | Copy-item -Destination "\\$CRMServer\c$\temp" -Force




#----------------------------------------------------------------------------------
# CRM Server

#----------------------------------------------------------------------------------

# ----- Add Octopus user account to Local Admin

# ----- Add Octopus user as DBO.Read to SQL Server MSCRM_Config


# ----- Install Octopus Deploy Client

invoke-command -ComputerName $CRMServer -ArgumentList $OctopusInstall.Name -ScriptBlock {
    Param (
        [String]$OctopusClient
    )

    Start-Process -FilePath c:\temp\$Octopusclient -ArgumentList '/Quiet' -Wait


    "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" create-instance --instance "Tentacle" --config "C:\Octopus\Tentacle\Tentacle.config"
    "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" new-certificate --instance "Tentacle" --if-blank
    "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" new-squid --instance "Tentacle"
    "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" configure --instance "Tentacle" --reset-trust
    "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" configure --instance "Tentacle" --home "C:\Octopus" --app "C:\Octopus\Applications" --port "10933" --noListen "False"
    "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" configure --instance "Tentacle" --trust "BDCC6F6BD41A84802170F69E85A9C771F029B765"
    "netsh" advfirewall firewall add rule "name=Octopus Deploy Tentacle" dir=in action=allow protocol=TCP localport=10933
    "C:\Program Files\Octopus Deploy\Tentacle\Tentacle.exe" service --instance "Tentacle" --install --start

}

# ----- Copy XRM CI Framework Powershell module to c:\program Files (x86)

# ----- Install SQL SMO 
#\\vaslnas.stratuslivedemo.com\Deploys\SLConfigs\Octopus Deploy CRM Prereqs
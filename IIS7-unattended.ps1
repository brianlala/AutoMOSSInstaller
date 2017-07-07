#Install IIS in Windows 2008 R2
#Brian Lalancette, 2009

Import-Module Servermanager
Add-WindowsFeature Web-Common-Http -IncludeAllSubFeature
Add-WindowsFeature Web-Asp-Net,Web-ISAPI-Ext,Web-ISAPI-Ext,Web-Http-Logging,Web-Log-Libraries
Add-WindowsFeature Web-Security,Web-Performance,Web-Mgmt-Console,Web-Scripting-Tools,Web-Mgmt-Service -IncludeAllSubFeature
Write-Host "- Finished installing IIS"

Add-WindowsFeature AS-NET-Framework
Write-Host "- Finished installing .Net Framework"


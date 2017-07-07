#Disable unneeded services in Windows 2008/2003
#Brian Lalancette, 2009

$QueryOS = Gwmi Win32_OperatingSystem -Comp localhost 
$QueryOS = $QueryOS.Caption 

If ($QueryOS.contains("2008") -or $QueryOS.contains("Vista"))
{
$OS = "Win2008"
$ServicesToSetManual = "Spooler","AudioSrv"
$ServicesToDisable = "WerSvc"
}

If ($QueryOS.contains("2003"))
{
$OS = "Win2003"
$ServicesToSetManual = "Spooler","WZCSVC","AudioSrv","Helpsvc"
$ServicesToDisable = "ERSvc","MpsSvc"
}

ForEach ($SvcName in $ServicesToSetManual)
{
$Svc = get-wmiobject win32_service | where-object {$_.Name -eq $SvcName} 
$SvcStartMode = $Svc.StartMode
$SvcState = $Svc.State
 If (($SvcState -eq "Running") -and ($SvcStartMode -eq "Auto"))
  {
  & net stop $SvcName
  Set-Service -name $SvcName -startupType Manual
  Write-Host "- Service $SvcName is now set to Manual start"
  }
 Else 
  {
  Write-Host "- $SvcName is already stopped and set to Manual start, no action required."
  }
Write-Host "-"
}

ForEach ($SvcName in $ServicesToDisable) 
{
$Svc = get-wmiobject win32_service | where-object {$_.Name -eq $SvcName} 
$SvcStartMode = $Svc.StartMode
$SvcState = $Svc.State
 If (($SvcState -eq "Running") -and (($SvcStartMode -eq "Auto") -or ($SvcStartMode -eq "Manual")))
  {
  & net stop $SvcName
  Set-Service -name $SvcName -startupType Disabled
  Write-Host "- Service $SvcName is now stopped and disabled."
  }
 Else 
  {
  Write-Host "- $SvcName is already stopped and disabled, no action required."
  }
Write-Host "-"
}

Write-Host "- Finished disabling services."
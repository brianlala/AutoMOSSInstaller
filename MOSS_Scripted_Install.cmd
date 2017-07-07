@echo off
CLS
SETLOCAL
IF NOT "%LaunchedFromBAT%"=="1" ECHO You must run this script using Launch.bat! & pause & EXIT
SET StartDateTime=%DATE% at %TIME%
ECHO --------------------------------------------
ECHO ^| Automated MOSS 2007 installation process ^|
ECHO ^| Started on %StartDateTime% ^|
ECHO --------------------------------------------

:: Get & Set Variables
SET InputFile=SetInputs.cmd
IF NOT EXIST "%~dp0\%InputFile%" ECHO - The input file seems to be missing. Please verify & ECHO - that %InputFile% is located in the current folder. & GOTO END
CALL "%~dp0\%InputFile%"

:: Check whether we're performing the installation as the Farm Service Account, if not, display informational warning
IF /I "%FarmAcct%" EQU "%USERDOMAIN%\%USERNAME%" GOTO DETECTOSTYPE
ECHO - You appear to be running this with non-Farm Service Account credentials. 
ECHO - In this case you should remember to change dbo for all MOSS databases 
ECHO - back to %FarmAcct% after the installation steps complete.
ECHO - 
pause
ECHO -

:DETECTOSTYPE
:: Determine if 32-bit or 64-bit install required
IF /i "%PROCESSOR_ARCHITECTURE%" == "X86" SET Platform=x86
IF /i not "%PROCESSOR_ARCHITECTURE%" == "X86" SET Platform=x64
:: Determine OS Type
:: This is not very sophisticated, but it seems to work well enough.
VER | find "6.0.600" >nul
IF %ERRORLEVEL% EQU 0 SET OSType=Win2008& GOTO SETPATH
VER | find "6.1.7" >nul
IF %ERRORLEVEL% EQU 0 SET OSType=Win2008R2& GOTO SETPATH
VER | find "5.2." >nul
IF %ERRORLEVEL% EQU 0 SET OSType=Win2003& GOTO SETPATH
ECHO - Couldn't determine type of OS & pause & GOTO END

:SETPATH
ECHO - %OSType% detected... 
:: Determine if we're running from a network path
ECHO %~dp0 | find ":\" >nul
IF NOT %ERRORLEVEL% EQU 0 ECHO - Script started from a network path; mapping drive letter... & pushd %~dp0
CD ..
SET MOSSPath=%CD%
IF NOT EXIST "%MOSSPath%\Extras\Pre-Requisites" ECHO - Please enter the path to the setup binaries (e.g. C:\Install\MOSS) & SET /P MOSSPath=
IF NOT DEFINED MOSSPath ECHO - path not specified, try again or Ctrl-C to exit... & GOTO SETPATH
IF NOT EXIST "%MOSSPath%\Extras\Pre-Requisites" ECHO - The path appears to be invalid, please try again & GOTO SETPATH

:CHOOSE
ECHO -
ECHO - You can start the process from different stages in the script:
ECHO - 
ECHO - 1: Run all steps, including installing prerequisites
ECHO - 2: Start at MOSS binary file installation
ECHO - 3: Start at Farm creation and configuration
ECHO - 4: Start at SharePoint service configuration
ECHO - 5: Start at SSP/My Sites creation and configuration
ECHO - 6: Start at Portal creation
ECHO - 7: Start at MOSS binary file installation ^& STOP (to add server to a Farm)
ECHO -
ECHO - Where would you like to start?
ECHO - 
CHOICE /C 1234567 /N /M "- 1 for ALL, 2 - MOSS, 3 - Farm, 4 - Services, 5 - SSP, 6 - Portal, 7 - Add:"
IF %ERRORLEVEL% == 7 SET InstallType=ADDTOFARM
IF %ERRORLEVEL% == 6 SET InstallType=PORTAL
IF %ERRORLEVEL% == 5 SET InstallType=SSP
IF %ERRORLEVEL% == 4 SET InstallType=SERVICES
IF %ERRORLEVEL% == 3 SET InstallType=FARM
IF %ERRORLEVEL% == 2 SET InstallType=MOSS
IF %ERRORLEVEL% == 1 SET InstallType=ALL

IF /I "%USERDOMAIN%\%USERNAME%" EQU "%FarmAcct%" GOTO INPUT

ECHO -
ECHO - Checking if Farm Account is a member of local Administrators group...
ECHO - 
IF NOT DEFINED FarmAcct ECHO - Farm Service Account not defined, please modify %InputFile% as required. & GOTO END
net localgroup "Administrators" | FINDSTR /I /C:"%FarmAcct%" >nul
IF %ERRORLEVEL% == 0 ECHO - %FarmAcct% is a local Administrator.
IF NOT %ERRORLEVEL% == 0 ECHO - Adding Farm Account %FarmAcct% to local Administrators group... & ECHO - & net localgroup "Administrators" /add "%FarmAcct%"

:INPUT
ECHO - 
ECHO - Collecting input, if required...

:INPUTPIDKEY

:INPUTDBServer
IF /I "%DBServer%" EQU "localhost" SET DBServer=
IF NOT DEFINED DBServer ECHO - Please enter the SQL Server name: (%COMPUTERNAME%) & SET /P DBServer=
IF /I "%DBServer%" EQU "localhost" ECHO - '%DBServer%' is not valid, to use local server just hit Enter & GOTO INPUTDBServer
IF NOT DEFINED DBServer ECHO - Assuming local server %COMPUTERNAME%... & SET DBServer=%COMPUTERNAME%

:INPUTFarmAcctPWD
IF NOT DEFINED FarmAcctPWD ECHO - Please enter the Farm Service Account password: & SET /P FarmAcctPWD=
IF NOT DEFINED FarmAcctPWD ECHO - Farm Account password must be provided... & GOTO INPUTFarmAcctPWD

:INPUTMOSSSearchAcctPWD
IF NOT DEFINED MOSSSearchAcctPWD ECHO - Please enter the MOSS Search Service Account password: & SET /P MOSSSearchAcctPWD=
IF NOT DEFINED MOSSSearchAcctPWD ECHO - MOSS Search Service Account password must be provided... & GOTO INPUTMOSSSearchAcctPWD

:INPUTWSSSearchAcctPWD
IF NOT DEFINED WSSSearchAcctPWD ECHO - Please enter the WSS Search Service Account password: & SET /P WSSSearchAcctPWD=
IF NOT DEFINED WSSSearchAcctPWD ECHO - WSS Search Account password must be provided... & GOTO INPUTWSSSearchAcctPWD

:INPUTSearchAccessAcctPWD
IF NOT DEFINED SearchAccessAcctPWD ECHO - Please enter the Search Access Account password: & SET /P SearchAccessAcctPWD=
IF NOT DEFINED SearchAccessAcctPWD ECHO - Search Access Account password must be provided... & GOTO INPUTSearchAccessAcctPWD

:INPUTSSPAppPoolAcctPWD
IF NOT DEFINED SSPAppPoolAcctPWD ECHO - Please enter the SSP App Pool Account password: & SET /P SSPAppPoolAcctPWD=
IF NOT DEFINED SSPAppPoolAcctPWD ECHO - SSP App Pool Account password must be provided... & GOTO INPUTSSPAppPoolAcctPWD

:INPUTMySitesAppPoolAcctPWD
IF NOT DEFINED MySitesAppPoolAcctPWD ECHO - Please enter the SSP App Pool Account password: & SET /P MySitesAppPoolAcctPWD=
IF NOT DEFINED MySitesAppPoolAcctPWD ECHO - SSP App Pool Account password must be provided... & GOTO INPUTMySitesAppPoolAcctPWD

:INPUTPortalAppPoolAcctPWD
IF NOT DEFINED PortalAppPoolAcctPWD ECHO - Please enter the Portal App Pool Account password: & SET /P PortalAppPoolAcctPWD=
IF NOT DEFINED PortalAppPoolAcctPWD ECHO - Portal App Pool Account password must be provided... & GOTO INPUTPortalAppPoolAcctPWD

ECHO - 
ECHO - Done collecting input.

IF %InstallType% == ALL GOTO PREREQ
IF %InstallType% == ADDTOFARM GOTO PREREQ
IF %InstallType% == MOSS GOTO MOSS
IF %InstallType% == FARM GOTO FARM
IF %InstallType% == SERVICES GOTO SERVICES
IF %InstallType% == SSP GOTO SSP
IF %InstallType% == PORTAL GOTO PORTAL

GOTO END

:PREREQ
IF %OSType% == Win2003 GOTO PREREQ-WIN2003
IF %OSType% == Win2008 GOTO PREREQ-WIN2008
IF %OSType% == Win2008R2 GOTO PREREQ-WIN2008R2


:PREREQ-WIN2003
ECHO - Installing Prerequisite Software:
ECHO -
ECHO - IIS 6...
TITLE IIS 6...
SYSOCMGR.exe /i:sysoc.inf /u:"%MOSSPath%\Scripted\IIS6-unattended.txt"
IF ERRORLEVEL 1 ECHO - An error occurred! & pause
ECHO - 
ECHO - .NET Framework 2.0 SP1...
TITLE .NET Framework 2.0 SP1...
IF EXIST "%SYSTEMROOT%\Microsoft.NET\Framework\v2.0.50727\aspnet_regiis.exe" ECHO - Already installed. & timeout 2
IF NOT EXIST "%SYSTEMROOT%\Microsoft.NET\Framework\v2.0.50727\aspnet_regiis.exe" "%MOSSPath%\Extras\Pre-Requisites\NetFx20SP1_%Platform%.exe" /qb /norestart
IF ERRORLEVEL 1 ECHO - An error occurred! & pause
ECHO - 
ECHO - Windows PowerShell 1.0...
TITLE Windows PowerShell 1.0...
IF EXIST "%SYSTEMROOT%\system32\windowspowershell\v1.0" ECHO - Already installed. & timeout 2
IF NOT EXIST "%SYSTEMROOT%\system32\windowspowershell\v1.0" IF %Platform% == x86 "%MOSSPath%\Extras\Pre-Requisites\WindowsServer2003-KB926139-v2-x86-ENU.exe" /passive /norestart
IF NOT EXIST "%SYSTEMROOT%\system32\windowspowershell\v1.0" IF %Platform% == x64 "%MOSSPath%\Extras\Pre-Requisites\WindowsServer2003.WindowsXP-KB926139-v2-x64-ENU.exe" /passive /norestart
IF ERRORLEVEL 1 ECHO - An error occurred! & pause
ECHO -
ECHO - .NET Framework 3.0...
TITLE .NET Framework 3.0...
IF EXIST "%SYSTEMROOT%\Microsoft.NET\Framework\v3.0\Windows Workflow Foundation" ECHO - Already installed. & timeout 2
IF NOT EXIST "%SYSTEMROOT%\Microsoft.NET\Framework\v3.0\Windows Workflow Foundation" "%MOSSPath%\Extras\Pre-Requisites\dotnetfx3_%Platform%.exe" /qb /norestart
IF ERRORLEVEL 1 ECHO - An error occurred! & pause
REM ECHO - Please wait for .NET 3.0 configuration to complete... & timeout 180
ECHO - 
REM ECHO - .NET Framework 3.0 SP1...
REM TITLE .NET Framework 3.0 SP1...
REM "%MOSSPath%\Extras\Pre-Requisites\dotnetfx30SP1setup.exe" /qb /norestart
REM IF ERRORLEVEL 1 ECHO - An error occurred! & pause
REM ECHO - 
ECHO - Install/Activate ASP.NET 2.0...
TITLE Install/Activate ASP.NET 2.0...
IF %Platform% == x86 "%WINDIR%\microsoft.net\framework\v2.0.50727\aspnet_regiis.exe" -i -enable
IF %Platform% == x64 "%WINDIR%\microsoft.net\framework64\v2.0.50727\aspnet_regiis.exe" -i -enable
IF ERRORLEVEL 1 ECHO - An error occurred! & pause
ECHO - 
ECHO - All Prerequisite Software installed successfully.
TITLE All Prerequisite Software installed successfully.
GOTO DISABLESERVICES

:PREREQ-WIN2008
ECHO - Installing Prerequisite Software:
ECHO - 
ECHO - Disabling IE Enhanced Security Configuration (ESC)
TITLE Disabling IE Enhanced Security Configuration (ESC)
:: From http://digitalformula.net/infrastructure/disable-internet-explorer-enhanced-security-configuration-ieesc-on-windows-2008/
:: Backup the registry keys - This is always a good idea before making registry changes
REG EXPORT "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" "%TEMP%.HKEY_LOCAL_MACHINE.SOFTWARE.Microsoft.Active Setup.Installed Components.A509B1A7-37EF-4b3f-8CFC-4F3A74704073.reg" /y
REG EXPORT "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" "%TEMP%.HKEY_LOCAL_MACHINE.SOFTWARE.Microsoft.Active Setup.Installed Components.A509B1A8-37EF-4b3f-8CFC-4F3A74704073.reg" /y
REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" /v "IsInstalled" /t REG_DWORD /d 0 /f
REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}" /v "IsInstalled" /t REG_DWORD /d 0 /f 
Rundll32 iesetup.dll, IEHardenLMSettings
Rundll32 iesetup.dll, IEHardenUser
Rundll32 iesetup.dll, IEHardenAdmin 
REG DELETE "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" /f /va
REG DELETE "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}" /f /va 
:: To remove the warning that shows on first IE run - this registry change will also set the default home page to about:blank
REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main" /v "First Home Page" /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main" /v "Default_Page_URL" /t REG_SZ /d "about:blank" /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main" /v "Start Page" /t REG_SZ /d "about:blank" /f
ECHO -
ECHO - Windows PowerShell 1.0...
TITLE Windows PowerShell 1.0...
IF EXIST "%SYSTEMROOT%\system32\windowspowershell\v1.0" ECHO - Already installed. & timeout 2
IF NOT EXIST "%SYSTEMROOT%\system32\windowspowershell\v1.0" start /w OCSetup.exe MicrosoftWindowsPowerShell
IF ERRORLEVEL 1 ECHO - An error occurred! & pause
ECHO - 
ECHO - IIS 7...
TITLE IIS 7...
ServerManagerCmd.exe -install Web-Server
ECHO -- IIS-ApplicationDevelopment
start /w OCSetup.exe IIS-ApplicationDevelopment
ECHO -- IIS-ISAPIExtensions
start /w OCSetup.exe IIS-ISAPIExtensions
ECHO -- IIS-ISAPIFilter
start /w OCSetup.exe IIS-ISAPIFilter
ECHO -- IIS-NetFxExtensibility
start /w OCSetup.exe IIS-NetFxExtensibility
ECHO -- WAS-NetFxEnvironment
start /w OCSetup.exe WAS-NetFxEnvironment
ECHO -- IIS-ASPNET
start /w OCSetup.exe IIS-ASPNET
ECHO -- IIS-BasicAuthentication
start /w OCSetup.exe IIS-BasicAuthentication
ECHO -- IIS-WindowsAuthentication
start /w OCSetup.exe IIS-WindowsAuthentication
ECHO -- IIS-HttpRedirect
start /w OCSetup.exe IIS-HttpRedirect
ECHO -- IIS-ManagementScriptingTools
start /w OCSetup.exe IIS-ManagementScriptingTools
ECHO - IIS 7 Installation done.
ECHO - 
ECHO - .NET Framework 3.0...
TITLE .NET Framework 3.0...
IF EXIST "%SYSTEMROOT%\Microsoft.NET\Framework\v3.0\Windows Workflow Foundation" ECHO - Already installed. & timeout 2
IF NOT EXIST "%SYSTEMROOT%\Microsoft.NET\Framework\v3.0\Windows Workflow Foundation" start /w OCSetup.exe NetFx3
IF ERRORLEVEL 1 ECHO - An error occurred! & pause
ECHO - 
ECHO - All Prerequisite Software installed successfully.
TITLE All Prerequisite Software installed successfully.
GOTO DISABLESERVICES

:PREREQ-WIN2008R2
ECHO - Installing Prerequisite Software:
ECHO - 
ECHO - Disabling IE Enhanced Security Configuration (ESC)
TITLE Disabling IE Enhanced Security Configuration (ESC)
:: From http://digitalformula.net/infrastructure/disable-internet-explorer-enhanced-security-configuration-ieesc-on-windows-2008/
:: Backup the registry keys - This is always a good idea before making registry changes
REG EXPORT "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" "%TEMP%.HKEY_LOCAL_MACHINE.SOFTWARE.Microsoft.Active Setup.Installed Components.A509B1A7-37EF-4b3f-8CFC-4F3A74704073.reg" /y
REG EXPORT "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" "%TEMP%.HKEY_LOCAL_MACHINE.SOFTWARE.Microsoft.Active Setup.Installed Components.A509B1A8-37EF-4b3f-8CFC-4F3A74704073.reg" /y
REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" /v "IsInstalled" /t REG_DWORD /d 0 /f
REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}" /v "IsInstalled" /t REG_DWORD /d 0 /f 
Rundll32 iesetup.dll, IEHardenLMSettings
Rundll32 iesetup.dll, IEHardenUser
Rundll32 iesetup.dll, IEHardenAdmin 
REG DELETE "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" /f /va
REG DELETE "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}" /f /va 
:: To remove the warning that shows on first IE run - this registry change will also set the default home page to about:blank
REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main" /v "First Home Page" /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main" /v "Default_Page_URL" /t REG_SZ /d "about:blank" /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main" /v "Start Page" /t REG_SZ /d "about:blank" /f
ECHO -
ECHO - Windows PowerShell 1.0...
TITLE Windows PowerShell 1.0...
IF EXIST "%SYSTEMROOT%\system32\windowspowershell\v1.0" ECHO - Already installed. & timeout 2
IF NOT EXIST "%SYSTEMROOT%\system32\windowspowershell\v1.0" start /w OCSetup.exe MicrosoftWindowsPowerShell
IF ERRORLEVEL 1 ECHO - An error occurred! & pause
ECHO - 
ECHO - IIS 7 and .Net Framework...
TITLE IIS 7 and .Net Framework...
:: Get existing Powershell ExecutionPolicy
FOR /F "tokens=*" %%x in ('powershell.exe Get-ExecutionPolicy') do (set ExecutionPolicy=%%x)
:: Temporarily set Unrestricted, in case we are running over a net share or UNC
IF NOT "%ExecutionPolicy%"=="Unrestricted" ECHO - PS ExecutionPolicy is %ExecutionPolicy%, setting ExecutionPolicy to Unrestricted. & powershell Set-ExecutionPolicy Unrestricted
powershell "%~dp0\IIS7-unattended.ps1"
:: Restore Powershell ExecutionPolicy
ECHO - Setting PS ExecutionPolicy back to %ExecutionPolicy%. & powershell Set-ExecutionPolicy %ExecutionPolicy%
ECHO - IIS 7 and .Net Framework Installation done.
ECHO - 
IF ERRORLEVEL 1 ECHO - An error occurred! & pause
ECHO - 
ECHO - All Prerequisite Software installed successfully.
TITLE All Prerequisite Software installed successfully.
GOTO DISABLESERVICES

:DISABLESERVICES
ECHO - 
ECHO - Disabling Unneeded Services...
TITLE Disabling Unneeded Services...
:: Get existing Powershell ExecutionPolicy
FOR /F "tokens=*" %%x in ('"%SYSTEMROOT%\system32\windowspowershell\v1.0\powershell.exe" Get-ExecutionPolicy') do (set ExecutionPolicy=%%x)
:: Temporarily set Unrestricted, in case we are running over a net share or UNC
IF NOT "%ExecutionPolicy%"=="Unrestricted" ECHO - PS ExecutionPolicy is %ExecutionPolicy%, setting ExecutionPolicy to Unrestricted. & "%SYSTEMROOT%\system32\windowspowershell\v1.0\powershell.exe" Set-ExecutionPolicy Unrestricted
"%SYSTEMROOT%\system32\windowspowershell\v1.0\powershell.exe" "%~dp0\DisableUnneededServices.ps1"
:: Restore Powershell ExecutionPolicy
ECHO - Setting PS ExecutionPolicy back to %ExecutionPolicy%. & "%SYSTEMROOT%\system32\windowspowershell\v1.0\powershell.exe" Set-ExecutionPolicy %ExecutionPolicy%
GOTO MOSS

:MOSS
ECHO - 
ECHO - Install MOSS binaries...
TITLE Install MOSS binaries...
"%MOSSPath%\%Platform%\setup.exe" /config "%~dp0\config.xml"
IF ERRORLEVEL 1 ECHO - An error occurred! & pause
IF %InstallType% == ADDTOFARM GOTO END
:: For some reason the SharePoint Products and Technologies Configuration Wizard
:: still pops up automatically, even on an unattended install.
:: We need to get rid it to avoid confusion,
:: since the wizard steps will instead be run automatically next, with no GUI interaction.
ECHO - Waiting for SharePoint Products and Technologies Wizard to launch
ECHO - We need to stop it as PSConfig.exe will run automatically instead.
timeout 15
"%~dp0\pskill.exe" psconfigui.exe /accepteula >nul
REM ECHO - Adding PowerShell ^& SharePoint [12 Hive]\BIN directory to system PATH
REM ECHO %PATH% | FIND /I "powershell" >nul
REM IF ERRORLEVEL 1 SET NEWPATH=%PATH%;%SYSTEMROOT%\system32\windowspowershell\v1.0 & SET PSINSTALLED=1
REM IF DEFINED NEWPATH REG ADD "HKLM\System\CurrentControlSet\Control\Session Manager\Environment" /v Path /f /d "%NEWPATH%" & GOTO FARM
REM IF DEFINED PSINSTALLED ECHO %NEWPATH% | FIND /I "web server extensions\12\BIN" >nul
REM IF ERRORLEVEL 1 IF DEFINED PSINSTALLED SET NEWPATH=%NEWPATH%;%COMMONPROGRAMFILES%\Microsoft Shared\web server extensions\12\BIN & SET SPINSTALLED=1
REM IF NOT DEFINED PSINSTALLED ECHO %PATH% | FIND /I "web server extensions\12\BIN" >nul
REM IF ERRORLEVEL 1 IF NOT DEFINED PSINSTALLED SET NEWPATH=%PATH%;%COMMONPROGRAMFILES%\Microsoft Shared\web server extensions\12\BIN & SET SPINSTALLED=1
REM IF DEFINED REG ADD "HKLM\System\CurrentControlSet\Control\Session Manager\Environment" /v Path /f /d "%NEWPATH%" & GOTO FARM
REM ECHO - [12 Hive]\BIN directory already in system PATH. 
GOTO FARM

:FARM
ECHO - 
ECHO - Create and configure MOSS farm...
TITLE Create and configure MOSS farm...
CD /D "%COMMONPROGRAMFILES%\Microsoft Shared\web server extensions\12\BIN\"
IF ERRORLEVEL 1 ECHO - MOSS Binaries are not installed yet! & pause & GOTO END
ECHO -
ECHO - Farm creation/config: Create Farm configuration DB
TITLE Farm creation/config: Create Farm configuration DB 
ECHO -
psconfig.exe -cmd ConfigDB -create -server "%DBServer%" -database %ConfigDB% -user "%FarmAcct%" -password "%FarmAcctPWD%" -admincontentdatabase %CentralAdminContentDB%
ECHO -
ECHO - Farm creation/config: Install Help collections
TITLE Farm creation/config: Install Help collections 
ECHO -
psconfig.exe -cmd helpcollections -installall
ECHO -
ECHO - Farm creation/config: Secure resources
TITLE Farm creation/config: Secure resources 
ECHO -
psconfig.exe -cmd secureresources
ECHO -
ECHO - Farm creation/config: Install services
TITLE Farm creation/config: Install services 
ECHO -
psconfig.exe -cmd services -install
ECHO -
ECHO - Farm creation/config: Install features
TITLE Farm creation/config: Install features
ECHO -
psconfig.exe -cmd installfeatures
ECHO -
ECHO - Farm creation/config: Provision central admininstration site
TITLE Farm creation/config: Provision central admininstration site 
ECHO -
psconfig.exe -cmd adminvs -provision -port %CentralAdminPort% -windowsauthprovider onlyusentlm
:: Set primary and secondary owners if running under different credentials than Farm Account
IF /I "%USERDOMAIN%\%USERNAME%" NEQ "%FarmAcct%" ECHO - Adding "%USERDOMAIN%\%USERNAME%" as secondary owner... & stsadm.exe -o siteowner -url http://%COMPUTERNAME%:%CentralAdminPort% -ownerlogin "%FarmAcct%" -secondarylogin "%USERDOMAIN%\%USERNAME%" 
ECHO -
ECHO - Farm creation/config: Install application content
TITLE Farm creation/config: Install application content 
ECHO -
psconfig.exe -cmd applicationcontent -install
ECHO -

ECHO - MOSS Farm created and configured successfully.
TITLE MOSS Farm created and configured successfully.
GOTO SERVICES

:SERVICES
ECHO -
ECHO - Services: WSS Search (Help Search)
TITLE Services: WSS Search (Help Search)
CD /D "%COMMONPROGRAMFILES%\Microsoft Shared\web server extensions\12\BIN\"
IF ERRORLEVEL 1 ECHO - MOSS Binaries are not installed yet! & pause & GOTO END
ECHO -
stsadm.exe -o spsearch -action start -farmperformancelevel PartlyReduced -farmserviceaccount "%WSSSearchAcct%" -farmservicepassword "%WSSSearchAcctPWD%" -farmcontentaccessaccount "%SearchAccessAcct%" -farmcontentaccesspassword "%SearchAccessAcctPWD%" -databaseserver "%DBServer%" -databasename %WSSSearchDB%
ECHO -
ECHO - Services: Office SharePoint Search (MOSS Search)
TITLE Services: Office SharePoint Search (MOSS Search)
ECHO -
stsadm.exe -o osearch -action start -role IndexQuery -farmcontactemail %FarmAcctEmail% -farmperformancelevel PartlyReduced -farmserviceaccount "%MOSSSearchAcct%" -farmservicepassword "%MOSSSearchAcctPWD%" -defaultindexlocation "%MOSSIndexLocation%"
ECHO -
ECHO - Services: Excel Services
TITLE Services: Excel Services
ECHO -
stsadm.exe -o provisionservice -action start -servicetype "Microsoft.Office.Excel.Server.ExcelServerSharedWebService, Microsoft.Office.Excel.Server, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
GOTO SSP

:SSP
ECHO -
ECHO - Create SSP and My Sites...
TITLE Create SSP and My Sites...
CD /D "%COMMONPROGRAMFILES%\Microsoft Shared\web server extensions\12\BIN\"
IF ERRORLEVEL 1 ECHO - MOSS Binaries are not installed yet! & pause & GOTO END
ECHO -
ECHO - SSP: Add Farm Account to Farm Administrators
TITLE SSP: Add Farm Account to Farm Administrators
ECHO -
stsadm.exe -o adduser -url http://%COMPUTERNAME%:%CentralAdminPort% -userlogin "%FarmAcct%" -useremail %FarmAcctEmail% -group "Farm Administrators" -username "%FarmAcct%" -siteadmin
ECHO -
ECHO - SSP: Create SSP admin web application
TITLE SSP: Create SSP admin web application
ECHO -
stsadm.exe -o extendvs -url %SSPURL%:%SSPPort% -ownerlogin "%FarmAcct%" -owneremail %FarmAcctEmail% -databasename %SSPAdminContentDB% -exclusivelyusentlm -description "%SSPDesc%" -apidname "%SSPDesc%" -apcreatenew -apidtype configurableid -apidlogin "%SSPAppPoolAcct%" -apidpwd "%SSPAppPoolAcctPWD%"
:: Set primary and secondary owners if running under different credentials than Farm Account
IF /I "%USERDOMAIN%\%USERNAME%" NEQ "%FarmAcct%" ECHO - Adding "%USERDOMAIN%\%USERNAME%" as secondary owner... & stsadm.exe -o siteowner -url %SSPURL%:%SSPPort% -ownerlogin "%FarmAcct%" -secondarylogin "%USERDOMAIN%\%USERNAME%" 
ECHO -
ECHO - SSP: Create My Sites web application
TITLE SSP: Create My Sites web application
ECHO -
stsadm.exe -o extendvs -url %MySitesURL%:%MySitesPort% -sethostheader -ownerlogin "%FarmAcct%" -owneremail %FarmAcctEmail% -databasename %MySitesContentDB% -exclusivelyusentlm -description "%MySitesDesc%" -apidname "%MySitesDesc%" -apcreatenew -apidtype configurableid -apidlogin "%MySitesAppPoolAcct%" -apidpwd "%MySitesAppPoolAcctPWD%"
:: Set primary and secondary owners if running under different credentials than Farm Account
IF /I "%USERDOMAIN%\%USERNAME%" NEQ "%FarmAcct%" ECHO - Adding "%USERDOMAIN%\%USERNAME%" as secondary owner... & stsadm.exe -o siteowner -url %MySitesURL%:%MySitesPort% -ownerlogin "%FarmAcct%" -secondarylogin "%USERDOMAIN%\%USERNAME%" 
ECHO -
ECHO - SSP: Create SSP
TITLE SSP: Create SSP
ECHO -
stsadm.exe -o createssp -title "%SSPDesc%" -url %SSPURL%:%SSPPort% -mysiteurl %MySitesURL%:%MySitesPort% -indexserver %SSPIndexServer% -indexlocation "%SSPIndexLocation%" -ssplogin "%SSPAppPoolAcct%" -ssppassword "%SSPAppPoolAcctPWD%" -sspdatabasename %SSPContentDB% -searchdatabasename %SSPSearchDB%
:: Set primary and secondary owners if running under different credentials than Farm Account
IF /I "%USERDOMAIN%\%USERNAME%" NEQ "%FarmAcct%" ECHO - Adding "%USERDOMAIN%\%USERNAME%" as secondary owner... & stsadm.exe -o siteowner -url %SSPURL%:%SSPPort%/ssp/admin -ownerlogin "%FarmAcct%" -secondarylogin "%USERDOMAIN%\%USERNAME%" 
GOTO PORTAL

:PORTAL
ECHO -
ECHO - Create Main Portal...
TITLE Create Main Portal...
CD /D "%COMMONPROGRAMFILES%\Microsoft Shared\web server extensions\12\BIN\"
IF ERRORLEVEL 1 ECHO - MOSS Binaries are not installed yet! & pause & GOTO END
stsadm.exe -o extendvs -url %PortalURL%:%PortalPort% -sethostheader -ownerlogin "%FarmAcct%" -owneremail %FarmAcctEmail% -databaseserver "%DBServer%" -databasename %PortalContentDB% -exclusivelyusentlm -sitetemplate %PortalTemplate% -description "%PortalDesc%" -apidname "%PortalDesc%" -apidtype configurableid -apidlogin "%PortalAppPoolAcct%" -apidpwd "%PortalAppPoolAcctPWD%"
:: Set primary and secondary owners if running under different credentials than Farm Account
IF /I "%USERDOMAIN%\%USERNAME%" NEQ "%FarmAcct%" ECHO - Adding "%USERDOMAIN%\%USERNAME%" as secondary owner... & stsadm.exe -o siteowner -url %PortalURL%:%PortalPort% -ownerlogin "%FarmAcct%" -secondarylogin "%USERDOMAIN%\%USERNAME%" 
ECHO - IISReset
TITLE IISReset
iisreset
ECHO -
ECHO - Launching Central Admin...
TITLE Launching Central Admin...
start /MIN "%PROGRAMFILES%\Internet Explorer\iexplore.exe" http://%COMPUTERNAME%:%CentralAdminPort%
timeout 60
ECHO -
ECHO - Launching Portal...
TITLE Launching Portal...
start "%PROGRAMFILES%\Internet Explorer\iexplore.exe" %PortalURL%:%PortalPort%
GOTO END

:END
ECHO -
ECHO - Done!
TITLE Done!
SET EndDateTime=%DATE% at %TIME%
ECHO --------------------------------------------
ECHO ^| Automated MOSS 2007 installation process ^|
ECHO ^| Started on %StartDateTime% ^|
ECHO ^| Ended on %EndDateTime%   ^|
ECHO --------------------------------------------

popd
ENDLOCAL
pause
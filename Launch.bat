@ECHO OFF
SETLOCAL
cls
FOR /F "tokens=2-4 delims=/ " %%i IN ('date /t') DO SET SHORTDATE=%%i-%%j-%%k
FOR /F "tokens=1-3 delims=: " %%i IN ('time /t') DO SET SHORTTIME=%%i-%%j%%k
SET LaunchedFromBAT=1
CALL "%~dp0\MOSS_Scripted_Install.cmd" | "%~dp0\tee.exe" "%TEMP%\MOSS_Scripted_Install_%SHORTDATE%_%SHORTTIME%.rtf"
START /MIN Explorer.exe "%TEMP%"
START %TEMP%\MOSS_Scripted_Install_%SHORTDATE%_%SHORTTIME%.rtf
ENDLOCAL
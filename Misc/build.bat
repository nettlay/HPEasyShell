@echo off

:: BatchGotAdmin
:: -------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"

echo =======================================================
echo    Config Testing
echo =======================================================
if not exist c:\svc md c:\svc
if not exist c:\svc\hpeasyshell md c:\svc\hpeasyshell
xcopy *.* c:\svc\hpeasyshell\ /s /e /r /c /y
copy .\services\svcconfig.ini c:\svc /y
copy .\services\runappasService.exe c:\svc /y
cd c:\svc
runappasService.exe --startup auto install


echo =======================================================
echo    Disable UAC
echo =======================================================
%windir%\System32\reg.exe add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v EnableLUA /t REG_DWORD /d 0 /f

echo =======================================================
echo       Disable firewall
echo =======================================================
netsh advfirewall set publicprofile state off
netsh advfirewall set privateprofile state off

echo =======================================================
echo 	Please Press any key to Reboot TC
echo =======================================================
pause
shutdown -r -t 5

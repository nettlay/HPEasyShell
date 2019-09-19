@echo off
set x=None
@for /f "tokens=1,3" %%i in ('netsh WLAN show interfaces') do (if [%%i]==[SSID] set x=%%j)
echo %x%
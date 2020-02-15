@echo off

SET PingResultFile=ping.txt
SET ArpResultFile=arp.txt

CALL :CLEAN

FOR /L %%i IN (1,1,254) DO CALL :CHECK 192.168.1 %%i
IF EXIST %ArpResultFile% type %ArpResultFile%
CALL :CLEAN
exit /B


:: ==============================================
:CHECK
SET IP=%1.%2
ping -n 1 -l 1 -w 100 %IP% >> %PingResultFile%
IF %ERRORLEVEL% GTR 0 GOTO :END
echo %IP%
arp -a %IP% >> %ArpResultFile%

:END
GOTO:EOF

:: ==============================================
:CLEAN
IF EXIST %PingResultFile% del /F /Q %PingResultFile%
IF EXIST %ArpResultFile% del /F /Q %ArpResultFile%
GOTO:EOF

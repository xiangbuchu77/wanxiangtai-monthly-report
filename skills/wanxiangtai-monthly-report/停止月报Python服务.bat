@echo off
setlocal EnableExtensions

rem Request elevation automatically; an elevated Python process cannot be stopped otherwise.
fltmc >nul 2>&1
if errorlevel 1 (
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

rem Remember the Skill path, then leave it so this cmd window cannot lock it.
set "SKILL_ROOT=%~dp0"
cd /d "%TEMP%"

echo Stopping the monthly-report service on port 8799...
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":8799" ^| findstr /I "LISTENING"') do (
    taskkill /F /T /PID %%P >nul 2>&1
)

rem Fallbacks for services started by older Skill packages.
taskkill /F /T /IM QClaw.exe >nul 2>&1
taskkill /F /T /IM openclaw-gateway.exe >nul 2>&1
taskkill /F /T /IM python.exe >nul 2>&1
taskkill /F /T /IM pythonw.exe >nul 2>&1
taskkill /F /T /IM python3.exe >nul 2>&1
taskkill /F /T /IM python3.10.exe >nul 2>&1
taskkill /F /T /IM python3.11.exe >nul 2>&1
taskkill /F /T /IM python3.12.exe >nul 2>&1
taskkill /F /T /IM python3.13.exe >nul 2>&1
taskkill /F /T /IM py.exe >nul 2>&1

timeout /t 2 /nobreak >nul

netstat -ano | findstr ":8799" | findstr /I "LISTENING" >nul
if not errorlevel 1 (
    echo.
    echo The service is still running. Right-click this file and choose Run as administrator.
    pause
    exit /b 1
)

rem Remove logs used by both current and legacy package layouts.
del /F /Q "%SKILL_ROOT%local-agent.log" >nul 2>&1
del /F /Q "%SKILL_ROOT%logs\local-agent.log" >nul 2>&1
del /F /Q "%SKILL_ROOT%assets\service\logs\local-agent.log" >nul 2>&1
del /F /Q "%TEMP%\wanxiangtai-monthly-report\logs\local-agent.log" >nul 2>&1

if exist "%SKILL_ROOT%local-agent.log" goto log_locked
if exist "%SKILL_ROOT%logs\local-agent.log" goto log_locked
if exist "%SKILL_ROOT%assets\service\logs\local-agent.log" goto log_locked
if exist "%TEMP%\wanxiangtai-monthly-report\logs\local-agent.log" goto log_locked

echo.
echo The service is stopped and local-agent.log has been removed.
echo You can now delete or replace the old Skill folder.
timeout /t 3 /nobreak >nul
exit /b 0

:log_locked
echo.
echo local-agent.log is still locked. Restart Windows, do not open QClaw, then run this file again.
pause
exit /b 1

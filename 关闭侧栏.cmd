@echo off
setlocal

echo.
echo Closing PPT Sidebar ...
echo.

taskkill /IM PowerPointSidebar.exe /F >nul 2>&1
if errorlevel 1 goto none

echo [OK] Sidebar closed.
goto end

:none
echo No running sidebar was found.
echo If it is actually running, try running this script as Administrator.
echo.
echo You can also use:  PowerPointSidebar.exe --kill
goto end

:end
endlocal

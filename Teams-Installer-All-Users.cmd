@echo off
:: Check for administrative privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Error: This script must be run as an Administrator.
    echo Please right-click the file and choose "Run as administrator".
    pause
    exit /b
)

echo Forcefully purging conflicting AppX/MSIX Teams records...
:: Force remove modern Teams packages across all users and provisioning logs
powershell -NoProfile -Command "Get-AppxPackage -AllUsers *MSTeams* | Remove-AppxPackage -AllUsers" >nul 2>&1
powershell -NoProfile -Command "Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like '*MSTeams*'} | Remove-AppxProvisionedPackage -Online" >nul 2>&1

echo.
echo Installing fresh Microsoft Teams copy for all users...
:: Re-attempt the global machine install using winget
winget install --id Microsoft.Teams --scope machine --override "/passive ALLUSERS=1" --accept-source-agreements --accept-package-agreements

if %errorLevel% equ 0 (
    echo.
    echo Microsoft Teams has been successfully installed globally for all users!
) else (
    echo.
    echo Installation failed with exit code %errorLevel%.
)

pause

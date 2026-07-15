@echo off
cd /d "%~dp0"
echo Starting Offline Office Installation...
setup.exe /configure configuration.xml
echo Offline installation complete.
pause

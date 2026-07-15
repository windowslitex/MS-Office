@echo off
cd /d "%~dp0"
echo Starting Online Office Installation...
setup.exe /configure configuration.xml
echo Installation complete or process finished.
pause

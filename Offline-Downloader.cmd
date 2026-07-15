@echo off
cd /d "%~dp0"
echo Downloading Office files for offline deployment...
setup.exe /download configuration.xml
echo Download complete. Check your source folder.
pause

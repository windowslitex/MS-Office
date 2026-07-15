@echo off
fltmc >nul || exit /b
call "%~dp0MAS_AIO.cmd" /Ohook
cd \
(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")

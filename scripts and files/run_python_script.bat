@echo off

REM Get the current directory of the script
set "script_folder=%~dp0"

REM Change the directory to the script's folder
cd /d "%script_folder%"

REM Run the Python script
python script_name.py

pause
@echo off
setlocal enabledelayedexpansion

REM Define variables
set "python_version=3.12.2"
set "install_path=C:\Python312"
set "libs=openpyxl psycopg2"

REM Install Python
echo Installing Python %python_version%...
curl -o python_installer.exe https://www.python.org/ftp/python/%python_version%/python-%python_version%-amd64.exe
start /wait python_installer.exe /quiet InstallAllUsers=1 PrependPath=1 TargetDir=%install_path%
del python_installer.exe

REM Install Python libraries
echo Installing Python libraries...
%install_path%\Scripts\pip install %libs%

REM Set environmental variables
echo Setting up environmental variables...
setx PATH "%install_path%;%install_path%\Scripts;" /M

echo Installation complete.
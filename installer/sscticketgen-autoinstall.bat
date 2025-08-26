@echo off

REM Check if Python is installed
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python is not installed.
    echo Please download and install Python from https://www.python.org/downloads/
    start https://www.python.org/downloads/
    echo After installing Python, re-run this script.
    pause
    exit /b
)

REM Install pip if not present
python -m ensurepip --default-pip

REM Ask if user wants to clone the repo
set /p clone_choice="Do you want to clone the SSCTicketGen repo from GitHub? (Y/N): "
if /I "%clone_choice%" NEQ "Y" (
    echo Skipping repo clone.
    goto skip_clone
)

REM Prompt user for directory to save the repo
set /p repo_dir="Enter the full path where you want to save SSCTicketGen (e.g. C:\MyProjects): "

REM Create the directory if it doesn't exist
if not exist "%repo_dir%" (
    mkdir "%repo_dir%"
)

cd /d "%repo_dir%"

REM Clone the GitHub repo
if not exist SSCTicketGen (
    git clone https://github.com/jamieisonline/SSCTicketGen
)

:skip_clone
REM If repo was not cloned, check if folder exists
if not exist "%repo_dir%\SSCTicketGen" (
    echo SSCTicketGen folder not found. Exiting.
    pause
    exit /b
)

cd SSCTicketGen

REM Install dependencies
python -m pip install --upgrade pip
python -m pip install PyQt6
python -m pip install Jinja2

REM Open the GitHub page in browser
start https://github.com/jamieisonline/SSCTicketGen

echo
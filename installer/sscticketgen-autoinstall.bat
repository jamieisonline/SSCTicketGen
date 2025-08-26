@echo off
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

REM Install Python if not present
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python not found. Downloading and installing Python...
    powershell -Command "Start-Process 'https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe' -Wait"
    echo Please install Python, then re-run this script.
    pause
    exit /b
)

REM Install pip if not present
python -m ensurepip --default-pip

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
python -m pip install -r requirements.txt

REM Open the GitHub page in browser
start https://github.com/jamieisonline/SSCTicketGen

echo
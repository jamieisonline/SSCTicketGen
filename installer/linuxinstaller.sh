#!/bin/bash

# Check if Python is installed
if ! command -v python3 &>/dev/null; then
    echo "Python is not installed."
    echo "Please download and install Python from https://www.python.org/downloads/"
    xdg-open https://www.python.org/downloads/
    echo "After installing Python, re-run this script."
    exit 1
fi

# Install pip if not present
python3 -m ensurepip --default-pip

# Ask if user wants to clone the repo
read -p "Do you want to clone the SSCTicketGen repo from GitHub? (Y/N): " clone_choice
if [[ ! "$clone_choice" =~ ^[Yy]$ ]]; then
    echo "Skipping repo clone."
    goto skip_clone
fi

# Prompt user for directory to save the repo
read -p "Enter the full path where you want to save SSCTicketGen (e.g. /home/user/MyProjects): " repo_dir

# Create the directory if it doesn't exist
if [ ! -d "$repo_dir" ]; then
    mkdir -p "$repo_dir"
fi

cd "$repo_dir" || exit

# Clone the GitHub repo if not already present
if [ ! -d "SSCTicketGen" ]; then
    git clone https://github.com/jamieisonline/SSCTicketGen
fi

skip_clone:

# If repo was not cloned, check if folder exists
if [ ! -d "$repo_dir/SSCTicketGen" ]; then
    echo "SSCTicketGen folder not found. Exiting."
    exit 1
fi

cd "$repo_dir/SSCTicketGen" || exit

# Install dependencies
python3 -m pip install --upgrade pip
python3 -m pip install PyQt6
python3 -m pip install Jinja2

# Open the GitHub page in browser
xdg-open https://github.com/jamieisonline/SSCTicketGen

echo "Script complete!"

# Get_OU PowerShell Script

## About
This PowerShell script fetches the Organizational Unit (OU) for specified computers from Active Directory. It offers two modes of operation: fetching just the OU or fetching the OU along with its modification history. You can either look up a single computer or use a file to look up multiple computers.

## Version
- v1.0: Basic features.
- v2.0: Enhanced user interface, modular code, customizable export options.
- v3.0: Added 'OU with History' mode, GUI for folder selection, and enhanced source selection options.

## Requirements
- PowerShell (Version 5 or higher)
- Administrative Privileges

## How to Run
1. Open PowerShell as Administrator.
2. Navigate to the directory containing the script.
3. Run the script.

### Alternatively
1. Navigate to the location of the script.
2. Right-click the `.ps1` script and click on "Run with PowerShell."

The script will automatically recognize if you have administrative privileges and will request the proper permissions if needed.

## New Features in v3.0
- Introduced 'OU with History' mode that fetches modification history along with the OU.
- GUI-based folder selection for saving the output file.
- Enhanced data source selection: you can now choose between a single computer or a list from a file.

## Output
The script will display a table with the computer names and their corresponding OUs, and optionally their modification history. You also have the option to export this data to a text or CSV file. With version 3.0, you can choose the folder where you'd like to save the file through a GUI.

## Author
Doron Bogomolov


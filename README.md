# Get_OU PowerShell Script

## About
This PowerShell script fetches the Organizational Unit (OU) for specified computers from Active Directory. You can either look up a single computer or use a file to look up multiple computers.

## Version
- v1.0: Basic features.
- v2.0: Enhanced user interface, modular code, customizable export options.

## Requirements
- PowerShell
- Administrative Privileges

## How to Run
1. Open PowerShell as Administrator.
2. Navigate to the directory containing the script.
3. Run the script.

### Alternatively
1. Navigate to the location of the script.
2. Right-click the `.ps1` script and click on "Run with PowerShell."

The script will automatically recognize if you have administrative privileges and will request the proper permissions if needed.

## New Features in v2.0
- User-friendly menu interface for selecting input options.
- The ability to export results to a custom-named file.
- The ability to export results to a custom directory.
- Default export naming and directory based on the current date and time.

## Output
The script will display a table with the computer names and their corresponding OUs. You also have the option to export this data to a text or CSV file. With version 2.0, you can customize the file name and directory.

## Author
Doron Bogomolov

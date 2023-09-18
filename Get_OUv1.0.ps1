# Get_OUv1.0.ps1
# This script fetches the Organizational Unit (OU) for specified computers from Active Directory.
# Options for data source: 1) Single computer 2) List from a file.
# Author: Doron Bogomolov
# Last Updated: 23-August-2023


# Check if running with admin privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    # Restart the script as admin
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Function to get the OU of a computer
function Get-ComputerOU {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$ComputerName
    )
    try {
        # Get the computer object using the Active Directory PowerShell module
        $computerObject = Get-ADComputer $ComputerName -ErrorAction Stop
        # Get the distinguished name (DN) of the computer object
        $ouDN = $computerObject.DistinguishedName
        # Return only the OU part of the DN
        $ou = $ouDN -replace "^CN=.*?,", ""
        return $ou
    }
    catch {
        return "Not Found"
    }
}

# Prompt the user to choose the input method
$choice = Read-Host "Choose input method:`n1. Single PC name`n2. Read from a file (text or CSV)`nEnter 1 or 2"
switch ($choice) {
    1 {
        # Prompt for a single PC name
        $pcName = Read-Host "Enter the PC name"
        $pcList = @($pcName)
    }
    2 {
        # Prompt for the file path
        $filePath = Read-Host "Enter the file path (text or CSV)"
        if (-not (Test-Path $filePath)) {
            Write-Host "File not found. Exiting script."
            exit
        }
        # Read the PC names from the file
        $pcList = Get-Content $filePath
    }
    default {
        Write-Host "Invalid choice. Exiting script."
        exit
    }
}

# Initialize the result array
$results = @()

# Loop through each PC in the list
foreach ($pc in $pcList) {
    # Get the OU of the PC
    $ou = Get-ComputerOU -ComputerName $pc
    # Add the PC and its OU to the result array
    $results += [PSCustomObject]@{
        "PC-Name" = $pc
        "OU" = $ou
    }
}

# Display the results
$results | Format-Table

# Prompt to export the results
$exportChoice = Read-Host "Export the results to a file? (Y/N)"
if ($exportChoice -eq "Y") {
    $exportPath = Read-Host "Enter the export file path (e.g., C:\output.txt or C:\output.csv)"
    $results | Export-Csv -Path $exportPath -NoTypeInformation
    Write-Host "Results exported to $exportPath"
}

# Pause to keep the window open
Read-Host "Press Enter to exit"

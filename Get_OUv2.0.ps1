<#
==== Summary ====
A script to find the Organizational Unit (OU) of a given computer or list of computers from Active Directory.

==== What it Does ====
    - You can either type in a single computer name.
    - Or, you can give it a file that has a bunch of computer names.
    The script will then look up the OU for each and show it to you.
    Also, you can save the results to a CSV if you want.

==== File Info ====
    File Name: Get_OUv2.0.ps1
    Written By: Doron Bogomolov
    Last Updated: 27-August-2023
#>


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

# Function to display the menu and get user input
function Show-Menu {
    Clear-Host
    Write-Host "1. Enter a single PC name"
    Write-Host "2. Use a list of PC names from a file"
    $choice = Read-Host "Please enter your choice (1 or 2):"
    return $choice
}

# Function to get PC names from user input
function Get-PCNames {
    $choice = Show-Menu
    while ($choice -ne "1" -and $choice -ne "2") {
        Write-Host "Invalid choice. Please try again."
        $choice = Show-Menu
    }

    if ($choice -eq "1") {
        $pcNames = @()
        $pcName = Read-Host "Enter the PC name:"
        $pcNames += $pcName
    } else {
        $pcNamesFile = Read-Host "Enter the path to the file containing PC names:"
        $pcNames = Get-Content $pcNamesFile
    }

    return $pcNames
}

# Get the list of PC names
$pcList = Get-PCNames

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

# Prompt the user whether to export the results to a file
$exportChoice = Read-Host "Export the results to a file? (Y/N)"
if ($exportChoice -eq "Y") {
    # Generate the default output file name based on current date and time
    $defaultOutputName = "OU_{0:yyyy-MM-dd_HH-mm-ss}.csv" -f (Get-Date)
    
    # Prompt for the desired file name
    $fileName = Read-Host "Enter the file name (default: $defaultOutputName):"
    
    # If no file name is provided, use the default name
    if ([string]::IsNullOrEmpty($fileName)) {
        $fileName = $defaultOutputName
    }
    
    # Prompt for the export path
    $exportPath = Read-Host "Enter the export path (default: current script directory):"
    
    # If no export path is provided, use the current script directory
    if ([string]::IsNullOrEmpty($exportPath)) {
        $scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Path
        $exportPath = Join-Path -Path $scriptDirectory -ChildPath $fileName
    } else {
        $exportPath = Join-Path -Path $exportPath -ChildPath $fileName
    }
    
    # Export the results to the specified path
    $results | Export-Csv -Path $exportPath -NoTypeInformation
    Write-Host "Results exported to $exportPath"
}

# Pause to keep the window open
Read-Host "Press Enter to exit"



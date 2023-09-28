<#
.SYNOPSIS
    This script fetches Organizational Unit (OU) information from Active Directory for specified computers.

.DESCRIPTION
    The script offers two modes:
    1. OU Only - Fetches just the OU information for the given computers.
    2. OU with History - Fetches OU information along with modification history.

    The data source can either be:
    1. A single computer specified by the user.
    2. A list of computers from a file (.txt or .csv).

.PARAMETER None
    The script prompts for all required information.

.EXAMPLE
    PS C:\> .\Get_OUv3.0.ps1

.NOTES
    File Name      : Get_OUv3.0.ps1
    Author         : Emperor Dor
    Prerequisite   : PowerShell V5, Run as Administrator
    Last Updated   : 23-August-2023

.LINK
    GitHub Repository - https://github.com/Doron-Bogomolov/Get_OU

#>

# Version Log
<#
    v1.0 - Initial release. Supports fetching OU for single computer.
    v2.0 - Added support for fetching OU from a list of computers.
    v3.0 - Added 'OU with History' mode.
#>

# [v1.0 Added Admin Privilege Check]
# Check if running with admin privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
    # Restart the script as admin
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# [v3.0 Added Windows Forms Assembly]
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')


# Step 1: User Input for Mode Selection


# [v3.0 Added Mode Selection]
# Initialize an empty variable for mode selection
$modeSelection = $null

# Loop until a valid mode is selected
do {
    # [v3.0 Added Mode Selection]
    # Display the menu for Mode Selection
    Write-Host "---------------- Mode Selection ----------------"
    Write-Host "1. OU Only"
    Write-Host "2. OU with History"
    Write-Host "------------------------------------------------"
    $modeSelection = Read-Host "Please enter the number corresponding to your choice"

    # Validate the user's choice
    if ($modeSelection -eq 1 -or $modeSelection -eq 2) {
        Write-Host "You selected mode $modeSelection."
        break # Exit the loop
    } else {
        Write-Host "Invalid selection. Please try again."
    }

} while ($modeSelection -ne 1 -and $modeSelection -ne 2)


# [v3.0 Added Source Selection]
# Initialize an empty variable for source selection
$sourceSelection = $null

# Loop until a valid source is selected
do {
    # [v3.0 Added Source Selection]
    # Display the menu for Data Source Selection
    Write-Host "------------- Data Source Selection -------------"
    Write-Host "1. Single Computer"
    Write-Host "2. List from File"
    Write-Host "------------------------------------------------"
    $sourceSelection = Read-Host "Please enter the number corresponding to your choice"

    # Validate the user's choice
    if ($sourceSelection -eq 1 -or $sourceSelection -eq 2) {
        Write-Host "You selected source $sourceSelection."
        break # Exit the loop
    } else {
        Write-Host "Invalid selection. Please try again."
    }

} while ($sourceSelection -ne 1 -and $sourceSelection -ne 2)





# [v1.0 Initial Function]
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


# [v3.0 Added Function]
# Function to get the OU Modification History of a computer
Function Get-OUModificationHistory {
    param (
        [string]$ComputerName
    )
    
    # Initialize an empty hash table to store the OU and modification details
    $OUHistory = @{}

try {
        # Fetch the AD computer object based on the ComputerName
        $ADComputer = Get-ADComputer $ComputerName -Properties "Created", "Modified", "DistinguishedName"
        
        # Extract the OU from the DistinguishedName
        $OU = ($ADComputer.DistinguishedName -split ",", 2)[1]

        # Populate the hash table with the OU and modification details
        $OUHistory['ComputerName'] = $ComputerName
        $OUHistory['OU'] = $OU
        $OUHistory['Created'] = $ADComputer.Created
        $OUHistory['Modified'] = $ADComputer.Modified

    } catch {
        Write-Host "An error occurred while fetching the OU and modification history for $ComputerName."
        $OUHistory['ComputerName'] = $ComputerName
        $OUHistory['Error'] = "Unable to fetch details"
    }

    # Return the hash table as output
    return $OUHistory
}

# [v3.0 Added Results Collection]
# Initialize an array to collect results
$resultsArray = @()


# Call Mode Selection and Data Source Selection functions (or code blocks)
# These should set the $modeSelection and $sourceSelection variables

# Initialize an array to store computer names
$computerNames = @()

# Determine the data source based on user selection
if ($sourceSelection -eq 1) {
    # Single Computer
    $computerNames += Read-Host "Please enter the computer name"
} elseif ($sourceSelection -eq 2) {
    # List from File
    $filePath = Read-Host "Please enter the path to the file containing the list of computer names"
    $computerNames = Get-Content $filePath
}

# Loop through each computer name and call the appropriate function based on mode selection
foreach ($computerName in $computerNames) {
    if ($modeSelection -eq 1) {
        # OU Only
        $result = Get-ComputerOU -ComputerName $computerName
        Write-Host "OU for $computerName is $result"
    } elseif ($modeSelection -eq 2) {
        # OU with History
        $result = Get-OUModificationHistory -ComputerName $computerName
        Write-Host "OU and modification history for $computerName is $result"
    }
    
    # Add the result to the results array
    $resultsArray += $result
}


# [v3.0 Added File Save Location]
# Ask the user where to save the file (GUI introduced in v3.0)
Write-Host "---------------- File Save Location ----------------"
Write-Host "1. Save in the same folder as this script"
Write-Host "2. Choose a different folder"
Write-Host "----------------------------------------------------"
$saveLocationChoice = Read-Host "Please enter the number corresponding to your choice"

# Initialize variable for export folder
$exportFolder = ""

# Determine the folder based on user selection (GUI introduced in v3.0)
if ($saveLocationChoice -eq 1) {
    $exportFolder = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
} elseif ($saveLocationChoice -eq 2) {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select a folder to save the exported file"
    $folderBrowser.ShowDialog() | Out-Null
    $exportFolder = $folderBrowser.SelectedPath
}


# Prompt for the file name
$defaultOutputName = "OU-List_{0:yyyy-MM-dd_HH-mm-ss}.csv" -f (Get-Date)
$exportFileName = Read-Host "Enter the export file name (default: $defaultOutputName):"
if ([string]::IsNullOrEmpty($exportFileName)) {
    $exportFileName = $defaultOutputName
}

# Combine the export folder and file name to get the full export path
$exportPath = Join-Path -Path $exportFolder -ChildPath $exportFileName

# Export the results to the specified path
$resultsArray | Export-Csv -Path $exportPath -NoTypeInformation
Write-Host "Results exported to $exportPath"

# Pause to keep the window open
Read-Host "Press Enter to exit"


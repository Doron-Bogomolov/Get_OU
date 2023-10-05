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
    PS C:\> .\Get_OUv3.3.ps1

.NOTES
    File Name      : Get_OUv3.4.ps1
    Author         : Emperor Dor
    Prerequisite   : PowerShell V5, Run as Administrator
    Last Updated   : 18-September-2023

.LINK
    GitHub Repository - https://github.com/Doron-Bogomolov/Get_OU

#>

# Version Log
<#
    v1.0 - Initial release. Supports fetching OU for single computer.
    v2.0 - Added support for fetching OU from a list of computers.
    v3.0 - Added 'OU with History' mode.
    v3.1 - Added GUI support for file and folder selection.
    v3.2 - Added error handling and screen clearing features.
    v3.3 - Implemented error logging to a hidden file.
    v3.4 - Added the option to not save to a file.
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


# [v3.3 - Implemented error logging to a hidden file]
$logFileLocation = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$logFile = Join-Path -Path $logFileLocation -ChildPath "logFile.txt"

# Check if the log file exists
if (-not (Test-Path $logFile)) {
    # Create the log file
    New-Item -Path $logFile -ItemType File
}

# Now set the item property
Set-ItemProperty -Path $logFile -Name Attributes -Value ([System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::Archive)

try {

    # [v3.0 Added Mode Selection]
    # Initialize an empty variable for mode selection
    $modeSelection = $null
    
    # [v3.0 Added Mode Selection]
    # Loop until a valid mode is selected
    do {
        # [v3.2 Added Screen Clearing]
        # Clear the screen to make the console output more readable
        Clear-Host
        
        # [v3.0 Added Mode Selection]
        # Display the menu for Mode Selection
        Write-Host "---------------- Mode Selection ----------------"
        Write-Host "1. OU Only"
        Write-Host "2. OU with History"
        Write-Host "------------------------------------------------"
        
        try {
        
            $modeSelection = Read-Host "Please enter the number corresponding to your choice"

            # [v3.0 Added Mode Selection]
            # Validate the user's choice
            if ($modeSelection -eq 1 -or $modeSelection -eq 2) {
                Write-Host "You selected mode $modeSelection."
                break # Exit the loop
            } else {
                Write-Host "Invalid selection. Please try again."
            }
        }
        catch {
            # [v3.2 Added Error Handling]
            # Added error handling to inform the user when an input error occurs
            Write-Host "An error occurred. Please try again."
            # [v3.3 - Enhanced Error Logging]
            $errorMessage = "An error occurred: $_"
            Add-Content -Path $logFile -Value $errorMessage
        }
    } while ($modeSelection -ne 1 -and $modeSelection -ne 2)



    # [v3.0 Added Source Selection]
    # Initialize an empty variable for source selection
    $sourceSelection = $null

    # [v3.0 Added Source Selection]
    # Loop until a valid source is selected
    do {
        # [v3.2 Added Screen Clearing]
        # Clear the screen
        Clear-Host
        
        # [v3.0 Added Source Selection]
        # Display the menu for Data Source Selection
        Write-Host "------------- Data Source Selection -------------"
        Write-Host "1. Single Computer"
        Write-Host "2. List from File"
        Write-Host "------------------------------------------------"
        
        
        # [v3.2 Added Error Handling] - Original
        # Wrap the input and validation logic in a try-catch block
        try {
            $sourceSelection = Read-Host "Please enter the number corresponding to your choice"

            # Validate the user's choice
            if ($sourceSelection -eq 1 -or $sourceSelection -eq 2) {
                Write-Host "You selected source $sourceSelection."
                break # Exit the loop
            } else {
                Write-Host "Invalid selection. Please try again."
            }
        }
        catch {
            Write-Host "An error occurred. Please try again."
            # [v3.3 - Enhanced Error Logging]
            $errorMessage = "An error occurred: $_"
            Add-Content -Path $logFile -Value $errorMessage
        }
    } while ($sourceSelection -ne 1 -and $sourceSelection -ne 2)





    # [v1.0 Initial Function]
    # Function to get the OU of a computer
    function Get-ComputerOU {
        param (
            [Parameter(Mandatory = $true, Position = 0)]
            [string]$ComputerName
        )
        
        # Initialize an empty hash table to store the OU and modification details
        $OUHash = @{}
        
        try {
            # Get the computer object using the Active Directory PowerShell module
            $computerObject = Get-ADComputer $ComputerName -ErrorAction Stop
            # Get the distinguished name (DN) of the computer object
            $ouDN = $computerObject.DistinguishedName
            # Return only the OU part of the DN
            $ou = $ouDN -replace "^CN=.*?,", ""
            
            # Populate the hash table with the OU and modification details
            $OUHash['ComputerName'] = $ComputerName
            $OUHash['OU'] = $ou
            return $OUHash
        }
        catch {
        # [v3.2 Enhanced Error Handling]
        # Added more comprehensive error handling to the function
        # Now returns a hash table with error details
        $OUHash['ComputerName'] = $ComputerName
        $OUHash['Error'] = "Not Found"
        # [v3.3 - Enhanced Error Logging]
        $errorMessage = "An error occurred: $_"
        Add-Content -Path $logFile -Value $errorMessage
        return $OUHash
        }
    }

    # [v3.0 Added Function]
    # Function to get the OU Modification History of a computer
    Function Get-OUModificationHistory {
        param (
            [string]$ComputerName
            )
        
        # [v3.0 Initialize an empty hash table]
        # Initialize an empty hash table to store the OU and modification details
        $OUHistory = @{}

    try {
            # [v3.0 Fetch the AD computer object]
            # Fetch the AD computer object based on the ComputerName
            $ADComputer = Get-ADComputer $ComputerName -Properties "Created", "Modified", "DistinguishedName"
            
            # [v3.0 Extract the OU]
            # Extract the OU from the DistinguishedName
            $OU = ($ADComputer.DistinguishedName -split ",", 2)[1]

            # [v3.0 Populate the hash table]
            # Populate the hash table with the OU and modification details
            $OUHistory['ComputerName'] = $ComputerName
            $OUHistory['OU'] = $OU
            $OUHistory['Created'] = $ADComputer.Created
            $OUHistory['Modified'] = $ADComputer.Modified

        } catch {
            # [v3.2 Enhanced Error Handling]
            # Improved error messages and populated the hash table with specific error details
            Write-Host "An error occurred while fetching the OU and modification history for $ComputerName."
            $OUHistory['ComputerName'] = $ComputerName
            $OUHistory['Error'] = "Unable to fetch details"
            # [v3.3 - Enhanced Error Logging]
            $errorMessage = "An error occurred: $_"
            Add-Content -Path $logFile -Value $errorMessage
        }
        
        # [v3.0 Return the hash table]
        # Return the hash table as output
        return $OUHistory
    }

    # [v3.0 Added Results Collection]
    # Initialize an array to collect results
    $resultsArray = @()

    # [v3.0 Call Mode Selection and Data Source Selection]
    # Call Mode Selection and Data Source Selection functions (or code blocks)
    # These should set the $modeSelection and $sourceSelection variables

    # [v1.0 Initial Data Source]
    # Initialize an array to store computer names
    $computerNames = @()

    # [v2.0 Added File Source]
    # Determine the data source based on user selection
    if ($sourceSelection -eq 1) {
        # Single Computer
        # [v3.2 Comment] User option for fetching OU for a single computer
        $computerNames += Read-Host "Please enter the computer name"
    } elseif ($sourceSelection -eq 2) {
        # [v3.1 Added GUI for File Selection]
        # [v3.2 Comment] User option for fetching OU for multiple computers from a file
        # Add the System.Windows.Forms assembly to use its classes
        Add-Type -AssemblyName System.Windows.Forms

        # Create a new OpenFileDialog object
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog

        # Set the title for the OpenFileDialog
        $openFileDialog.Title = "Select a PC names list file"

        # Set the filter to only show .txt and .csv files
        $openFileDialog.Filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv"

        # Disable multiple file selection
        $openFileDialog.Multiselect = $false

        # Show the OpenFileDialog and capture the result
        if ($openFileDialog.ShowDialog() -eq "OK") {
            # Get the full path of the selected file
            $filePath = $openFileDialog.FileName
            
            # Read the content of the selected file into the $computerNames array
            $computerNames = Get-Content $filePath
        } else {
            # If no file is selected, display a message and exit the script
            Write-Host "No file selected, exiting."
            exit
        }
    }

    # [v3.0 Added Loop Logic]
    # Loop through each computer name and call the appropriate function based on mode selection
    foreach ($computerName in $computerNames) {
        $result = @{}
        if ($modeSelection -eq 1) {
            # Execute the function for fetching OU in 'OU Only' mode
            $result = Get-ComputerOU -ComputerName $computerName
            # Write-Host "OU for $computerName is $result"  # Commented out
        } elseif ($modeSelection -eq 2) {
            # Execute the function for fetching OU in 'OU with History' mode
            $result = Get-OUModificationHistory -ComputerName $computerName
            # Write-Host "OU and modification history for $computerName is $result"  # Commented out
        }

        # [v3.0 Added Result Conversion]
        # Convert the hashtable to a custom object
        $resultObject = New-Object PSObject -Property $result
        
        # [v3.0 Added Result Collection]
        # Add the custom object to the results array
        $resultsArray += $resultObject
    }


    # [v3.0 Added Display Results]
    # Display the results
    $resultsArray | Format-Table


    # Initialize variable for save location choice
    $saveLocationChoice = $null

    # [v3.2 Added Results Storing]
    # Store the results in a string for re-display if needed
    $resultsString = $resultsArray | Format-Table | Out-String
    
        # Clear the screen
        Clear-Host
        
    # Output the stored results
    Write-Host $resultsString

    # Loop until a valid save location is selected
    do {
        # [v3.0 Added Save Location Prompt]
        # [v3.4 Added option 3 to exit without saving the output]        
        # Prompt the user to specify where the results should be saved 
        Write-Host "---------------- File Save Location ----------------"
        Write-Host "1. Save in the same folder as this script"
        Write-Host "2. Choose a different folder"
        Write-Host "3. Exit without saving"  #[v3.4 addon]
        Write-Host "----------------------------------------------------"
        # Capture the user's choice for save location
        $saveLocationChoice = Read-Host "Please enter the number corresponding to your choice"

        # Validate the user's choice for saving location
        if ($saveLocationChoice -eq 1 -or $saveLocationChoice -eq 2) {
            Write-Host "You selected source $saveLocationChoice."
            break # Exit the loop
        }

        # [v3.4 Added option to exit without saving the output]
        elseif ($saveLocationChoice -eq 3) {
            #Clearing screen
            Clear-Host

            #Print the results again for last display
            Write-Host $resultsString
            Write-Host "Exiting without saving."
            Write-Host "Press any key to exit immediately or wait 5 seconds."
            
            #Make the system wait for 5 seconds before exiting the instance
            Start-Sleep -Seconds 5
            exit # Exit the entire script
        }

        else {
        
            # Clear the screen if invalid choice is made
            Clear-Host
        
            Write-Host "`nInvalid selection. Please try again.`n"
            
            # Re-display the results and prompt for choice again
            Write-Host $resultsString
        }
        
        # Continue looping until a valid choice is made
    } while ($saveLocationChoice -ne 1 -and $saveLocationChoice -ne 2)


    # [v3.0 Added Export Folder Initialization]
    # Initialize a variable to hold the export folder path6
    $exportFolder = ""

    # [v3.0 Added Folder Selection]
    # Determine the export folder based on user's save location choice
    if ($saveLocationChoice -eq 1) {
        # Use the folder where the script is located
        $exportFolder = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
    } elseif ($saveLocationChoice -eq 2) {
        # Open folder browser dialog
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select a folder to save the exported file"
        $folderBrowser.ShowDialog() | Out-Null
        $exportFolder = $folderBrowser.SelectedPath
    }

    # [v3.0 Added File Name Prompt]
    # Prompt the user for the name of the export file with a default name based on the current date and time
    $defaultOutputName = "OU-List_{0:yyyy-MM-dd_HH-mm-ss}.csv" -f (Get-Date)
    $exportFileName = Read-Host "Enter the export file name (default: $defaultOutputName):"
    
    # Validate the entered export file name. If it is null or empty, use the default name
    if ([string]::IsNullOrEmpty($exportFileName)) {
        $exportFileName = $defaultOutputName
    } else {
        # Validate the file extension of the user-provided file name
        $fileExtension = [System.IO.Path]::GetExtension($exportFileName)
        
        # If the extension is neither .csv nor .txt, default to .csv
        if ($fileExtension -ne ".csv" -and $fileExtension -ne ".txt") {
            Write-Host "Invalid file extension. Defaulting to .csv."
            $exportFileName += ".csv"
        }
    }

    # [v3.0 Added Export]
    # Combine the export folder path and file name to form the full path where the CSV will be saved
    $exportPath = Join-Path -Path $exportFolder -ChildPath $exportFileName

    # Export the collected results to a CSV file at the specified path
    $resultsArray | Export-Csv -Path $exportPath -NoTypeInformation
    Write-Host "Results exported to $exportPath"

    # [v3.0 Added Pause]
    # Pause the script to allow the user to read the final output message
    Read-Host "Press Enter to exit"
    
    
} catch {
    # [v3.3 - Added Time-Stamped Error Logging in Final Catch]
    # Log the time of the error first
    $errorTime = Get-Date
    Add-Content -Path $logFile -Value "Time of Error: $errorTime"

    # Then log the error message
    $errorMessage = "Error occurred: " + $_.Exception.Message
    Add-Content -Path $logFile -Value $errorMessage
}


<#
.SYNOPSIS
    Automated Excel workbook processor that refreshes data connections in multiple Excel files.

.DESCRIPTION
    This PowerShell script automates the process of refreshing data connections in Excel workbooks (.xlsm files).
    It includes robust error handling, retry logic, and logging capabilities. The script processes all Excel
    files in a specified directory, handling each file by opening it, refreshing all data connections,
    saving changes, and properly closing the file.

.FEATURES
    - Single instance enforcement to prevent multiple script executions
    - Comprehensive logging system with timestamp
    - Automatic handling of native database query dialogs
    - Retry logic for failed operations
    - Proper COM object cleanup
    - File locking detection
    - Existing Excel process management

.PARAMETERS
    None - All configuration is handled through constants defined in the script

.CONSTANTS
    WAIT_FOR_EXCEL_TIMEOUT_SECONDS = 40    # Maximum time to wait for Excel to be ready
    REFRESH_DELAY_SECONDS = 30             # Delay before refreshing data
    REFRESH_WAIT_SECONDS = 40              # Wait time after initiating refresh
    SAVE_CLOSE_DELAY_SECONDS = 30          # Delay between save and close operations
    ITERATION_DELAY_SECONDS = 15           # Delay between processing different files
    MAX_RETRY_ATTEMPTS = 3                 # Maximum number of retry attempts per file
    RETRY_DELAY_SECONDS = 10               # Delay between retry attempts

.INPUTS
    None - The script processes Excel files from a predefined directory path

.OUTPUTS
    Log entries are written to:
    "[LOG_FILE_PATH]"

.REQUIREMENTS
    - Windows PowerShell 5.1 or later
    - Microsoft Excel
    - OpenXML SDK V2.5 (DocumentFormat.OpenXml.dll)
    - Appropriate permissions to access and modify Excel files
    - Write permissions for the log file location

.EXAMPLE
    .\ExcelFileProcessor.ps1

.NOTES
    File Name      : ExcelFileProcessor.ps1
    Author         : [Your Name]
    Prerequisite   : PowerShell 5.1 or later
    Copyright      : [Your Organization]
    Version        : 1.0
    Created Date   : [Creation Date]
    Modified Date  : [Last Modified Date]

.LINK
    [Repository URL]
#>
# Log file path
$logFilePath = "[LOG_FILE_PATH]"

# Function to log messages to a file
function Write-Log {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Write-Host $logMessage
    Add-Content -Path $logFilePath -Value $logMessage
}

# Constants
$WAIT_FOR_EXCEL_TIMEOUT_SECONDS = 40
$REFRESH_DELAY_SECONDS = 30
$REFRESH_WAIT_SECONDS = 40
$SAVE_CLOSE_DELAY_SECONDS = 30
$ITERATION_DELAY_SECONDS = 15
$MAX_RETRY_ATTEMPTS = 3
$RETRY_DELAY_SECONDS = 10

# Log the start of the script
Write-Log "Script starting..."

# Load the OpenXML assembly for manipulating Excel files
try {
    Add-Type -Path "[OPENXML_SDK_PATH]"
    Write-Log "Loaded OpenXML assembly successfully."
} catch {
    Write-Log "Failed to load OpenXML assembly: $_"
}

# Function to wait until Excel is ready
function WaitForExcelReady {
    param($excel)
    $timeout = New-TimeSpan -Seconds $WAIT_FOR_EXCEL_TIMEOUT_SECONDS
    $startTime = Get-Date
    while (-not $excel.Ready) {
        if ((Get-Date) - $startTime -ge $timeout) {
            Write-Log "Timeout: Excel did not become ready within $($WAIT_FOR_EXCEL_TIMEOUT_SECONDS) seconds."
            return $false
        }
        Start-Sleep -Seconds 1
    }
    return $true
}

# Function to handle the native database query message box
function HandleDatabaseQueryMessage {
    param($excel)
    $nativeQueryMessageBox = $excel.Windows | Where-Object { $_.Caption -eq "Microsoft Excel" -and $_.Text -like "*Native Database Query*" }
    if ($null -ne $nativeQueryMessageBox) {
        $nativeQueryMessageBox.Activate()
        $nativeQueryMessageBox.SendKeys("{ENTER}")
    }
}

# Function to check if a file is open
function IsFileOpen {
    param ($file)
    try {
        $fileStream = [System.IO.File]::Open($file, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
        $fileStream.Close()
        return $false
    } catch {
        return $true
    }
}

# Function to close any existing Excel instances
function CloseExistingExcelInstances {
    $excelProcesses = Get-Process -Name Excel -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        $excelProcesses | ForEach-Object { $_.Kill() }
        Write-Log "Closed existing Excel instances."
    }
}

# Function to process an individual Excel file with retry logic
function ProcessExcelFile {
    param ($excelFile)
    $retryCount = 0
    $success = $false

    while (-not $success -and $retryCount -lt $MAX_RETRY_ATTEMPTS) {
        try {
            Write-Log "Processing Excel file: $($excelFile.FullName)"
            
            # Check if the file is open
            if (IsFileOpen($excelFile.FullName)) {
                Write-Log "Skipping file: $($excelFile.FullName) - The file is open or in use by another process."
                return
            }

            # Close any existing Excel instances
            CloseExistingExcelInstances

            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false

            try {
                $workbook = $excel.Workbooks.Open($excelFile.FullName)
            } catch {
                Write-Log "Skipping file: $($excelFile.FullName) - The file is open or in use by another process."
                $excel.Quit()
                return
            }

            # Wait for Excel to be ready
            if (-not (WaitForExcelReady $excel)) {
                Write-Log "Skipping file: $($excelFile.FullName) - Excel is not ready."
                $excel.Quit()
                return
            }

            Start-Sleep -Seconds $REFRESH_DELAY_SECONDS

            $refreshStartTime = Get-Date

            # Refresh the data in the workbook
            $workbook.RefreshAll()

            Start-Sleep -Seconds $REFRESH_WAIT_SECONDS

            # Handle any database query messages
            HandleDatabaseQueryMessage $excel

            $refreshEndTime = Get-Date

            Write-Log "Refresh operation took: $((New-TimeSpan -Start $refreshStartTime -End $refreshEndTime).Seconds) seconds."

            Start-Sleep -Seconds $SAVE_CLOSE_DELAY_SECONDS

            # Save and close the workbook
            $workbook.Save()
            $null = $workbook.Close()

            Start-Sleep -Seconds $SAVE_CLOSE_DELAY_SECONDS

            $excel.Quit()

            $success = $true
        } catch {
            Write-Log "Error occurred while processing Excel file: $($excelFile.FullName)"
            Write-Log "Error message: $_"
            $retryCount++
            Write-Log "Retrying... ($retryCount/$MAX_RETRY_ATTEMPTS)"
            Start-Sleep -Seconds $RETRY_DELAY_SECONDS
        } finally {
            # Release COM objects to free up memory
            if ($null -ne $workbook) {
                $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
                $workbook = $null
            }
            if ($null -ne $excel) {
                $excel.Quit()
                $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
                $excel = $null
            }
            Start-Sleep -Seconds $ITERATION_DELAY_SECONDS
        }
    }

    if (-not $success) {
        Write-Log "Failed to process file: $($excelFile.FullName) after $MAX_RETRY_ATTEMPTS attempts."
    }
}

# Function to process all Excel files in the specified folder
function ProcessExcelFiles {
    $rootFolder = "[EXCEL_FILES_PATH]"
    Write-Log "Looking for Excel files in $rootFolder"
    
    $excelFiles = Get-ChildItem -Path $rootFolder -Filter *.xlsm

    foreach ($excelFile in $excelFiles) {
        ProcessExcelFile -excelFile $excelFile
    }
}

# Ensure only a single instance of the script is running
$scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$runningProcesses = Get-Process | Where-Object { $_.ProcessName -eq $scriptName -and $_.Id -ne $PID }
if ($runningProcesses.Count -gt 0) {
    Write-Log "Another instance of the script is already running. Exiting."
    exit
}

# Main execution
try {
    $startTime = Get-Date
    Write-Log "Script started at $startTime"

    # Process all Excel files
    ProcessExcelFiles

    $endTime = Get-Date
    Write-Log "Total script execution time: $((New-TimeSpan -Start $startTime -End $endTime).Seconds) seconds."
    
    Write-Log "Script completed successfully."

} catch {
    Write-Log "Unexpected error: $_"
}


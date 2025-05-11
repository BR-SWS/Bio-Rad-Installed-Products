# Define the Excel file path and log file path
$excelFilePath = "C:\Users\u108298\OneDrive - Bio-Rad Laboratories Inc\Bio-Rad-Installed-Products\InstalledProducts.xlsx"
$logFilePath = "C:\Users\u108298\OneDrive - Bio-Rad Laboratories Inc\Bio-Rad-Installed-Products\AUTOREFESH\Refresh_log.txt"

# Function to write messages to the log file and screen
function Write-Log {
    param (
        [string]$message
    )

    # Get the current date and time
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Format the log message with timestamp
    $logMessage = "$timestamp - $message"

    # Append the log message to the log file
    Add-Content -Path $logFilePath -Value "$logMessage"

    # Write message to PowerShell console
    Write-Host $logMessage
}

# Start logging
Write-Log "Starting the process to open Excel and refresh data..."

# Function to check if file is locked
function Test-FileLock {
    param (
        [string]$filePath
    )

    try {
        $fileStream = [System.IO.File]::Open($filePath, 'Open', 'ReadWrite', 'None')
        $fileStream.Close()
        return $false  # File is not locked
    }
    catch {
        return $true  # File is locked
    }
}

# Wait for the file to become available (up to 5 minutes)
$maxRetries = 10
$retryCount = 0

while (Test-FileLock -filePath $excelFilePath -eq $true -and $retryCount -lt $maxRetries) {
    Write-Log "File is locked. Retrying in 30 seconds..."
    Start-Sleep -Seconds 30
    $retryCount++
}

# If the file is still locked after retries, forcefully kill Excel (or other processes)
if ($retryCount -ge $maxRetries) {
    Write-Log "File is still locked after retries. Attempting to kill the locking process..."

    # Identify and kill the process locking the file
    $lockingProcess = Get-Process | Where-Object { $_.Modules.FileName -eq $excelFilePath }

    if ($lockingProcess) {
        Write-Log "Killing process: $($lockingProcess.Name) (PID: $($lockingProcess.Id))"
        Stop-Process -Id $lockingProcess.Id -Force
    }
    else {
        Write-Log "No locking process found. Please manually check the file."
        exit
    }
}

# Proceed with opening Excel
Write-Log "Opening Excel workbook: $excelFilePath..."

# Create Excel COM object to open and interact with Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true  # Set to $true for visible Excel window while debugging

try {
    # Open the workbook
    $workbook = $excel.Workbooks.Open($excelFilePath)

    # Log refreshing data from Power Query
    Write-Log "Refreshing data from Power Query..."

    # Refresh all data connections
    $workbook.RefreshAll()

    # Wait for 3 minutes to allow refresh
    Write-Log "Waiting for data refresh..."
    Start-Sleep -Seconds 180

    # Save and close the workbook
    Write-Log "Saving and closing the workbook..."
    $workbook.Close($true)
    $excel.Quit()

    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Log the completion of the process
    Write-Log "Process complete. Excel workbook refreshed and closed."
}
catch {
    $errorMsg = $_.Exception.Message
    Write-Log "Error occurred: $errorMsg"
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    exit 1
}

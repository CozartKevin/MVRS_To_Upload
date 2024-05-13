# Define paths
$backupFolder = "C:\Scripts\TransferredZips"
$extractFolder = "C:\Scripts\TransferredZips"
$uploadFolder = "C:\MVRS\XFER\Upload"
$uploadFile = "UPLOAD.dat"
$logFolder = "C:\Scripts\Logs"
$errorLog = "error.log"
$scriptLog = "script.log"
$downloadsFolder = "C:\Users\$env:USERNAME\Downloads"
$datFileNotFound = $false

# Function to ensure directory exists
function Test-EnsureDirectory {
    param (
        [string]$Path
    )
    if (-not (Test-Path -Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force
        Write-Verbose "Directory created: $Path"
        Write-Log -Message "Directory created: $Path" -LogFile $scriptLog
    }
    else {
        Write-Verbose "Directory already exists: $Path"
        Write-Log -Message "Directory already exists: $Path" -LogFile $scriptLog
    
    }
}

# Function to write log messages
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFile
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    $logPath = Join-Path -Path $logFolder -ChildPath $LogFile

    if (Test-Path -Path $logPath) {
        $existingLogs = Get-Content -Path $logPath
        $newLogs = $logEntry, $existingLogs
        Set-Content -Path $logPath -Value $newLogs
    }
    else {
        Add-Content -Path $logPath -Value $logEntry
    }
}

# Ensure all required directories exist
Test-EnsureDirectory -Path $logFolder
Test-EnsureDirectory -Path $backupFolder
Test-EnsureDirectory -Path $extractFolder
Test-EnsureDirectory -Path $uploadFolder

# Log script start
Write-Log -Message "Script started." -LogFile $scriptLog

## ZIP Checks and Moving to Backup Folder

# Find the most recent zip file in the Downloads folder
$recentZip = Get-ChildItem -Path $downloadsFolder -Filter *.zip | Sort-Object LastWriteTime -Descending | Select-Object -First 1

# Check if a zip file was found
if (-not $recentZip) {
    $errorMsg = "No zip file found in Downloads folder."
    Write-Output $errorMsg
    Write-Log -Message $errorMsg -LogFile $errorLog
    exit 1
}

Write-Verbose "Processing '$recentZip'."
Write-Log -Message "Processing '$recentZip'." -LogFile $scriptLog

# Move zip file to backup folder
try {
    $destinationPath = Join-Path -Path $backupFolder -ChildPath $recentZip.Name
    Move-Item -Path $recentZip.FullName -Destination $destinationPath -Force
    Write-Verbose "ZIP file '$($recentZip.Name)' moved to '$backupFolder'."
    Write-Log -Message "ZIP file '$($recentZip.Name)' moved to '$backupFolder'." -LogFile $scriptLog
}
catch {
    $errorMsg = "Failed to move zip file. Error: $_"
    Write-Output $errorMsg
    Write-Log -Message $errorMsg -LogFile $errorLog
    exit 1
}

# Ensure the zip file exists in new location
if (-not (Test-Path -Path $destinationPath -PathType Leaf)) {
    $errorMsg = "Zip file does not exist or is not accessible. Path: $destinationPath"
    Write-Output $errorMsg
    Write-Log -Message $errorMsg -LogFile $errorLog
    exit 1
}

## ZIP Extract to Backup Folder and .dat file check then Move to UploadFolder


try {
    # Try to extract zip contents to Extract folder
    Expand-Archive -LiteralPath $destinationPath -DestinationPath $extractFolder -Force -ErrorAction Stop
    
    # Output a message indicating successful extraction
    Write-Verbose "ZIP file '$($recentZip.Name)' successfully extracted to '$extractFolder'."
    Write-Log -Message "ZIP file '$($recentZip.Name)' successfully extracted to '$extractFolder'." -LogFile $scriptLog

    # Find the most recent .dat file in the extracted folder
    $recentDat = Get-ChildItem -Path $extractFolder -Filter *.dat | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    # Check if a .dat file was found
    if (-not $recentDat) {
        $datFileNotFound = $true
        throw "No .dat file found in the extracted contents."
    }

    # Rename the most recent .dat file to Upload.dat
    Rename-Item -Path $recentDat.FullName -NewName $uploadFile -Force

    # Move the renamed Upload.dat file to the upload folder
    Move-Item -Path (Join-Path -Path $extractFolder -ChildPath $uploadFile) -Destination $uploadFolder -Force -ErrorAction Stop

    # Output a message indicating successful movement
    Write-Verbose "File '$recentDat' was successfully renamed '$uploadFile' and moved to '$uploadFolder'."
    Write-Log -Message "File '$recentDat' was successfully renamed '$uploadFile' and moved to '$uploadFolder'." -LogFile $scriptLog

}
catch {
    $errorMsg = "Failed to process zip file '$($recentZip.Name)'. Error: $_"
    Write-Output $errorMsg
    Write-Log -Message $errorMsg -LogFile $errorLog
    exit 1
}
finally {

    # Cleanup of any non .zip, .dat or .bat files including folders
    try {
        # Remove files
        Get-ChildItem -Path $extractFolder -Exclude *.zip, *.dat, *.bat -Recurse | Where-Object { -not $_.PSIsContainer } | Remove-Item -Force -ErrorAction Stop
    
        # Remove all directories
        Get-ChildItem -Path $extractFolder -Directory -Recurse | Remove-Item -Recurse -Force -ErrorAction Stop

        Write-Verbose "Non-.zip and non-.bat files and all directories removed from '$extractFolder'."
        Write-Log -Message "Non-.zip and non-.bat files and all directories removed from '$extractFolder'." -LogFile $scriptLog


        if ($datFileNotFound) {
            Remove-Item -Path $destinationPath -Force -ErrorAction Stop  
            Write-Verbose "ZIP file '$($recentZip.Name)' removed from '$backupFolder' due to missing .dat file."
            Write-Log -Message "ZIP file '$($recentZip.Name)' removed from '$backupFolder' due to missing .dat file." -LogFile $scriptLog
        }


    }
    catch {
        $errorMsg = "Failed to remove non-.zip and non-.bat files from the extraction folder. Error: $_"
        Write-Output $errorMsg
        Write-Log -Message $errorMsg -LogFile $errorLog
        exit 1
    }
}

# Log script end
Write-Output "File '$recentDat' successfully extracted to '$uploadFolder'.  Script executed successfully. "
Write-Log -Message "File '$recentDat' successfully extracted to '$uploadFolder'.  Script executed successfully." -LogFile $scriptLog

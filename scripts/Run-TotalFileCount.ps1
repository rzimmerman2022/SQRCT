<#
.SYNOPSIS
    Counts all files in a target directory and logs the result.

.DESCRIPTION
    This script recursively scans a target folder to get a total count of files.
    It logs the start, end, and result to a transcript file.
    It attempts to use a primary network location for logs, falling back to an alternative
    if the primary is not accessible.

.NOTES
    Version: 1.1
    Author: AI Assistant
    Last Modified: 2025-05-06

    Ensure the account running this script has read access to the target folder
    and write access to the chosen log directory.
    The Get-ChildItem command can be time-intensive for very large directories.
#>

# --- Configuration ---
$TargetFolderPath = "W:\" # The folder to analyze

# Define the primary log directory (network path)
$LogDirectoryPrimary = "\\sc-file02.ad.sciencecare.com\UEMProfiles\Ryan.Zimmerman\Desktop\PS_FileAnalysis_Logs"

# Define an alternative log directory (local or mapped drive)
# <<<< IMPORTANT: ADJUST THIS ALTERNATIVE PATH IF NEEDED >>>>
$LogDirectoryAlternative = "R:\Ryan Zimmerman\PS_FileAnalysis_Logs"
# Example local path:
# $LogDirectoryAlternative = "C:\Temp\PS_FileAnalysis_Logs"

$LogDirectory = "" # This will be set to the chosen valid path

# --- Function to Select and Create Log Directory ---
Function Set-LogDirectory {
    param(
        [string]$PrimaryPath,
        [string]$AlternativePath
    )
    Write-Host "Attempting to set up log directory..."
    # Try Primary Path
    if (Test-Path -Path $PrimaryPath -PathType Container) {
        Write-Host "Primary log directory found: $PrimaryPath"
        return $PrimaryPath
    } elseif (Test-Path -Path (Split-Path -Path $PrimaryPath -Parent) -PathType Container) {
        Write-Host "Parent of primary log directory found. Attempting to create log directory: $PrimaryPath"
        try {
            New-Item -ItemType Directory -Path $PrimaryPath -ErrorAction Stop -Force | Out-Null
            Write-Host "Successfully created primary log directory: $PrimaryPath"
            return $PrimaryPath
        } catch { Write-Warning "Could not create primary log directory '$PrimaryPath'. Error: $($_.Exception.Message)" }
    } else { Write-Warning "Primary log directory or its parent not accessible: $PrimaryPath" }

    Write-Warning "Attempting to use alternative log directory..."
    # Try Alternative Path
    if (Test-Path -Path $AlternativePath -PathType Container) {
        Write-Host "Alternative log directory found: $AlternativePath"
        return $AlternativePath
    } elseif (Test-Path -Path (Split-Path -Path $AlternativePath -Parent) -PathType Container) {
        Write-Host "Parent of alternative log directory found. Attempting to create log directory: $AlternativePath"
        try {
            New-Item -ItemType Directory -Path $AlternativePath -ErrorAction Stop -Force | Out-Null
            Write-Host "Successfully created alternative log directory: $AlternativePath"
            return $AlternativePath
        } catch { Write-Error "Could not create alternative log directory '$AlternativePath'. Error: $($_.Exception.Message)" }
    } else { Write-Error "Alternative log directory or its parent not accessible: $AlternativePath" }
    return $null
}

# --- Set the Log Directory ---
$LogDirectory = Set-LogDirectory -PrimaryPath $LogDirectoryPrimary -AlternativePath $LogDirectoryAlternative
if (-not $LogDirectory) {
    Write-Error "FATAL: No valid log directory could be established. Please check paths and permissions. Exiting."
    exit 1
}

# --- Define Log File Paths ---
$TimestampSuffix = Get-Date -Format 'yyyyMMdd_HHmmss'
$TranscriptLogFile = Join-Path -Path $LogDirectory -ChildPath "TotalFileCount_Transcript_$TimestampSuffix.log"
# No separate data CSV for this simple count, transcript is sufficient.

# --- Initialize Logging ---
try {
    Start-Transcript -Path $TranscriptLogFile -Append -ErrorAction Stop
} catch {
    Write-Error "FATAL: Could not start transcript at '$TranscriptLogFile'. Error: $($_.Exception.Message). Exiting."
    exit 1
}

Function Write-LogMessage ($Message) {
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] $Message"
    Write-Host $LogEntry # Output to console (captured by transcript)
}

Write-LogMessage "SCRIPT START: Total File Count"
Write-LogMessage "Target Folder: $TargetFolderPath"
Write-LogMessage "Using Log Directory: $LogDirectory"
Write-LogMessage "Transcript Log: $TranscriptLogFile"

# --- Count Files ---
$TotalFiles = 0
$FileCollectionTimeSeconds = 0
Write-LogMessage "Attempting to count all files... This may take a while."

$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
try {
    # Get-ChildItem itself is the main operation. Progress is implicit to its execution.
    # For extremely large directories, PowerShell might appear to hang here.
    # The -Force parameter can help include hidden or system files if needed, but usually not for general counts.
    $FileCollection = Get-ChildItem -Path $TargetFolderPath -File -Recurse -ErrorAction SilentlyContinue
    $TotalFiles = $FileCollection.Count # Counting after collection
    # Alternative for potentially lower memory on very large sets, but slower:
    # $TotalFiles = (Get-ChildItem -Path $TargetFolderPath -File -Recurse -ErrorAction SilentlyContinue | Measure-Object).Count
    $Stopwatch.Stop()
    $FileCollectionTimeSeconds = $Stopwatch.Elapsed.TotalSeconds
    Write-LogMessage ("Total files found: {0} (Collection took {1:N2} seconds)" -f $TotalFiles, $FileCollectionTimeSeconds)
} catch {
    $Stopwatch.Stop()
    $FileCollectionTimeSeconds = $Stopwatch.Elapsed.TotalSeconds
    Write-LogMessage ("ERROR: Could not complete file count (ran for {0:N2} seconds). Error: {1}" -f $FileCollectionTimeSeconds, $_.Exception.Message)
}

Write-LogMessage "SCRIPT END: Total File Count"
Stop-Transcript

Write-Host "Script finished. Log is available in $TranscriptLogFile"

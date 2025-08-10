<#
.SYNOPSIS
    Analyzes file path lengths in a target directory, logs results, and shows progress.

.DESCRIPTION
    This script recursively scans a target folder to identify files with path lengths
    greater than or equal to 247 characters and those with shorter paths.
    It provides progress updates and logs detailed information to transcript and data files.
    It attempts to use a primary network location for logs, falling back to an alternative
    if the primary is not accessible.

.NOTES
    Version: 1.1
    Author: AI Assistant
    Last Modified: 2025-05-06

    Ensure the account running this script has read access to the target folder
    and write access to the chosen log directory.
    The initial step of gathering all file items can be memory and time-intensive
    for very large directories.
#>

# --- Configuration ---
$TargetFolderPath = "W:\" # The folder to analyze (e.g., "W:\", "C:\Users\YourUser\Documents")

# Define the primary log directory (network path)
# IMPORTANT: Ensure the user running the script has WRITE PERMISSIONS to this network path.
$LogDirectoryPrimary = "\\sc-file02.ad.sciencecare.com\UEMProfiles\Ryan.Zimmerman\Desktop\PS_FileAnalysis_Logs"

# Define an alternative log directory (local or mapped drive)
# <<<< IMPORTANT: ADJUST THIS ALTERNATIVE PATH IF NEEDED >>>>
# Based on your image, if R: is "Client Relations (R:)" and you have a "Ryan Zimmerman" folder in it:
$LogDirectoryAlternative = "R:\Ryan Zimmerman\PS_FileAnalysis_Logs"
# If R: is just R: and you want a folder on its root:
# $LogDirectoryAlternative = "R:\PS_FileAnalysis_Logs"
# Or a local path:
# $LogDirectoryAlternative = "C:\Temp\PS_FileAnalysis_Logs"


$LogDirectory = "" # This will be set to the chosen valid path

$ProgressUpdateInterval = 1000  # Update progress bar every N files
$SummaryUpdateIntervalMinutes = 2 # Write interim summary to data log every N minutes

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
        } catch {
            Write-Warning "Could not create primary log directory '$PrimaryPath'. Error: $($_.Exception.Message)"
        }
    } else {
         Write-Warning "Primary log directory or its parent not accessible: $PrimaryPath"
    }

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
        } catch {
            Write-Error "Could not create alternative log directory '$AlternativePath'. Error: $($_.Exception.Message)"
        }
    } else {
        Write-Error "Alternative log directory or its parent not accessible: $AlternativePath"
    }

    # Fallback to script's current directory if all else fails (optional)
    # Write-Warning "Falling back to script's current directory for logs."
    # $scriptDirLog = Join-Path -Path $PSScriptRoot -ChildPath "PS_FileAnalysis_Logs_Fallback"
    # try {
    #    New-Item -ItemType Directory -Path $scriptDirLog -ErrorAction Stop -Force | Out-Null
    #    Write-Host "Successfully created fallback log directory: $scriptDirLog"
    #    return $scriptDirLog
    # } catch {
    #    Write-Error "Could not create fallback log directory '$scriptDirLog'. Error: $($_.Exception.Message)"
    # }

    return $null # Return null if no directory could be set
}

# --- Set the Log Directory ---
$LogDirectory = Set-LogDirectory -PrimaryPath $LogDirectoryPrimary -AlternativePath $LogDirectoryAlternative

if (-not $LogDirectory) {
    Write-Error "FATAL: No valid log directory could be established. Please check paths and permissions. Exiting."
    exit 1
}

# --- Define Log File Paths ---
$TimestampSuffix = Get-Date -Format 'yyyyMMdd_HHmmss'
$TranscriptLogFile = Join-Path -Path $LogDirectory -ChildPath "PathLengthAnalysis_Transcript_$TimestampSuffix.log"
$DataLogFile = Join-Path -Path $LogDirectory -ChildPath "PathLengthAnalysis_Data_$TimestampSuffix.csv" # Changed to CSV for easier import

# --- Initialize Logging ---
try {
    Start-Transcript -Path $TranscriptLogFile -Append -ErrorAction Stop
} catch {
    Write-Error "FATAL: Could not start transcript at '$TranscriptLogFile'. Error: $($_.Exception.Message). Exiting."
    exit 1
}

Function Write-LogMessage ($Message, [OutputType([System.Management.Automation.PSCustomObject])] $DataForCsv = $null) {
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] $Message"
    Write-Host $LogEntry # Output to console (captured by transcript)

    if ($null -ne $DataForCsv) {
        try {
            # Append to CSV data log file
            # Check if file exists to write header only once
            if (-not (Test-Path $DataLogFile)) {
                # Create header based on the properties of the first data object
                ($DataForCsv | Select-Object * | Get-Member -MemberType NoteProperty).Name | Sort-Object | ForEach-Object {$_ -join ","} | Set-Content -Path $DataLogFile
            }
            $DataForCsv | Export-Csv -Path $DataLogFile -Append -NoTypeInformation -Force
        } catch {
            Write-Warning "Failed to write data point to $DataLogFile. Error: $($_.Exception.Message)"
        }
    } elseif ($Message -notlike "---*" -and $Message -notlike "SCRIPT*") { # Avoid writing headers/footers as data if no specific data object
         # If you want general messages in the data log too, uncomment below
         # Add-Content -Path $DataLogFile -Value $LogEntry
    }
}

Write-LogMessage "SCRIPT START: Path Length Analysis"
Write-LogMessage "Target Folder: $TargetFolderPath"
Write-LogMessage "Using Log Directory: $LogDirectory"
Write-LogMessage "Transcript Log: $TranscriptLogFile"
Write-LogMessage "Data Log (CSV): $DataLogFile"
Write-LogMessage "Progress Update Interval: $ProgressUpdateInterval files"
Write-LogMessage "Summary Update Interval (for CSV): $SummaryUpdateIntervalMinutes minutes"

# --- Stage 1: Get all file items and total count (can take time and be memory intensive) ---
$AllFiles = @()
$TotalFiles = 0
Write-LogMessage "Attempting to retrieve all file items and count... This may take a while."
try {
    # For very large directories, consider Get-ChildItem with -PipelineVariable for streaming if memory is an issue,
    # but then getting $TotalFiles accurately upfront for progress is harder.
    $AllFiles = Get-ChildItem -Path $TargetFolderPath -File -Recurse -ErrorAction SilentlyContinue
    $TotalFiles = $AllFiles.Count
    Write-LogMessage "Successfully retrieved file list. Total files to process: $TotalFiles"
} catch {
    Write-LogMessage "ERROR: Could not retrieve file list. $($_.Exception.Message)"
    Stop-Transcript
    exit 1
}

if ($TotalFiles -eq 0) {
    Write-LogMessage "No files found in the target directory. Exiting."
    Stop-Transcript
    exit 0
}

# --- Stage 2: Analyze Path Lengths ---
$LongPathCount = 0
$ShortPathCount = 0
$ProcessedFilesCount = 0
$ErrorDuringProcessingCount = 0
$LastSummaryWriteTime = Get-Date

Write-LogMessage "Starting path length analysis loop..."
# Write-LogMessage "InterimSummaryTimestamp,LongPathCount,ShortPathCount,ProcessedFilesCount,ErrorsDuringProcessing" -IsDataPoint $true # CSV Header for data

foreach ($FileItem in $AllFiles) {
    $ProcessedFilesCount++
    try {
        if ($null -eq $FileItem.FullName) { # Handle potential null FullName property
            Write-Warning "File item with null FullName encountered. Name: $($FileItem.Name), Path: $($FileItem.DirectoryName)"
            $ErrorDuringProcessingCount++
            continue # Skip this item
        }

        if ($FileItem.FullName.Length -ge 247) {
            $LongPathCount++
            # Optional: Log individual long paths to a separate detailed log if needed.
            # This would make the main data log cleaner (summary only).
            # Add-Content -Path (Join-Path $LogDirectory "LongPaths_Detailed.log") -Value $FileItem.FullName
        } else {
            $ShortPathCount++
        }
    } catch {
        Write-LogMessage "ERROR: Could not process file '$($FileItem.Name)' (Path: $($FileItem.DirectoryName)) - $($_.Exception.Message)"
        $ErrorDuringProcessingCount++
    }

    # Update Progress Bar
    if ($ProcessedFilesCount % $ProgressUpdateInterval -eq 0 -or $ProcessedFilesCount -eq $TotalFiles) {
        Write-Progress -Activity "Analyzing File Paths" `
                       -Status "Processed $ProcessedFilesCount of $TotalFiles files. Long: $LongPathCount Short: $ShortPathCount Errors: $ErrorDuringProcessingCount" `
                       -PercentComplete (($ProcessedFilesCount / $TotalFiles) * 100) `
                       -Id 1

        # Periodically write summary to data CSV log
        if (((Get-Date) - $LastSummaryWriteTime).TotalMinutes -ge $SummaryUpdateIntervalMinutes -or $ProcessedFilesCount -eq $TotalFiles) {
            $summaryData = [PSCustomObject]@{
                InterimSummaryTimestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                LongPathCount           = $LongPathCount
                ShortPathCount          = $ShortPathCount
                ProcessedFilesCount     = $ProcessedFilesCount
                ErrorsDuringProcessing  = $ErrorDuringProcessingCount
            }
            Write-LogMessage "Writing interim summary to CSV..." -DataForCsv $summaryData
            $LastSummaryWriteTime = Get-Date
        }
    }
} # End of foreach loop

Write-Progress -Activity "Analyzing File Paths" -Completed -Id 1
Write-LogMessage "Path length analysis loop finished."

# --- Final Summary ---
Write-LogMessage "--- FINAL PATH LENGTH SUMMARY ---"
$finalSummaryData = [PSCustomObject]@{
    FinalSummaryTimestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    TotalFilesProcessed     = $ProcessedFilesCount
    FilesWithPathLengthGe247 = $LongPathCount
    FilesWithPathLengthLt247 = $ShortPathCount
    ErrorsDuringProcessingLoop = $ErrorDuringProcessingCount
}
Write-LogMessage "Writing final summary to CSV..." -DataForCsv $finalSummaryData

Write-LogMessage "Total files processed: $ProcessedFilesCount"
Write-LogMessage "Files with path length >= 247: $LongPathCount"
Write-LogMessage "Files with path length < 247: $ShortPathCount"
Write-LogMessage "Errors during file processing loop: $ErrorDuringProcessingCount"

Write-LogMessage "SCRIPT END: Path Length Analysis"
Stop-Transcript

Write-Host "Script finished. Logs are available in $LogDirectory"

<#
.SYNOPSIS
    Analyzes the first character of file names in a target directory, logs results, and shows progress.

.DESCRIPTION
    This script recursively scans a target folder to categorize files based on the
    first character of their names (e.g., A-Z, 0-9, symbols).
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
$TargetFolderPath = "W:\" # The folder to analyze

# Define the primary log directory (network path)
$LogDirectoryPrimary = "\\sc-file02.ad.sciencecare.com\UEMProfiles\Ryan.Zimmerman\Desktop\PS_FileAnalysis_Logs"

# Define an alternative log directory (local or mapped drive)
# <<<< IMPORTANT: ADJUST THIS ALTERNATIVE PATH IF NEEDED >>>>
$LogDirectoryAlternative = "R:\Ryan Zimmerman\PS_FileAnalysis_Logs"
# Example local path:
# $LogDirectoryAlternative = "C:\Temp\PS_FileAnalysis_Logs"

$LogDirectory = "" # This will be set to the chosen valid path

$ProgressUpdateInterval = 1000  # Update progress bar every N files
$SummaryUpdateIntervalMinutes = 5 # Write interim summary to data log every N minutes

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
$TranscriptLogFile = Join-Path -Path $LogDirectory -ChildPath "FirstCharAnalysis_Transcript_$TimestampSuffix.log"
$DataLogFile = Join-Path -Path $LogDirectory -ChildPath "FirstCharAnalysis_Data_$TimestampSuffix.csv"

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
    Write-Host $LogEntry
    if ($null -ne $DataForCsv) {
        try {
            if (-not (Test-Path $DataLogFile)) {
                # Dynamically create header for summary objects
                 ($DataForCsv | Select-Object * | Get-Member -MemberType NoteProperty).Name | Sort-Object | ForEach-Object {$_ -join ","} | Set-Content -Path $DataLogFile
            }
            $DataForCsv | Export-Csv -Path $DataLogFile -Append -NoTypeInformation -Force
        } catch { Write-Warning "Failed to write data point to $DataLogFile. Error: $($_.Exception.Message)" }
    }
}
# Function to write detailed character counts to a separate section or file if needed
Function Write-DetailedCharCountsToLog ($CharCountsHashtable) {
    Write-LogMessage "--- Detailed First Character Counts (in Transcript Log) ---"
    $CharCountsHashtable.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-LogMessage ("Character '{0}': {1} occurrences" -f $_.Name, $_.Value)
    }
    # Optionally, write this to the CSV as well, perhaps in a different format or a separate CSV
    # For simplicity, current CSV is for summary stats.
}


Write-LogMessage "SCRIPT START: First Character Analysis"
Write-LogMessage "Target Folder: $TargetFolderPath"
Write-LogMessage "Using Log Directory: $LogDirectory"
Write-LogMessage "Transcript Log: $TranscriptLogFile"
Write-LogMessage "Data Log (CSV): $DataLogFile"
Write-LogMessage "Progress Update Interval: $ProgressUpdateInterval files"
Write-LogMessage "Summary Update Interval (for CSV): $SummaryUpdateIntervalMinutes minutes"

# --- Stage 1: Get all file items and total count ---
$AllFiles = @()
$TotalFiles = 0
Write-LogMessage "Attempting to retrieve all file items and count... This may take a while."
try {
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

# --- Stage 2: Analyze First Characters ---
$FirstCharCounts = @{} # Hashtable to store counts of all first characters
$FilesStartingWithLetter = 0
$FilesNotStartingWithLetter = 0 # Includes numbers, symbols, empty names
$FilesStartingWithNumber = 0
$FilesStartingWithSymbol = 0 # Non-letter, non-number
$EmptyOrNullNameCount = 0
$ProcessedFilesCount = 0
$ErrorDuringProcessingCount = 0
$LastSummaryWriteTime = Get-Date

Write-LogMessage "Starting first character analysis loop..."

foreach ($FileItem in $AllFiles) {
    $ProcessedFilesCount++
    try {
        if ($null -ne $FileItem.Name -and $FileItem.Name.Length -gt 0) {
            $FirstChar = $FileItem.Name.Substring(0,1).ToUpper() # Case-insensitive for A-Z check
            $FirstCharCounts[$FirstChar] = $FirstCharCounts[$FirstChar] + 1

            if ($FirstChar -match "[A-Z]") {
                $FilesStartingWithLetter++
            } elseif ($FirstChar -match "[0-9]") {
                $FilesStartingWithNumber++
                $FilesNotStartingWithLetter++
            } else {
                $FilesStartingWithSymbol++ # Any other character
                $FilesNotStartingWithLetter++
            }
        } else {
            Write-Warning "File item with null or empty name encountered. FullPath: $($FileItem.FullName)"
            $FirstCharCounts["<EMPTY_OR_NULL_NAME>"] = $FirstCharCounts["<EMPTY_OR_NULL_NAME>"] + 1
            $EmptyOrNullNameCount++
            $FilesNotStartingWithLetter++ # Count it as not starting with a letter for the primary metric
        }
    } catch {
        Write-LogMessage "ERROR: Could not process file '$($FileItem.Name)' (Path: $($FileItem.DirectoryName)) - $($_.Exception.Message)"
        $ErrorDuringProcessingCount++
    }

    # Update Progress Bar
    if ($ProcessedFilesCount % $ProgressUpdateInterval -eq 0 -or $ProcessedFilesCount -eq $TotalFiles) {
        Write-Progress -Activity "Analyzing First Characters" `
                       -Status ("Processed {0} of {1} files. Letters: {2} Non-Letters: {3} Numbers: {4} Symbols: {5} Errors: {6}" -f $ProcessedFilesCount, $TotalFiles, $FilesStartingWithLetter, $FilesNotStartingWithLetter, $FilesStartingWithNumber, $FilesStartingWithSymbol, $ErrorDuringProcessingCount) `
                       -PercentComplete (($ProcessedFilesCount / $TotalFiles) * 100) `
                       -Id 2
        
        # Periodically write summary to data CSV log
        if (((Get-Date) - $LastSummaryWriteTime).TotalMinutes -ge $SummaryUpdateIntervalMinutes -or $ProcessedFilesCount -eq $TotalFiles) {
            $summaryData = [PSCustomObject]@{
                InterimSummaryTimestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                ProcessedFilesCount     = $ProcessedFilesCount
                FilesStartingWithLetter = $FilesStartingWithLetter
                FilesNotStartingWithLetter = $FilesNotStartingWithLetter
                FilesStartingWithNumber = $FilesStartingWithNumber
                FilesStartingWithSymbol = $FilesStartingWithSymbol
                EmptyOrNullNameCount    = $EmptyOrNullNameCount
                ErrorsDuringProcessing  = $ErrorDuringProcessingCount
            }
            Write-LogMessage "Writing interim summary to CSV..." -DataForCsv $summaryData
            $LastSummaryWriteTime = Get-Date
        }
    }
} # End of foreach loop

Write-Progress -Activity "Analyzing First Characters" -Completed -Id 2
Write-LogMessage "First character analysis loop finished."

# --- Final Summary ---
Write-LogMessage "--- FINAL FIRST CHARACTER SUMMARY ---"
$finalSummaryData = [PSCustomObject]@{
    FinalSummaryTimestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    TotalFilesProcessed     = $ProcessedFilesCount
    FilesStartingWithLetter = $FilesStartingWithLetter
    FilesNotStartingWithLetter = $FilesNotStartingWithLetter
    FilesStartingWithNumber = $FilesStartingWithNumber
    FilesStartingWithSymbol = $FilesStartingWithSymbol
    EmptyOrNullNameCount    = $EmptyOrNullNameCount
    ErrorsDuringProcessingLoop = $ErrorDuringProcessingCount
}
Write-LogMessage "Writing final summary to CSV..." -DataForCsv $finalSummaryData

Write-LogMessage "Total files processed: $ProcessedFilesCount"
Write-LogMessage "Files starting with a letter (A-Z): $FilesStartingWithLetter"
Write-LogMessage "Files NOT starting with a letter (incl. numbers, symbols, empty): $FilesNotStartingWithLetter"
Write-LogMessage "  Specifically, files starting with a number (0-9): $FilesStartingWithNumber"
Write-LogMessage "  Specifically, files starting with a symbol (non-letter, non-number): $FilesStartingWithSymbol"
Write-LogMessage "  Specifically, files with empty or null names: $EmptyOrNullNameCount"
Write-LogMessage "Errors during file processing loop: $ErrorDuringProcessingCount"

# Log detailed character counts to transcript
Write-DetailedCharCountsToLog -CharCountsHashtable $FirstCharCounts

Write-LogMessage "SCRIPT END: First Character Analysis"
Stop-Transcript

Write-Host "Script finished. Logs are available in $LogDirectory"

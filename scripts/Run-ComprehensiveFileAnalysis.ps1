<#
.SYNOPSIS
    Analyzes files in a target directory by processing its top-level subfolders and root files individually.
    Provides counts for total files, path lengths (short/long), and first character of filenames (alpha/non-alpha).
    This is a memory-efficient approach for large file sets. Logs results to transcript and CSV.

.DESCRIPTION
    This script lists top-level folders and processes files at the root of the target path.
    For each folder and the root, it recursively (for folders) or directly (for root) scans files,
    categorizing them by path length and the first character of their name.
    It uses a streaming method within each processed unit to maintain low memory usage.
    Results, including per-folder/root stats and grand totals, are logged.
    Includes parameterization, advanced logging (primary, alternative, fallback), and progress reporting.

.PARAMETER TargetFolderPath
    The root folder to analyze. Defaults to "W:\".

.PARAMETER LogDirectoryPrimary
    The primary network path for log files.
    Defaults to "\\sc-file02.ad.sciencecare.com\UEMProfiles\Ryan.Zimmerman\Desktop\PS_FileAnalysis_Logs".

.PARAMETER LogDirectoryAlternative
    An alternative path for log files (e.g., mapped drive or local path) if the primary is unavailable.
    Defaults to "R:\Ryan Zimmerman\PS_FileAnalysis_Logs".

.NOTES
    Version: 3.0
    Author: AI Assistant & User Collaboration
    Last Modified: 2025-05-06

    Ensure the account running this script has read access to the target folder
    and write access to the chosen log directory.
#>
param(
    [string]$TargetFolderPath = "W:\",
    [string]$LogDirectoryPrimary = "\\sc-file02.ad.sciencecare.com\UEMProfiles\Ryan.Zimmerman\Desktop\PS_FileAnalysis_Logs",
    [string]$LogDirectoryAlternative = "R:\Ryan Zimmerman\PS_FileAnalysis_Logs" # <<<< ADJUST IF NEEDED
)

# --- Global Stopwatch ---
$GlobalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# --- Configuration ---
$LogDirectory = "" # This will be set to the chosen valid path
$FallbackLogDirName = "PS_FileAnalysis_Logs_Fallback"

# --- Function to Select and Create Log Directory ---
Function Set-LogDirectory {
    param(
        [string]$PrimaryPath,
        [string]$AlternativePath,
        [string]$TempFallbackParentPath, # e.g., $env:TEMP
        [string]$TempFallbackFolderName
    )
    Write-Host "Attempting to set up log directory..."
    # Try Primary Path
    if (Test-Path -Path $PrimaryPath -PathType Container) {
        Write-Host "Primary log directory found: $PrimaryPath"
        return $PrimaryPath
    } elseif (Test-Path -Path (Split-Path -Path $PrimaryPath -Parent) -PathType Container) {
        Write-Host "Parent of primary log directory found. Attempting to create log directory: $PrimaryPath"
        try { New-Item -ItemType Directory -Path $PrimaryPath -ErrorAction Stop -Force | Out-Null; Write-Host "Successfully created primary log directory: $PrimaryPath"; return $PrimaryPath } catch { Write-Warning "Could not create primary log directory '$PrimaryPath'. Error: $($_.Exception.Message)" }
    } else { Write-Warning "Primary log directory or its parent not accessible: $PrimaryPath" }

    Write-Warning "Attempting to use alternative log directory..."
    # Try Alternative Path
    if (Test-Path -Path $AlternativePath -PathType Container) {
        Write-Host "Alternative log directory found: $AlternativePath"
        return $AlternativePath
    } elseif (Test-Path -Path (Split-Path -Path $AlternativePath -Parent) -PathType Container) {
        Write-Host "Parent of alternative log directory found. Attempting to create log directory: $AlternativePath"
        try { New-Item -ItemType Directory -Path $AlternativePath -ErrorAction Stop -Force | Out-Null; Write-Host "Successfully created alternative log directory: $AlternativePath"; return $AlternativePath } catch { Write-Warning "Could not create alternative log directory '$AlternativePath'. Error: $($_.Exception.Message)" }
    } else { Write-Warning "Alternative log directory or its parent not accessible: $AlternativePath" }

    Write-Warning "Attempting to use temporary fallback log directory..."
    # Try Temp Fallback Path
    $TempLogPath = Join-Path -Path $TempFallbackParentPath -ChildPath $TempFallbackFolderName
    if (Test-Path -Path $TempLogPath -PathType Container) {
        Write-Host "Temporary fallback log directory found: $TempLogPath"
        return $TempLogPath
    } else {
        Write-Host "Attempting to create temporary fallback log directory: $TempLogPath"
        try { New-Item -ItemType Directory -Path $TempLogPath -ErrorAction Stop -Force | Out-Null; Write-Host "Successfully created temporary fallback log directory: $TempLogPath"; return $TempLogPath } catch { Write-Error "Could not create temporary fallback log directory '$TempLogPath'. Error: $($_.Exception.Message)" }
    }
    return $null
}

# --- Set the Log Directory ---
$LogDirectory = Set-LogDirectory -PrimaryPath $LogDirectoryPrimary -AlternativePath $LogDirectoryAlternative -TempFallbackParentPath $env:TEMP -TempFallbackFolderName $FallbackLogDirName
if (-not $LogDirectory) {
    Write-Error "FATAL: No valid log directory could be established. Please check paths and permissions. Exiting."
    $GlobalStopwatch.Stop()
    exit 1
}

# --- Define Log File Paths ---
$TimestampSuffix = Get-Date -Format 'yyyyMMdd_HHmmss'
$ScriptName = $MyInvocation.MyCommand.Name -replace ".ps1", ""
$TranscriptLogFile = Join-Path -Path $LogDirectory -ChildPath "${ScriptName}_Transcript_$TimestampSuffix.log"
$DataLogFile = Join-Path -Path $LogDirectory -ChildPath "${ScriptName}_Data_$TimestampSuffix.csv"

# --- Initialize Logging ---
try {
    Start-Transcript -Path $TranscriptLogFile -Append -ErrorAction Stop
} catch {
    Write-Error "FATAL: Could not start transcript at '$TranscriptLogFile'. Error: $($_.Exception.Message). Exiting."
    $GlobalStopwatch.Stop()
    exit 1
}

Function Write-LogMessage ($Message, [OutputType([System.Management.Automation.PSCustomObject])] $DataForCsv = $null) {
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] $Message"
    Write-Host $LogEntry

    if ($null -ne $DataForCsv) {
        try {
            if (-not (Test-Path $DataLogFile)) {
                # Dynamically create header from the PSCustomObject properties
                $DataForCsv | Select-Object * | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name | Sort-Object | ForEach-Object {$_ -join ","} | Set-Content -Path $DataLogFile
            }
            $DataForCsv | Export-Csv -Path $DataLogFile -Append -NoTypeInformation -Force
        } catch { Write-Warning "Failed to write data point to $DataLogFile. Error: $($_.Exception.Message)" }
    }
}

# --- Core Folder/File Processing Function ---
Function Invoke-DetailedFileAnalysis {
    param(
        [string]$PathToProcess,
        [bool]$RecurseScan = $true # True for folders, false for root files
    )
    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Initialize counters
    $countShortPath      = 0
    $countLongPath       = 0
    $countAlphaStart     = 0
    $countNonAlphaStart  = 0
    $countTotalFiles     = 0
    $errorMessage        = $null

    try {
        Write-LogMessage "Analyzing path: $PathToProcess (Recursive: $RecurseScan)"
        $childItemParams = @{
            Path = $PathToProcess
            File = $true
            ErrorAction = 'SilentlyContinue'
        }
        if ($RecurseScan) {
            $childItemParams.Recurse = $true
        }

        Get-ChildItem @childItemParams | ForEach-Object {
            $countTotalFiles++
            try {
                # Path Length Analysis
                if ($_.FullName.Length -ge 247) { $countLongPath++ } else { $countShortPath++ }

                # First Character Analysis
                if ($_.Name -and $_.Name.Length -gt 0) {
                    $firstChar = $_.Name.Substring(0,1).ToUpper()
                    if ($firstChar -ge 'A' -and $firstChar -le 'Z') { $countAlphaStart++ } else { $countNonAlphaStart++ }
                } else {
                    # Consider files with no name or empty name as NonAlphaStart
                    $countNonAlphaStart++
                    Write-Warning "File with no name or empty name encountered: $($_.FullName)"
                }
            } catch {
                Write-Warning "Error processing individual file '$($_.FullName)': $($_.Exception.Message)"
                # Decide how to count this error: perhaps an 'erroredFileInLoop' counter
            }
        }
    } catch {
        # This catch is for errors in Get-ChildItem itself (e.g., path not found, major access denied)
        $errorMessage = "ERROR during Get-ChildItem for '$PathToProcess'. Details: $($_.Exception.Message)"
        Write-LogMessage $errorMessage
    }

    $Stopwatch.Stop()

    return [pscustomobject]@{
        ProcessedPath       = $PathToProcess
        ShortPathFiles      = $countShortPath
        LongPathFiles       = $countLongPath
        AlphaStartFiles     = $countAlphaStart
        NonAlphaStartFiles  = $countNonAlphaStart
        TotalFilesAnalyzed  = $countTotalFiles # This should equal ShortPathFiles + LongPathFiles
        ProcessingSeconds   = [math]::Round($Stopwatch.Elapsed.TotalSeconds,2)
        Error               = $errorMessage
    }
}


# --- Main Script Logic ---
Write-LogMessage "SCRIPT START: Comprehensive Snapshot File Analysis (Optimized)"
Write-LogMessage "Script Name: $ScriptName"
Write-LogMessage "Target Root Folder: $TargetFolderPath"
Write-LogMessage "Using Log Directory: $LogDirectory"
Write-LogMessage "Transcript Log: $TranscriptLogFile"
Write-LogMessage "Data Log (CSV for per-item stats): $DataLogFile"

# Initialize Grand Totals
$GrandTotalFilesAnalyzed = 0
$GrandTotalShortPathFiles = 0
$GrandTotalLongPathFiles = 0
$GrandTotalAlphaStartFiles = 0
$GrandTotalNonAlphaStartFiles = 0
$TotalProcessingTimeOverallSeconds = 0

# --- Analyze Files Directly at the Root ---
Write-LogMessage "--- Analyzing files directly at root: $TargetFolderPath ---"
$RootStats = Invoke-DetailedFileAnalysis -PathToProcess $TargetFolderPath -RecurseScan $false # Process only files directly in root

$GrandTotalFilesAnalyzed += $RootStats.TotalFilesAnalyzed
$GrandTotalShortPathFiles += $RootStats.ShortPathFiles
$GrandTotalLongPathFiles += $RootStats.LongPathFiles
$GrandTotalAlphaStartFiles += $RootStats.AlphaStartFiles
$GrandTotalNonAlphaStartFiles += $RootStats.NonAlphaStartFiles
$TotalProcessingTimeOverallSeconds += $RootStats.ProcessingSeconds

$rootData = [PSCustomObject]@{
    Timestamp               = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    ItemType                = "RootFiles"
    Path                    = $RootStats.ProcessedPath
    TotalFiles              = $RootStats.TotalFilesAnalyzed
    ShortPathFiles          = $RootStats.ShortPathFiles
    LongPathFiles           = $RootStats.LongPathFiles
    AlphaStartFiles         = $RootStats.AlphaStartFiles
    NonAlphaStartFiles      = $RootStats.NonAlphaStartFiles
    ProcessingTimeSeconds   = $RootStats.ProcessingSeconds
    ErrorMessage            = $RootStats.Error
}
Write-LogMessage "Logging root file analysis to CSV..." -DataForCsv $rootData

# --- Get Top-Level Folders ---
$TopLevelFolders = @()
$TotalFoldersFound = 0
Write-LogMessage "--- Processing Top-Level Folders under: $TargetFolderPath ---"
Write-LogMessage "Attempting to retrieve top-level folders..."
try {
    $TopLevelFolders = Get-ChildItem -Path $TargetFolderPath -Directory -ErrorAction SilentlyContinue
    $TotalFoldersFound = $TopLevelFolders.Count
    if ($TotalFoldersFound -eq 0) {
        Write-LogMessage "No top-level folders found directly under '$TargetFolderPath'."
    } else {
        Write-LogMessage ("Found {0} top-level folders to process." -f $TotalFoldersFound)
    }
} catch {
    Write-LogMessage "ERROR: Could not retrieve top-level folders. $($_.Exception.Message)"
}

# --- Process Each Top-Level Folder ---
$FoldersProcessedCount = 0
if ($TotalFoldersFound -gt 0) {
    Write-LogMessage "Starting to process each top-level folder..."
    foreach ($Folder in $TopLevelFolders) {
        $FoldersProcessedCount++
        $CurrentFolderName = $Folder.FullName
        
        Write-Progress -Activity "Processing Top-Level Folders" `
                       -Status ("Folder {0} of {1}: {2}" -f $FoldersProcessedCount, $TotalFoldersFound, $Folder.Name) `
                       -PercentComplete (($FoldersProcessedCount / $TotalFoldersFound) * 100) `
                       -Id 1

        $FolderStats = Invoke-DetailedFileAnalysis -PathToProcess $CurrentFolderName -RecurseScan $true
        
        $GrandTotalFilesAnalyzed += $FolderStats.TotalFilesAnalyzed
        $GrandTotalShortPathFiles += $FolderStats.ShortPathFiles
        $GrandTotalLongPathFiles += $FolderStats.LongPathFiles
        $GrandTotalAlphaStartFiles += $FolderStats.AlphaStartFiles
        $GrandTotalNonAlphaStartFiles += $FolderStats.NonAlphaStartFiles
        $TotalProcessingTimeOverallSeconds += $FolderStats.ProcessingSeconds
        
        $folderData = [PSCustomObject]@{
            Timestamp               = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            ItemType                = "TopLevelFolder"
            Path                    = $FolderStats.ProcessedPath
            TotalFiles              = $FolderStats.TotalFilesAnalyzed
            ShortPathFiles          = $FolderStats.ShortPathFiles
            LongPathFiles           = $FolderStats.LongPathFiles
            AlphaStartFiles         = $FolderStats.AlphaStartFiles
            NonAlphaStartFiles      = $FolderStats.NonAlphaStartFiles
            ProcessingTimeSeconds   = $FolderStats.ProcessingSeconds
            ErrorMessage            = $FolderStats.Error
        }
        Write-LogMessage "Logging data for $CurrentFolderName to CSV..." -DataForCsv $folderData
    }
    Write-Progress -Activity "Processing Top-Level Folders" -Completed -Id 1
} else {
    Write-LogMessage "Skipping top-level folder processing as none were found or an error occurred retrieving them."
}

# --- Final Summary ---
$GlobalStopwatch.Stop()
$TotalScriptExecutionTimeSeconds = [math]::Round($GlobalStopwatch.Elapsed.TotalSeconds, 2)

Write-LogMessage "--- FINAL SUMMARY ---"
$finalSummaryData = [PSCustomObject]@{
    FinalSummaryTimestamp           = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    TargetRootFolder                = $TargetFolderPath
    TotalFilesAnalyzed_GrandTotal   = $GrandTotalFilesAnalyzed
    ShortPathFiles_GrandTotal       = $GrandTotalShortPathFiles
    LongPathFiles_GrandTotal        = $GrandTotalLongPathFiles    # Files PQ would filter by path length
    AlphaStartFiles_GrandTotal      = $GrandTotalAlphaStartFiles
    NonAlphaStartFiles_GrandTotal   = $GrandTotalNonAlphaStartFiles # Files PQ would filter by non-alpha start
    TotalTopLevelFoldersFound       = $TotalFoldersFound
    TotalTopLevelFoldersProcessed   = $FoldersProcessedCount
    TotalProcessingTimeOverallSeconds = [math]::Round($TotalProcessingTimeOverallSeconds, 2) # Time spent in Invoke-DetailedFileAnalysis
    TotalScriptExecutionTimeSeconds = $TotalScriptExecutionTimeSeconds
}
Write-LogMessage "Writing final summary to CSV..." -DataForCsv $finalSummaryData

Write-LogMessage ("Target Root Folder: {0}" -f $TargetFolderPath)
Write-LogMessage ("--- Grand Totals ---")
Write-LogMessage ("Total Files Analyzed (Root + Folders): {0}" -f $GrandTotalFilesAnalyzed)
Write-LogMessage ("  Files with Short Paths (<247 chars): {0}" -f $GrandTotalShortPathFiles)
Write-LogMessage ("  Files with Long Paths (>=247 chars): {0}  <-- POTENTIALLY FILTERED BY POWER QUERY" -f $GrandTotalLongPathFiles)
Write-LogMessage ("  Files Starting with a Letter (A-Z): {0}" -f $GrandTotalAlphaStartFiles)
Write-LogMessage ("  Files NOT Starting with a Letter: {0}  <-- POTENTIALLY FILTERED BY POWER QUERY (A-F, G-L, M-Z bands)" -f $GrandTotalNonAlphaStartFiles)
Write-LogMessage ("Total Top-Level Folders Found: {0}" -f $TotalFoldersFound)
Write-LogMessage ("Total Top-Level Folders Processed: {0}" -f $FoldersProcessedCount)
Write-LogMessage ("Total time spent in detailed file analysis loops: {0:N2} seconds" -f $TotalProcessingTimeOverallSeconds)
Write-LogMessage ("Total script execution time: {0:N2} seconds" -f $TotalScriptExecutionTimeSeconds)

Write-LogMessage "SCRIPT END: Comprehensive Snapshot File Analysis (Optimized)"
Stop-Transcript

Write-Host "Script finished. Logs are available in '$LogDirectory'. Total execution time: $($TotalScriptExecutionTimeSeconds)s"

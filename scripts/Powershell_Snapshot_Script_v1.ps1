<#
.SYNOPSIS
    Generates a daily snapshot CSV of all files within specified order management folders.
    Streams file processing with buffered writes for memory and I/O efficiency,
    and uses isolated regex matches for reliability.
    The CSV contains file metadata and derived identifiers (QuoteNumber, FormNumber, FormType)
    in a format directly consumable by Power Query for efficient data loading and analysis.

.DESCRIPTION
    This script iterates through a predefined list of source folders using UNC paths.
    For each file found, it extracts:
        - QuoteNumber (SC/BS numbers from filename)
        - FormNumber (TRF/ERF numbers from filename)
        - FormType (TRF/ERF from filename)
        - FolderLocation (derived from the source folder's defined status)
        - FileName
        - FileExtension
        - FileDateModified (UTC ISO 8601 format)
        - FileDateCreated (UTC ISO 8601 format)

    The script outputs a single CSV file, overwriting the previous day's snapshot.
    It uses an atomic write pattern (writing to a temporary file then renaming) to ensure
    data integrity. The CSV is UTF-8 encoded and uses standard quoting (conditionally for PS5.1).
    A header row is pre-written to the temporary file.
    Includes concurrent run protection and buffered CSV writes.

.NOTES
    Version: 1.7
    Author: AI Assistant (In collaboration with User)
    Last Modified: 2025-05-08

    - Ensure the account running this script has read access to all $OrderRoots paths via UNC.
    - Ensure the account running this script has write/modify access to the $SnapshotDir.
    - Schedule this script to run daily (e.g., via Windows Task Scheduler).
    - Output CSV column order and names are critical for Power Query compatibility.
    - Paths use UNC for reliability.
    - Export-Csv uses a dynamic hashtable for -UseQuotes parameter for PS5.1 & PS7+ compatibility.
    - Corrected "Closed" folder name to "5. Closed Files".
    - Refined hashtable cloning for Export-Csv parameters.
#>

#Requires -Version 5.1

[CmdletBinding()]
param()

# --- Script Start Time ---
$ScriptStartTime = Get-Date
Write-Host "Script started at: $ScriptStartTime"

# --- Configuration ---
# Define the base UNC path for order management folders
$OrderBaseUNC = '\\scfiles\SCFILES\Client Services\Order Management' # VERIFY THIS UNC PATH IS 100% ACCURATE

# Define the specific subfolders for order statuses and their corresponding names.
# The 'Status' value will be used for the 'FolderLocation' column in the CSV.
$OrderRoots = @(
    @{ Status='1. New Orders';          Path= (Join-Path -Path $OrderBaseUNC -ChildPath '1. New Orders') },
    @{ Status='2. Open Orders';         Path= (Join-Path -Path $OrderBaseUNC -ChildPath '2. Open Orders') },
    @{ Status='3. As Available Orders'; Path= (Join-Path -Path $OrderBaseUNC -ChildPath '3. As Available Orders') },
    @{ Status='4. Hold Orders';         Path= (Join-Path -Path $OrderBaseUNC -ChildPath '4. Hold Orders') },
    @{ Status='5. Closed Files';        Path= (Join-Path -Path $OrderBaseUNC -ChildPath '5. Closed Files') }, # Corrected name
    @{ Status='6. Declined Orders';     Path= (Join-Path -Path $OrderBaseUNC -ChildPath '6. Declined Orders') }
    # Add or remove folders as needed. Ensure the ChildPath values are 100% correct.
)

# Define the output directory and filename for the snapshot CSV.
$SnapshotDir  = 'C:\Data\OrderCatalogue' # Example: Can be a local path on a server or a shared path.
$SnapshotFile = Join-Path -Path $SnapshotDir -ChildPath 'OrderCatalogue_LATEST.csv'
$TemporarySnapshotFile = $SnapshotFile + '.tmp'
$WriteBufferSize = 1000 # Number of records to buffer before writing to CSV

# --- Preparation ---
# Ensure the snapshot directory exists.
if (-not (Test-Path -Path $SnapshotDir -PathType Container)) {
    try {
        Write-Host "Snapshot directory '$SnapshotDir' does not exist. Attempting to create it..."
        New-Item -Path $SnapshotDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Host "Successfully created snapshot directory: $SnapshotDir"
    } catch {
        Write-Error "FATAL: Could not create snapshot directory '$SnapshotDir'. Error: $($_.Exception.Message). Please create it manually and ensure permissions. Exiting."
        exit 1
    }
}

# Concurrent Run Protection: Check if the temporary file exists from a potentially ongoing run.
if (Test-Path -Path $TemporarySnapshotFile) {
    Write-Error "FATAL: Temporary snapshot file '$TemporarySnapshotFile' already exists. Another instance of the script might be running or a previous run failed to clean up. Please investigate. Exiting."
    exit 1
}

# Pre-write the header row to the temporary file.
try {
    '"QuoteNumber","FormNumber","FormType","FolderLocation","FileName","FileExtension","FileDateModified","FileDateCreated"' |
        Set-Content -Path $TemporarySnapshotFile -Encoding UTF8 -ErrorAction Stop
    Write-Host "Header row written to temporary file: $TemporarySnapshotFile"
} catch {
    Write-Error "FATAL: Could not write header row to temporary file '$TemporarySnapshotFile'. Error: $($_.Exception.Message). Exiting."
    exit 1
}

# --- Main Processing Loop with Buffered Writes ---
Write-Host "Starting file enumeration and data extraction (buffered streaming to CSV)..."
$TotalFilesProcessed = 0
$OverallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$Buffer = New-Object -TypeName "System.Collections.Generic.List[psobject]" -ArgumentList $WriteBufferSize
$LastFileProcessedInLoop = $null 

# Prepare Export-Csv parameters dynamically for compatibility
$BaseCsvParams = [ordered]@{ # Ensure base is ordered if downstream processes might care, though for Export-Csv it's less critical
    Append            = $true
    NoTypeInformation = $true
    Encoding          = 'UTF8' 
    ErrorAction       = 'Stop' 
}
if ((Get-Command Export-Csv).Parameters.ContainsKey('UseQuotes')) {
    $BaseCsvParams.UseQuotes = 'AsNeeded'
}

foreach ($RootFolderInfo in $OrderRoots) {
    Write-Host "Processing folder: '$($RootFolderInfo.Path)' for status: '$($RootFolderInfo.Status)'"
    if (-not (Test-Path -Path $RootFolderInfo.Path -PathType Container)) {
        Write-Warning "WARNING: Source folder path not found or not a directory: '$($RootFolderInfo.Path)'. Skipping."
        continue
    }

    try {
        Get-ChildItem -Path $RootFolderInfo.Path -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
            $TotalFilesProcessed++
            $LastFileProcessedInLoop = $_.FullName 
            $FileRecord = [pscustomobject]@{
                QuoteNumber      = ([regex]::Match($_.Name, '((?:SC|BS)\d{3,})')).Groups[1].Value
                FormNumber       = ([regex]::Match($_.Name, '(?:TRF|ERF)[ _\-#]*?(\d{6,10})')).Groups[1].Value
                FormType         = ([regex]::Match($_.Name, '(TRF|ERF)')).Groups[1].Value
                FolderLocation   = $RootFolderInfo.Status
                FileName         = $_.Name
                FileExtension    = $_.Extension
                FileDateModified = $_.LastWriteTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
                FileDateCreated  = $_.CreationTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
            }
            $Buffer.Add($FileRecord)

            if ($Buffer.Count -ge $WriteBufferSize) {
                try {
                    # Correctly create a new hashtable for current parameters and add Path
                    $CurrentCsvParams = [ordered]@{} + $BaseCsvParams
                    $CurrentCsvParams.Path = $TemporarySnapshotFile
                    $Buffer | Export-Csv @CurrentCsvParams
                } catch {
                    Write-Warning "ERROR writing buffer to CSV for files in '$($RootFolderInfo.Path)'. Last file processed before potential write error: '$LastFileProcessedInLoop'. Details: $($_.Exception.Message). Some records may not have been written."
                }
                $Buffer.Clear()
            }
        }
    } catch { 
        Write-Warning "ERROR during Get-ChildItem for folder '$($RootFolderInfo.Path)'. Details: $($_.Exception.Message)."
    }
}

# Flush any remaining records in the buffer
if ($Buffer.Count -gt 0) {
    Write-Host "Flushing remaining $($Buffer.Count) records from buffer..."
    try {
        # Correctly create a new hashtable for current parameters and add Path
        $CurrentCsvParams = [ordered]@{} + $BaseCsvParams
        $CurrentCsvParams.Path = $TemporarySnapshotFile
        $Buffer | Export-Csv @CurrentCsvParams
    } catch {
         Write-Warning "ERROR writing final buffer to CSV. Last file processed overall: '$LastFileProcessedInLoop'. Details: $($_.Exception.Message). Some records may not have been written."
    }
    $Buffer.Clear()
}

# --- Atomic Switch-Over ---
Write-Host "Attempting to finalize snapshot file..."
try {
    if (Test-Path -Path $TemporarySnapshotFile) {
        Move-Item -LiteralPath $TemporarySnapshotFile -Destination $SnapshotFile -Force -ErrorAction Stop
        Write-Host "Successfully created/updated snapshot file: $SnapshotFile"
        
        $FinalFileSize = (Get-Item -Path $SnapshotFile).Length
        Write-Host "Final file size: $([math]::Round($FinalFileSize / 1MB, 2)) MB"

        if ($FinalFileSize -lt 200) { 
            Write-Warning "Snapshot file is very small (size: $FinalFileSize bytes). It might be header-only or contain very few records, indicating few or no files were processed."
        }
    } else {
        Write-Error "CRITICAL: Temporary snapshot file '$TemporarySnapshotFile' was not found. The final snapshot '$SnapshotFile' has NOT been updated. This indicates a problem earlier in the script (e.g., header write failure or no data processed and buffer never flushed). Exiting."
        exit 1
    }
} catch {
    Write-Error "FATAL: Could not move temporary file '$TemporarySnapshotFile' to '$SnapshotFile'. Error: $($_.Exception.Message). Manual cleanup may be required. Exiting."
    exit 1
}

# --- Script End ---
$OverallStopwatch.Stop()
$ScriptEndTime = Get-Date
Write-Host "Script finished at: $ScriptEndTime"
Write-Host "Total file objects processed: $TotalFilesProcessed"
Write-Host "Total script execution time: $($OverallStopwatch.Elapsed.ToString())"

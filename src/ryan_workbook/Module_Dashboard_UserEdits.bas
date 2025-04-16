Attribute VB_Name = "Module_Dashboard_UserEdits"
Option Explicit

'===============================================================================
' MODULE_DASHBOARD_USEREDITS
' Contains functions related to managing the UserEdits sheet, logging,
' and backups for the SQRCT Dashboard.
' Requires constants defined in Module_Dashboard_Core.
'===============================================================================

'===============================================================================
' USER EDITS LOGGING: Tracks operations and errors
'===============================================================================
Public Sub LogUserEditsOperation(message As String)
    Dim wsLog As Worksheet
    Dim lastRow As Long
    Const MAX_LOG_ROWS As Long = 5000 ' Optional: Limit log size

    On Error Resume Next ' Use minimal error handling within logging itself
    Set wsLog = ThisWorkbook.Sheets(Module_Dashboard_Core.USEREDITSLOG_SHEET_NAME) ' Use constant from Core module

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = Module_Dashboard_Core.USEREDITSLOG_SHEET_NAME
        wsLog.Range("A1:C1").Value = Array("Timestamp", "Workbook", "Operation")
        wsLog.Range("A1:C1").Font.Bold = True
        wsLog.Visible = xlSheetHidden ' Keep hidden unless debugging
        wsLog.Columns("A").ColumnWidth = 20
        wsLog.Columns("B").ColumnWidth = 15
        wsLog.Columns("C").ColumnWidth = 100
        wsLog.Columns("C").WrapText = False
    End If
    Err.Clear ' Clear any error from sheet check/creation

    Dim identity As String
    On Error Resume Next
    identity = Module_Identity.GetWorkbookIdentity() ' Assumes Module_Identity exists
    If Err.Number <> 0 Then identity = "Unknown (Module_Identity Error)"
    Err.Clear
    On Error GoTo 0 ' Restore default error handling

    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    ' Optional: Trim log if it exceeds max rows
    If lastRow >= MAX_LOG_ROWS + 1 Then ' Check if adding one more row exceeds limit
        wsLog.Rows("2:" & (lastRow - MAX_LOG_ROWS + 2)).Delete ' Delete oldest rows
        lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row ' Recalculate last row
        If lastRow < 1 Then lastRow = 1
    End If

    wsLog.Cells(lastRow + 1, "A").Value = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    wsLog.Cells(lastRow + 1, "B").Value = identity
    wsLog.Cells(lastRow + 1, "C").Value = message

ErrorHandler: ' Label for potential future GoTo, though not used currently
    If Err.Number <> 0 Then Debug.Print "Error within LogUserEditsOperation: " & Err.Description
End Sub

'===============================================================================
' USER EDITS BACKUP: Creates a timestamped backup of the UserEdits sheet
'===============================================================================
Public Function CreateUserEditsBackup(Optional backupSuffix As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim wsEdits As Worksheet
    Dim wsBackup As Worksheet
    Dim backupName As String

    ' Set backup name using Core constant
    If backupSuffix = "" Then backupName = Module_Dashboard_Core.USEREDITS_SHEET_NAME & "_Backup_" & Format(Now, "yyyymmdd")
    Else: backupName = Module_Dashboard_Core.USEREDITS_SHEET_NAME & "_Backup_" & backupSuffix
    End If
    backupName = Left(backupName, 31) ' Ensure name length <= 31 chars

    ' Check if UserEdits sheet exists
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets(Module_Dashboard_Core.USEREDITS_SHEET_NAME)
    On Error GoTo ErrorHandler
    If wsEdits Is Nothing Then LogUserEditsOperation "Backup skipped: UserEdits sheet not found.": CreateUserEditsBackup = False: Exit Function

    ' Delete existing backup sheet with the same name if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(backupName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler ' Restore error handling after potential delete

    ' Create new backup sheet
    Set wsBackup = ThisWorkbook.Sheets.Add(After:=wsEdits) ' Add after UserEdits
    wsBackup.Name = backupName
    If Err.Number <> 0 Then GoTo ErrorHandler ' Error if name is invalid/duplicate after delete failed

    ' Copy data
    wsEdits.UsedRange.Copy wsBackup.Range("A1")
    If Err.Number <> 0 Then GoTo ErrorHandler

    ' Hide backup sheet
    wsBackup.Visible = xlSheetHidden
    LogUserEditsOperation "Created UserEdits backup: " & backupName
    CreateUserEditsBackup = True
    Exit Function

ErrorHandler:
    Debug.Print "Error creating backup '" & backupName & "': [" & Err.Number & "] " & Err.Description
    LogUserEditsOperation "ERROR creating UserEdits backup '" & backupName & "': " & Err.Description
    ' Attempt to clean up partially created backup sheet on error
    On Error Resume Next: Application.DisplayAlerts = False: If Not wsBackup Is Nothing Then wsBackup.Delete: Application.DisplayAlerts = True: On Error GoTo 0
    CreateUserEditsBackup = False
End Function

'===============================================================================
' USER EDITS RESTORE: Restores UserEdits data from the most recent backup
'===============================================================================
Public Function RestoreUserEditsFromBackup(Optional specificBackupName As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim wsEdits As Worksheet
    Dim wsBackup As Worksheet
    Dim backupNameToRestore As String: backupNameToRestore = ""
    Dim baseBackupName As String: baseBackupName = Module_Dashboard_Core.USEREDITS_SHEET_NAME & "_Backup_" ' Use Core constant

    ' 1. Prioritize specific backup name if provided
    If specificBackupName <> "" Then
        On Error Resume Next: Set wsBackup = ThisWorkbook.Sheets(specificBackupName): On Error GoTo ErrorHandler
        If Not wsBackup Is Nothing Then backupNameToRestore = specificBackupName Else LogUserEditsOperation "Specified backup '" & specificBackupName & "' not found."
    End If

    ' 2. Find most recent backup if specific not found/provided
    If backupNameToRestore = "" Then
        Dim i As Integer, mostRecentDate As Date, tempName As String
        mostRecentDate = DateSerial(1900, 1, 1) ' Initialize to a very old date
        For i = 1 To ThisWorkbook.Sheets.Count
            tempName = ThisWorkbook.Sheets(i).Name
            If InStr(1, tempName, baseBackupName, vbTextCompare) = 1 Then ' Check if name starts with base name
                Dim datePart As String: datePart = Mid$(tempName, Len(baseBackupName) + 1)
                Dim suffixDateTime As Date: On Error Resume Next ' Attempt to parse date/time from suffix
                If InStr(datePart, "_") > 0 And Len(datePart) >= 15 Then ' Format yyyymmdd_hhmmss
                    suffixDateTime = CDate(Format(Left$(datePart, 8), "0000-00-00") & " " & Format(Mid$(datePart, 10, 6), "00:00:00"))
                ElseIf Len(datePart) = 8 Then ' Format yyyymmdd
                    suffixDateTime = CDate(Format(datePart, "0000-00-00"))
                Else ' Unrecognized format
                    suffixDateTime = DateSerial(1900, 1, 1)
                End If
                If Err.Number = 0 Then ' If date parsing succeeded
                    If suffixDateTime >= mostRecentDate Then ' Check if this backup is newer
                        mostRecentDate = suffixDateTime
                        backupNameToRestore = tempName
                    End If
                End If
                Err.Clear: On Error GoTo ErrorHandler ' Clear parsing error and restore main handler
            End If
        Next i
    End If

    ' 3. Check if a backup was found
    If backupNameToRestore = "" Then LogUserEditsOperation "No suitable backup found to restore.": RestoreUserEditsFromBackup = False: Exit Function

    ' 4. Get the backup sheet object
    On Error Resume Next: Set wsBackup = ThisWorkbook.Sheets(backupNameToRestore): On Error GoTo ErrorHandler
    If wsBackup Is Nothing Then LogUserEditsOperation "Backup sheet '" & backupNameToRestore & "' could not be accessed.": RestoreUserEditsFromBackup = False: Exit Function

    ' 5. Ensure target UserEdits sheet exists (call Setup)
    SetupUserEditsSheet
    Set wsEdits = ThisWorkbook.Sheets(Module_Dashboard_Core.USEREDITS_SHEET_NAME) ' Use Core constant
    If wsEdits Is Nothing Then LogUserEditsOperation "UserEdits sheet missing for restore.": RestoreUserEditsFromBackup = False: Exit Function

    ' 6. Unprotect target sheet and clear/copy data
    On Error Resume Next: wsEdits.Unprotect: On Error GoTo ErrorHandler ' Unprotect target

    wsEdits.Cells.Clear: If Err.Number <> 0 Then GoTo ErrorHandler
    wsBackup.UsedRange.Copy wsEdits.Range("A1"): If Err.Number <> 0 Then GoTo ErrorHandler

    ' 7. Log success and exit
    LogUserEditsOperation "Restored UserEdits from backup: " & backupNameToRestore
    RestoreUserEditsFromBackup = True
    Exit Function

ErrorHandler:
    Debug.Print "Error restoring from backup: [" & Err.Number & "] " & Err.Description
    LogUserEditsOperation "ERROR restoring UserEdits from backup: " & Err.Description
    RestoreUserEditsFromBackup = False
End Function

'===============================================================================
' SAVEUSEREDITSFROMDASHBOARD: Captures edits from Dashboard (L:N) -> UserEdits sheet
' Uses Dictionary lookup for performance. Reads/Writes using arrays.
'===============================================================================
Public Sub SaveUserEditsFromDashboard()
    Dim wsSrc As Worksheet, wsEdits As Worksheet, lastRowSrc As Long, lastRowEdits As Long
    Dim i As Long, destRow As Long, docNum As String, hasUserEdits As Boolean, wasChanged As Boolean
    Dim userEditsDict As Object, editRow As Variant, dashboardDataArray As Variant
    Dim dashboardRange As Range, dashRows As Long, dashCols As Long
    Dim userEditsChanges As Collection, newItemData As Variant
    Dim workbookID As String, timeStampStr As String

    LogUserEditsOperation "Starting SaveUserEditsFromDashboard"
    On Error GoTo ErrorHandler

    ' Get sheets
    Set wsSrc = Module_Dashboard_Core.GetOrCreateDashboardSheet(Module_Dashboard_Core.DASHBOARD_SHEET_NAME) ' Use Public constant from Core
    If wsSrc Is Nothing Then LogUserEditsOperation Module_Dashboard_Core.DASHBOARD_SHEET_NAME & " sheet not found.": Exit Sub
    SetupUserEditsSheet ' Ensure UserEdits exists
    Set wsEdits = ThisWorkbook.Sheets(Module_Dashboard_Core.USEREDITS_SHEET_NAME) ' Use Public constant from Core
    If wsEdits Is Nothing Then LogUserEditsOperation "UserEdits sheet not found.": Exit Sub

    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, Module_Dashboard_Core.DB_COL_DOCNUM).End(xlUp).Row ' Use Core constant
    If lastRowSrc < 4 Then LogUserEditsOperation "No data rows on dashboard.": Exit Sub

    ' Unprotect sheets
    On Error Resume Next: wsSrc.Unprotect: wsEdits.Unprotect: On Error GoTo ErrorHandler

    ' Load dictionary {DocNum: RowNum} from UserEdits
    Set userEditsDict = LoadUserEditsToDictionary(wsEdits) ' Call Private helper in this module
    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, Module_Dashboard_Core.UE_COL_DOCNUM).End(xlUp).Row ' Use Core constant
    If lastRowEdits < 1 Then lastRowEdits = 1

    ' Read dashboard data array (A:N)
    Set dashboardRange = wsSrc.Range("A4", wsSrc.Cells(lastRowSrc, Module_Dashboard_Core.DB_COL_COMMENTS)) ' Use Core constant for N
    On Error Resume Next: dashboardDataArray = dashboardRange.Value: If Err.Number <> 0 Or Not IsArray(dashboardDataArray) Then GoTo ReadError: On Error GoTo ErrorHandler
    dashRows = UBound(dashboardDataArray, 1): dashCols = UBound(dashboardDataArray, 2)
    If dashCols < 14 Then GoTo ColumnError ' Check if we have at least A-N

    ' Process dashboard rows
    Set userEditsChanges = New Collection ' To store changes {destRow, docNum, phase, lastContact, comments}
    workbookID = Module_Identity.GetWorkbookIdentity: timeStampStr = Format$(Now(), "yyyy-mm-dd hh:mm:ss")

    For i = 1 To dashRows
        docNum = Trim(CStr(dashboardDataArray(i, Module_Dashboard_Core.DB_COL_DOCNUM))) ' Col A index = 1
        If docNum <> "" And docNum <> "Document Number" Then
            ' Get values from dashboard array (L, M, N -> indices 12, 13, 14)
            Dim dbPhase As Variant: dbPhase = dashboardDataArray(i, 12)
            Dim dbLastContact As Variant: dbLastContact = dashboardDataArray(i, 13)
            Dim dbComments As Variant: dbComments = dashboardDataArray(i, 14)
            hasUserEdits = (CStr(dbPhase) <> "" Or CStr(dbLastContact) <> "" Or CStr(dbComments) <> "")

            If userEditsDict.Exists(docNum) Then editRow = userEditsDict(docNum) Else editRow = 0 ' editRow is the SHEET row number

            ' Process if dashboard has edits OR if it exists in UserEdits (to check for cleared edits)
            If hasUserEdits Or editRow > 0 Then
                wasChanged = False
                If editRow > 0 Then ' Existing entry - check for changes vs UserEdits sheet
                    destRow = editRow
                    On Error Resume Next ' Handle read errors from UserEdits sheet
                    Dim uePhase As Variant: uePhase = wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_PHASE).Value ' Col B
                    Dim ueLastContact As Variant: ueLastContact = wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_LASTCONTACT).Value ' Col C
                    Dim ueComments As Variant: ueComments = wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_COMMENTS).Value ' Col D
                    If Err.Number <> 0 Then wasChanged = True: Err.Clear Else ' Compare values robustly
                        Dim date1 As Date, date2 As Date, datesDiffer As Boolean: datesDiffer = False
                        On Error Resume Next: date1 = CDate(ueLastContact): date2 = CDate(dbLastContact)
                        If Err.Number <> 0 Then If CStr(ueLastContact) <> CStr(dbLastContact) Then datesDiffer = True Else If date1 <> date2 Then datesDiffer = True
                        Err.Clear: On Error GoTo ErrorHandler
                        If CStr(uePhase) <> CStr(dbPhase) Or datesDiffer Or CStr(ueComments) <> CStr(dbComments) Then wasChanged = True
                    End If
                    On Error GoTo ErrorHandler
                Else ' New entry needed in UserEdits
                    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, Module_Dashboard_Core.UE_COL_DOCNUM).End(xlUp).Row + 1 ' Find next row
                    If lastRowEdits < 2 Then lastRowEdits = 2
                    destRow = lastRowEdits
                    userEditsDict.Add docNum, destRow ' Add to dictionary for this run
                    wasChanged = True
                End If

                If wasChanged Then ' Queue update if new or changed
                    newItemData = Array(destRow, docNum, dbPhase, dbLastContact, dbComments)
                    userEditsChanges.Add newItemData
                End If
            End If
        End If
    Next i

    ' Batch Write changes to UserEdits sheet
    If userEditsChanges.Count > 0 Then
        LogUserEditsOperation "Writing " & userEditsChanges.Count & " changes to UserEdits sheet."
        Dim item As Variant
        Application.EnableEvents = False ' Prevent triggering other events during write
        For Each item In userEditsChanges
             destRow = item(0)
             ' Write data using Core constants for UserEdits columns
             If wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_DOCNUM).Value <> item(1) Then wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_DOCNUM).Value = item(1) ' A
             wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_PHASE).Value = item(2) ' B
             wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_LASTCONTACT).Value = item(3) ' C
             wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_COMMENTS).Value = item(4) ' D
             wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_SOURCE).Value = workbookID ' E
             wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_TIMESTAMP).Value = timeStampStr ' F
        Next item
        Application.EnableEvents = True
        LogUserEditsOperation "Finished writing changes to UserEdits."
    Else
         LogUserEditsOperation "No changes detected on dashboard to save to UserEdits."
    End If

ErrorHandler: ' Handles errors and cleanup
    On Error Resume Next ' Ignore errors during cleanup/protection
    If Not wsEdits Is Nothing Then wsEdits.Protect UserInterfaceOnly:=True
    If Not wsSrc Is Nothing Then wsSrc.Protect UserInterfaceOnly:=True
    On Error GoTo 0
    Set userEditsDict = Nothing: Set userEditsChanges = Nothing
    If IsArray(dashboardDataArray) Then Erase dashboardDataArray
    If Err.Number <> 0 Then LogUserEditsOperation "ERROR in SaveUserEditsFromDashboard: [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")" Else LogUserEditsOperation "Completed SaveUserEditsFromDashboard"
    Application.EnableEvents = True ' Ensure events are re-enabled
    Exit Sub ' Exit after handling error or normal completion

ReadError: LogUserEditsOperation "Error reading dashboard data array.": GoTo ErrorHandler
ColumnError: LogUserEditsOperation "Error: Dashboard data array column count mismatch (Expected 14+).": GoTo ErrorHandler
End Sub


'===============================================================================
' SETUPUSEREDITSSHEET: Creates or ensures existence/structure of "UserEdits" sheet
' Uses Core constants for sheet name and columns. Handles structure updates.
'===============================================================================
Public Sub SetupUserEditsSheet()
    Dim wsEdits As Worksheet, wsBackup As Worksheet, currentHeaders As Variant, expectedHeaders As Variant
    Dim structureCorrect As Boolean, backupSuccess As Boolean, i As Long, h As Long
    Dim emergencyFlag As Boolean, backupName As String, structureNeedsUpdate As Boolean

    On Error GoTo ErrorHandler
    ' Define expected headers using Core constants (A-F)
    expectedHeaders = Array("DocNumber", "Engagement Phase", "Last Contact Date", "User Comments", "ChangeSource", "Timestamp")
    Const NUM_EXPECTED_HEADERS As Long = 6

    ' Check if sheet exists
    On Error Resume Next: Set wsEdits = ThisWorkbook.Sheets(Module_Dashboard_Core.USEREDITS_SHEET_NAME): Dim sheetExists As Boolean: sheetExists = (Err.Number = 0 And Not wsEdits Is Nothing): Err.Clear: On Error GoTo ErrorHandler

    If Not sheetExists Then ' Create New Sheet
        LogUserEditsOperation "UserEdits sheet doesn't exist - creating new"
        On Error Resume Next: Set wsEdits = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)): If Err.Number <> 0 Then GoTo FatalError
        wsEdits.Name = Module_Dashboard_Core.USEREDITS_SHEET_NAME: If Err.Number <> 0 Then LogUserEditsOperation "CRITICAL ERROR: Failed to name new UserEdits sheet '" & Module_Dashboard_Core.USEREDITS_SHEET_NAME & "'.": Err.Clear
        ' Set up headers A1:F1
        With wsEdits.Range("A1").Resize(1, NUM_EXPECTED_HEADERS)
             .Value = expectedHeaders: .Font.Bold = True: .Interior.Color = RGB(16, 107, 193): .Font.Color = RGB(255, 255, 255): .EntireColumn.AutoFit
        End With
        wsEdits.Visible = xlSheetHidden: If Err.Number <> 0 Then LogUserEditsOperation "Warning: Error formatting/hiding new UserEdits sheet: " & Err.Description: Err.Clear
        On Error GoTo ErrorHandler: LogUserEditsOperation "Created new UserEdits sheet.": Exit Sub
    End If

    ' Verify Structure of Existing Sheet
    structureNeedsUpdate = True: On Error Resume Next: wsEdits.Unprotect: Err.Clear
    currentHeaders = wsEdits.Range("A1").Resize(1, NUM_EXPECTED_HEADERS).Value ' Read A1:F1
    If Err.Number = 0 And IsArray(currentHeaders) Then
        If UBound(currentHeaders, 1) = 1 And UBound(currentHeaders, 2) = NUM_EXPECTED_HEADERS Then
            structureCorrect = True: For i = 0 To NUM_EXPECTED_HEADERS - 1: If CStr(currentHeaders(1, i + 1)) <> expectedHeaders(i) Then structureCorrect = False: Exit For: End If: Next i
            If structureCorrect Then structureNeedsUpdate = False
        End If
    Else: LogUserEditsOperation "WARNING: Error reading UserEdits headers: " & Err.Description: Err.Clear
    End If: On Error GoTo ErrorHandler
    If Not structureNeedsUpdate Then LogUserEditsOperation "UserEdits sheet structure verified.": Exit Sub ' Structure OK

    ' Backup and Rebuild Incorrect Structure
    LogUserEditsOperation "UserEdits structure needs update - creating backup."
    backupName = Module_Dashboard_Core.USEREDITS_SHEET_NAME & "_StructureUpdateBackup_" & Format(Now, "yyyymmdd_hhmmss"): backupName = Left(backupName, 31)
    On Error Resume Next: Application.DisplayAlerts = False: ThisWorkbook.Sheets(backupName).Delete: Application.DisplayAlerts = True: On Error GoTo ErrorHandler
    Set wsBackup = ThisWorkbook.Sheets.Add(After:=wsEdits): If Err.Number <> 0 Then GoTo FatalError
    wsBackup.Name = backupName: If Err.Number <> 0 Then LogUserEditsOperation "CRITICAL: Failed to name backup sheet '" & backupName & "'.": Err.Clear: GoTo FatalError ' Treat naming failure as fatal
    On Error GoTo ErrorHandler
    On Error Resume Next: wsEdits.UsedRange.Copy wsBackup.Range("A1")
    If Err.Number <> 0 Then LogUserEditsOperation "CRITICAL: Failed to copy data to backup sheet '" & wsBackup.Name & "': " & Err.Description: Application.DisplayAlerts = False: wsBackup.Delete: Application.DisplayAlerts = True: GoTo FatalError
    wsBackup.Visible = xlSheetHidden: LogUserEditsOperation "Created structure update backup: " & wsBackup.Name: On Error GoTo ErrorHandler

    ' Rebuild Sheet
    wsEdits.Cells.Clear: If Err.Number <> 0 Then LogUserEditsOperation "Warning: Error clearing original UserEdits sheet: " & Err.Description: Err.Clear
    With wsEdits.Range("A1").Resize(1, NUM_EXPECTED_HEADERS) ' A1:F1
         .Value = expectedHeaders: .Font.Bold = True: .Interior.Color = RGB(16, 107, 193): .Font.Color = RGB(255, 255, 255): .EntireColumn.AutoFit
    End With
    If Err.Number <> 0 Then LogUserEditsOperation "Warning: Error setting up new headers on UserEdits: " & Err.Description: Err.Clear: On Error GoTo ErrorHandler

    ' Migrate Data from Backup
    Dim lastRowBackup As Long, backupData As Variant, migratedRowCount As Long
    On Error Resume Next: lastRowBackup = wsBackup.Cells(wsBackup.Rows.Count, 1).End(xlUp).Row
    If lastRowBackup > 1 Then backupData = wsBackup.Range("A1", wsBackup.UsedRange.Cells(wsBackup.UsedRange.Rows.Count, wsBackup.UsedRange.Columns.Count)).Value
    If Err.Number <> 0 Or Not IsArray(backupData) Then LogUserEditsOperation "Warning: Could not read data from backup sheet '" & wsBackup.Name & "'.": Err.Clear Else
        ' Map old headers to new indices (0-5 correspond to A-F)
        Dim colMap(0 To NUM_EXPECTED_HEADERS - 1) As Long, headerText As String
        For h = 1 To UBound(backupData, 2)
            headerText = CStr(backupData(1, h))
            Select Case headerText
                Case "DocNumber": colMap(0) = h ' A
                Case "UserStageOverride", "EngagementPhase", "Engagement Phase": colMap(1) = h ' B
                Case "LastContactDate", "Last Contact Date": colMap(2) = h ' C
                Case "UserComments", "User Comments": colMap(3) = h ' D
                Case "ChangeSource": colMap(4) = h ' E
                Case "Timestamp": colMap(5) = h ' F
                ' Ignore "EmailContact" or other old columns
            End Select
        Next h

        ' Prepare migration array (A-F)
        Dim migrationData() As Variant: ReDim migrationData(1 To UBound(backupData, 1) - 1, 1 To NUM_EXPECTED_HEADERS)
        migratedRowCount = 0
        Dim currentIdentity As String: currentIdentity = Module_Identity.GetWorkbookIdentity
        Dim currentTime As Date: currentTime = Now()

        For i = 2 To UBound(backupData, 1) ' Start from row 2 of backup data
             Dim docNumValue As String: docNumValue = ""
             If colMap(0) > 0 Then On Error Resume Next: docNumValue = Trim(CStr(backupData(i, colMap(0)))): On Error GoTo ErrorHandler ' Get DocNum safely
             If docNumValue <> "" Then
                migratedRowCount = migratedRowCount + 1
                migrationData(migratedRowCount, 1) = docNumValue ' Col A
                If colMap(1) > 0 Then On Error Resume Next: migrationData(migratedRowCount, 2) = backupData(i, colMap(1)): On Error GoTo ErrorHandler Else migrationData(migratedRowCount, 2) = vbNullString ' Col B
                If colMap(2) > 0 Then On Error Resume Next: migrationData(migratedRowCount, 3) = backupData(i, colMap(2)): On Error GoTo ErrorHandler Else migrationData(migratedRowCount, 3) = vbNullString ' Col C
                If colMap(3) > 0 Then On Error Resume Next: migrationData(migratedRowCount, 4) = backupData(i, colMap(3)): On Error GoTo ErrorHandler Else migrationData(migratedRowCount, 4) = vbNullString ' Col D
                If colMap(4) > 0 Then On Error Resume Next: migrationData(migratedRowCount, 5) = backupData(i, colMap(4)): On Error GoTo ErrorHandler Else migrationData(migratedRowCount, 5) = currentIdentity ' Col E
                If colMap(5) > 0 Then On Error Resume Next: migrationData(migratedRowCount, 6) = backupData(i, colMap(5)): On Error GoTo ErrorHandler Else migrationData(migratedRowCount, 6) = currentTime ' Col F
             End If
        Next i

        ' Write migrated data if any
        If migratedRowCount > 0 Then
            wsEdits.Range("A2").Resize(migratedRowCount, NUM_EXPECTED_HEADERS).Value = migrationData
            If Err.Number <> 0 Then LogUserEditsOperation "ERROR writing migrated data: " & Err.Description: Err.Clear Else LogUserEditsOperation "Migrated " & migratedRowCount & " records."
        Else
            LogUserEditsOperation "No valid records found in backup to migrate."
        End If
        If IsArray(migrationData) Then Erase migrationData ' Clean up array
    End If: On Error GoTo ErrorHandler

    wsEdits.Visible = xlSheetHidden: LogUserEditsOperation "UserEdits sheet structure update completed."
    Exit Sub

ErrorHandler: LogUserEditsOperation "CRITICAL ERROR in SetupUserEditsSheet: [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")": MsgBox "Error setting up UserEdits: " & Err.Description, vbExclamation: Exit Sub
FatalError: MsgBox "Fatal error during UserEdits setup. Backup may have failed or sheet could not be rebuilt. Manual intervention required.", vbCritical: Exit Sub ' User must intervene
End Sub


'===============================================================================
' LOADUSEREDITSTODICTIONARY: Loads UserEdits into dictionary {DocNum: RowNum}
' Uses Core constant for DocNum column. Returns case-insensitive dictionary.
'===============================================================================
Public Function LoadUserEditsToDictionary(wsEdits As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive keys

    If wsEdits Is Nothing Then GoTo ExitEarly ' Return empty dict if sheet is invalid

    On Error GoTo LoadDictErrorHandler
    Dim lastRow As Long: lastRow = wsEdits.Cells(wsEdits.Rows.Count, Module_Dashboard_Core.UE_COL_DOCNUM).End(xlUp).Row ' Use Core constant
    If lastRow <= 1 Then GoTo ExitEarly ' No data rows

    Dim i As Long, docNum As String, dataRange As Variant
    dataRange = wsEdits.Range(Module_Dashboard_Core.UE_COL_DOCNUM & "2:" & Module_Dashboard_Core.UE_COL_DOCNUM & lastRow).Value ' Read Col A

    If IsArray(dataRange) Then ' Handle multi-row data
        For i = 1 To UBound(dataRange, 1)
            docNum = Trim(CStr(dataRange(i, 1)))
            If docNum <> "" And Not dict.Exists(docNum) Then dict.Add docNum, i + 1 ' Store SHEET row number (i+1)
        Next i
    ElseIf lastRow = 2 And Not IsEmpty(dataRange) Then ' Handle single data row case
        docNum = Trim(CStr(dataRange))
        If docNum <> "" And Not dict.Exists(docNum) Then dict.Add docNum, 2 ' Store SHEET row number 2
    End If

ExitEarly:
    Set LoadUserEditsToDictionary = dict
    Exit Function
LoadDictErrorHandler:
     LogUserEditsOperation "ERROR in LoadUserEditsToDictionary: [" & Err.Number & "] " & Err.Description
     Set LoadUserEditsToDictionary = dict ' Return potentially partial dictionary
End Function

'===============================================================================
' CLEANUPOLDBACKUPS: Deletes UserEdits backup sheets older than a defined period
' Uses Core constant for sheet name.
'===============================================================================
Public Sub CleanupOldBackups()
    On Error GoTo CleanupErrorHandler
    Application.DisplayAlerts = False

    Const DAYS_TO_KEEP As Long = 7 ' Keep backups for 7 days
    Dim cutoffDate As Date: cutoffDate = Date - DAYS_TO_KEEP
    Dim baseBackupName As String: baseBackupName = Module_Dashboard_Core.USEREDITS_SHEET_NAME & "_Backup_" ' Use Core constant
    Dim oldSheets As New Collection, sh As Worksheet, i As Long

    For Each sh In ThisWorkbook.Sheets
        If InStr(1, sh.Name, baseBackupName, vbTextCompare) = 1 Then ' Check if name starts with base name
             Dim datePart As String: datePart = Mid$(sh.Name, Len(baseBackupName) + 1)
             Dim backupDate As Date: On Error Resume Next ' Attempt to parse date
             If InStr(datePart, "_") > 0 And Len(datePart) >= 15 Then ' Format yyyymmdd_hhmmss
                 backupDate = CDate(Format(Left$(datePart, 8), "0000-00-00")) ' Use only date part for comparison
             ElseIf Len(datePart) = 8 Then ' Format yyyymmdd
                 backupDate = CDate(Format(datePart, "0000-00-00"))
             Else ' Unrecognized format
                 backupDate = DateSerial(1900, 1, 1)
             End If
             On Error GoTo CleanupErrorHandler ' Restore main handler
             If IsDate(backupDate) And backupDate <> DateSerial(1900, 1, 1) Then ' Check if valid date parsed
                 If backupDate < cutoffDate Then oldSheets.Add sh ' Add sheet to collection if older than cutoff
             End If
        End If
    Next sh

    If oldSheets.Count > 0 Then
        For i = 1 To oldSheets.Count: oldSheets(i).Delete: Next i
        LogUserEditsOperation "Deleted " & oldSheets.Count & " old backup sheets older than " & Format(cutoffDate, "yyyy-mm-dd") & "."
    Else
        LogUserEditsOperation "No old backup sheets found to delete."
    End If

    Set oldSheets = Nothing: Application.DisplayAlerts = True
    Exit Sub
CleanupErrorHandler:
     LogUserEditsOperation "ERROR during old backup cleanup: [" & Err.Number & "] " & Err.Description
     Application.DisplayAlerts = True: Set oldSheets = Nothing
End Sub

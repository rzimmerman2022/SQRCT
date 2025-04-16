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
    ' This function remains the same as provided in Module_Dashboard_Core
    ' It uses constants defined there (USEREDITSLOG_SHEET_NAME)
    ' and calls Module_Identity.GetWorkbookIdentity()
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
    On Error GoTo 0

    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    If lastRow > MAX_LOG_ROWS Then
        wsLog.Rows("2:" & (lastRow - MAX_LOG_ROWS + 1)).Delete
        lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
        If lastRow < 1 Then lastRow = 1
    End If

    wsLog.Cells(lastRow + 1, "A").Value = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    wsLog.Cells(lastRow + 1, "B").Value = identity
    wsLog.Cells(lastRow + 1, "C").Value = message

ErrorHandler:
    If Err.Number <> 0 Then Debug.Print "Error within LogUserEditsOperation: " & Err.Description
End Sub

'===============================================================================
' USER EDITS BACKUP: Creates a timestamped backup of the UserEdits sheet
'===============================================================================
Public Function CreateUserEditsBackup(Optional backupSuffix As String = "") As Boolean
    ' This function remains the same as provided in Module_Dashboard_Core
    ' It uses constants defined there (USEREDITS_SHEET_NAME)
    On Error GoTo ErrorHandler

    Dim wsEdits As Worksheet
    Dim wsBackup As Worksheet
    Dim backupName As String

    If backupSuffix = "" Then backupName = Module_Dashboard_Core.USEREDITS_SHEET_NAME & "_Backup_" & Format(Now, "yyyymmdd")
    Else: backupName = Module_Dashboard_Core.USEREDITS_SHEET_NAME & "_Backup_" & backupSuffix
    End If
    backupName = Left(backupName, 31) ' Ensure name length

    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets(Module_Dashboard_Core.USEREDITS_SHEET_NAME)
    On Error GoTo ErrorHandler
    If wsEdits Is Nothing Then LogUserEditsOperation "Backup skipped: UserEdits sheet not found.": CreateUserEditsBackup = False: Exit Function

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(backupName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    Set wsBackup = ThisWorkbook.Sheets.Add(After:=wsEdits)
    wsBackup.Name = backupName
    If Err.Number <> 0 Then GoTo ErrorHandler

    wsEdits.UsedRange.Copy wsBackup.Range("A1")
    If Err.Number <> 0 Then GoTo ErrorHandler

    wsBackup.Visible = xlSheetHidden
    LogUserEditsOperation "Created UserEdits backup: " & backupName
    CreateUserEditsBackup = True
    Exit Function

ErrorHandler:
    Debug.Print "Error creating backup '" & backupName & "': [" & Err.Number & "] " & Err.Description
    LogUserEditsOperation "ERROR creating UserEdits backup '" & backupName & "': " & Err.Description
    On Error Resume Next: Application.DisplayAlerts = False: If Not wsBackup Is Nothing Then wsBackup.Delete: Application.DisplayAlerts = True: On Error GoTo 0
    CreateUserEditsBackup = False
End Function

'===============================================================================
' USER EDITS RESTORE: Restores UserEdits data from the most recent backup
'===============================================================================
Public Function RestoreUserEditsFromBackup(Optional specificBackupName As String = "") As Boolean
    ' This function remains the same as provided in Module_Dashboard_Core
    ' It uses constants defined there (USEREDITS_SHEET_NAME)
    On Error GoTo ErrorHandler

    Dim wsEdits As Worksheet
    Dim wsBackup As Worksheet
    Dim backupNameToRestore As String: backupNameToRestore = ""

    If specificBackupName <> "" Then ' Prioritize specific backup name if provided
        On Error Resume Next: Set wsBackup = ThisWorkbook.Sheets(specificBackupName): On Error GoTo ErrorHandler
        If Not wsBackup Is Nothing Then backupNameToRestore = specificBackupName Else LogUserEditsOperation "Specified backup '" & specificBackupName & "' not found."
    End If

    If backupNameToRestore = "" Then ' Find most recent if specific not found/provided
        Dim i As Integer, mostRecentDate As Date, tempName As String
        mostRecentDate = DateSerial(1900, 1, 1)
        For i = 1 To ThisWorkbook.Sheets.Count
            tempName = ThisWorkbook.Sheets(i).Name
            If InStr(1, tempName, Module_Dashboard_Core.USEREDITS_SHEET_NAME & "_Backup_") > 0 Then
                Dim datePart As String: datePart = Mid$(tempName, Len(Module_Dashboard_Core.USEREDITS_SHEET_NAME & "_Backup_") + 1)
                Dim suffixDateTime As Date: On Error Resume Next ' Parse date/time
                If InStr(datePart, "_") > 0 And Len(datePart) >= 15 Then suffixDateTime = CDate(Format(Left$(datePart, 8), "0000-00-00") & " " & Format(Mid$(datePart, 10, 6), "00:00:00"))
                ElseIf Len(datePart) = 8 Then suffixDateTime = CDate(Format(datePart, "0000-00-00"))
                Else: suffixDateTime = DateSerial(1900, 1, 1)
                End If
                If Err.Number = 0 Then If suffixDateTime >= mostRecentDate Then mostRecentDate = suffixDateTime: backupNameToRestore = tempName
                Err.Clear: On Error GoTo ErrorHandler
            End If
        Next i
    End If

    If backupNameToRestore = "" Then LogUserEditsOperation "No suitable backup found to restore.": RestoreUserEditsFromBackup = False: Exit Function

    On Error Resume Next: Set wsBackup = ThisWorkbook.Sheets(backupNameToRestore): On Error GoTo ErrorHandler
    If wsBackup Is Nothing Then LogUserEditsOperation "Backup sheet '" & backupNameToRestore & "' could not be accessed.": RestoreUserEditsFromBackup = False: Exit Function

    SetupUserEditsSheet ' Ensure target sheet exists
    Set wsEdits = ThisWorkbook.Sheets(Module_Dashboard_Core.USEREDITS_SHEET_NAME)
    If wsEdits Is Nothing Then LogUserEditsOperation "UserEdits sheet missing for restore.": RestoreUserEditsFromBackup = False: Exit Function

    On Error Resume Next: wsEdits.Unprotect: On Error GoTo ErrorHandler ' Unprotect target

    wsEdits.Cells.Clear: If Err.Number <> 0 Then GoTo ErrorHandler
    wsBackup.UsedRange.Copy wsEdits.Range("A1"): If Err.Number <> 0 Then GoTo ErrorHandler

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
'===============================================================================
Public Sub SaveUserEditsFromDashboard()
    ' This function remains the same as provided in Module_Dashboard_Core
    ' It uses constants defined there (DB_COL_..., UE_COL_...)
    Dim wsSrc As Worksheet, wsEdits As Worksheet, lastRowSrc As Long, lastRowEdits As Long
    Dim i As Long, destRow As Long, docNum As String, hasUserEdits As Boolean, wasChanged As Boolean
    Dim userEditsDict As Object, editRow As Variant, dashboardDataArray As Variant
    Dim dashboardRange As Range, dashRows As Long, dashCols As Long
    Dim userEditsChanges As Collection, newItemData As Variant
    Dim workbookID As String, timeStampStr As String

    LogUserEditsOperation "Starting SaveUserEditsFromDashboard"
    On Error GoTo ErrorHandler

    ' Get sheets
    Set wsSrc = Module_Dashboard_Core.GetOrCreateDashboardSheet(Module_Dashboard_Core.DASHBOARD_SHEET_NAME) ' Use Public constant
    If wsSrc Is Nothing Then LogUserEditsOperation Module_Dashboard_Core.DASHBOARD_SHEET_NAME & " sheet not found.": Exit Sub
    SetupUserEditsSheet ' Ensure UserEdits exists
    Set wsEdits = ThisWorkbook.Sheets(Module_Dashboard_Core.USEREDITS_SHEET_NAME) ' Use Public constant
    If wsEdits Is Nothing Then LogUserEditsOperation "UserEdits sheet not found.": Exit Sub

    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    If lastRowSrc < 4 Then LogUserEditsOperation "No data rows on dashboard.": Exit Sub

    ' Unprotect sheets
    On Error Resume Next: wsSrc.Unprotect: wsEdits.Unprotect: On Error GoTo ErrorHandler

    ' Load dictionary
    Set userEditsDict = LoadUserEditsToDictionary(wsEdits) ' Call Private helper in this module
    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, Module_Dashboard_Core.UE_COL_DOCNUM).End(xlUp).Row
    If lastRowEdits < 1 Then lastRowEdits = 1

    ' Read dashboard data array
    Set dashboardRange = wsSrc.Range("A4", wsSrc.Cells(lastRowSrc, Module_Dashboard_Core.DB_COL_COMMENTS)) ' A:N
    On Error Resume Next: dashboardDataArray = dashboardRange.Value: If Err.Number <> 0 Or Not IsArray(dashboardDataArray) Then GoTo ReadError: On Error GoTo ErrorHandler
    dashRows = UBound(dashboardDataArray, 1): dashCols = UBound(dashboardDataArray, 2)
    If dashCols < 14 Then GoTo ColumnError

    ' Process dashboard rows
    Set userEditsChanges = New Collection
    workbookID = Module_Identity.GetWorkbookIdentity: timeStampStr = Format$(Now(), "yyyy-mm-dd hh:mm:ss")

    For i = 1 To dashRows
        docNum = Trim(CStr(dashboardDataArray(i, 1))) ' Col A
        If docNum <> "" And docNum <> "Document Number" Then
            Dim dbPhase As Variant: dbPhase = dashboardDataArray(i, 12) ' Col L
            Dim dbLastContact As Variant: dbLastContact = dashboardDataArray(i, 13) ' Col M
            Dim dbComments As Variant: dbComments = dashboardDataArray(i, 14) ' Col N
            hasUserEdits = (CStr(dbPhase) <> "" Or CStr(dbLastContact) <> "" Or CStr(dbComments) <> "")

            If userEditsDict.Exists(docNum) Then editRow = userEditsDict(docNum) Else editRow = 0

            If hasUserEdits Or editRow > 0 Then
                wasChanged = False
                If editRow > 0 Then ' Existing entry - check for changes
                    destRow = editRow
                    On Error Resume Next ' Handle read errors from UserEdits
                    Dim uePhase As Variant: uePhase = wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_PHASE).Value
                    Dim ueLastContact As Variant: ueLastContact = wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_LASTCONTACT).Value
                    Dim ueComments As Variant: ueComments = wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_COMMENTS).Value
                    If Err.Number <> 0 Then wasChanged = True: Err.Clear Else ' Compare values robustly
                        Dim date1 As Date, date2 As Date, datesDiffer As Boolean: datesDiffer = False
                        On Error Resume Next: date1 = CDate(ueLastContact): date2 = CDate(dbLastContact)
                        If Err.Number <> 0 Then If CStr(ueLastContact) <> CStr(dbLastContact) Then datesDiffer = True Else If date1 <> date2 Then datesDiffer = True
                        Err.Clear: On Error GoTo ErrorHandler
                        If CStr(uePhase) <> CStr(dbPhase) Or datesDiffer Or CStr(ueComments) <> CStr(dbComments) Then wasChanged = True
                    End If
                    On Error GoTo ErrorHandler
                Else ' New entry
                    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, Module_Dashboard_Core.UE_COL_DOCNUM).End(xlUp).Row + 1
                    If lastRowEdits < 2 Then lastRowEdits = 2
                    destRow = lastRowEdits
                    userEditsDict.Add docNum, destRow ' Add to dictionary
                    wasChanged = True
                End If

                If wasChanged Then ' Queue update
                    newItemData = Array(destRow, docNum, dbPhase, dbLastContact, dbComments)
                    userEditsChanges.Add newItemData
                End If
            End If
        End If
    Next i

    ' Batch Write changes
    If userEditsChanges.Count > 0 Then
        LogUserEditsOperation "Writing " & userEditsChanges.Count & " changes to UserEdits sheet."
        Dim item As Variant
        Application.EnableEvents = False
        For Each item In userEditsChanges
             destRow = item(0)
             If wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_DOCNUM).Value <> item(1) Then wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_DOCNUM).Value = item(1)
             wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_PHASE).Value = item(2)
             wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_LASTCONTACT).Value = item(3)
             wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_COMMENTS).Value = item(4)
             wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_SOURCE).Value = workbookID
             wsEdits.Cells(destRow, Module_Dashboard_Core.UE_COL_TIMESTAMP).Value = timeStampStr
        Next item
        Application.EnableEvents = True
        LogUserEditsOperation "Finished writing changes to UserEdits."
    Else
         LogUserEditsOperation "No changes detected on dashboard to save to UserEdits."
    End If

ErrorHandler: ' Handles errors and cleanup
    On Error Resume Next ' Ignore errors during cleanup/protection
    If Not wsEdits Is Nothing Then wsEdits.Protect UserInterfaceOnly:=True
    On Error GoTo 0
    Set userEditsDict = Nothing: Set userEditsChanges = Nothing
    If IsArray(dashboardDataArray) Then Erase dashboardDataArray
    If Err.Number <> 0 Then LogUserEditsOperation "ERROR in SaveUserEditsFromDashboard: [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")" Else LogUserEditsOperation "Completed SaveUserEditsFromDashboard"
    Application.EnableEvents = True
    Exit Sub ' Exit after handling error or normal completion

ReadError: LogUserEditsOperation "Error reading dashboard data array.": GoTo ErrorHandler
ColumnError: LogUserEditsOperation "Error: Dashboard data array column count mismatch.": GoTo ErrorHandler
End Sub


'===============================================================================
' SETUPUSEREDITSSHEET: Creates or ensures existence/structure of "UserEdits" sheet
'===============================================================================
Public Sub SetupUserEditsSheet()
    ' This function remains the same as provided in Module_Dashboard_Core
    ' It uses constants defined there (USEREDITS_SHEET_NAME, UE_COL_...)
    Dim wsEdits As Worksheet, wsBackup As Worksheet, currentHeaders As Variant, expectedHeaders As Variant
    Dim structureCorrect As Boolean, backupSuccess As Boolean, i As Long, h As Long
    Dim emergencyFlag As Boolean, backupName As String, structureNeedsUpdate As Boolean

    On Error GoTo ErrorHandler
    expectedHeaders = Array("DocNumber", "Engagement Phase", "Last Contact Date", "User Comments", "ChangeSource", "Timestamp")

    On Error Resume Next: Set wsEdits = ThisWorkbook.Sheets(Module_Dashboard_Core.USEREDITS_SHEET_NAME): Dim sheetExists As Boolean: sheetExists = (Err.Number = 0 And Not wsEdits Is Nothing): Err.Clear: On Error GoTo ErrorHandler

    If Not sheetExists Then ' Create New Sheet
        LogUserEditsOperation "UserEdits sheet doesn't exist - creating new"
        On Error Resume Next: Set wsEdits = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)): If Err.Number <> 0 Then GoTo FatalError
        wsEdits.Name = Module_Dashboard_Core.USEREDITS_SHEET_NAME: If Err.Number <> 0 Then LogUserEditsOperation "CRITICAL ERROR: Failed to name new UserEdits sheet '" & Module_Dashboard_Core.USEREDITS_SHEET_NAME & "'.": Err.Clear
        With wsEdits.Range("A1").Resize(1, UBound(expectedHeaders) + 1): .Value = expectedHeaders: .Font.Bold = True: .Interior.Color = RGB(16, 107, 193): .Font.Color = RGB(255, 255, 255): .EntireColumn.AutoFit: End With
        wsEdits.Visible = xlSheetHidden: If Err.Number <> 0 Then LogUserEditsOperation "Warning: Error formatting/hiding new UserEdits sheet: " & Err.Description: Err.Clear
        On Error GoTo ErrorHandler: LogUserEditsOperation "Created new UserEdits sheet.": Exit Sub
    End If

    ' Verify Structure of Existing Sheet
    structureNeedsUpdate = True: On Error Resume Next: wsEdits.Unprotect: Err.Clear
    currentHeaders = wsEdits.Range("A1").Resize(1, UBound(expectedHeaders) + 1).Value
    If Err.Number = 0 And IsArray(currentHeaders) Then
        If UBound(currentHeaders, 1) = 1 And UBound(currentHeaders, 2) = UBound(expectedHeaders) + 1 Then
            structureCorrect = True: For i = 0 To UBound(expectedHeaders): If CStr(currentHeaders(1, i + 1)) <> expectedHeaders(i) Then structureCorrect = False: Exit For: End If: Next i
            If structureCorrect Then structureNeedsUpdate = False
        End If
    Else: LogUserEditsOperation "WARNING: Error reading UserEdits headers: " & Err.Description: Err.Clear
    End If: On Error GoTo ErrorHandler
    If Not structureNeedsUpdate Then Exit Sub ' Structure OK

    ' Backup and Rebuild Incorrect Structure
    LogUserEditsOperation "UserEdits structure needs update - creating backup."
    backupName = Module_Dashboard_Core.USEREDITS_SHEET_NAME & "_StructureUpdateBackup_" & Format(Now, "yyyymmdd_hhmmss"): backupName = Left(backupName, 31)
    On Error Resume Next: Application.DisplayAlerts = False: ThisWorkbook.Sheets(backupName).Delete: Application.DisplayAlerts = True: On Error GoTo ErrorHandler
    Set wsBackup = ThisWorkbook.Sheets.Add(After:=wsEdits): If Err.Number <> 0 Then GoTo FatalError
    wsBackup.Name = backupName: If Err.Number <> 0 Then LogUserEditsOperation "CRITICAL: Failed to name backup sheet '" & backupName & "'.": Err.Clear
    On Error GoTo ErrorHandler
    On Error Resume Next: wsEdits.UsedRange.Copy wsBackup.Range("A1")
    If Err.Number <> 0 Then LogUserEditsOperation "CRITICAL: Failed to copy data to backup sheet '" & wsBackup.Name & "': " & Err.Description: Application.DisplayAlerts = False: wsBackup.Delete: Application.DisplayAlerts = True: GoTo FatalError
    wsBackup.Visible = xlSheetHidden: LogUserEditsOperation "Created structure update backup: " & wsBackup.Name: On Error GoTo ErrorHandler

    ' Rebuild Sheet
    wsEdits.Cells.Clear: If Err.Number <> 0 Then LogUserEditsOperation "Warning: Error clearing original UserEdits sheet: " & Err.Description: Err.Clear
    With wsEdits.Range("A1").Resize(1, UBound(expectedHeaders) + 1): .Value = expectedHeaders: .Font.Bold = True: .Interior.Color = RGB(16, 107, 193): .Font.Color = RGB(255, 255, 255): .EntireColumn.AutoFit: End With
    If Err.Number <> 0 Then LogUserEditsOperation "Warning: Error setting up new headers on UserEdits: " & Err.Description: Err.Clear: On Error GoTo ErrorHandler

    ' Migrate Data
    Dim lastRowBackup As Long, backupData As Variant, migratedRowCount As Long
    On Error Resume Next: lastRowBackup = wsBackup.Cells(wsBackup.Rows.Count, 1).End(xlUp).Row
    If lastRowBackup > 1 Then backupData = wsBackup.Range("A1", wsBackup.UsedRange.Cells(wsBackup.UsedRange.Rows.Count, wsBackup.UsedRange.Columns.Count)).Value
    If Err.Number <> 0 Or Not IsArray(backupData) Then LogUserEditsOperation "Warning: Could not read data from backup sheet '" & wsBackup.Name & "'.": Err.Clear Else
        Dim colMap(0 To 5) As Long, headerText As String: For h = 1 To UBound(backupData, 2): headerText = CStr(backupData(1, h)): Select Case headerText: Case "DocNumber": colMap(0) = h: Case "UserStageOverride", "EngagementPhase", "Engagement Phase": colMap(1) = h: Case "LastContactDate", "Last Contact Date": colMap(2) = h: Case "UserComments", "User Comments": colMap(3) = h: Case "ChangeSource": colMap(4) = h: Case "Timestamp": colMap(5) = h: End Select: Next h
        Dim migrationData() As Variant: ReDim migrationData(1 To UBound(backupData, 1) - 1, 1 To 6): migratedRowCount = 0
        Dim currentIdentity As String: currentIdentity = Module_Identity.GetWorkbookIdentity
        For i = 2 To UBound(backupData, 1)
             Dim docNumValue As String: docNumValue = "": If colMap(0) > 0 Then docNumValue = Trim(CStr(backupData(i, colMap(0))))
             If docNumValue <> "" Then
                migratedRowCount = migratedRowCount + 1
                migrationData(migratedRowCount, 1) = docNumValue
                If colMap(1) > 0 Then migrationData(migratedRowCount, 2) = backupData(i, colMap(1)) Else migrationData(migratedRowCount, 2) = vbNullString
                If colMap(2) > 0 Then migrationData(migratedRowCount, 3) = backupData(i, colMap(2)) Else migrationData(migratedRowCount, 3) = vbNullString
                If colMap(3) > 0 Then migrationData(migratedRowCount, 4) = backupData(i, colMap(3)) Else migrationData(migratedRowCount, 4) = vbNullString
                If colMap(4) > 0 Then migrationData(migratedRowCount, 5) = backupData(i, colMap(4)) Else migrationData(migratedRowCount, 5) = currentIdentity
                If colMap(5) > 0 Then migrationData(migratedRowCount, 6) = backupData(i, colMap(5)) Else migrationData(migratedRowCount, 6) = Now()
             End If
        Next i
        If migratedRowCount > 0 Then wsEdits.Range("A2").Resize(migratedRowCount, 6).Value = migrationData: If Err.Number <> 0 Then LogUserEditsOperation "ERROR writing migrated data: " & Err.Description: Err.Clear: LogUserEditsOperation "Migrated " & migratedRowCount & " records." Else LogUserEditsOperation "No valid records in backup."
    End If: On Error GoTo ErrorHandler

    wsEdits.Visible = xlSheetHidden: LogUserEditsOperation "UserEdits sheet structure update completed."
    Exit Sub

ErrorHandler: LogUserEditsOperation "CRITICAL ERROR in SetupUserEditsSheet: [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")": MsgBox "Error setting up UserEdits: " & Err.Description, vbExclamation: Exit Sub
FatalError: MsgBox "Fatal error during UserEdits setup. Backup failed.", vbCritical: Exit Sub ' User must intervene
End Sub


'===============================================================================
' LOADUSEREDITSTODICTIONARY: Loads UserEdits into dictionary {DocNum: RowNum}
'===============================================================================
Public Function LoadUserEditsToDictionary(wsEdits As Worksheet) As Object
    ' This function remains the same as provided in Module_Dashboard_Core
    ' It uses constants defined there (UE_COL_DOCNUM)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive keys

    If wsEdits Is Nothing Then GoTo ExitEarly

    On Error GoTo LoadDictErrorHandler
    Dim lastRow As Long: lastRow = wsEdits.Cells(wsEdits.Rows.Count, Module_Dashboard_Core.UE_COL_DOCNUM).End(xlUp).Row
    If lastRow <= 1 Then GoTo ExitEarly

    Dim i As Long, docNum As String, dataRange As Variant
    dataRange = wsEdits.Range(Module_Dashboard_Core.UE_COL_DOCNUM & "2:" & Module_Dashboard_Core.UE_COL_DOCNUM & lastRow).Value

    If IsArray(dataRange) Then
        For i = 1 To UBound(dataRange, 1)
            docNum = Trim(CStr(dataRange(i, 1)))
            If docNum <> "" And Not dict.Exists(docNum) Then dict.Add docNum, i + 1
        Next i
    ElseIf lastRow = 2 And Not IsEmpty(dataRange) Then
        docNum = Trim(CStr(dataRange))
        If docNum <> "" And Not dict.Exists(docNum) Then dict.Add docNum, 2
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
'===============================================================================
Public Sub CleanupOldBackups()
    ' This function remains the same as provided in Module_Dashboard_Core
    ' It uses constants defined there (USEREDITS_SHEET_NAME)
    On Error GoTo CleanupErrorHandler
    Application.DisplayAlerts = False

    Const DAYS_TO_KEEP As Long = 7
    Dim cutoffDate As Date: cutoffDate = Date - DAYS_TO_KEEP
    Dim backupBaseName As String: backupBaseName = Module_Dashboard_Core.USEREDITS_SHEET_NAME & "_Backup_"
    Dim oldSheets As New Collection, sh As Worksheet, i As Long

    For Each sh In ThisWorkbook.Sheets
        If InStr(1, sh.Name, backupBaseName) > 0 Then
             Dim datePart As String: datePart = Mid$(sh.Name, Len(backupBaseName) + 1)
             Dim backupDate As Date: On Error Resume Next
             If InStr(datePart, "_") > 0 And Len(datePart) >= 15 Then backupDate = CDate(Format(Left$(datePart, 8), "0000-00-00"))
             ElseIf Len(datePart) = 8 Then backupDate = CDate(Format(datePart, "0000-00-00"))
             Else: backupDate = DateSerial(1900, 1, 1)
             End If: On Error GoTo CleanupErrorHandler
             If IsDate(backupDate) Then If backupDate < cutoffDate Then oldSheets.Add sh
        End If
    Next sh

    If oldSheets.Count > 0 Then
        For i = 1 To oldSheets.Count: oldSheets(i).Delete: Next i
        LogUserEditsOperation "Deleted " & oldSheets.Count & " old backup sheets."
    End If

    Set oldSheets = Nothing: Application.DisplayAlerts = True
    Exit Sub
CleanupErrorHandler:
     LogUserEditsOperation "ERROR during old backup cleanup: [" & Err.Number & "] " & Err.Description
     Application.DisplayAlerts = True: Set oldSheets = Nothing
End Sub


Attribute VB_Name = "Module_Dashboard"
Option Explicit

' --- Constants ---
Private Const DASHBOARD_SHEET_NAME As String = "SQRCT Dashboard"
Private Const USEREDITS_SHEET_NAME As String = "UserEdits"
Private Const USEREDITSLOG_SHEET_NAME As String = "UserEditsLog"
Private Const MASTER_QUOTES_FINAL_SOURCE As String = "MasterQuotes_Final" ' Name of the PQ query/table for A-J

' *** NEW: Constants for Power Query Output ***
Private Const PQ_LATEST_LOCATION_SHEET As String = "DocNum_LatestLocation" ' Sheet where DocNum_LatestLocation loads
Private Const PQ_LATEST_LOCATION_TABLE As String = "DocNum_LatestLocation" ' Assumes table name matches query name

' UserEdits Columns (Remain A-F as Email was already removed)
Private Const UE_COL_DOCNUM As String = "A"
Private Const UE_COL_PHASE As String = "B"
Private Const UE_COL_LASTCONTACT As String = "C"
Private Const UE_COL_COMMENTS As String = "D"
Private Const UE_COL_SOURCE As String = "E"
Private Const UE_COL_TIMESTAMP As String = "F"

' Dashboard Columns (Adjusted for NEW Column K)
' A-J remain the same (populated by MasterQuotes_Final)
' *** NEW ***
Private Const DB_COL_LATEST_LOCATION As String = "K" ' New column for PQ Location lookup
' *** SHIFTED ***
Private Const DB_COL_PHASE As String = "L"         ' Shifted from K
Private Const DB_COL_LASTCONTACT As String = "M"     ' Shifted from L
Private Const DB_COL_COMMENTS As String = "N"        ' Shifted from M
' --- End Constants ---


'===============================================================================
' MODULE_DASHBOARD
' Contains functions for managing the SQRCT Dashboard, including refresh operations,
' user edits management, and UI interactions.
' Refactored for performance (Dictionary lookup), security (no passwords), and robustness (constants).
' Protection logic revised: Applied after modifications in relevant subs.
' Added Latest Location lookup from Power Query.
'===============================================================================

'===============================================================================
' BUTTON FUNCTIONS: These can be assigned to buttons in the UI
'===============================================================================

' Standard workflow - saves dashboard edits first, then restores after refresh
Public Sub Button_RefreshDashboard_SaveAndRestoreEdits()
    RefreshDashboard_TwoWaySync
End Sub

' Special workflow - only updates dashboard from UserEdits without saving current edits
Public Sub Button_RefreshDashboard_PreserveUserEdits()
    RefreshDashboard_OneWayFromUserEdits
End Sub

'===============================================================================
' MAIN FUNCTIONS: with clear names indicating data flow direction
'===============================================================================

' Standard two-way sync: Dashboard -> UserEdits -> Refresh -> UserEdits -> Dashboard
Public Sub RefreshDashboard_TwoWaySync()
    Call RefreshDashboard(PreserveUserEdits:=False)
End Sub

' One-way sync: Refresh -> UserEdits -> Dashboard (preserves manual UserEdits)
Public Sub RefreshDashboard_OneWayFromUserEdits()
    Call RefreshDashboard(PreserveUserEdits:=True)
End Sub

'===============================================================================
' USER EDITS LOGGING: Tracks operations performed on UserEdits
'===============================================================================
Public Sub LogUserEditsOperation(message As String)
    Dim wsLog As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets(USEREDITSLOG_SHEET_NAME) ' Use Constant

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = USEREDITSLOG_SHEET_NAME ' Use Constant
        wsLog.Range("A1:C1").Value = Array("Timestamp", "Workbook", "Operation")
        wsLog.Range("A1:C1").Font.Bold = True
        wsLog.Visible = xlSheetHidden
    End If
    On Error GoTo 0 ' Restore default error handling before proceeding

    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    wsLog.Cells(lastRow + 1, "A").Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    wsLog.Cells(lastRow + 1, "B").Value = Module_Identity.GetWorkbookIdentity() ' Using Module_Identity function
    wsLog.Cells(lastRow + 1, "C").Value = message

    On Error GoTo 0
End Sub

'===============================================================================
' USER EDITS BACKUP: Creates a backup of the UserEdits sheet
'===============================================================================
Public Function CreateUserEditsBackup(Optional backupSuffix As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim wsEdits As Worksheet
    Dim wsBackup As Worksheet
    Dim backupName As String

    ' Set backup name
    If backupSuffix = "" Then
        backupName = USEREDITS_SHEET_NAME & "_Backup_" & Format(Now, "yyyymmdd") ' Use Constant
    Else
        backupName = USEREDITS_SHEET_NAME & "_Backup_" & backupSuffix ' Use Constant
    End If

    ' Create backup only if UserEdits exists and has data
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Use Constant
    On Error GoTo ErrorHandler

    If wsEdits Is Nothing Then
        CreateUserEditsBackup = False
        Exit Function
    End If

    ' Check if the backup sheet already exists
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets(backupName)
    If wsBackup Is Nothing Then
        Set wsBackup = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsBackup.Name = backupName
    Else
        wsBackup.Cells.Clear
    End If
    On Error GoTo ErrorHandler

    ' Copy data
    wsEdits.UsedRange.Copy wsBackup.Range("A1")

    ' Hide backup sheet
    wsBackup.Visible = xlSheetHidden

    ' Log the backup creation
    LogUserEditsOperation "Created UserEdits backup: " & backupName

    CreateUserEditsBackup = True
    Exit Function

ErrorHandler:
    Debug.Print "Error creating backup: " & Err.Description
    LogUserEditsOperation "ERROR creating UserEdits backup: " & Err.Description
    CreateUserEditsBackup = False
End Function

'===============================================================================
' USER EDITS RESTORE: Restores UserEdits data from a backup
'===============================================================================
Public Function RestoreUserEditsFromBackup(Optional backupName As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim wsEdits As Worksheet
    Dim wsBackup As Worksheet

    ' Find most recent backup if name not provided
    If backupName = "" Then
        Dim i As Integer
        For i = 1 To ThisWorkbook.Sheets.Count
            If InStr(1, ThisWorkbook.Sheets(i).Name, USEREDITS_SHEET_NAME & "_Backup_") > 0 Then ' Use Constant
                backupName = ThisWorkbook.Sheets(i).Name
                ' If multiple backups exist, we'll get the last one alphabetically
            End If
        Next i
    End If

    If backupName = "" Then
        RestoreUserEditsFromBackup = False
        Exit Function
    End If

    ' Get backup sheet
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets(backupName)
    On Error GoTo ErrorHandler

    If wsBackup Is Nothing Then
        RestoreUserEditsFromBackup = False
        Exit Function
    End If

    ' Ensure UserEdits sheet exists
    SetupUserEditsSheet
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Use Constant

    ' Clear current data and restore from backup
    wsEdits.Cells.Clear
    wsBackup.UsedRange.Copy wsEdits.Range("A1")

    ' Log the restoration
    LogUserEditsOperation "Restored UserEdits from backup: " & backupName

    RestoreUserEditsFromBackup = True
    Exit Function

ErrorHandler:
    Debug.Print "Error restoring from backup: " & Err.Description
    LogUserEditsOperation "ERROR restoring from backup: " & Err.Description
    RestoreUserEditsFromBackup = False
End Function

'===============================================================================
' MAIN SUB: Creates or refreshes the SQRCT Dashboard (Standardized Version)
' Optimized Restore Edits section using Array Method
' Includes logic to create/update a Text-Only version of the dashboard
' Adjusted for NEW Latest Location column K, shifted user columns L-N
'===============================================================================
Public Sub RefreshDashboard(Optional PreserveUserEdits As Boolean = False)
    Dim ws As Worksheet, wsEdits As Worksheet
    Dim lastRow As Long, lastRowEdits As Long
    Dim backupCreated As Boolean
    Dim t_start As Single, t_location As Single, t_format As Single, t_textOnly As Single
    Dim userEditsDict As Object
    Dim dashboardDocNumArray As Variant, userEditsDataArray As Variant, outputEditsArray As Variant
    Dim numDashboardRows As Long
    Dim wsValues As Worksheet, srcRange As Range, currentSheet As Worksheet
    Const TEXT_ONLY_SHEET_NAME As String = "SQRCT Dashboard (Text-Only)"
    
    ' 1) Backup & log
    backupCreated = CreateUserEditsBackup("RefreshDashboard_" & Format(Now, "yyyymmdd_hhmmss"))
    LogUserEditsOperation "Starting dashboard refresh. PreserveUserEdits=" & PreserveUserEdits & ", Backup created: " & backupCreated
    t_start = Timer
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' 2) Ensure UserEdits sheet
    SetupUserEditsSheet
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    
    ' 3) Save edits if needed
    If Not PreserveUserEdits Then SaveUserEditsFromDashboard
    
    ' 4) Get or create Dashboard
    Set ws = GetOrCreateDashboardSheet(DASHBOARD_SHEET_NAME)
    ws.Unprotect
    
    ' 5) Clean and init layout
    CleanupDashboardLayout ws
    InitializeDashboardLayout ws
    
    ' 6) Populate A�J
    If IsMasterQuotesFinalPresent Then
        PopulateMasterQuotesData ws
    Else
        MsgBox "Warning: " & MASTER_QUOTES_FINAL_SOURCE & " not found.", vbInformation, "Data Source Not Found"
        GoTo Cleanup
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 7) Sort data
    SortDashboardData ws, lastRow
    
    ' 8) Populate Latest Location (Workflow Location) AFTER sort
    If lastRow >= 4 Then
        t_location = Timer
        PopulateWorkflowLocation ws, lastRow
        Debug.Print "PopulateWorkflowLocation Time: " & Timer - t_location
    End If
    
    ' 9) Autofit & widths
    With ws
        .Columns("A:J").AutoFit
        .Columns(DB_COL_LATEST_LOCATION).ColumnWidth = 20
        .Columns(DB_COL_PHASE & ":" & DB_COL_COMMENTS).AutoFit
        .Columns("C").ColumnWidth = 25
        .Columns(DB_COL_COMMENTS).ColumnWidth = 40
    End With
    
    ' 10) Restore user edits via arrays
    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row
    On Error Resume Next
    dashboardDocNumArray = ws.Range("A4:A" & lastRow).Value
    On Error GoTo ErrorHandler
    
    If lastRowEdits > 1 Then
        Dim userEditsRange As Range, singleRowData As Variant, cIdx As Long
        Set userEditsRange = wsEdits.Range(UE_COL_DOCNUM & "2:" & UE_COL_TIMESTAMP & lastRowEdits)
        If userEditsRange.Rows.Count = 1 Then
            ReDim singleRowData(1 To 1, 1 To 6)
            For cIdx = 1 To 6
                singleRowData(1, cIdx) = userEditsRange.Cells(1, cIdx).Value
            Next cIdx
            userEditsDataArray = singleRowData
        Else
            userEditsDataArray = userEditsRange.Value
        End If
    End If
    
    Set userEditsDict = LoadUserEditsToDictionary(wsEdits)
    
    numDashboardRows = UBound(dashboardDocNumArray, 1)
    ReDim outputEditsArray(1 To numDashboardRows, 1 To 3)
    
    Dim i As Long, j As Long, editSheetRow As Long, editArrayRow As Long, ubUED As Long
    For i = 1 To numDashboardRows
        For j = 1 To 3
            outputEditsArray(i, j) = vbNullString
        Next j
    Next i
    
    If Not IsEmpty(userEditsDataArray) And lastRowEdits > 1 Then
        ubUED = UBound(userEditsDataArray, 1)
        For i = 1 To numDashboardRows
            Dim docNum As String
            docNum = Trim(CStr(dashboardDocNumArray(i, 1)))
            If docNum <> "" Then
                If userEditsDict.Exists(docNum) Then
                    editSheetRow = userEditsDict(docNum)
                    editArrayRow = editSheetRow - 1
                    If editArrayRow >= 1 And editArrayRow <= ubUED Then
                        On Error Resume Next
                        outputEditsArray(i, 1) = userEditsDataArray(editArrayRow, 2)
                        outputEditsArray(i, 2) = userEditsDataArray(editArrayRow, 3)
                        outputEditsArray(i, 3) = userEditsDataArray(editArrayRow, 4)
                        On Error GoTo ErrorHandler
                    End If
                End If
            End If
        Next i
    End If
    
    ws.Range(DB_COL_PHASE & "4").Resize(numDashboardRows, 3).Value = outputEditsArray
    
    ' 11) Freeze panes
    FreezeDashboard ws
    
    ' 12) Formatting & protection
    t_format = Timer
    ApplyColorFormatting ws
    ProtectUserColumns ws
    Debug.Print "Format/Protect Time: " & Timer - t_format
    
    ' 13) Timestamp
    With ws.Range("G2:I2")
        .Merge
        .Value = "Last Refreshed: " & Format$(Now, "mm/dd/yyyy h:mm") & " MST"
        .HorizontalAlignment = xlCenter
        .Font.Size = 9: .Font.Name = "Segoe UI"
        .Font.Color = RGB(80, 80, 80)
    End With
    
    ' 14) Buttons
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.TopLeftCell.Row = 2 Then shp.Delete
    Next shp
    On Error GoTo 0
    
    ModernButton ws, "C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits"
    ModernButton ws, "E2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits"
    
    MsgBox IIf(PreserveUserEdits,
        DASHBOARD_SHEET_NAME & " refreshed!" & vbCrLf & USEREDITS_SHEET_NAME & " preserved.",
        DASHBOARD_SHEET_NAME & " refreshed!" & vbCrLf & "Edits saved & restored."
    ), vbInformation, "Dashboard Refresh Complete"
    
    LogUserEditsOperation "Dashboard refresh completed successfully."
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    LogUserEditsOperation "ERROR in RefreshDashboard: " & Err.Description
    MsgBox "Error during refresh: " & Err.Description, vbCritical, "Dashboard Refresh Error"
    If backupCreated Then RestoreUserEditsFromBackup
    Resume Cleanup
End Sub


'===============================================================================
' *** NEW SUBROUTINE: Populate Latest Location from Power Query ***
'===============================================================================
Private Sub PopulateLatestLocation(ws As Worksheet, lastRow As Long)
    Dim wsPQ As Worksheet
    Dim tblPQ As ListObject
    Dim pqDataRange As Range
    Dim pqDocNumCol As Range, pqLocationCol As Range
    Dim dashboardDocNumArray As Variant
    Dim outputLocationArray As Variant
    Dim i As Long
    Dim docNum As String
    Dim locationResult As Variant

    Dim t_start As Single
    t_start = Timer
    LogUserEditsOperation "Starting PopulateLatestLocation"

    ' --- Error Handling Setup ---
    On Error GoTo LocationErrorHandler

    ' --- Get References to PQ Output Sheet and Table ---
    Set wsPQ = Nothing ' Reset object variable
    On Error Resume Next ' Temporarily ignore error if sheet doesn't exist
    Set wsPQ = ThisWorkbook.Sheets(PQ_LATEST_LOCATION_SHEET)
    On Error GoTo LocationErrorHandler ' Restore error handling
    If wsPQ Is Nothing Then
        LogUserEditsOperation "ERROR: Sheet '" & PQ_LATEST_LOCATION_SHEET & "' not found."
        Exit Sub ' Cannot proceed without the PQ output sheet
    End If

    Set tblPQ = Nothing ' Reset object variable
    On Error Resume Next ' Temporarily ignore error if table doesn't exist
    Set tblPQ = wsPQ.ListObjects(PQ_LATEST_LOCATION_TABLE) ' Assumes table name matches query name
    If tblPQ Is Nothing Then Set tblPQ = wsPQ.ListObjects(1) ' Fallback: try first table on sheet
    On Error GoTo LocationErrorHandler ' Restore error handling
    If tblPQ Is Nothing Then
        LogUserEditsOperation "ERROR: Table '" & PQ_LATEST_LOCATION_TABLE & "' (or first table) not found on sheet '" & PQ_LATEST_LOCATION_SHEET & "'."
        Exit Sub ' Cannot proceed without the PQ output table
    End If

    ' --- Verify Required Columns Exist in PQ Table ---
    Dim pqDocNumColIndex As Long
    Dim pqLocationColIndex As Long
    On Error Resume Next ' Check for column existence
    pqDocNumColIndex = tblPQ.ListColumns("PrimaryDocNumber").Index
    pqLocationColIndex = tblPQ.ListColumns("MostRecent_FolderLocation").Index
    On Error GoTo LocationErrorHandler ' Restore error handling
    If pqDocNumColIndex = 0 Or pqLocationColIndex = 0 Then
        LogUserEditsOperation "ERROR: Required columns ('PrimaryDocNumber' or 'MostRecent_FolderLocation') not found in PQ table '" & tblPQ.Name & "'."
        Exit Sub
    End If

    ' --- Read Data into Arrays for Performance ---
    ' Read Dashboard Document Numbers (Column A, rows 4 to lastRow)
    dashboardDocNumArray = ws.Range("A4:A" & lastRow).Value

    ' Read PQ Table Data (only the two needed columns) into an array
    ' This is generally faster than repeated lookups on the sheet/table object
    Dim pqTableArray As Variant
    Set pqDocNumCol = tblPQ.ListColumns(pqDocNumColIndex).DataBodyRange
    Set pqLocationCol = tblPQ.ListColumns(pqLocationColIndex).DataBodyRange
    ' Create a temporary array to hold the two columns
    ReDim pqTableArray(1 To pqDocNumCol.Rows.Count, 1 To 2)
    Dim r As Long
    For r = 1 To pqDocNumCol.Rows.Count
        pqTableArray(r, 1) = pqDocNumCol.Cells(r, 1).Value ' DocNum
        pqTableArray(r, 2) = pqLocationCol.Cells(r, 1).Value ' Location
    Next r


    ' --- Prepare Output Array ---
    ReDim outputLocationArray(1 To UBound(dashboardDocNumArray, 1), 1 To 1)

    ' --- Perform Lookup using Application.Match on the Array ---
    Dim matchRow As Variant ' Needs to be Variant to handle errors from Match
    For i = 1 To UBound(dashboardDocNumArray, 1)
        docNum = Trim(CStr(dashboardDocNumArray(i, 1)))
        locationResult = "Not Found" ' Default value

        If Len(docNum) > 0 Then
            ' Use Application.Match for potentially faster lookup within the array
            matchRow = Application.Match(docNum, Application.Index(pqTableArray, 0, 1), 0) ' Match in first column of pqTableArray

            If Not IsError(matchRow) Then
                ' Found a match, get the location from the second column of pqTableArray
                locationResult = pqTableArray(matchRow, 2)
            End If
        End If
        outputLocationArray(i, 1) = locationResult
    Next i

    ' --- Write Results Back to Dashboard Column K ---
    ws.Range(DB_COL_LATEST_LOCATION & "4").Resize(UBound(outputLocationArray, 1), 1).Value = outputLocationArray

    LogUserEditsOperation "Successfully populated Latest Location column."
    Debug.Print "PopulateLatestLocation (Array Method) Time: " & Timer - t_start
    Exit Sub ' Normal exit

LocationErrorHandler:
    LogUserEditsOperation "ERROR in PopulateLatestLocation: " & Err.Description
    ' Optionally display a message to the user
    ' MsgBox "An error occurred while updating the 'Latest Location' column.", vbWarning
    ' Resume Next or Exit Sub depending on desired error handling
    Exit Sub

End Sub


'===============================================================================
' CreateSQRCTDashboard - Legacy name for backward compatibility
'===============================================================================
Public Sub CreateSQRCTDashboard()
    RefreshDashboard_TwoWaySync
End Sub

'===============================================================================
' SAVEUSEREDITSFROMDASHBOARD: Captures any user edits from Dashboard -> UserEdits
' Refactored to use Dictionary lookup for performance.
' Adjusted for SHIFTED user columns L-N on Dashboard.
'===============================================================================
Public Sub SaveUserEditsFromDashboard()
    Dim wsSrc As Worksheet, wsEdits As Worksheet
    Dim lastRowSrc As Long, lastRowEdits As Long
    Dim i As Long, destRow As Long
    Dim docNum As String
    Dim hasUserEdits As Boolean
    Dim wasChanged As Boolean
    Dim userEditsDict As Object ' Dictionary for UserEdits lookup
    Dim editRow As Variant      ' To store row number or data from dictionary

    LogUserEditsOperation "Starting SaveUserEditsFromDashboard"

    ' Use GetOrCreateDashboardSheet to ensure wsSrc is valid
    Set wsSrc = GetOrCreateDashboardSheet(DASHBOARD_SHEET_NAME) ' Use Constant and helper function
    If wsSrc Is Nothing Then
        LogUserEditsOperation DASHBOARD_SHEET_NAME & " sheet not found" ' Use Constant
        Exit Sub
    End If

    SetupUserEditsSheet
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Use Constant

    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row

    ' Load UserEdits into dictionary for faster lookup
    Set userEditsDict = LoadUserEditsToDictionary(wsEdits)
    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row ' Get last row based on DocNum column
    If lastRowEdits < 1 Then lastRowEdits = 1

    For i = 4 To lastRowSrc
        docNum = Trim(CStr(wsSrc.Cells(i, "A").Value)) ' col A = Document Number
        If docNum <> "" And docNum <> "Document Number" Then

            ' Check if this row has any user edits (Columns L-N now)
            hasUserEdits = False
            If wsSrc.Cells(i, DB_COL_PHASE).Value <> "" Or _
               wsSrc.Cells(i, DB_COL_LASTCONTACT).Value <> "" Or _
               wsSrc.Cells(i, DB_COL_COMMENTS).Value <> "" Then ' Check new columns L, M, N
                hasUserEdits = True
            End If

            ' Find existing row using dictionary
            If userEditsDict.Exists(docNum) Then
                editRow = userEditsDict(docNum) ' Get existing row number
            Else
                editRow = 0 ' Flag as not found
            End If

            ' Process this document number if:
            ' 1. It has user edits in columns L-N, OR
            ' 2. It already exists in UserEdits (to potentially clear edits)
            If hasUserEdits Or editRow > 0 Then
                ' Determine destination row
                If editRow > 0 Then
                    destRow = editRow
                Else
                    ' Find the next available row dynamically
                    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row + 1
                    If lastRowEdits < 2 Then lastRowEdits = 2 ' Ensure it's at least row 2
                    destRow = lastRowEdits
                    wsEdits.Cells(destRow, UE_COL_DOCNUM).Value = docNum
                    ' Add to dictionary immediately to handle duplicates within the dashboard itself
                    userEditsDict.Add docNum, destRow
                End If

                ' Track if we're making changes to determine if timestamp needs updating
                wasChanged = False

                ' Get current values from dashboard (Columns L, M, N)
                Dim dbPhase, dbLastContact, dbComments
                dbPhase = wsSrc.Cells(i, DB_COL_PHASE).Value         ' Read from L
                dbLastContact = wsSrc.Cells(i, DB_COL_LASTCONTACT).Value ' Read from M
                dbComments = wsSrc.Cells(i, DB_COL_COMMENTS).Value    ' Read from N

                ' Only update UserEdits if either:
                ' 1. This is a new entry (editRow was 0 initially), or
                ' 2. The value in the dashboard is different from what's in UserEdits

                If editRow = 0 Or wsEdits.Cells(destRow, UE_COL_PHASE).Value <> dbPhase Then
                    wsEdits.Cells(destRow, UE_COL_PHASE).Value = dbPhase ' Write to UserEdits B
                    wasChanged = True
                End If

                If editRow = 0 Or wsEdits.Cells(destRow, UE_COL_LASTCONTACT).Value <> dbLastContact Then
                    wsEdits.Cells(destRow, UE_COL_LASTCONTACT).Value = dbLastContact ' Write to UserEdits C
                    wasChanged = True
                End If

                If editRow = 0 Or wsEdits.Cells(destRow, UE_COL_COMMENTS).Value <> dbComments Then
                    wsEdits.Cells(destRow, UE_COL_COMMENTS).Value = dbComments ' Write to UserEdits D
                    wasChanged = True
                End If

                ' Set ChangeSource to workbook identity and update timestamp only if something changed
                If wasChanged Then ' Update timestamp if any field was modified or if it's a new entry with edits
                    wsEdits.Cells(destRow, UE_COL_SOURCE).Value = Module_Identity.GetWorkbookIdentity() ' Write to UserEdits E
                    wsEdits.Cells(destRow, UE_COL_TIMESTAMP).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss") ' Write to UserEdits F
                    LogUserEditsOperation "Updated UserEdits for DocNumber " & docNum & " with attribution " & Module_Identity.GetWorkbookIdentity()
                End If
            End If
        End If
    Next i
    Set userEditsDict = Nothing ' Clean up dictionary

    LogUserEditsOperation "Completed SaveUserEditsFromDashboard"
    Exit Sub

ErrorHandler:
    LogUserEditsOperation "ERROR in SaveUserEditsFromDashboard: " & Err.Description
    Set userEditsDict = Nothing ' Clean up dictionary on error
    Application.EnableEvents = True ' Re-enable events on error exit
End Sub


'===============================================================================
' SETUPUSEREDITSSHEET: Creates or ensures existence of "UserEdits" with standard structure
' Adjusted for removed Email column. No changes needed from previous version.
'===============================================================================
Public Sub SetupUserEditsSheet()
    Dim wsEdits As Worksheet
    Dim needsBackup As Boolean
    Dim wsBackup As Worksheet
    Dim currentHeaders As Variant
    Dim expectedHeaders As Variant
    Dim structureCorrect As Boolean
    Dim backupSuccess As Boolean
    Dim i As Long
    Dim emergencyFlag As Boolean

    LogUserEditsOperation "Starting SetupUserEditsSheet"

    ' Define expected headers for the 6-column structure (A-F)
    expectedHeaders = Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                            "User Comments", "ChangeSource", "Timestamp")

    ' SAFETY CHECK 1: Verify UserEdits sheet exists
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    If Err.Number <> 0 Then
        LogUserEditsOperation "WARNING: Error accessing UserEdits sheet: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    If wsEdits Is Nothing Then
        ' Creating new sheet - no backup needed
        LogUserEditsOperation "UserEdits sheet doesn't exist - creating new"
        Set wsEdits = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsEdits.Name = USEREDITS_SHEET_NAME

        ' Set up headers (A1:F1)
        With wsEdits.Range(wsEdits.Cells(1, UE_COL_DOCNUM), wsEdits.Cells(1, UE_COL_TIMESTAMP)) ' A:F
            .Value = expectedHeaders
            .Font.Bold = True
            .Interior.Color = RGB(16, 107, 193)
            .Font.Color = RGB(255, 255, 255)
        End With
        wsEdits.Visible = xlSheetHidden
        LogUserEditsOperation "Created new UserEdits sheet with standard headers"
        Exit Sub ' Exit early - nothing else to do for new sheet
    End If

    ' SAFETY CHECK 2: Verify header structure
    structureCorrect = False
    needsBackup = False

    On Error Resume Next
    currentHeaders = wsEdits.Range("A1:" & UE_COL_TIMESTAMP & "1").Value ' Check A:F
    If Err.Number <> 0 Then
        LogUserEditsOperation "WARNING: Error reading UserEdits headers: " & Err.Description
        Err.Clear
        needsBackup = True ' Assume we need backup if headers can't be read
    Else
        ' Check if headers match expected structure (6 columns)
        If UBound(currentHeaders, 2) = UBound(expectedHeaders) + 1 Then
            structureCorrect = True ' Tentatively correct
            For i = 0 To UBound(expectedHeaders)
                If CStr(currentHeaders(1, i + 1)) <> expectedHeaders(i) Then
                    structureCorrect = False
                    Exit For
                End If
            Next i
        End If

        needsBackup = Not structureCorrect
    End If
    On Error GoTo ErrorHandler

    ' Early exit if structure is already correct
    If structureCorrect Then
        LogUserEditsOperation "UserEdits sheet structure verified - no changes needed"
        Exit Sub
    End If

    ' SAFETY CHECK 3: Create backup BEFORE any modifications
    backupSuccess = False
    LogUserEditsOperation "UserEdits structure needs update - creating backup"

    On Error Resume Next
    ' Try first backup name
    Set wsBackup = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME & "_Backup")
    If wsBackup Is Nothing Then
        Set wsBackup = ThisWorkbook.Sheets.Add(After:=wsEdits)
        If Err.Number <> 0 Then
            LogUserEditsOperation "WARNING: Failed to create primary backup: " & Err.Description
            Err.Clear
        Else
            wsBackup.Name = USEREDITS_SHEET_NAME & "_Backup"
            If Err.Number <> 0 Then
                LogUserEditsOperation "WARNING: Failed to name primary backup: " & Err.Description
                Err.Clear
            End If
        End If
    Else
        wsBackup.Cells.Clear
    End If

    ' CRITICAL: If first backup attempt failed, try emergency backup with timestamp
    If wsBackup Is Nothing Then
        emergencyFlag = True
        Set wsBackup = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        If Err.Number <> 0 Then
            LogUserEditsOperation "CRITICAL: Emergency backup sheet creation failed: " & Err.Description
            Err.Clear
        Else
            wsBackup.Name = "EMERGENCY_" & USEREDITS_SHEET_NAME & "_" & Format(Now, "yyyymmdd_hhmmss")
            If Err.Number <> 0 Then
                LogUserEditsOperation "CRITICAL: Emergency backup naming failed: " & Err.Description
                Err.Clear
            End If
        End If
    End If
    On Error GoTo ErrorHandler

    ' CRITICAL SAFETY: Verify backup exists before proceeding
    If wsBackup Is Nothing Then
        LogUserEditsOperation "CRITICAL: All backup attempts failed - aborting to prevent data loss"
        MsgBox "WARNING: Could not create backup of UserEdits. No changes will be made to avoid data loss.", _
               vbExclamation, "UserEdits Protection"
        Exit Sub
    End If

    ' Copy data to backup
    On Error Resume Next
    wsEdits.UsedRange.Copy wsBackup.Range("A1")
    If Err.Number <> 0 Then
        LogUserEditsOperation "CRITICAL: Failed to copy data to backup: " & Err.Description
        Err.Clear
        Set wsBackup = Nothing ' Mark backup as invalid
    Else
        wsBackup.Visible = xlSheetHidden
        backupSuccess = True
        If emergencyFlag Then
            LogUserEditsOperation "Created emergency backup: " & wsBackup.Name
        Else
            LogUserEditsOperation "Created standard backup: " & wsBackup.Name
        End If
    End If
    On Error GoTo ErrorHandler

    ' CRITICAL SAFETY: Final verification that backup succeeded
    If Not backupSuccess Or wsBackup Is Nothing Then
        LogUserEditsOperation "CRITICAL: Backup verification failed - aborting to prevent data loss"
        MsgBox "WARNING: Backup verification failed. No changes will be made to avoid data loss.", _
               vbExclamation, "UserEdits Protection"
        Exit Sub
    End If

    ' CRITICAL: ONLY NOW is it safe to clear the sheet
    On Error Resume Next
    wsEdits.Cells.Clear
    If Err.Number <> 0 Then
        LogUserEditsOperation "ERROR clearing UserEdits: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' Set up new headers (A:F)
    With wsEdits.Range(wsEdits.Cells(1, UE_COL_DOCNUM), wsEdits.Cells(1, UE_COL_TIMESTAMP)) ' A:F
        .Value = expectedHeaders
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Migrate data from backup
    Dim lastRowBackup As Long
    lastRowBackup = wsBackup.Cells(wsBackup.Rows.Count, "A").End(xlUp).Row

    If lastRowBackup > 1 Then
        Dim oldColPhase As Long, oldColLastContact As Long, oldColEmail As Long
        Dim oldColComments As Long, oldColSource As Long, oldColTimestamp As Long
        Dim h As Long, headerText As String

        ' Find old column indices by header text in backup
        For h = 1 To wsBackup.UsedRange.Columns.Count
            headerText = CStr(wsBackup.Cells(1, h).Value)
            Select Case headerText
                Case "UserStageOverride", "EngagementPhase", "Engagement Phase": oldColPhase = h
                Case "LastContactDate", "Last Contact Date": oldColLastContact = h
                Case "EmailContact", "Email Contact": oldColEmail = h ' Find but don't use
                Case "UserComments", "User Comments": oldColComments = h
                Case "ChangeSource": oldColSource = h
                Case "Timestamp": oldColTimestamp = h
            End Select
        Next h

        ' Migrate data row by row
        For i = 2 To lastRowBackup
            ' Only migrate if there's a document number
            If Not IsEmpty(wsBackup.Cells(i, 1).Value) And wsBackup.Cells(i, 1).Value <> "" Then
                ' Copy DocNumber (always column A)
                wsEdits.Cells(i, UE_COL_DOCNUM).Value = wsBackup.Cells(i, 1).Value ' Write to A

                ' Map remaining fields safely
                If oldColPhase > 0 Then wsEdits.Cells(i, UE_COL_PHASE).Value = wsBackup.Cells(i, oldColPhase).Value Else wsEdits.Cells(i, UE_COL_PHASE).Value = vbNullString ' Write to B
                If oldColLastContact > 0 Then wsEdits.Cells(i, UE_COL_LASTCONTACT).Value = wsBackup.Cells(i, oldColLastContact).Value Else wsEdits.Cells(i, UE_COL_LASTCONTACT).Value = vbNullString ' Write to C
                If oldColComments > 0 Then wsEdits.Cells(i, UE_COL_COMMENTS).Value = wsBackup.Cells(i, oldColComments).Value Else wsEdits.Cells(i, UE_COL_COMMENTS).Value = vbNullString ' Write to D
                If oldColSource > 0 Then wsEdits.Cells(i, UE_COL_SOURCE).Value = wsBackup.Cells(i, oldColSource).Value Else wsEdits.Cells(i, UE_COL_SOURCE).Value = Module_Identity.GetWorkbookIdentity() ' Write to E
                If oldColTimestamp > 0 Then wsEdits.Cells(i, UE_COL_TIMESTAMP).Value = wsBackup.Cells(i, oldColTimestamp).Value Else wsEdits.Cells(i, UE_COL_TIMESTAMP).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss") ' Write to F
            End If
        Next i

        LogUserEditsOperation "Successfully migrated " & (lastRowBackup - 1) & " records from backup to new structure"
    End If

    LogUserEditsOperation "UserEdits sheet structure update completed successfully"
    Exit Sub

ErrorHandler:
    LogUserEditsOperation "CRITICAL ERROR in SetupUserEditsSheet: " & Err.Description
    MsgBox "Error updating UserEdits structure: " & Err.Description & vbCrLf & _
           "Your data may have been preserved in a backup sheet.", vbExclamation, "Structure Update Error"
End Sub


'===============================================================================
' GETORCREATEDASHBOARDSHEET: Returns or creates the SQRCT Dashboard
'===============================================================================
Private Function GetOrCreateDashboardSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName) ' Use Constant DASHBOARD_SHEET_NAME
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
        ' Optionally set up row 1 & 2 (title, timestamp)
        SetupDashboard ws
    End If

    Set GetOrCreateDashboardSheet = ws
End Function

'===============================================================================
' CLEANUPDASHBOARDLAYOUT: Clean up duplicate rows while preserving the core structure
' Adjusted for new column structure A:N
'===============================================================================
Private Sub CleanupDashboardLayout(ws As Worksheet)
    Application.ScreenUpdating = False

    ' Step 1: Save data from row 4 onward (Range A:N)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim dataRange As Range
    Dim tempData As Variant

    If lastRow >= 4 Then
        ' Capture all data below row 3
        Set dataRange = ws.Range("A4:" & DB_COL_COMMENTS & lastRow) ' Use Comments constant (N)
        tempData = dataRange.Value
    End If

    ' Step 2: Find rows to preserve (rows 1-3)
    Dim hasTitle As Boolean
    hasTitle = False

    ' Check for title text in each cell of row 1 individually (Range A:N)
    Dim cell As Range
    For Each cell In ws.Range("A1:" & DB_COL_COMMENTS & "1").Cells ' Use Comments constant (N)
        If InStr(1, CStr(cell.Value), "STRATEGIC QUOTE RECOVERY", vbTextCompare) > 0 Then
            hasTitle = True
            Exit For
        End If
    Next cell

    ' Step 3: Clear the entire sheet EXCEPT rows 1-3 (Range A:N)
    ws.Range("A4:" & DB_COL_COMMENTS & ws.Rows.Count).Clear ' Use Comments constant (N)

    ' Step 4: If the title row (row 1) is missing, recreate it (Range A:N)
    If Not hasTitle Then
        With ws.Range("A1:" & DB_COL_COMMENTS & "1") ' Use Comments constant (N)
            .Merge
            .Value = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 18
            .Font.Bold = True
            .Interior.Color = RGB(16, 107, 193)  ' Slightly more vibrant blue
            .Font.Color = RGB(255, 255, 255)
            .RowHeight = 32  ' Slightly taller for a more spacious look
        End With
    End If

    ' Step 5: Ensure row 2 has control panel with professional styling (Range A:N)
    With ws.Range("A2:" & DB_COL_COMMENTS & "2") ' Use Comments constant (N)
        .Interior.Color = RGB(245, 245, 245)  ' Very light gray
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).Color = RGB(200, 200, 200)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
        .RowHeight = 28  ' Taller for better button spacing
    End With

    With ws.Range("A2")
        .Value = "CONTROL PANEL"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Name = "Segoe UI"  ' More modern font if available
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(70, 130, 180)  ' Steel blue - professional but distinct
        .Font.Color = RGB(255, 255, 255)
        .ColumnWidth = 16  ' Slightly wider for better proportions
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(200, 200, 200)
    End With

    ' Question mark in corner for help (Column N)
    With ws.Range(DB_COL_COMMENTS & "2") ' Use Comments constant (N)
        .Value = "?"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(70, 130, 180)  ' Matching steel blue
    End With

    ' Step 6: Ensure row 3 has column headers with improved styling (Range A:N)
    With ws.Range("A3:" & DB_COL_COMMENTS & "3") ' Use Comments constant (N)
        .Clear
        .Value = Array( _
            "Document Number", "Client ID", "Customer Name", "Document Amount", "Document Date", _
            "First Date Pulled", "Salesperson ID", "Entered By", "Pull Count", "Missing Quote Alert", _
            "Latest Location", "Engagement Phase", "Last Contact Date", "User Comments") ' Added Latest Location, shifted others
        ' Headers correspond to columns:
        ' K - DB_COL_LATEST_LOCATION (NEW)
        ' L - DB_COL_PHASE (Shifted)
        ' M - DB_COL_LASTCONTACT (Shifted)
        ' N - DB_COL_COMMENTS (Shifted)
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)  ' Match title row color
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' Step 7: Restore data if we had any (Range A:N)
    If Not IsEmpty(tempData) Then
        Dim numCols As Long
        On Error Resume Next
        numCols = UBound(tempData, 2)
        On Error GoTo 0 ' Or appropriate error handler

        If numCols = 13 Then ' Previous structure (A:M)
            Dim restoredData As Variant
            Dim r As Long, c As Long, targetCol As Long
            ReDim restoredData(1 To UBound(tempData, 1), 1 To 14) ' Resize to new structure (A:N)
            For r = 1 To UBound(tempData, 1)
                targetCol = 1
                For c = 1 To 13 ' Loop through old columns
                    If targetCol = 11 Then targetCol = 12 ' Skip new column K (index 11) when writing
                    restoredData(r, targetCol) = tempData(r, c)
                    targetCol = targetCol + 1
                Next c
                ' Column K (index 11) in restoredData will be empty
            Next r
            ws.Range("A4").Resize(UBound(restoredData, 1), UBound(restoredData, 2)).Value = restoredData
        ElseIf numCols = 14 Then ' Already new structure (A:N)
             ws.Range("A4").Resize(UBound(tempData, 1), UBound(tempData, 2)).Value = tempData
        Else ' Unexpected number of columns
            ' Handle error or log warning
        End If
    End If

    Application.ScreenUpdating = True
End Sub


'===============================================================================
' INITIALIZEDASHBOARDLAYOUT: Clears rows 4+ in A-N, sets up header row in A3:N3
' Adjusted for NEW Latest Location column K, shifted user columns L-N
'===============================================================================
Private Sub InitializeDashboardLayout(ws As Worksheet)
    ' Only clear rows 4+ to preserve header/control panel (Range A:N)
    ws.Range("A4:" & DB_COL_COMMENTS & ws.Rows.Count).Clear ' Use Comments constant (N)

    ' Delete extra columns O+ if needed
    On Error Resume Next
    ws.Range("O:" & ws.Columns.Count).Delete ' Delete columns beyond N
    On Error GoTo 0

    ' Ensure row 3 has correct headers with improved styling (Range A:N)
    With ws.Range("A3:" & DB_COL_COMMENTS & "3") ' Use Comments constant (N)
        .Clear
        .Value = Array( _
            "Document Number", "Client ID", "Customer Name", "Document Amount", "Document Date", _
            "First Date Pulled", "Salesperson ID", "Entered By", "Pull Count", "Missing Quote Alert", _
            "Workflow Location", "Engagement Phase", "Last Contact Date", "User Comments") ' Added Latest Location, shifted others
        ' Headers correspond to columns:
        ' K - DB_COL_LATEST_LOCATION (NEW)
        ' L - DB_COL_PHASE (Shifted)
        ' M - DB_COL_LASTCONTACT (Shifted)
        ' N - DB_COL_COMMENTS (Shifted)
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)  ' Match title row color
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' Set initial column widths (Added K, adjusted L-N)
    With ws
        .Columns("A").ColumnWidth = 15
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 25
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 12
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 15
        .Columns("J").ColumnWidth = 20
        .Columns(DB_COL_LATEST_LOCATION).ColumnWidth = 20 ' K (NEW)
        .Columns(DB_COL_PHASE).ColumnWidth = 20          ' L (was K)
        .Columns(DB_COL_LASTCONTACT).ColumnWidth = 15    ' M (was L)
        .Columns(DB_COL_COMMENTS).ColumnWidth = 40       ' N (was M)
    End With
End Sub


'===============================================================================
' POPULATEMASTERQUOTESDATA: Pulls columns A-J from MasterQuotes_Final
' No changes needed here as it only populates A-J
'===============================================================================
Private Sub PopulateMasterQuotesData(ws As Worksheet)
    ' Use constant for the source name
    Dim sourceName As String
    sourceName = MASTER_QUOTES_FINAL_SOURCE

    ' Check if source exists before attempting formulas
    If Not IsMasterQuotesFinalPresent Then Exit Sub ' Exit if source not found

    Dim lastMasterRow As Long
    On Error Resume Next ' Handle error if source is empty or invalid
    lastMasterRow = Application.WorksheetFunction.CountA(Range(sourceName & "[Document Number]"))
    If Err.Number <> 0 Or lastMasterRow = 0 Then
        Debug.Print "MasterQuotes_Final source is empty or invalid. Cannot populate A-J."
        Exit Sub ' Exit if source is empty
    End If
    On Error GoTo 0 ' Restore error handling

    Dim targetRowCount As Long
    targetRowCount = lastMasterRow ' Number of data rows to populate

    With ws
        ' A: Document Number
        .Range("A4").Resize(targetRowCount, 1).Formula = _
            "=IF(ROWS($A$4:A4)<=ROWS(" & sourceName & "[Document Number])," & _
            "IFERROR(INDEX(" & sourceName & "[Document Number],ROWS($A$4:A4)),""""),"""")"

        ' B: Client ID -> from Customer Number
        .Range("B4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Customer Number],ROWS($A$4:A4)),""""),"""")"

        ' C: Customer Name
        .Range("C4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Customer Name],ROWS($A$4:A4)),""""),"""")"

        ' D: Document Amount
        .Range("D4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[Document Amount],ROWS($A$4:A4)),""""),"""")"

        ' E: Document Date
        .Range("E4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[Document Date],ROWS($A$4:A4)),""""),"""")"

        ' F: First Date Pulled
        .Range("F4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[First Date Pulled],ROWS($A$4:A4)),""""),"""")"

        ' G: Salesperson ID
        .Range("G4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Salesperson ID],ROWS($A$4:A4)),""""),"""")"

        ' H: Entered By (was User To Enter)
        .Range("H4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[User To Enter],ROWS($A$4:A4)),""""),"""")"

        ' I: Pull Count (was Occurrence Counter / AutoStage)
        .Range("I4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Pull Count],ROWS($A$4:A4)),""""),"""")" ' Changed from AutoStage to Pull Count

        ' J: Missing Quote Alert (was Auto Note)
        .Range("J4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[AutoNote],ROWS($A$4:A4)),""""),"""")"

        ' Convert formulas to values for performance (optional but recommended for large datasets)
        '.Range("A4:J" & 3 + targetRowCount).Value = .Range("A4:J" & 3 + targetRowCount).Value

        ' Format numeric/date columns
        .Range("D4:D" & 3 + targetRowCount).NumberFormat = "$#,##0.00"   ' Document Amount
        .Range("E4:E" & 3 + targetRowCount).NumberFormat = "mm/dd/yyyy" ' Document Date
        .Range("F4:F" & 3 + targetRowCount).NumberFormat = "mm/dd/yyyy" ' First Date Pulled
    End With
End Sub


'===============================================================================
' SORTDASHBOARDDATA: Sort by First Date Pulled (F asc), then Document Amount (D desc)
' Adjusted range to A:N
'===============================================================================
Private Sub SortDashboardData(ws As Worksheet, lastRow As Long)
    If lastRow < 5 Then Exit Sub ' Need at least one data row + header

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("F4:F" & lastRow), _
                          SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range("D4:D" & lastRow), _
                          SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange ws.Range("A3:" & DB_COL_COMMENTS & lastRow) ' Use Comments constant (N) for full range
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'===============================================================================
' FREEZEDASHBOARD: Freezes rows 1-3
'===============================================================================
Private Sub FreezeDashboard(ws As Worksheet)
    ws.Activate
    ActiveWindow.FreezePanes = False ' Unfreeze first
    ws.Range("A4").Select          ' Select cell below freeze row
    ActiveWindow.FreezePanes = True ' Freeze above selected cell
End Sub

'===============================================================================
' SETUPDASHBOARD: Professional row 1 & 2 design (title & control panel)
' Adjusted ranges to A:N
'===============================================================================
Public Sub SetupDashboard(ws As Worksheet)
    ' Merge & style title in row 1 (Range A:N)
    With ws.Range("A1:" & DB_COL_COMMENTS & "1") ' Use Comments constant (N)
        .Merge
        .Value = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 18
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)  ' Slightly more vibrant blue
        .Font.Color = RGB(255, 255, 255)
        .RowHeight = 32  ' Slightly taller for a more spacious look
    End With

    ' Set up control panel in row 2 with modern styling (Range A:N)
    With ws.Range("A2:" & DB_COL_COMMENTS & "2") ' Use Comments constant (N)
        .Interior.Color = RGB(245, 245, 245)  ' Very light gray
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).Color = RGB(200, 200, 200)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
        .RowHeight = 28  ' Taller for better button spacing
    End With

    ' Add control panel label with professional styling
    With ws.Range("A2")
        .Value = "CONTROL PANEL"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Name = "Segoe UI"  ' More modern font if available
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(70, 130, 180)  ' Steel blue - professional but distinct
        .Font.Color = RGB(255, 255, 255)
        .ColumnWidth = 16  ' Slightly wider for better proportions
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(200, 200, 200)
    End With

    ' Question mark in corner for help (Column N)
    With ws.Range(DB_COL_COMMENTS & "2") ' Use Comments constant (N)
        .Value = "?"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(70, 130, 180)  ' Matching steel blue
    End With

    ' Last refreshed timestamp - positioned and styled more elegantly
    With ws.Range("G2:I2") ' Position remains G:I
        .Merge
        .Value = "Last Refreshed: " & Format$(Now(), "mm/dd/yyyy h:mm") & " MST"
        .HorizontalAlignment = xlCenter
        .Font.Size = 9
        .Font.Name = "Segoe UI"
        .Font.Color = RGB(80, 80, 80)  ' Dark gray for subtle elegance
    End With

    ' Create buttons with improved spacing and styling (Positions remain C, E)
    ModernButton ws, "C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits"
    ModernButton ws, "E2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits"

End Sub


'===============================================================================
' MODERNBUTTON: Creates professional, modern-looking buttons with proper spacing
'===============================================================================
Public Sub ModernButton(ws As Worksheet, cellRef As String, buttonText As String, macroName As String)
    Dim btn As Object ' Changed from Shape to Object for broader compatibility
    Dim buttonTop As Double, buttonLeft As Double
    Dim buttonWidth As Double, buttonHeight As Double
    Dim targetCell As Range

    On Error Resume Next ' Handle cases where cellRef might be invalid
    Set targetCell = ws.Range(cellRef)
    If targetCell Is Nothing Then Exit Sub ' Exit if cellRef is invalid
    On Error GoTo 0

    ' Get the position and size based on the cell
    buttonLeft = targetCell.Left
    buttonTop = targetCell.Top

    ' Calculate precise width and position to prevent overlap
    buttonWidth = targetCell.Width * 1.6
    buttonHeight = targetCell.Height * 0.75

    ' Better centering with more space between buttons
    buttonTop = buttonTop + (targetCell.Height - buttonHeight) / 2

    ' Create the button using Shapes.AddShape (msoShapeRoundedRectangle = 5)
    On Error Resume Next
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
    If Err.Number <> 0 Then
        Debug.Print "Error creating button shape: " & Err.Description
        Exit Sub ' Exit if shape creation fails
    End If
    On Error GoTo 0

    ' Style the button with modern, professional appearance
    With btn
        ' Modern gradient effect using solid color
        .Fill.ForeColor.RGB = RGB(42, 120, 180)  ' Professional blue
        .Fill.Visible = msoTrue
        .Fill.Solid

        ' More subtle border
        .Line.ForeColor.RGB = RGB(25, 95, 150)
        .Line.Weight = 0.75
        .Line.Visible = msoTrue

        ' Set the text with improved font
        On Error Resume Next ' Handle potential errors setting text properties
        .TextFrame2.TextRange.Text = buttonText
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Name = "Segoe UI"
        .TextFrame2.TextRange.Font.Bold = msoTrue

        ' Text positioning using TextFrame2 for better control
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.WordWrap = msoFalse ' Prevent text wrapping

        ' Add a subtle shadow effect if possible
        .Shadow.Type = msoShadow21
        .Shadow.Transparency = 0.7
        .Shadow.Visible = msoTrue
        On Error GoTo 0 ' Restore error handling

        ' Assign macro
        .OnAction = macroName
    End With
End Sub


'===============================================================================
' PROTECTUSERCOLUMNS: Lock A-K, unlock L-N
' Adjusted for NEW Latest Location column K, shifted user columns L-N
'===============================================================================
Public Sub ProtectUserColumns(ws As Worksheet)
    ' Assumes sheet is already unprotected by the calling procedure
    On Error GoTo 0 ' Resume default error handling

    ws.Cells.Locked = True
    ' Unlock columns L:N (Phase, LastContact, Comments) using constants
    ws.Range(DB_COL_PHASE & "4:" & DB_COL_COMMENTS & ws.Rows.Count).Locked = False ' Use Comments constant (N)

    ' Re-apply protection here after setting Locked status
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub


'===============================================================================
' APPLYCOLORFORMATTING: For coloring column L (Engagement Phase)
' Adjusted target column to L (DB_COL_PHASE)
' Added startDataRow parameter for flexibility (e.g., Text-Only sheet)
'===============================================================================
Public Sub ApplyColorFormatting(ws As Worksheet, Optional startDataRow As Long = 4)
    ' Assumes sheet is unprotected by the calling procedure
    On Error GoTo 0 ' Resume default error handling

    Dim endRow As Long
    endRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Use actual last row in Col A
    If endRow < startDataRow Then Exit Sub ' No data rows to format

    ' Define the range for applying formatting (Column L)
    Dim rngPhase As Range
    Set rngPhase = ws.Range(DB_COL_PHASE & startDataRow & ":" & DB_COL_PHASE & endRow) ' Column L

    ' Apply conditional formatting ONLY to the Engagement Phase (L) column.
    ' The ApplyStageFormatting helper sub will clear existing rules from rngPhase first.
    ApplyStageFormatting rngPhase

    ' Protection is no longer handled here
End Sub


' Helper for detailed color rules implementing the evidence-based color system
' Applies formatting to targetRng based on the corresponding value in the Engagement Phase column (L).
Private Sub ApplyStageFormatting(targetRng As Range)
    If targetRng Is Nothing Then Exit Sub
    If targetRng.Cells.Count = 0 Then Exit Sub

    Dim formulaBase As String
    Dim firstCellAddress As String

    ' Get address of the first cell in the target range (e.g., L4)
    firstCellAddress = targetRng.Cells(1).Address(RowAbsolute:=False, ColumnAbsolute:=True) ' e.g., $L4

    ' Formula now checks the value in the cell itself
    formulaBase = "=EXACT(" & firstCellAddress & ",""{PHASE}"")"

    ' Clear existing rules first to ensure a clean slate
    targetRng.FormatConditions.Delete

    With targetRng
        ' --- Follow-up Stages (Sequential Process) ---
        ' First F/U: Light blue (#D0E6F5)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "First F/U")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(208, 230, 245)

        ' Second F/U: Medium blue (#92C6ED)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Second F/U")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(146, 198, 237)

        ' Third F/U: Yellow (#F5E171)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Third F/U")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(245, 225, 113)

        ' Long-Term F/U: Orange (#FF9636)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Long-Term F/U")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 150, 54)

        ' --- Queue/Processing Status ---
        ' Requoting: Light lavender (#E3D7E8)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Requoting")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(227, 215, 232)

        ' Pending: Pale yellow (#FFF7D1)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Pending")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 247, 209)

        ' No Response: Beige (#F5EEE0)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "No Response")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(245, 238, 224)

        ' Texas (No F/U): Tan (#E6D9CC)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Texas (No F/U)")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(230, 217, 204)

        ' --- Team Member Assignments ---
        ' AF: Teal (#A2D9D2)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "AF")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(162, 217, 210)

        ' RZ: Periwinkle (#8A9BD4)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "RZ")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(138, 155, 212)

        ' KMH: Salmon (#F7C4AF)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "KMH")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(247, 196, 175)

        ' RI: Sky blue (#BFE1F3)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "RI")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(191, 225, 243)

        ' WW/OM: Deep purple (#9B7CB9)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "WW/OM")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(155, 124, 185)

        ' --- Outcome Statuses ---
        ' Converted: Green (Adjusted x6) (Bold)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Converted")
        With .FormatConditions(.FormatConditions.Count)
            .Interior.Color = RGB(120, 235, 120) ' Final final final green adjustment
            .Font.Bold = True
        End With

        ' Declined: True red (#D12F2F)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Declined")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(209, 47, 47)

        ' Closed (Extra Order): Medium-dark red (#B82727)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Closed (Extra Order)")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(184, 39, 39)

        ' Closed: Dark red (#A61C1C)
        .FormatConditions.Add Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Closed")
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(166, 28, 28)

    End With
End Sub


'===============================================================================
' ISMASTERQUOTESFINALPRESENT: Checks for PQ, Table, or Named Range "MasterQuotes_Final"
'===============================================================================
Public Function IsMasterQuotesFinalPresent() As Boolean
    Dim lo As ListObject
    Dim nm As Name
    Dim queryObj As Object ' Changed from WorkbookQuery for broader compatibility

    IsMasterQuotesFinalPresent = False
    On Error Resume Next ' Ignore errors if queries collection doesn't exist or other errors

    ' 1) Power Query named MASTER_QUOTES_FINAL_SOURCE
    Set queryObj = Nothing ' Reset before check
    Err.Clear
    Set queryObj = ActiveWorkbook.Queries(MASTER_QUOTES_FINAL_SOURCE) ' Use Constant
    If Err.Number = 0 And Not queryObj Is Nothing Then
        IsMasterQuotesFinalPresent = True
    End If
    Err.Clear

    ' 2) ListObject named MASTER_QUOTES_FINAL_SOURCE
    If Not IsMasterQuotesFinalPresent Then
        Set lo = Nothing ' Reset before check
        Err.Clear
        Set lo = ActiveSheet.ListObjects(MASTER_QUOTES_FINAL_SOURCE) ' Check ActiveSheet first for performance
        If Err.Number <> 0 Then ' If not on active sheet, check all sheets
            Dim ws As Worksheet
            For Each ws In ActiveWorkbook.Worksheets
                Err.Clear
                Set lo = ws.ListObjects(MASTER_QUOTES_FINAL_SOURCE) ' Use Constant
                If Err.Number = 0 And Not lo Is Nothing Then Exit For
            Next ws
        End If
        If Not lo Is Nothing Then IsMasterQuotesFinalPresent = True
        Err.Clear
    End If

    ' 3) Named Range MASTER_QUOTES_FINAL_SOURCE
    If Not IsMasterQuotesFinalPresent Then
        Set nm = Nothing ' Reset before check
        Err.Clear
        Set nm = ActiveWorkbook.Names(MASTER_QUOTES_FINAL_SOURCE) ' Use Constant
        If Err.Number = 0 And Not nm Is Nothing Then
            IsMasterQuotesFinalPresent = True
        End If
        Err.Clear
    End If

    On Error GoTo 0 ' Restore default error handling
End Function


'===============================================================================
' LOADUSEREDITSTODICTIONARY: Loads UserEdits DocNumbers and SHEET row numbers into a dictionary
'===============================================================================
Public Function LoadUserEditsToDictionary(wsEdits As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive keys

    If wsEdits Is Nothing Then
        Set LoadUserEditsToDictionary = dict
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row

    If lastRow > 1 Then ' Check if there's data beyond the header row
        Dim i As Long
        Dim docNum As String
        Dim dataRange As Variant

        On Error Resume Next ' Handle potential errors reading range
        dataRange = wsEdits.Range(UE_COL_DOCNUM & "2:" & UE_COL_DOCNUM & lastRow).Value
        If Err.Number <> 0 Then
             LogUserEditsOperation "Error reading UserEdits DocNum column: " & Err.Description
             Set LoadUserEditsToDictionary = dict ' Return empty dictionary
             Exit Function
        End If
        On Error GoTo 0 ' Restore error handling

        If Not IsArray(dataRange) Then ' Handle single data row case
            If lastRow = 2 Then
                docNum = Trim(CStr(dataRange))
                If docNum <> "" Then
                    If Not dict.Exists(docNum) Then dict.Add docNum, 2 ' Store SHEET row number
                End If
            End If
        Else ' Process the 2D array
            For i = 1 To UBound(dataRange, 1)
                docNum = Trim(CStr(dataRange(i, 1)))
                If docNum <> "" Then
                    If Not dict.Exists(docNum) Then
                        dict.Add docNum, i + 1 ' Store SHEET row number (i+1 because array starts from sheet row 2)
                    End If
                End If
            Next i
        End If
    End If

    Set LoadUserEditsToDictionary = dict
End Function






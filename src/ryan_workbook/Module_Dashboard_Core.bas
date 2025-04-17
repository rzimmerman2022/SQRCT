Attribute VB_Name = "Module_Dashboard_Core"
Option Explicit

'===============================================================================
' MODULE_DASHBOARD_CORE
' Contains core functions for managing the SQRCT Dashboard refresh operations,
' data population, formatting, and UI interactions.
'===============================================================================

' --- Public Constants ---
Public Const DASHBOARD_SHEET_NAME As String = "SQRCT Dashboard"
Public Const USEREDITS_SHEET_NAME As String = "UserEdits" ' Referenced by UserEdits module
Public Const USEREDITSLOG_SHEET_NAME As String = "UserEditsLog" ' Referenced by UserEdits module
Public Const MASTER_QUOTES_FINAL_SOURCE As String = "MasterQuotes_Final" ' Name of the PQ query/table for A-I

' Power Query Output for Workflow Location
Public Const PQ_LATEST_LOCATION_SHEET As String = "DocNum_LatestLocation" ' Sheet where DocNum_LatestLocation loads
Public Const PQ_LATEST_LOCATION_TABLE As String = "DocNum_LatestLocation" ' Assumes table name matches query name

' UserEdits Columns (Referenced by UserEdits module)
Public Const UE_COL_DOCNUM As String = "A"
Public Const UE_COL_PHASE As String = "B"
Public Const UE_COL_LASTCONTACT As String = "C"
Public Const UE_COL_COMMENTS As String = "D"
Public Const UE_COL_SOURCE As String = "E"
Public Const UE_COL_TIMESTAMP As String = "F"

' Dashboard Columns (UPDATED LAYOUT)
' A-I populated by MasterQuotes_Final
Public Const DB_COL_DOCNUM As String = "A"
Public Const DB_COL_CLIENTID As String = "B" ' From Customer Number
Public Const DB_COL_CUSTNAME As String = "C"
Public Const DB_COL_DOCAMT As String = "D"
Public Const DB_COL_DOCDATE As String = "E"
Public Const DB_COL_FIRSTPULL As String = "F"
Public Const DB_COL_SALESID As String = "G"
Public Const DB_COL_ENTEREDBY As String = "H" ' From User To Enter
Public Const DB_COL_PULLCOUNT As String = "I" ' From Pull Count
' J-N populated/managed by VBA
Public Const DB_COL_WORKFLOW_LOCATION As String = "J" ' NEW - Populated by PopulateWorkflowLocation
Public Const DB_COL_MISSING_QUOTE_ALERT As String = "K" ' NEW - Static Formula
Public Const DB_COL_PHASE As String = "L"         ' Shifted from original K
Public Const DB_COL_LASTCONTACT As String = "M"     ' Shifted from original L
Public Const DB_COL_COMMENTS As String = "N"        ' Shifted from original M
' --- End Constants ---


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
' MAIN REFRESH SUB: Creates or refreshes the SQRCT Dashboard
' Incorporates fixes for column layout, workflow location timing/method,
' column K formula, column widths, and button loop.
' Calls UserEdits module for saving/loading/logging.
'===============================================================================
Public Sub RefreshDashboard(Optional PreserveUserEdits As Boolean = False)
    Dim ws As Worksheet, wsEdits As Worksheet
    Dim lastRow As Long, lastRowEdits As Long
    Dim backupCreated As Boolean
    Dim t_start As Single, t_location As Single, t_format As Single
    Dim userEditsDict As Object
    Dim dashboardDocNumArray As Variant, userEditsDataArray As Variant, outputEditsArray As Variant
    Dim numDashboardRows As Long

    ' 1) Backup & log (Call UserEdits module)
    backupCreated = Module_Dashboard_UserEdits.CreateUserEditsBackup("RefreshDashboard_" & Format(Now, "yyyymmdd_hhmmss"))
    Module_Dashboard_UserEdits.LogUserEditsOperation "Starting dashboard refresh. PreserveUserEdits=" & PreserveUserEdits & ", Backup created: " & backupCreated
    t_start = Timer

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 2) Ensure UserEdits sheet (Call UserEdits module)
    Module_Dashboard_UserEdits.SetupUserEditsSheet
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)

    ' 3) Save edits if needed (Call UserEdits module)
    If Not PreserveUserEdits Then Module_Dashboard_UserEdits.SaveUserEditsFromDashboard

    ' 4) Get or create Dashboard
    Set ws = GetOrCreateDashboardSheet(DASHBOARD_SHEET_NAME)
    If ws Is Nothing Then GoTo Cleanup ' Exit if sheet couldn't be created/found
    On Error Resume Next: ws.Unprotect: On Error GoTo ErrorHandler ' Unprotect for modifications

    ' 5) Clean and init layout (A:N)
    CleanupDashboardLayout ws
    InitializeDashboardLayout ws

    ' 6) Populate A-I from MasterQuotes_Final
    If IsMasterQuotesFinalPresent Then
        PopulateMasterQuotesData ws ' Populates A-I
    Else
        MsgBox "Warning: " & MASTER_QUOTES_FINAL_SOURCE & " not found. Cannot populate base data.", vbInformation, "Data Source Not Found"
        ' Continue to allow workflow/user edits population if possible
    End If

    lastRow = ws.Cells(ws.Rows.Count, DB_COL_DOCNUM).End(xlUp).Row ' Last row based on Col A

    ' 7) Sort data (A:N) based on F (asc) then D (desc)
    If lastRow >= 4 Then SortDashboardData ws, lastRow

    ' --- Data Population AFTER Sort ---
    If lastRow >= 4 Then
        ' 8) Populate Workflow Location (Column J) AFTER sort using Dictionary lookup
        t_location = Timer
        PopulateWorkflowLocation ws, lastRow ' NEW METHOD
        Debug.Print "PopulateWorkflowLocation Time: " & Timer - t_location

        ' 9) Populate Missing Quote Alert (Column K) with static formula AFTER sort
        ws.Range(DB_COL_MISSING_QUOTE_ALERT & "4:" & DB_COL_MISSING_QUOTE_ALERT & lastRow).Formula = _
            "=IF(" & DB_COL_DOCNUM & "4<>"""",""Confirm Converted/Voided"","""")"
        ' Optional: Convert formula to values if preferred
        ' ws.Range(DB_COL_MISSING_QUOTE_ALERT & "4:" & DB_COL_MISSING_QUOTE_ALERT & lastRow).Value = ws.Range(DB_COL_MISSING_QUOTE_ALERT & "4:" & DB_COL_MISSING_QUOTE_ALERT & lastRow).Value
    End If

    ' 10) Autofit & Set Column Widths (Adjusted)
    With ws
        .Columns("A:I").AutoFit ' Autofit PQ columns
        .Columns(DB_COL_WORKFLOW_LOCATION).ColumnWidth = 25 ' J: Set width
        .Columns(DB_COL_MISSING_QUOTE_ALERT).AutoFit ' K: Autofit
        If .Columns(DB_COL_MISSING_QUOTE_ALERT).ColumnWidth < 25 Then .Columns(DB_COL_MISSING_QUOTE_ALERT).ColumnWidth = 25 ' K: Min width 25
        .Columns(DB_COL_PHASE).AutoFit ' L: Autofit
        .Columns(DB_COL_LASTCONTACT).AutoFit ' M: Autofit
        .Columns(DB_COL_COMMENTS).ColumnWidth = 40 ' N: Fixed width 40
        ' Adjust specific PQ columns if needed after autofit
        If .Columns(DB_COL_CUSTNAME).ColumnWidth > 40 Then .Columns(DB_COL_CUSTNAME).ColumnWidth = 40 ' Limit Customer Name width
        If .Columns(DB_COL_DOCAMT).ColumnWidth < 12 Then .Columns(DB_COL_DOCAMT).ColumnWidth = 12 ' Min width for Amount
    End With

    ' 11) Restore user edits (L:N) via arrays
    If lastRow >= 4 Then
        lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row
        On Error Resume Next ' Handle empty dashboard range
        dashboardDocNumArray = ws.Range(DB_COL_DOCNUM & "4:" & DB_COL_DOCNUM & lastRow).Value
        Dim dashboardReadError As Boolean: dashboardReadError = (Err.Number <> 0)
        On Error GoTo ErrorHandler
        If dashboardReadError Or Not IsArray(dashboardDocNumArray) Then
             Module_Dashboard_UserEdits.LogUserEditsOperation "Warning: Could not read dashboard DocNums for restoring edits."
        Else
            If lastRowEdits > 1 Then
                Dim userEditsRange As Range, singleRowData As Variant, cIdx As Long
                Set userEditsRange = wsEdits.Range(UE_COL_DOCNUM & "2:" & UE_COL_TIMESTAMP & lastRowEdits) ' A:F
                On Error Resume Next ' Handle single row case
                If userEditsRange.Rows.Count = 1 Then
                    ReDim singleRowData(1 To 1, 1 To 6) ' A-F
                    For cIdx = 1 To 6: singleRowData(1, cIdx) = userEditsRange.Cells(1, cIdx).Value: Next cIdx
                    userEditsDataArray = singleRowData
                Else
                    userEditsDataArray = userEditsRange.Value
                End If
                Dim userEditsReadError As Boolean: userEditsReadError = (Err.Number <> 0)
                On Error GoTo ErrorHandler
                If userEditsReadError Then Module_Dashboard_UserEdits.LogUserEditsOperation "Warning: Could not read UserEdits data array."
            End If

            Set userEditsDict = Module_Dashboard_UserEdits.LoadUserEditsToDictionary(wsEdits) ' Load {DocNum: RowNum}

            numDashboardRows = UBound(dashboardDocNumArray, 1)
            ReDim outputEditsArray(1 To numDashboardRows, 1 To 3) ' For columns L, M, N

            Dim i As Long, j As Long, editSheetRow As Long, editArrayRow As Long, ubUED As Long
            ' Initialize output array
            For i = 1 To numDashboardRows: For j = 1 To 3: outputEditsArray(i, j) = vbNullString: Next j: Next i

            If Not IsEmpty(userEditsDataArray) And lastRowEdits > 1 And Not userEditsReadError Then
                ubUED = UBound(userEditsDataArray, 1)
                For i = 1 To numDashboardRows
                    Dim docNum As String
                    docNum = Trim(CStr(dashboardDocNumArray(i, 1)))
                    If docNum <> "" Then
                        If userEditsDict.Exists(docNum) Then
                            editSheetRow = userEditsDict(docNum) ' Get sheet row number from UserEdits
                            editArrayRow = editSheetRow - 1 ' Adjust for 1-based array read from row 2
                            If editArrayRow >= 1 And editArrayRow <= ubUED Then
                                On Error Resume Next ' Handle potential errors reading specific array elements
                                outputEditsArray(i, 1) = userEditsDataArray(editArrayRow, 2) ' Col B (Phase) -> Output Col 1 (L)
                                outputEditsArray(i, 2) = userEditsDataArray(editArrayRow, 3) ' Col C (LastContact) -> Output Col 2 (M)
                                outputEditsArray(i, 3) = userEditsDataArray(editArrayRow, 4) ' Col D (Comments) -> Output Col 3 (N)
                                On Error GoTo ErrorHandler
                            End If
                        End If
                    End If
                Next i
            End If

            ' Write the restored edits to columns L:N
            ws.Range(DB_COL_PHASE & "4").Resize(numDashboardRows, 3).Value = outputEditsArray
            Module_Dashboard_UserEdits.LogUserEditsOperation "Restored user edits to dashboard columns L:N."
        End If
    End If ' End check if lastRow >= 4 for restoring edits

    ' 12) Freeze panes
    FreezeDashboard ws

    ' 13) Formatting & protection
    t_format = Timer
    ApplyColorFormatting ws ' Applies to Col L
    ProtectUserColumns ws ' Locks A-K, Unlocks L-N
    Debug.Print "Format/Protect Time: " & Timer - t_format

    ' 14) Timestamp
    With ws.Range("G2:I2") ' Position remains G:I
        .Merge
        .Value = "Last Refreshed: " & Format$(Now, "mm/dd/yyyy h:mm") & " MST"
        .HorizontalAlignment = xlCenter
        .Font.Size = 9: .Font.Name = "Segoe UI"
        .Font.Color = RGB(80, 80, 80)
    End With

    ' 15) Buttons (Using simple loop workaround)
    On Error Resume Next ' Ignore errors if shape doesn't exist or isn't deletable
    Dim shp As Shape
    For Each shp In ws.Shapes
        ' Delete shapes anchored anywhere in row 2
        If shp.TopLeftCell.Row = 2 Then shp.Delete
    Next shp
    On Error GoTo ErrorHandler ' Restore error handling

    ModernButton ws, "C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits"
    ModernButton ws, "E2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits"

    ' 16) Final message
    Dim msgText As String
    If PreserveUserEdits Then
        msgText = DASHBOARD_SHEET_NAME & " refreshed!" & vbCrLf & USEREDITS_SHEET_NAME & " preserved."
    Else
        msgText = DASHBOARD_SHEET_NAME & " refreshed!" & vbCrLf & "Edits saved & restored."
    End If
    
    MsgBox msgText, vbInformation, "Dashboard Refresh Complete"

    Module_Dashboard_UserEdits.LogUserEditsOperation "Dashboard refresh completed successfully. Total time: " & Timer - t_start & "s"

Cleanup:
    ' Ensure sheet is protected on exit/error
    If Not ws Is Nothing Then On Error Resume Next: ws.Protect UserInterfaceOnly:=True: On Error GoTo 0
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Set ws = Nothing: Set wsEdits = Nothing: Set userEditsDict = Nothing
    If IsArray(dashboardDocNumArray) Then Erase dashboardDocNumArray
    If IsArray(userEditsDataArray) Then Erase userEditsDataArray
    If IsArray(outputEditsArray) Then Erase outputEditsArray
    Exit Sub

ErrorHandler:
    Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR in RefreshDashboard: [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")"
    MsgBox "Error during refresh: " & Err.Description, vbCritical, "Dashboard Refresh Error"
    ' Attempt restore only if backup was confirmed created
    If backupCreated Then Module_Dashboard_UserEdits.RestoreUserEditsFromBackup
    Resume Cleanup ' Go to cleanup routine after logging/restoring
End Sub


'===============================================================================
' GETORCREATEDASHBOARDSHEET: Returns or creates the SQRCT Dashboard sheet
'===============================================================================
Public Function GetOrCreateDashboardSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error GoTo CreateSheetError ' Handle error during creation
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1)) ' Add as first sheet
        ws.Name = sheetName
        Module_Dashboard_UserEdits.LogUserEditsOperation "Created new dashboard sheet: " & sheetName
        ' Setup initial title/control panel rows
        SetupDashboard ws
        On Error GoTo 0 ' Resume default error handling
    End If

    Set GetOrCreateDashboardSheet = ws
    Exit Function

CreateSheetError:
    Module_Dashboard_UserEdits.LogUserEditsOperation "FATAL ERROR: Could not create dashboard sheet '" & sheetName & "'. Error: " & Err.Description
    MsgBox "Fatal Error: Could not create the required dashboard sheet '" & sheetName & "'.", vbCritical, "Sheet Creation Failed"
    Set GetOrCreateDashboardSheet = Nothing ' Return Nothing on failure
End Function

'===============================================================================
' CLEANUPDASHBOARDLAYOUT: Clears data rows, ensures header rows 1-3 are correct.
' Adjusted for A:N layout.
'===============================================================================
Private Sub CleanupDashboardLayout(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    On Error Resume Next: ws.Unprotect: On Error GoTo 0 ' Ensure unprotected

    ' Clear data area (Row 4 downwards, Columns A:N)
    ws.Range("A4:" & DB_COL_COMMENTS & ws.Rows.Count).Clear

    ' Verify/Recreate Row 1: Title (A:N)
    Dim titleCell As Range: Set titleCell = ws.Range("A1")
    If InStr(1, CStr(titleCell.Value), "STRATEGIC QUOTE RECOVERY", vbTextCompare) = 0 Or titleCell.MergeCells = False Or titleCell.MergeArea.Address <> ws.Range("A1:" & DB_COL_COMMENTS & "1").Address Then
        With ws.Range("A1:" & DB_COL_COMMENTS & "1")
            .UnMerge
            .ClearContents
            .Merge
            .Value = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER"
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Font.Size = 18: .Font.Bold = True
            .Interior.Color = RGB(16, 107, 193): .Font.Color = RGB(255, 255, 255)
            .RowHeight = 32
        End With
    End If

    ' Verify/Recreate Row 2: Control Panel (A:N)
    With ws.Range("A2:" & DB_COL_COMMENTS & "2")
        .Interior.Color = RGB(245, 245, 245)
        With .Borders(xlEdgeTop): .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 200, 200): End With
        With .Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 200, 200): End With
        .RowHeight = 28
    End With
    With ws.Range("A2")
        .Value = "CONTROL PANEL"
        .Font.Bold = True: .Font.Size = 10: .Font.Name = "Segoe UI"
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = RGB(70, 130, 180): .Font.Color = RGB(255, 255, 255)
        .ColumnWidth = 16
        With .Borders(xlEdgeRight): .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 200, 200): End With
    End With
    With ws.Range(DB_COL_COMMENTS & "2") ' Help '?' in Col N
        .Value = "?"
        .Font.Bold = True: .Font.Size = 14: .HorizontalAlignment = xlCenter
        .Font.Color = RGB(70, 130, 180)
    End With
    ' Clear any old buttons/timestamp in row 2 before recreating
    ws.Range("B2:M2").ClearContents ' Clear area between label and help
    On Error Resume Next
    Dim shp As Shape: For Each shp In ws.Shapes: If shp.TopLeftCell.Row = 2 Then shp.Delete: Next shp
    On Error GoTo 0

    ' Verify/Recreate Row 3: Headers (A:N)
    Dim expectedHeaders As Variant
    expectedHeaders = Array( _
        "Document Number", "Client ID", "Customer Name", "Document Amount", "Document Date", _
        "First Date Pulled", "Salesperson ID", "Entered By", "Pull Count", _
        "Workflow Location", "Missing Quote Alert", "Engagement Phase", "Last Contact Date", "User Comments") ' Updated J, K
    Dim currentHeaders As Variant, headersMatch As Boolean: headersMatch = True
    On Error Resume Next
    currentHeaders = ws.Range("A3:" & DB_COL_COMMENTS & "3").Value
    If Err.Number <> 0 Or Not IsArray(currentHeaders) Then headersMatch = False Else
        If UBound(currentHeaders, 2) <> UBound(expectedHeaders) + 1 Then headersMatch = False Else
            Dim h As Long: For h = 0 To UBound(expectedHeaders): If CStr(currentHeaders(1, h + 1)) <> expectedHeaders(h) Then headersMatch = False: Exit For: End If: Next h
        End If
    End If
    On Error GoTo 0

    If Not headersMatch Then
        With ws.Range("A3:" & DB_COL_COMMENTS & "3")
            .ClearContents
            .Value = expectedHeaders
            .Font.Bold = True: .Interior.Color = RGB(16, 107, 193)
            .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlCenter
        End With
    End If

    ' Protection applied later by RefreshDashboard
    Application.ScreenUpdating = True
End Sub


'===============================================================================
' INITIALIZEDASHBOARDLAYOUT: Clears rows 4+, sets up header row A3:N3
' Adjusted for A:N layout.
'===============================================================================
Private Sub InitializeDashboardLayout(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next: ws.Unprotect: On Error GoTo 0 ' Ensure unprotected

    ' Clear data area (Row 4 downwards, Columns A:N)
    ws.Range("A4:" & DB_COL_COMMENTS & ws.Rows.Count).Clear

    ' Delete extra columns O+ if they exist
    On Error Resume Next
    If ws.Columns.Count > 14 Then ws.Range(ws.Columns(15), ws.Columns(ws.Columns.Count)).Delete
    On Error GoTo 0

    ' Ensure row 3 has correct headers (A:N)
    Dim expectedHeaders As Variant
    expectedHeaders = Array( _
        "Document Number", "Client ID", "Customer Name", "Document Amount", "Document Date", _
        "First Date Pulled", "Salesperson ID", "Entered By", "Pull Count", _
        "Workflow Location", "Missing Quote Alert", "Engagement Phase", "Last Contact Date", "User Comments") ' Updated J, K
    With ws.Range("A3:" & DB_COL_COMMENTS & "3")
        .ClearContents
        .Value = expectedHeaders
        .Font.Bold = True: .Interior.Color = RGB(16, 107, 193)
        .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlCenter
    End With

    ' Set initial column widths (A:N)
    With ws
        .Columns(DB_COL_DOCNUM).ColumnWidth = 15     ' A
        .Columns(DB_COL_CLIENTID).ColumnWidth = 15    ' B
        .Columns(DB_COL_CUSTNAME).ColumnWidth = 30    ' C
        .Columns(DB_COL_DOCAMT).ColumnWidth = 14      ' D
        .Columns(DB_COL_DOCDATE).ColumnWidth = 12     ' E
        .Columns(DB_COL_FIRSTPULL).ColumnWidth = 15   ' F
        .Columns(DB_COL_SALESID).ColumnWidth = 12     ' G
        .Columns(DB_COL_ENTEREDBY).ColumnWidth = 15   ' H
        .Columns(DB_COL_PULLCOUNT).ColumnWidth = 10   ' I
        .Columns(DB_COL_WORKFLOW_LOCATION).ColumnWidth = 25 ' J (NEW)
        .Columns(DB_COL_MISSING_QUOTE_ALERT).ColumnWidth = 25 ' K (NEW)
        .Columns(DB_COL_PHASE).ColumnWidth = 20          ' L (Shifted)
        .Columns(DB_COL_LASTCONTACT).ColumnWidth = 15    ' M (Shifted)
        .Columns(DB_COL_COMMENTS).ColumnWidth = 40       ' N (Shifted)
    End With
    ' Protection applied later by RefreshDashboard
End Sub


'===============================================================================
' POPULATEMASTERQUOTESDATA: Pulls columns A-I from MasterQuotes_Final
' Adjusted to only populate A-I. Column J (Workflow) and K (Alert) handled separately.
'===============================================================================
Private Sub PopulateMasterQuotesData(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim sourceName As String: sourceName = MASTER_QUOTES_FINAL_SOURCE
    If Not IsMasterQuotesFinalPresent Then Exit Sub ' Exit if source not found

    Dim lastMasterRow As Long, targetRowCount As Long
    On Error Resume Next ' Handle error if source is empty or invalid
    lastMasterRow = Application.WorksheetFunction.CountA(Range(sourceName & "[Document Number]"))
    If Err.Number <> 0 Or lastMasterRow = 0 Then
        Module_Dashboard_UserEdits.LogUserEditsOperation "Warning: " & sourceName & " source is empty or invalid. Cannot populate A-I."
        Exit Sub
    End If
    On Error GoTo 0
    targetRowCount = lastMasterRow

    With ws
        ' A: Document Number
        .Range(DB_COL_DOCNUM & "4").Resize(targetRowCount, 1).Formula = _
            "=IF(ROWS($A$4:A4)<=ROWS(" & sourceName & "[Document Number])," & _
            "IFERROR(INDEX(" & sourceName & "[Document Number],ROWS($A$4:A4)),""""),"""")"
        ' B: Client ID (from Customer Number)
        .Range(DB_COL_CLIENTID & "4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Customer Number],ROWS($A$4:A4)),""""),"""")"
        ' C: Customer Name
        .Range(DB_COL_CUSTNAME & "4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Customer Name],ROWS($A$4:A4)),""""),"""")"
        ' D: Document Amount
        .Range(DB_COL_DOCAMT & "4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[Document Amount],ROWS($A$4:A4)),""""),"""")"
        ' E: Document Date
        .Range(DB_COL_DOCDATE & "4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[Document Date],ROWS($A$4:A4)),""""),"""")"
        ' F: First Date Pulled
        .Range(DB_COL_FIRSTPULL & "4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[First Date Pulled],ROWS($A$4:A4)),""""),"""")"
        ' G: Salesperson ID
        .Range(DB_COL_SALESID & "4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Salesperson ID],ROWS($A$4:A4)),""""),"""")"
        ' H: Entered By (from User To Enter)
        .Range(DB_COL_ENTEREDBY & "4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[User To Enter],ROWS($A$4:A4)),""""),"""")"
        ' I: Pull Count
        .Range(DB_COL_PULLCOUNT & "4").Resize(targetRowCount, 1).Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Pull Count],ROWS($A$4:A4)),""""),"""")"

        ' Convert formulas A-I to values for performance
        Dim dataRange As Range: Set dataRange = .Range("A4:" & DB_COL_PULLCOUNT & (3 + targetRowCount))
        dataRange.Value = dataRange.Value

        ' Format numeric/date columns (A-I)
        .Range(DB_COL_DOCAMT & "4:" & DB_COL_DOCAMT & (3 + targetRowCount)).NumberFormat = "$#,##0.00"   ' D
        .Range(DB_COL_DOCDATE & "4:" & DB_COL_DOCDATE & (3 + targetRowCount)).NumberFormat = "mm/dd/yyyy" ' E
        .Range(DB_COL_FIRSTPULL & "4:" & DB_COL_FIRSTPULL & (3 + targetRowCount)).NumberFormat = "mm/dd/yyyy" ' F
    End With
    Module_Dashboard_UserEdits.LogUserEditsOperation "Populated dashboard columns A-I from " & sourceName
End Sub


'===============================================================================
' SORTDASHBOARDDATA: Sort by First Date Pulled (F asc), then Document Amount (D desc)
' Sorts the full data range A:N.
'===============================================================================
Private Sub SortDashboardData(ws As Worksheet, lastRow As Long)
    If ws Is Nothing Or lastRow < 4 Then Exit Sub ' Need header + data

    Module_Dashboard_UserEdits.LogUserEditsOperation "Sorting dashboard data A3:" & DB_COL_COMMENTS & lastRow
    With ws.Sort
        .SortFields.Clear
        ' Sort by First Date Pulled (F) Ascending
        .SortFields.Add Key:=ws.Range(DB_COL_FIRSTPULL & "4:" & DB_COL_FIRSTPULL & lastRow), _
                          SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ' Then by Document Amount (D) Descending
        .SortFields.Add Key:=ws.Range(DB_COL_DOCAMT & "4:" & DB_COL_DOCAMT & lastRow), _
                          SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        ' Set the full range including headers and all columns (A:N)
        .SetRange ws.Range("A3:" & DB_COL_COMMENTS & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


'===============================================================================
' POPULATEWORKFLOWLOCATION: Populates Column J based on DocNum_LatestLocation PQ.
' REWRITTEN to use Scripting.Dictionary for lookup AFTER the main data sort.
'===============================================================================
Private Sub PopulateWorkflowLocation(ws As Worksheet, lastRow As Long)
    If ws Is Nothing Or lastRow < 4 Then Exit Sub

    Dim wsPQ As Worksheet
    Dim tblPQ As ListObject
    Dim pqDataRange As Range
    Dim pqDocNumColIndex As Long, pqLocationColIndex As Long
    Dim pqTableArray As Variant
    Dim locationDict As Object ' Scripting.Dictionary
    Dim dashboardDocNumArray As Variant
    Dim outputLocationArray As Variant
    Dim i As Long, r As Long
    Dim docNum As String, locationResult As String
    Dim t_start As Single: t_start = Timer
    Const DEFAULT_LOCATION As String = "Quote Only" ' Default if not found in PQ

    Module_Dashboard_UserEdits.LogUserEditsOperation "Starting PopulateWorkflowLocation (Dictionary Method)"
    On Error GoTo LocationErrorHandler

    ' --- Get References to PQ Output Sheet and Table ---
    Set wsPQ = Nothing: On Error Resume Next
    Set wsPQ = ThisWorkbook.Sheets(PQ_LATEST_LOCATION_SHEET): On Error GoTo LocationErrorHandler
    If wsPQ Is Nothing Then Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR: PQ Sheet '" & PQ_LATEST_LOCATION_SHEET & "' not found.": Exit Sub

    Set tblPQ = Nothing: On Error Resume Next
    Set tblPQ = wsPQ.ListObjects(PQ_LATEST_LOCATION_TABLE)
    If tblPQ Is Nothing Then Set tblPQ = wsPQ.ListObjects(1) ' Fallback
    On Error GoTo LocationErrorHandler
    If tblPQ Is Nothing Then Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR: PQ Table '" & PQ_LATEST_LOCATION_TABLE & "' not found.": Exit Sub

    ' --- Verify Required Columns Exist and Get Data ---
    On Error Resume Next
    pqDocNumColIndex = tblPQ.ListColumns("PrimaryDocNumber").Index
    pqLocationColIndex = tblPQ.ListColumns("MostRecent_FolderLocation").Index
    On Error GoTo LocationErrorHandler
    If pqDocNumColIndex = 0 Or pqLocationColIndex = 0 Then Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR: Required columns not found in PQ table '" & tblPQ.Name & "'." : Exit Sub

    If tblPQ.DataBodyRange Is Nothing Then Module_Dashboard_UserEdits.LogUserEditsOperation "Warning: PQ Table '" & tblPQ.Name & "' is empty." : Exit Sub ' Exit if no data rows

    ' Read only the necessary columns from PQ into an array
    Dim numPQRows As Long: numPQRows = tblPQ.DataBodyRange.Rows.Count
    ReDim pqTableArray(1 To numPQRows, 1 To 2)
    Dim pqDocCol As Range: Set pqDocCol = tblPQ.ListColumns(pqDocNumColIndex).DataBodyRange
    Dim pqLocCol As Range: Set pqLocCol = tblPQ.ListColumns(pqLocationColIndex).DataBodyRange
    For r = 1 To numPQRows
        pqTableArray(r, 1) = pqDocCol.Cells(r, 1).Value ' DocNum
        pqTableArray(r, 2) = pqLocCol.Cells(r, 1).Value ' Location
    Next r

    ' --- Build the Dictionary (Case-Insensitive) ---
    Set locationDict = CreateObject("Scripting.Dictionary")
    locationDict.CompareMode = vbTextCompare ' Case-insensitive keys

    For r = 1 To numPQRows
        docNum = Trim(CStr(pqTableArray(r, 1)))
        locationResult = Trim(CStr(pqTableArray(r, 2)))
        If docNum <> "" Then
            If Not locationDict.Exists(docNum) Then
                locationDict.Add docNum, locationResult
            Else
                ' Optional: Handle duplicates in PQ source if necessary (e.g., log warning)
                ' Currently, first one encountered wins due to dictionary behavior
            End If
        End If
    Next r
    Erase pqTableArray ' Free memory

    ' --- Read Dashboard Document Numbers (Column A, rows 4 to lastRow) ---
    On Error Resume Next ' Handle empty dashboard range
    dashboardDocNumArray = ws.Range(DB_COL_DOCNUM & "4:" & DB_COL_DOCNUM & lastRow).Value
    Dim dashboardReadError As Boolean: dashboardReadError = (Err.Number <> 0)
    On Error GoTo LocationErrorHandler
    If dashboardReadError Or Not IsArray(dashboardDocNumArray) Then Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR: Could not read dashboard DocNums for location lookup." : Exit Sub

    ' --- Prepare Output Array ---
    Dim numDashboardRows As Long: numDashboardRows = UBound(dashboardDocNumArray, 1)
    ReDim outputLocationArray(1 To numDashboardRows, 1 To 1)

    ' --- Perform Lookup using Dictionary ---
    For i = 1 To numDashboardRows
        docNum = Trim(CStr(dashboardDocNumArray(i, 1)))
        If docNum <> "" And locationDict.Exists(docNum) Then
            outputLocationArray(i, 1) = locationDict(docNum) ' Get location from dictionary
        ElseIf docNum <> "" Then
            outputLocationArray(i, 1) = DEFAULT_LOCATION ' Use default if not found
        Else
            outputLocationArray(i, 1) = "" ' Blank if no DocNum on dashboard row
        End If
    Next i
    Erase dashboardDocNumArray ' Free memory

    ' --- Write Results Back to Dashboard Column J ---
    ws.Range(DB_COL_WORKFLOW_LOCATION & "4").Resize(numDashboardRows, 1).Value = outputLocationArray

    Module_Dashboard_UserEdits.LogUserEditsOperation "Successfully populated Workflow Location (Col J). Time: " & Timer - t_start & "s"
    Set locationDict = Nothing ' Clean up
    Exit Sub ' Normal exit

LocationErrorHandler:
    Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR in PopulateWorkflowLocation: [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")"
    Set locationDict = Nothing ' Clean up
    ' Optionally display a message to the user
    ' MsgBox "An error occurred while updating the 'Workflow Location' column.", vbWarning
    Exit Sub
End Sub


'===============================================================================
' FREEZEDASHBOARD: Freezes rows 1-3
'===============================================================================
Private Sub FreezeDashboard(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Activate ' Required for ActiveWindow operations
    ActiveWindow.FreezePanes = False ' Unfreeze first
    ws.Range("A4").Select          ' Select cell below freeze row
    ActiveWindow.FreezePanes = True ' Freeze above selected cell
    ws.Range("A1").Select ' Select A1 after freezing
End Sub

'===============================================================================
' SETUPDASHBOARD: Sets up rows 1 & 2 (title & control panel)
' Called by GetOrCreateDashboardSheet. Ensures A:N layout.
'===============================================================================
Public Sub SetupDashboard(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    On Error Resume Next: ws.Unprotect: On Error GoTo 0 ' Ensure unprotected

    ' Row 1: Title (A:N)
    With ws.Range("A1:" & DB_COL_COMMENTS & "1")
        If .MergeCells = False Or .MergeArea.Address <> ws.Range("A1:" & DB_COL_COMMENTS & "1").Address Then .UnMerge
        .ClearContents: .Merge
        .Value = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER"
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Font.Size = 18: .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193): .Font.Color = RGB(255, 255, 255)
        .RowHeight = 32
    End With

    ' Row 2: Control Panel (A:N)
    With ws.Range("A2:" & DB_COL_COMMENTS & "2")
        .Interior.Color = RGB(245, 245, 245)
        With .Borders(xlEdgeTop): .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 200, 200): End With
        With .Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 200, 200): End With
        .RowHeight = 28
    End With
    With ws.Range("A2") ' Label
        .Value = "CONTROL PANEL"
        .Font.Bold = True: .Font.Size = 10: .Font.Name = "Segoe UI"
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = RGB(70, 130, 180): .Font.Color = RGB(255, 255, 255)
        .ColumnWidth = 16
        With .Borders(xlEdgeRight): .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 200, 200): End With
    End With
    With ws.Range(DB_COL_COMMENTS & "2") ' Help '?' in Col N
        .Value = "?"
        .Font.Bold = True: .Font.Size = 14: .HorizontalAlignment = xlCenter
        .Font.Color = RGB(70, 130, 180)
    End With
    With ws.Range("G2:I2") ' Timestamp
        .Merge
        .Value = "Last Refreshed: " & Format$(Now(), "mm/dd/yyyy h:mm") & " MST"
        .HorizontalAlignment = xlCenter: .Font.Size = 9: .Font.Name = "Segoe UI"
        .Font.Color = RGB(80, 80, 80)
    End With

    ' Create buttons (will be done by RefreshDashboard after clearing)
    ' ModernButton ws, "C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits"
    ' ModernButton ws, "E2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits"

    ' Protection applied later by RefreshDashboard
    Application.ScreenUpdating = True
End Sub


'===============================================================================
' MODERNBUTTON: Creates professional, modern-looking buttons
'===============================================================================
Public Sub ModernButton(ws As Worksheet, cellRef As String, buttonText As String, macroName As String)
    If ws Is Nothing Then Exit Sub
    Dim btn As Button ' Use Button object type
    Dim targetCell As Range
    Dim btnLeft As Double, btnTop As Double, btnWidth As Double, btnHeight As Double

    On Error Resume Next: Set targetCell = ws.Range(cellRef): On Error GoTo 0
    If targetCell Is Nothing Then Exit Sub ' Invalid cell reference

    ' Calculate position and size based on cell, with padding
    btnLeft = targetCell.Left + 2
    btnTop = targetCell.Top + 2
    btnWidth = targetCell.Width * 1.6 - 4 ' Adjust width, leave padding
    btnHeight = targetCell.Height - 4    ' Adjust height, leave padding

    On Error Resume Next ' Handle potential errors during button creation/modification
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    If btn Is Nothing Then
        Module_Dashboard_UserEdits.LogUserEditsOperation "Error: Failed to create button in " & cellRef
        Exit Sub
    End If

    With btn
        .Caption = buttonText
        .OnAction = macroName
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        ' Note: Standard buttons don't have easy gradient/shadow options like Shapes.
        ' For more complex styling, Shapes.AddShape would be needed, but Buttons are simpler.
    End With
    On Error GoTo 0
End Sub


'===============================================================================
' PROTECTUSERCOLUMNS: Lock A-K, unlock L-N
' Adjusted for A:N layout. Called by RefreshDashboard AFTER all modifications.
'===============================================================================
Public Sub ProtectUserColumns(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next ' Ignore error if already protected/unprotected

    ' Ensure sheet is unprotected before changing lock status
    ws.Unprotect

    ' Lock all cells by default
    ws.Cells.Locked = True

    ' Unlock the user-editable range (L4:N down to last row or max rows)
    Dim lastRowProt As Long
    lastRowProt = ws.Cells(ws.Rows.Count, DB_COL_DOCNUM).End(xlUp).Row
    If lastRowProt < 4 Then lastRowProt = 4 ' Ensure at least row 4 is included
    ws.Range(DB_COL_PHASE & "4:" & DB_COL_COMMENTS & lastRowProt).Locked = False

    ' Re-apply protection
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True

    On Error GoTo 0 ' Resume default error handling
End Sub


'===============================================================================
' APPLYCOLORFORMATTING: Applies conditional formatting to Engagement Phase (Col L)
'===============================================================================
Public Sub ApplyColorFormatting(ws As Worksheet, Optional startDataRow As Long = 4)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next: ws.Unprotect: On Error GoTo 0 ' Ensure unprotected

    Dim endRow As Long
    endRow = ws.Cells(ws.Rows.Count, DB_COL_DOCNUM).End(xlUp).Row ' Use actual last row in Col A
    If endRow < startDataRow Then Exit Sub ' No data rows to format

    ' Define the range for applying formatting (Column L)
    Dim rngPhase As Range
    Set rngPhase = ws.Range(DB_COL_PHASE & startDataRow & ":" & DB_COL_PHASE & endRow)

    ' Apply the formatting rules
    ApplyStageFormatting rngPhase

    ' Protection is reapplied by the main RefreshDashboard sub
    Module_Dashboard_UserEdits.LogUserEditsOperation "Applied conditional formatting to " & rngPhase.Address
End Sub


'===============================================================================
' APPLYSTAGEFORMATTING: Helper for detailed color rules for Engagement Phase (Col L)
' Applies formatting rules directly to the provided targetRng (expected to be Col L).
'===============================================================================
Private Sub ApplyStageFormatting(targetRng As Range)
    If targetRng Is Nothing Then Exit Sub
    If targetRng.Cells.Count = 0 Then Exit Sub

    Dim formulaBase As String
    Dim firstCellAddress As String

    ' Get address of the first cell in the target range (e.g., L4) relative for row, absolute for column
    firstCellAddress = targetRng.Cells(1).Address(RowAbsolute:=False, ColumnAbsolute:=True) ' e.g., $L4

    ' Formula checks the value in the cell itself
    formulaBase = "=EXACT(" & firstCellAddress & ",""{PHASE}"")"

    ' Clear existing rules from the target range first
    targetRng.FormatConditions.Delete

    With targetRng.FormatConditions
        ' --- Follow-up Stages ---
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "First F/U"))
            .Interior.Color = RGB(208, 230, 245) ' Light blue
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Second F/U"))
            .Interior.Color = RGB(146, 198, 237) ' Medium blue
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Third F/U"))
            .Interior.Color = RGB(245, 225, 113) ' Yellow
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Long-Term F/U"))
            .Interior.Color = RGB(255, 150, 54)  ' Orange
        End With

        ' --- Queue/Processing Status ---
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Requoting"))
            .Interior.Color = RGB(227, 215, 232) ' Light lavender
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Pending"))
            .Interior.Color = RGB(255, 247, 209) ' Pale yellow
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "No Response"))
            .Interior.Color = RGB(245, 238, 224) ' Beige
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Texas (No F/U)"))
            .Interior.Color = RGB(230, 217, 204) ' Tan
        End With

        ' --- Team Member Assignments ---
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "AF"))
            .Interior.Color = RGB(162, 217, 210) ' Teal
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "RZ"))
            .Interior.Color = RGB(138, 155, 212) ' Periwinkle
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "KMH"))
            .Interior.Color = RGB(247, 196, 175) ' Salmon
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "RI"))
            .Interior.Color = RGB(191, 225, 243) ' Sky blue
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "WW/OM"))
            .Interior.Color = RGB(155, 124, 185) ' Deep purple
        End With

        ' --- Outcome Statuses ---
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Converted"))
            .Interior.Color = RGB(120, 235, 120) ' Green
            .Font.Bold = True
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Declined"))
            .Interior.Color = RGB(209, 47, 47)   ' True red
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Closed (Extra Order)"))
            .Interior.Color = RGB(184, 39, 39)   ' Medium-dark red
        End With
        With .Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Closed"))
            .Interior.Color = RGB(166, 28, 28)   ' Dark red
        End With
    End With
End Sub


'===============================================================================
' ISMASTERQUOTESFINALPRESENT: Checks for PQ, Table, or Named Range
'===============================================================================
Public Function IsMasterQuotesFinalPresent() As Boolean
    Dim lo As ListObject, nm As Name, queryObj As Object ' WorkbookQuery
    Dim ws As Worksheet
    IsMasterQuotesFinalPresent = False
    On Error Resume Next ' Ignore errors during checks

    ' 1) Power Query
    Set queryObj = Nothing: Err.Clear
    Set queryObj = ActiveWorkbook.Queries(MASTER_QUOTES_FINAL_SOURCE)
    If Err.Number = 0 And Not queryObj Is Nothing Then IsMasterQuotesFinalPresent = True: GoTo ExitCheck

    ' 2) ListObject (Table)
    Set lo = Nothing: Err.Clear
    For Each ws In ActiveWorkbook.Worksheets
        Set lo = Nothing: Err.Clear
        Set lo = ws.ListObjects(MASTER_QUOTES_FINAL_SOURCE)
        If Err.Number = 0 And Not lo Is Nothing Then IsMasterQuotesFinalPresent = True: GoTo ExitCheck
    Next ws

    ' 3) Named Range
    Set nm = Nothing: Err.Clear
    Set nm = ActiveWorkbook.Names(MASTER_QUOTES_FINAL_SOURCE)
    If Err.Number = 0 And Not nm Is Nothing Then IsMasterQuotesFinalPresent = True: GoTo ExitCheck

ExitCheck:
    On Error GoTo 0 ' Restore error handling
    If Not IsMasterQuotesFinalPresent Then Module_Dashboard_UserEdits.LogUserEditsOperation "Warning: Data source '" & MASTER_QUOTES_FINAL_SOURCE & "' not found as Query, Table, or Named Range."
End Function

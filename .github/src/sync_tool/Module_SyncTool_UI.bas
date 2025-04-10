Attribute VB_Name = "Module_SyncTool_UI"
Option Explicit

'===============================================================================
' MODULE_SYNCTOOL_UI
' Contains functions for the SyncTool user interface.
' Manages dashboard layout and UI interactions.
'===============================================================================

'===============================================================================
' CREATE_SYNCTOOL_DASHBOARD - Creates or refreshes the SyncTool dashboard.
'
' Purpose:
' Ensures the dashboard exists by attempting to get the sheet defined by
' SYNCTOOL_DASHBOARD_SHEET. If it doesn't exist, a new sheet is created.
' The dashboard layout (column widths, row heights, header, labels, and buttons)
' is then set up.
'
' Returns:
' Nothing.
'===============================================================================
Public Sub CreateSyncToolDashboard()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim btn As Shape
    
    ' Attempt to get the dashboard sheet; create it if it doesn't exist.
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SYNCTOOL_DASHBOARD_SHEET)
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SYNCTOOL_DASHBOARD_SHEET
    End If
    
    On Error GoTo ErrorHandler
    
    ' Clear existing content
    ws.Cells.Clear
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 90 ' Extra wide for file paths
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 20
    ws.Columns("F").ColumnWidth = 20
    
    ' Set row heights
    ws.Rows("3:5").RowHeight = 22 ' For browse buttons
    ws.Rows("8").RowHeight = 30 ' For action buttons
    
    ' Create the blue header
    With ws.Range("A1:F1")
        .Merge
        .value = "SQRCT Sync Tool Dashboard"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 16
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Interior.Color = COLOR_HEADER_BLUE
        .Font.Color = COLOR_HEADER_TEXT
        .RowHeight = 32
    End With
    
    ' Set up file path labels
    ws.Range("A3").value = "Ally's Working File:"
    ws.Range("A3").Font.Bold = True
    ws.Range("A3").Font.Name = "Calibri"
    
    ws.Range("A4").value = "Ryan's Working File:"
    ws.Range("A4").Font.Bold = True
    ws.Range("A4").Font.Name = "Calibri"
    
    ws.Range("A5").value = "Automated Master File:"
    ws.Range("A5").Font.Bold = True
    ws.Range("A5").Font.Name = "Calibri"
    
    ' Status and sync information labels
    ws.Range("A10").value = "Status:"
    ws.Range("A10").Font.Bold = True
    ws.Range("A10").Font.Name = "Calibri"
    
    ws.Range("A11").value = "Last Successful Sync:"
    ws.Range("A11").Font.Bold = True
    ws.Range("A11").Font.Name = "Calibri"
    
    ' Instructions section
    ws.Range("A15").value = "Instructions:"
    ws.Range("A15").Font.Bold = True
    ws.Range("A15").Font.Name = "Calibri"
    
    ws.Range("A16").value = "1. Select Ally's, Ryan's, and Master files using the Browse buttons."
    ws.Range("A16").Font.Name = "Calibri"
    
    ws.Range("A17").value = "2. Click 'Sync' to merge edits from all three files."
    ws.Range("A17").Font.Name = "Calibri"
    
    ws.Range("A18").value = "3. Check the SyncLog sheet for a history of synchronizations."
    ws.Range("A18").Font.Name = "Calibri"
    
    ws.Range("A19").value = "Note: Source tracking uses AF=Ally, RZ=Ryan, MASTER=Master file."
    ws.Range("A19").Font.Name = "Calibri"
    ws.Range("A19").Font.Italic = True
    
    ' Add Action label
    ws.Range("A8").value = "Action:"
    ws.Range("A8").Font.Bold = True
    ws.Range("A8").Font.Name = "Calibri"
    
    ' Create buttons using the helper routine
    CreateDashboardButton ws, "B3", "Browse", "BrowseAllyFile_Click"
    CreateDashboardButton ws, "B4", "Browse", "BrowseRyanFile_Click"
    CreateDashboardButton ws, "B5", "Browse", "BrowseMasterFile_Click"
    CreateDashboardButton ws, "B8", "Sync", "Sync_Click"
    CreateDashboardButton ws, "C8", "Diagnose Files", "DiagnoseFiles_Click"
    CreateDashboardButton ws, "D8", "Show Conflicts", "ShowConflicts_Click"
    CreateDashboardButton ws, "E8", "View Log", "ViewLog_Click"
    
    ' Restore any existing file paths from an older dashboard sheet, if available
    Dim oldSheet As Worksheet
    
    On Error Resume Next
    Set oldSheet = ThisWorkbook.Sheets("SQRCT Dashboard")
    
    If Not oldSheet Is Nothing Then
        If oldSheet.Range("F3").value <> "" Then ws.Range(CELL_ALLY_PATH).value = oldSheet.Range("F3").value
        If oldSheet.Range("F4").value <> "" Then ws.Range(CELL_RYAN_PATH).value = oldSheet.Range("F4").value
        If oldSheet.Range("F5").value <> "" Then ws.Range(CELL_MASTER_PATH).value = oldSheet.Range("F5").value
    End If
    
    On Error GoTo ErrorHandler
    
    ' Activate the dashboard sheet
    ws.Activate
    
    MsgBox "SyncTool Dashboard has been created! All buttons are connected to appropriate functions.", _
           vbInformation, "Dashboard Created"
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "Error in CreateSyncToolDashboard: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "Error in CreateSyncToolDashboard: " & err.Description, vbExclamation, "Dashboard Error"
End Sub

'===============================================================================
' CREATE_DASHBOARD_BUTTON - Creates a consistent modern button on the dashboard.
'
' Parameters:
' ws - The worksheet on which to create the button.
' cellRef - A cell reference (as text) indicating the button's starting position.
' buttonText - The text to display on the button.
' macroName - The name of the macro to assign to the button.
'
' Returns:
' Nothing.
'
' Purpose:
' Creates (or updates) a button shape with consistent styling and positions it
' relative to the cell referenced by cellRef.
'===============================================================================
Private Sub CreateDashboardButton(ws As Worksheet, cellRef As String, buttonText As String, macroName As String)
    Dim btn As Shape
    Dim buttonLeft As Double, buttonTop As Double
    Dim buttonWidth As Double, buttonHeight As Double
    
    ' Validate inputs
    If ws Is Nothing Then Exit Sub
    If Module_Utilities.IsNullOrEmpty(cellRef) Then Exit Sub
    If Module_Utilities.IsNullOrEmpty(buttonText) Then Exit Sub
    If Module_Utilities.IsNullOrEmpty(macroName) Then Exit Sub
    
    ' Calculate button position based on the referenced cell
    buttonLeft = ws.Range(cellRef).Left + 2
    buttonTop = ws.Range(cellRef).Top + 2
    
    ' Determine button width based on the length of the button text
    If Len(buttonText) <= 5 Then
        buttonWidth = 70
    ElseIf Len(buttonText) <= 10 Then
        buttonWidth = 90
    Else
        buttonWidth = 110
    End If
    
    buttonHeight = ws.Range(cellRef).height - 4
    
    ' Build a unique button name using the button text and cell reference
    Dim buttonName As String
    buttonName = "btn" & Replace(buttonText, " ", "") & cellRef
    
    On Error Resume Next
    Set btn = ws.Shapes(buttonName)
    
    If btn Is Nothing Then
        ' Create new button if one doesn't exist
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
        btn.Name = buttonName
    Else
        ' Update existing button's position and size
        btn.Left = buttonLeft
        btn.Top = buttonTop
        btn.width = buttonWidth
        btn.height = buttonHeight
    End If
    
    On Error GoTo 0
    
    ' Apply consistent styling to the button
    With btn
        .Fill.ForeColor.RGB = COLOR_HEADER_BLUE
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = buttonText
        .TextFrame.Characters.Font.Name = "Calibri"
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = COLOR_HEADER_TEXT
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = xlCenter
        .OnAction = macroName
    End With
End Sub

'===============================================================================
' GET_SYNCTOOL_DASHBOARD - Returns a reference to the SyncTool dashboard sheet.
'
' Purpose:
' Retrieves the dashboard sheet based on SYNCTOOL_DASHBOARD_SHEET. If not found,
' it attempts to find an older sheet name or defaults to the active sheet.
'
' Returns:
' A Worksheet object representing the SyncTool dashboard.
'===============================================================================
Public Function GetSyncToolDashboard() As Worksheet
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SYNCTOOL_DASHBOARD_SHEET)
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("SQRCT Dashboard")
    End If
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    On Error GoTo ErrorHandler
    
    Set GetSyncToolDashboard = ws
    
    Exit Function
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "Error in GetSyncToolDashboard: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "Error in GetSyncToolDashboard: " & err.Description, vbExclamation, "Dashboard Error"
    Set GetSyncToolDashboard = Nothing
End Function

'===============================================================================
' UPDATE_STATUS_DISPLAY - Updates the status display on the dashboard.
'
' Parameters:
' statusText - The status message to display.
'
' Purpose:
' Updates both a specific cell on the dashboard and the application's status bar
' with the provided message, prefixed by the current time.
'
' Returns:
' Nothing.
'===============================================================================
Public Sub UpdateStatusDisplay(statusText As String)
    On Error Resume Next
    
    Dim wsDashboard As Worksheet
    Set wsDashboard = GetSyncToolDashboard()
    
    If Not wsDashboard Is Nothing Then
        Dim statusWithTime As String
        statusWithTime = Format$(Now(), "hh:mm:ss") & " - " & statusText
        
        wsDashboard.Range(CELL_STATUS_DISPLAY).value = statusWithTime
    End If
    
    Application.StatusBar = statusText
    
    ' Log the status update using the logger module (explicit reference)
    Module_SyncTool_Logger.LogMessage statusText
    
    ' Allow Excel UI to update
    DoEvents
    
    On Error GoTo 0  ' Restore normal error handling
End Sub


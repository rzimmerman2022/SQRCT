Attribute VB_Name = "Module_Dashboard"
Option Explicit

'===============================================================================
' MODULE_DASHBOARD
' Contains functions for managing the SQRCT Dashboard, including refresh operations,
' user edits management, and UI interactions.
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

' Standard two-way sync: Dashboard ? UserEdits ? Refresh ? UserEdits ? Dashboard
Public Sub RefreshDashboard_TwoWaySync()
    Call RefreshDashboard(PreserveUserEdits:=False)
End Sub

' One-way sync: Refresh ? UserEdits ? Dashboard (preserves manual UserEdits)
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
    Set wsLog = ThisWorkbook.Sheets("UserEditsLog")
    
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "UserEditsLog"
        wsLog.Range("A1:C1").Value = Array("Timestamp", "Workbook", "Operation")
        wsLog.Range("A1:C1").Font.Bold = True
        wsLog.Visible = xlSheetHidden
    End If
    
    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    
    wsLog.Cells(lastRow + 1, "A").Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    wsLog.Cells(lastRow + 1, "B").Value = GetWorkbookIdentity()  ' Using Module_Identity function
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
        backupName = "UserEdits_Backup_" & Format(Now, "yyyymmdd")
    Else
        backupName = "UserEdits_Backup_" & backupSuffix
    End If
    
    ' Create backup only if UserEdits exists and has data
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets("UserEdits")
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
            If InStr(1, ThisWorkbook.Sheets(i).Name, "UserEdits_Backup_") > 0 Then
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
    Set wsEdits = ThisWorkbook.Sheets("UserEdits")
    
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
'===============================================================================
Public Sub RefreshDashboard(Optional PreserveUserEdits As Boolean = False)
    Dim ws As Worksheet, wsEdits As Worksheet
    Dim lastRow As Long, lastRowEdits As Long
    Dim docNum As String
    Dim i As Long
    Dim backupCreated As Boolean
    
    ' Create error recovery backup before any operations
    backupCreated = CreateUserEditsBackup("RefreshDashboard_" & Format(Now, "yyyymmdd_hhmmss"))
    LogUserEditsOperation "Starting dashboard refresh. PreserveUserEdits=" & PreserveUserEdits & ", Backup created: " & backupCreated
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 1. Ensure UserEdits sheet exists with standardized structure
    SetupUserEditsSheet
    Set wsEdits = ThisWorkbook.Sheets("UserEdits")
    
    ' 2. Save any current user edits from the dashboard to UserEdits
    '    ONLY if we're not prioritizing manually edited UserEdits
    If Not PreserveUserEdits Then
        SaveUserEditsFromDashboard
    End If
    
    ' 3. Locate/create "SQRCT Dashboard"
    Set ws = GetOrCreateDashboardSheet("SQRCT Dashboard")
    
    ' 4. Clean up any duplicate headers/layout issues
    CleanupDashboardLayout ws
    
    ' 5. Clear old data from dashboard & rebuild layout (row 3 header, etc.)
    InitializeDashboardLayout ws
    
    ' 6. Populate columns A–J with data from MasterQuotes_Final
    If IsMasterQuotesFinalPresent Then
        PopulateMasterQuotesData ws
    Else
        MsgBox "Warning: MasterQuotes_Final not found. Dashboard created but no data pulled." & vbCrLf & _
               "Please ensure the data source exists.", vbInformation, "Data Source Not Found"
        GoTo Cleanup
    End If
    
    ' 7. Determine how many rows of data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 8. Sort by First Date Pulled (F) ascending, then Document Amount (D) descending
    SortDashboardData ws, lastRow
    
    ' 9. AutoFit columns & fix column widths
    With ws
        .Columns("A:J").AutoFit   ' Protected columns
        .Columns("K:N").AutoFit   ' User columns
        .Columns("C").ColumnWidth = 25  ' Customer Name
        .Columns("N").ColumnWidth = 40  ' Widen User Comments
    End With
    
    ' 10. Restore user data from UserEdits to Dashboard
    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, "A").End(xlUp).Row
    
    For i = 4 To lastRow
        docNum = Trim(CStr(ws.Cells(i, "A").Value))
        If docNum <> "" And docNum <> "Document Number" Then
            ' More robust search for matching document number
            Dim foundRow As Long
            foundRow = 0
            
            If lastRowEdits > 1 Then
                Dim j As Long
                For j = 2 To lastRowEdits
                    ' Compare trimmed, lowercase document numbers for exact match
                    If Trim(LCase(wsEdits.Cells(j, "A").Value)) = Trim(LCase(docNum)) Then
                        foundRow = j
                        Exit For
                    End If
                Next j
            End If
            
            If foundRow > 0 Then
                ' Map UserEdits data back to Dashboard using the direct column mapping:
                ' UserEdits B ? Dashboard K (Engagement Phase)
                ' UserEdits C ? Dashboard L (Last Contact Date)
                ' UserEdits D ? Dashboard M (Email Contact)
                ' UserEdits E ? Dashboard N (User Comments)
                ws.Cells(i, "K").Value = wsEdits.Cells(foundRow, "B").Value  ' Engagement Phase
                ws.Cells(i, "L").Value = wsEdits.Cells(foundRow, "C").Value  ' Last Contact Date
                ws.Cells(i, "M").Value = wsEdits.Cells(foundRow, "D").Value  ' Email Contact
                ws.Cells(i, "N").Value = wsEdits.Cells(foundRow, "E").Value  ' User Comments
            End If
        End If
    Next i
    
    ' 11. Freeze header rows
    FreezeDashboard ws
    
    ' 12. Apply color conditional formatting (optional)
    ApplyColorFormatting ws
    
    ' 13. Protect columns A–J, allow K–N
    ProtectUserColumns ws
    
    ' 14. Update the timestamp with improved styling
    With ws.Range("G2:I2")
        .Merge
        .Value = "Last Refreshed: " & Format$(Now(), "mm/dd/yyyy h:mm") & " MST"
        .HorizontalAlignment = xlCenter
        .Font.Size = 9
        .Font.Name = "Segoe UI"
        .Font.Color = RGB(80, 80, 80)
    End With
    
    ' 15. Re-create buttons with modern styling
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.TopLeftCell.Row = 2 Then
            shp.Delete
        End If
    Next shp
    On Error GoTo 0
    
    ' Create buttons with improved styling
    ModernButton ws, "C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits"
    ModernButton ws, "E2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits"
    
    ' Build message based on which mode was used
    Dim msgText As String
    If PreserveUserEdits Then
        msgText = "SQRCT Dashboard refreshed successfully!" & vbCrLf & _
                  "UserEdits were preserved and applied to the dashboard." & vbCrLf & _
                  "No changes from the dashboard were saved to UserEdits."
    Else
        msgText = "SQRCT Dashboard refreshed successfully!" & vbCrLf & _
                  "Dashboard edits were saved to UserEdits before refresh." & vbCrLf & _
                  "UserEdits were then restored to the dashboard."
    End If
    
    MsgBox msgText, vbInformation, "Dashboard Refresh Complete"
    
    ' Log successful completion
    LogUserEditsOperation "Dashboard refresh completed successfully. Mode: " & IIf(PreserveUserEdits, "PreserveUserEdits", "StandardRefresh")
    
    ' Clean up old backups if refresh was successful
    If backupCreated Then
        On Error Resume Next
        Application.DisplayAlerts = False
        Dim oldSheets As New Collection
        Dim sh As Worksheet
        For Each sh In ThisWorkbook.Sheets
            If InStr(1, sh.Name, "UserEdits_Backup_") > 0 And sh.Name <> "UserEdits_Backup_" & Format(Now, "yyyymmdd") Then
                oldSheets.Add sh
            End If
        Next sh
        
        For i = 1 To oldSheets.Count
            oldSheets(i).Delete
        Next i
        Application.DisplayAlerts = True
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    LogUserEditsOperation "ERROR in RefreshDashboard: " & Err.Description
    MsgBox "An error occurred during dashboard refresh. Your data has been backed up to a recovery sheet." & _
           vbCrLf & "Error: " & Err.Description, vbCritical, "Dashboard Refresh Error"
    ' Attempt to restore from backup if available
    If backupCreated Then RestoreUserEditsFromBackup
    Resume Cleanup
End Sub

'===============================================================================
' CreateSQRCTDashboard - Legacy name for backward compatibility
'===============================================================================
Public Sub CreateSQRCTDashboard()
    RefreshDashboard_TwoWaySync
End Sub

'===============================================================================
' SAVEUSEREDITSFROMDASHBOARD: Captures any user edits from Dashboard -> UserEdits
'===============================================================================
Public Sub SaveUserEditsFromDashboard()
    Dim wsSrc As Worksheet, wsEdits As Worksheet
    Dim lastRowSrc As Long, lastRowEdits As Long
    Dim i As Long, destRow As Long
    Dim docNum As String
    Dim hasUserEdits As Boolean
    Dim wasChanged As Boolean
    
    LogUserEditsOperation "Starting SaveUserEditsFromDashboard"
    
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Sheets("SQRCT Dashboard")
    On Error GoTo ErrorHandler
    If wsSrc Is Nothing Then
        LogUserEditsOperation "SQRCT Dashboard sheet not found"
        Exit Sub
    End If
    
    SetupUserEditsSheet
    Set wsEdits = ThisWorkbook.Sheets("UserEdits")
    
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastRowEdits = wsEdits.Cells(wsEdits.Rows.Count, "A").End(xlUp).Row
    If lastRowEdits < 1 Then lastRowEdits = 1
    
    ' Create a collection to track processed document numbers
    Dim processedDocs As New Collection
    On Error Resume Next ' In case of duplicate key errors
    
    For i = 4 To lastRowSrc
        docNum = Trim(CStr(wsSrc.Cells(i, "A").Value))  ' col A = Document Number
        If docNum <> "" And docNum <> "Document Number" Then
            ' Skip if we've already processed this document number
            On Error Resume Next
            processedDocs.Add docNum, docNum
            If Err.Number <> 0 Then
                ' Already processed this document, skip it
                Err.Clear
                GoTo NextIteration
            End If
            On Error GoTo ErrorHandler
            
            ' Check if this row has any user edits (K-N columns)
            hasUserEdits = False
            If wsSrc.Cells(i, "K").Value <> "" Or _
               wsSrc.Cells(i, "L").Value <> "" Or _
               wsSrc.Cells(i, "M").Value <> "" Or _
               wsSrc.Cells(i, "N").Value <> "" Then
                hasUserEdits = True
            End If
            
            ' More robust method to find existing document number
            Dim foundRow As Long
            foundRow = 0
            
            If lastRowEdits > 1 Then
                Dim j As Long
                For j = 2 To lastRowEdits
                    ' Compare cleaned document numbers for exact match
                    If Trim(LCase(wsEdits.Cells(j, "A").Value)) = Trim(LCase(docNum)) Then
                        foundRow = j
                        Exit For
                    End If
                Next j
            End If
            
            ' Process this document number if:
            ' 1. It has user edits in columns K-N, OR
            ' 2. It already exists in UserEdits
            If hasUserEdits Or foundRow > 0 Then
                ' Determine destination row
                If foundRow > 0 Then
                    destRow = foundRow
                Else
                    destRow = lastRowEdits + 1
                    wsEdits.Cells(destRow, "A").Value = docNum
                    lastRowEdits = destRow
                End If
                
                ' Track if we're making changes to determine if timestamp needs updating
                wasChanged = False
                
                ' Only update UserEdits if either:
                ' 1. This is a new entry, or
                ' 2. The value in the dashboard is different from what's in UserEdits
                
                If foundRow = 0 Or wsEdits.Cells(destRow, "B").Value <> wsSrc.Cells(i, "K").Value Then
                    wsEdits.Cells(destRow, "B").Value = wsSrc.Cells(i, "K").Value  ' Engagement Phase
                    wasChanged = True
                End If
                
                If foundRow = 0 Or wsEdits.Cells(destRow, "C").Value <> wsSrc.Cells(i, "L").Value Then
                    wsEdits.Cells(destRow, "C").Value = wsSrc.Cells(i, "L").Value  ' Last Contact Date
                    wasChanged = True
                End If
                
                If foundRow = 0 Or wsEdits.Cells(destRow, "D").Value <> wsSrc.Cells(i, "M").Value Then
                    wsEdits.Cells(destRow, "D").Value = wsSrc.Cells(i, "M").Value  ' Email Contact
                    wasChanged = True
                End If
                
                If foundRow = 0 Or wsEdits.Cells(destRow, "E").Value <> wsSrc.Cells(i, "N").Value Then
                    wsEdits.Cells(destRow, "E").Value = wsSrc.Cells(i, "N").Value  ' User Comments
                    wasChanged = True
                End If
                
                ' Set ChangeSource to workbook identity and update timestamp only if something changed
                If hasUserEdits And wasChanged Then
                    wsEdits.Cells(destRow, "F").Value = GetWorkbookIdentity()  ' Use workbook identity
                    wsEdits.Cells(destRow, "G").Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")  ' Timestamp
                    LogUserEditsOperation "Updated UserEdits for DocNumber " & docNum & " with attribution " & GetWorkbookIdentity()
                End If
            End If
        End If
NextIteration:
    Next i
    
    LogUserEditsOperation "Completed SaveUserEditsFromDashboard"
    Exit Sub
    
ErrorHandler:
    LogUserEditsOperation "ERROR in SaveUserEditsFromDashboard: " & Err.Description
    Resume NextIteration ' Try to continue with the next document
End Sub

'===============================================================================
' SETUPUSEREDITSSHEET: Creates or ensures existence of "UserEdits" with standard structure
'===============================================================================
Public Sub SetupUserEditsSheet()
    Dim wsEdits As Worksheet
    Dim needsBackup As Boolean
    Dim wsBackup As Worksheet
    
    LogUserEditsOperation "Starting SetupUserEditsSheet"
    
    ' First check if the sheet exists
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets("UserEdits")
    On Error GoTo ErrorHandler
    
    ' If it exists, determine if we need to create a backup before modifying
    If Not wsEdits Is Nothing Then
        ' Check if we need to restructure (missing Timestamp column or wrong order)
        If wsEdits.Cells(1, 7).Value <> "Timestamp" Or _
           wsEdits.Cells(1, 6).Value <> "ChangeSource" Or _
           wsEdits.Cells(1, 2).Value <> "Engagement Phase" Then
            needsBackup = True
            LogUserEditsOperation "UserEdits sheet structure needs update - will create backup"
        End If
        
        ' Create backup if needed
        If needsBackup Then
            On Error Resume Next
            Set wsBackup = ThisWorkbook.Sheets("UserEdits_Backup")
            If wsBackup Is Nothing Then
                Set wsBackup = ThisWorkbook.Sheets.Add(After:=wsEdits)
                wsBackup.Name = "UserEdits_Backup"
                LogUserEditsOperation "Created UserEdits_Backup sheet"
            End If
            On Error GoTo ErrorHandler
            
            ' Copy current data to backup
            wsEdits.UsedRange.Copy wsBackup.Range("A1")
            wsBackup.Visible = xlSheetHidden
        End If
    End If
    
    ' Now create or update the UserEdits sheet
    If wsEdits Is Nothing Then
        ' Sheet doesn't exist, create it new
        Set wsEdits = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsEdits.Name = "UserEdits"
        LogUserEditsOperation "Created new UserEdits sheet"
        
        ' Set up headers with improved styling
        With wsEdits.Range("A1:G1")
            .Value = Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                           "Email Contact", "User Comments", "ChangeSource", "Timestamp")
            .Font.Bold = True
            .Interior.Color = RGB(16, 107, 193)  ' Match dashboard title
            .Font.Color = RGB(255, 255, 255)
        End With
        
        wsEdits.Visible = xlSheetHidden
    Else
        ' Sheet exists, make sure it has the right structure
        ' First preserve existing data if it's in the old format
        If needsBackup Then
            ' Migrate old data to new structure if needed
            If wsBackup Is Nothing Then Exit Sub  ' Safety check
            
            ' Capture old data structure
            Dim lastRow As Long
            lastRow = wsEdits.Cells(wsEdits.Rows.Count, "A").End(xlUp).Row
            
            ' Clear existing data
            wsEdits.Cells.Clear
            
            ' Set up new headers
            With wsEdits.Range("A1:G1")
                .Value = Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                               "Email Contact", "User Comments", "ChangeSource", "Timestamp")
                .Font.Bold = True
                .Interior.Color = RGB(16, 107, 193)  ' Match dashboard title
                .Font.Color = RGB(255, 255, 255)
            End With
            
            ' Migrate data from backup to new structure
            If lastRow > 1 Then
                Dim i As Long
                For i = 2 To lastRow
                    ' Only migrate if there's a document number
                    If wsBackup.Cells(i, 1).Value <> "" Then
                        ' Map old structure to new structure
                        wsEdits.Cells(i, "A").Value = wsBackup.Cells(i, 1).Value  ' DocNumber
                        
                        ' Map based on headers in backup
                        Dim colE As String, colB As String, colC As String, colD As String
                        colB = "" ' Engagement Phase
                        colC = "" ' Last Contact Date
                        colD = "" ' Email Contact
                        colE = "" ' User Comments
                        
                        ' Look for corresponding columns in backup
                        If wsBackup.Cells(1, 3).Value = "UserStageOverride" Or _
                           wsBackup.Cells(1, 3).Value = "EngagementPhase" Then
                            colB = wsBackup.Cells(i, 3).Value  ' Engagement Phase
                        End If
                        
                        If wsBackup.Cells(1, 4).Value = "LastContactDate" Then
                            colC = wsBackup.Cells(i, 4).Value  ' Last Contact Date
                        End If
                        
                        If wsBackup.Cells(1, 5).Value = "EmailContact" Then
                            colD = wsBackup.Cells(i, 5).Value  ' Email Contact
                        End If
                        
                        If wsBackup.Cells(1, 2).Value = "UserComments" Then
                            colE = wsBackup.Cells(i, 2).Value  ' User Comments
                        End If
                        
                        ' Set values in new structure
                        wsEdits.Cells(i, "B").Value = colB
                        wsEdits.Cells(i, "C").Value = colC
                        wsEdits.Cells(i, "D").Value = colD
                        wsEdits.Cells(i, "E").Value = colE
                        wsEdits.Cells(i, "F").Value = GetWorkbookIdentity()  ' Use workbook identity
                        wsEdits.Cells(i, "G").Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")  ' Current timestamp
                    End If
                Next i
            End If
            
            LogUserEditsOperation "Migrated UserEdits data to new structure"
        Else
            ' Just ensure headers are correct
            wsEdits.Range("A1:G1").Value = Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                                               "Email Contact", "User Comments", "ChangeSource", "Timestamp")
            wsEdits.Range("A1:G1").Font.Bold = True
            wsEdits.Range("A1:G1").Interior.Color = RGB(16, 107, 193)
            wsEdits.Range("A1:G1").Font.Color = RGB(255, 255, 255)
        End If
    End If
    
    LogUserEditsOperation "Completed SetupUserEditsSheet"
    Exit Sub
    
ErrorHandler:
    LogUserEditsOperation "ERROR in SetupUserEditsSheet: " & Err.Description
    Resume Next ' Try to continue
End Sub

'===============================================================================
' GETORCREATEDASHBOARDSHEET: Returns or creates the SQRCT Dashboard
'===============================================================================
Private Function GetOrCreateDashboardSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
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
'===============================================================================
Private Sub CleanupDashboardLayout(ws As Worksheet)
    Application.ScreenUpdating = False
    
    ' Unprotect sheet for modifications
    On Error Resume Next
    ws.Unprotect Password:="password"
    On Error GoTo 0
    
    ' Step 1: Save data from row 4 onward
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim dataRange As Range
    Dim tempData As Variant
    
    If lastRow >= 4 Then
        ' Capture all data below row 3
        Set dataRange = ws.Range("A4:N" & lastRow)
        tempData = dataRange.Value
    End If
    
    ' Step 2: Find rows to preserve (rows 1-3)
    Dim hasTitle As Boolean
    hasTitle = False
    
    ' Check for title text in each cell of row 1 individually
    Dim cell As Range
    For Each cell In ws.Range("A1:N1").Cells
        If InStr(1, CStr(cell.Value), "STRATEGIC QUOTE RECOVERY", vbTextCompare) > 0 Then
            hasTitle = True
            Exit For
        End If
    Next cell
    
    ' Step 3: Clear the entire sheet EXCEPT rows 1-3
    ws.Range("A4:N" & ws.Rows.Count).Clear
    
    ' Step 4: If the title row (row 1) is missing, recreate it
    If Not hasTitle Then
        With ws.Range("A1:N1")
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
    
    ' Step 5: Ensure row 2 has control panel with professional styling
    With ws.Range("A2:N2")
        ' Light silver-gray gradient effect
        .Interior.Color = RGB(245, 245, 245)  ' Very light gray
        
        ' Add a subtle top and bottom border for definition
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).Color = RGB(200, 200, 200)
        
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
        
        ' Set row height for better proportions
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
        
        ' Add a subtle right border instead of surrounding border
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(200, 200, 200)
    End With
    
    ' Question mark in corner for help
    With ws.Range("N2")
        .Value = "?"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(70, 130, 180)  ' Matching steel blue
    End With
    
    ' Step 6: Ensure row 3 has column headers with improved styling
    With ws.Range("A3:N3")
        .Clear
        .Value = Array( _
            "Document Number", _
            "Client ID", _
            "Customer Name", _
            "Document Amount", _
            "Document Date", _
            "First Date Pulled", _
            "Salesperson ID", _
            "Entered By", _
            "Occurrence Counter", _
            "Missing Quote Alert", _
            "Engagement Phase", _
            "Last Contact Date", _
            "Email Contact", _
            "User Comments")
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)  ' Match title row color
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Step 7: Restore data if we had any
    If Not IsEmpty(tempData) Then
        ws.Range("A4").Resize(UBound(tempData), UBound(tempData, 2)).Value = tempData
    End If
    
    Application.ScreenUpdating = True
End Sub

'===============================================================================
' INITIALIZEDASHBOARDLAYOUT: Clears rows 4+ in A–N, sets up header row in A3:N3
'===============================================================================
Private Sub InitializeDashboardLayout(ws As Worksheet)
    ' Only clear rows 4+ to preserve header/control panel
    ws.Range("A4:N" & ws.Rows.Count).Clear
    
    ' Delete extra columns O:Z if needed
    On Error Resume Next
    ws.Range("O:Z").Delete
    On Error GoTo 0
    
    ' Ensure row 3 has correct headers with improved styling
    With ws.Range("A3:N3")
        .Clear
        .Value = Array( _
            "Document Number", _
            "Client ID", _
            "Customer Name", _
            "Document Amount", _
            "Document Date", _
            "First Date Pulled", _
            "Salesperson ID", _
            "Entered By", _
            "Occurrence Counter", _
            "Missing Quote Alert", _
            "Engagement Phase", _
            "Last Contact Date", _
            "Email Contact", _
            "User Comments")
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)  ' Match title row color
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Set initial column widths
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
        
        .Columns("K").ColumnWidth = 20
        .Columns("L").ColumnWidth = 15
        .Columns("M").ColumnWidth = 25
        .Columns("N").ColumnWidth = 40
    End With
End Sub

'===============================================================================
' POPULATEMASTERQUOTESDATA: Pulls columns A–J from MasterQuotes_Final
'===============================================================================
Private Sub PopulateMasterQuotesData(ws As Worksheet)
    With ws
        ' A: Document Number
        .Range("A4").Formula = _
            "=IF(ROWS($A$4:A4)<=ROWS(MasterQuotes_Final[Document Number])," & _
            "IFERROR(INDEX(MasterQuotes_Final[Document Number],ROWS($A$4:A4)),""""),"""")"
        
        ' B: Client ID -> from Customer Number
        .Range("B4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(MasterQuotes_Final[Customer Number],ROWS($A$4:A4)),""""),"""")"
        
        ' C: Customer Name
        .Range("C4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(MasterQuotes_Final[Customer Name],ROWS($A$4:A4)),""""),"""")"
        
        ' D: Document Amount
        .Range("D4").Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(MasterQuotes_Final[Document Amount],ROWS($A$4:A4)),""""),"""")"
        
        ' E: Document Date
        .Range("E4").Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(MasterQuotes_Final[Document Date],ROWS($A$4:A4)),""""),"""")"
        
        ' F: First Date Pulled
        .Range("F4").Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(MasterQuotes_Final[First Date Pulled],ROWS($A$4:A4)),""""),"""")"
        
        ' G: Salesperson ID
        .Range("G4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(MasterQuotes_Final[Salesperson ID],ROWS($A$4:A4)),""""),"""")"
        
        ' H: Entered By (was User To Enter)
        .Range("H4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(MasterQuotes_Final[User To Enter],ROWS($A$4:A4)),""""),"""")"
        
        ' I: Occurrence Counter (was Auto Stage)
        .Range("I4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(MasterQuotes_Final[AutoStage],ROWS($A$4:A4)),""""),"""")"
        
        ' J: Missing Quote Alert (was Auto Note)
        .Range("J4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(MasterQuotes_Final[AutoNote],ROWS($A$4:A4)),""""),"""")"
        
        ' Autofill down
        .Range("A4:J4").AutoFill Destination:=.Range("A4:J1000"), Type:=xlFillDefault
        
        ' Format numeric/date columns
        .Range("D4:D1000").NumberFormat = "$#,##0.00"   ' Document Amount
        .Range("E4:E1000").NumberFormat = "mm/dd/yyyy" ' Document Date
        .Range("F4:F1000").NumberFormat = "mm/dd/yyyy" ' First Date Pulled
    End With
End Sub

'===============================================================================
' SORTDASHBOARDDATA: Sort by First Date Pulled (F asc), then Document Amount (D desc)
'===============================================================================
Private Sub SortDashboardData(ws As Worksheet, lastRow As Long)
    If lastRow < 5 Then Exit Sub
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("F4:F" & lastRow), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range("D4:D" & lastRow), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange ws.Range("A3:N" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'===============================================================================
' FREEZEDASHBOARD: Freezes rows 1–3
'===============================================================================
Private Sub FreezeDashboard(ws As Worksheet)
    ws.Activate
    
    ' First unfreeze any existing splits
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 0
    ActiveWindow.SplitColumn = 0
    
    ' Freeze rows 1–3
    ActiveWindow.SplitRow = 3
    ActiveWindow.SplitColumn = 0
    ActiveWindow.FreezePanes = True
End Sub

'===============================================================================
' SETUPDASHBOARD: Professional row 1 & 2 design (title & control panel)
'===============================================================================
Public Sub SetupDashboard(ws As Worksheet)
    ' Merge & style title in row 1 - updated to a more vibrant blue
    With ws.Range("A1:N1")
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
    
    ' Set up control panel in row 2 with modern styling
    With ws.Range("A2:N2")
        ' Light silver-gray gradient effect
        .Interior.Color = RGB(245, 245, 245)  ' Very light gray
        
        ' Add a subtle top and bottom border for definition
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).Color = RGB(200, 200, 200)
        
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
        
        ' Set row height for better proportions
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
        
        ' Add a subtle right border instead of surrounding border
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(200, 200, 200)
    End With
    
    ' Question mark in corner for help
    With ws.Range("N2")
        .Value = "?"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(70, 130, 180)  ' Matching steel blue
    End With
    
    ' Last refreshed timestamp - positioned and styled more elegantly
    With ws.Range("G2:I2")
        .Merge
        .Value = "Last Refreshed: " & Format$(Now(), "mm/dd/yyyy h:mm") & " MST"
        .HorizontalAlignment = xlCenter
        .Font.Size = 9
        .Font.Name = "Segoe UI"
        .Font.Color = RGB(80, 80, 80)  ' Dark gray for subtle elegance
    End With
    
    ' Unprotect sheet temporarily to add buttons
    On Error Resume Next
    ws.Unprotect Password:="password"
    On Error GoTo 0
    
    ' Create buttons with improved spacing and styling
    ModernButton ws, "C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits"
    ModernButton ws, "E2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits"
    
    ' Protect sheet again
    ws.Protect Password:="password", DrawingObjects:=True, Contents:=True, _
                Scenarios:=True, UserInterfaceOnly:=True, AllowFormattingCells:=False, _
                AllowFormattingColumns:=False, AllowFormattingRows:=False
End Sub

'===============================================================================
' MODERNBUTTON: Creates professional, modern-looking buttons with proper spacing
'===============================================================================
Public Sub ModernButton(ws As Worksheet, cellRef As String, buttonText As String, macroName As String)
    Dim btn As Object
    Dim buttonTop As Double, buttonLeft As Double
    Dim buttonWidth As Double, buttonHeight As Double
    
    ' Get the position and size based on the cell
    buttonLeft = ws.Range(cellRef).Left
    buttonTop = ws.Range(cellRef).Top
    
    ' Calculate precise width and position to prevent overlap
    buttonWidth = ws.Range(cellRef & ":" & cellRef).Width * 1.6
    buttonHeight = ws.Range(cellRef).Height * 0.75
    
    ' Better centering with more space between buttons
    buttonTop = buttonTop + (ws.Range(cellRef).Height - buttonHeight) / 2
    
    ' Create the button using simple AddShape (9 = rounded rectangle)
    On Error Resume Next
    Set btn = ws.Shapes.AddShape(9, buttonLeft, buttonTop, buttonWidth, buttonHeight)
    
    If Err.Number <> 0 Then
        ' Try alternative approach for older Excel versions
        Err.Clear
        On Error Resume Next
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
    End If
    On Error GoTo 0
    
    ' Exit if button creation failed
    If btn Is Nothing Then Exit Sub
    
    ' Style the button with modern, professional appearance
    With btn
        ' Modern gradient effect using solid color
        .Fill.ForeColor.RGB = RGB(42, 120, 180)  ' Professional blue
        
        ' More subtle border
        .Line.ForeColor.RGB = RGB(25, 95, 150)
        .Line.Weight = 0.75
        
        ' Set the text with improved font
        On Error Resume Next
        .TextFrame.Characters.Text = buttonText
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Size = 10
        
        ' Try to use a modern font if available
        .TextFrame.Characters.Font.Name = "Segoe UI"
        .TextFrame.Characters.Font.Bold = True
        
        ' Text positioning
        On Error Resume Next
        .TextFrame.HorizontalAlignment = 2  ' Center = 2
        .TextFrame.VerticalAlignment = 3    ' Middle = 3
        
        ' Add a subtle shadow effect if possible
        On Error Resume Next
        .Shadow.Type = msoShadow21
        .Shadow.Transparency = 0.7
        
        ' Assign macro
        .OnAction = macroName
    End With
End Sub

'===============================================================================
' PROTECTUSERCOLUMNS: Lock A–J, unlock K–N
'===============================================================================
Public Sub ProtectUserColumns(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:="password"
    On Error GoTo 0
    
    ws.Cells.Locked = True
    ws.Range("K4:N" & ws.Rows.Count).Locked = False
    
    ws.Protect Password:="password", UserInterfaceOnly:=True, DrawingObjects:=True
End Sub

'===============================================================================
' APPLYCOLORFORMATTING: For coloring columns I (Occurrence) and K (Engagement)
'===============================================================================
Public Sub ApplyColorFormatting(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:="password"
    On Error GoTo 0
    
    ' Clear existing rules in I4:I1000 and K4:K1000
    ws.Range("I4:I1000,K4:K1000").FormatConditions.Delete
    
    Dim rngOccur As Range, rngPhase As Range
    Set rngOccur = ws.Range("I4:I1000")
    Set rngPhase = ws.Range("K4:K1000")
    
    ' Apply conditional formatting to both columns
    ApplyStageFormatting rngOccur
    ApplyStageFormatting rngPhase
    
    ws.Protect Password:="password", UserInterfaceOnly:=True
End Sub

' Helper for detailed color rules with more conditions added
Private Sub ApplyStageFormatting(rng As Range)
    With rng
        ' First F/U (Light Blue)
        .FormatConditions.Add Type:=xlTextString, String:="First F/U", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(197, 217, 241)
        
        ' Second F/U (Light Green)
        .FormatConditions.Add Type:=xlTextString, String:="Second F/U", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(198, 239, 206)
        
        ' Third F/U (Light Yellow)
        .FormatConditions.Add Type:=xlTextString, String:="Third F/U", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 235, 156)
        
        ' Long-Term F/U (Orange)
        .FormatConditions.Add Type:=xlTextString, String:="Long-Term F/U", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 192, 0)
        
        ' WW/OM (Pink)
        .FormatConditions.Add Type:=xlTextString, String:="WW/OM", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 192, 203)
        
        ' Converted (Bright Green, Bold)
        .FormatConditions.Add Type:=xlTextString, String:="Converted", TextOperator:=xlContains
        With .FormatConditions(.FormatConditions.Count)
            .Interior.Color = RGB(146, 208, 80)
            .Font.Bold = True
        End With
        
        ' Declined (Light Red)
        .FormatConditions.Add Type:=xlTextString, String:="Declined", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 199, 206)
        
        ' Closed (Extra Order) (Light Gray)
        .FormatConditions.Add Type:=xlTextString, String:="Closed (Extra Order)", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(217, 217, 217)
        
        ' Added: Closed (any type) (Medium Gray)
        .FormatConditions.Add Type:=xlTextString, String:="Closed", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(217, 217, 217)
    End With
End Sub

'===============================================================================
' ISMASTERQUOTESFINALPRESENT: Checks for PQ, Table, or Named Range "MasterQuotes_Final"
'===============================================================================
Public Function IsMasterQuotesFinalPresent() As Boolean
    Dim lo As ListObject
    Dim nm As Name
    Dim queryObj As Object
    
    IsMasterQuotesFinalPresent = False
    
    On Error Resume Next
    
    ' 1) Power Query named "MasterQuotes_Final"
    For Each queryObj In ActiveWorkbook.Queries
        If queryObj.Name = "MasterQuotes_Final" Then
            IsMasterQuotesFinalPresent = True
            Exit For
        End If
    Next queryObj
    
    ' 2) ListObject named "MasterQuotes_Final"
    If Not IsMasterQuotesFinalPresent Then
        For Each lo In ActiveWorkbook.ListObjects
            If lo.Name = "MasterQuotes_Final" Then
                IsMasterQuotesFinalPresent = True
                Exit For
            End If
        Next lo
    End If
    
    ' 3) Named Range "MasterQuotes_Final"
    If Not IsMasterQuotesFinalPresent Then
        For Each nm In ActiveWorkbook.Names
            If nm.Name = "MasterQuotes_Final" Then
                IsMasterQuotesFinalPresent = True
                Exit For
            End If
        Next nm
    End If
    
    On Error GoTo 0
End Function


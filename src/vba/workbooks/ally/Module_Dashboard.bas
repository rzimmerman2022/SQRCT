Attribute VB_Name = "Module_Dashboard"
Option Explicit

' --- Constants ---
Private Const DASHBOARD_SHEET_NAME As String = "SQRCT Dashboard"
Private Const USEREDITS_SHEET_NAME As String = "UserEdits"
Private Const USEREDITSLOG_SHEET_NAME As String = "UserEditsLog"
Private Const MASTER_QUOTES_FINAL_SOURCE As String = "MasterQuotes_Final" ' Name of the PQ query/table

' UserEdits Columns
Private Const UE_COL_DOCNUM As String = "A"
Private Const UE_COL_PHASE As String = "B"
Private Const UE_COL_LASTCONTACT As String = "C"
Private Const UE_COL_EMAIL As String = "D"
Private Const UE_COL_COMMENTS As String = "E"
Private Const UE_COL_SOURCE As String = "F"
Private Const UE_COL_TIMESTAMP As String = "G"

' Dashboard Columns (Editable)
Private Const DB_COL_PHASE As String = "K"
Private Const DB_COL_LASTCONTACT As String = "L"
Private Const DB_COL_EMAIL As String = "M"
Private Const DB_COL_COMMENTS As String = "N"
' --- End Constants ---


'===============================================================================
' MODULE_DASHBOARD
' Contains functions for managing the SQRCT Dashboard, including refresh operations,
' user edits management, and UI interactions.
' Refactored for performance (Dictionary lookup), security (no passwords), and robustness (constants).
' Protection logic revised: Applied after modifications in relevant subs.
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
'===============================================================================
Public Sub RefreshDashboard(Optional PreserveUserEdits As Boolean = False)
    Dim ws As Worksheet, wsEdits As Worksheet
    Dim lastRow As Long, lastRowEdits As Long
    Dim docNum As String
    Dim i As Long
    Dim backupCreated As Boolean
    Dim userEditsDict As Object ' Dictionary for UserEdits lookup
    Dim editRow As Variant      ' To store row number or data from dictionary

    ' Create error recovery backup before any operations
    backupCreated = CreateUserEditsBackup("RefreshDashboard_" & Format(Now, "yyyymmdd_hhmmss"))
    LogUserEditsOperation "Starting dashboard refresh. PreserveUserEdits=" & PreserveUserEdits & ", Backup created: " & backupCreated

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 1. Ensure UserEdits sheet exists with standardized structure
    SetupUserEditsSheet
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Use Constant

    ' 2. Save any current user edits from the dashboard to UserEdits
    '    ONLY if we're not prioritizing manually edited UserEdits
    If Not PreserveUserEdits Then
        SaveUserEditsFromDashboard
    End If

    ' 3. Locate/create "SQRCT Dashboard" using name constant
    Set ws = GetOrCreateDashboardSheet(DASHBOARD_SHEET_NAME) ' Use Constant and helper function

    ' 4. Clean up any duplicate headers/layout issues
    CleanupDashboardLayout ws

    ' 5. Clear old data from dashboard & rebuild layout (row 3 header, etc.)
    InitializeDashboardLayout ws

    ' 6. Populate columns A-J with data from MasterQuotes_Final
    If IsMasterQuotesFinalPresent Then
        PopulateMasterQuotesData ws ' Uses MASTER_QUOTES_FINAL_SOURCE constant internally
    Else
        MsgBox "Warning: " & MASTER_QUOTES_FINAL_SOURCE & " not found. Dashboard created but no data pulled." & vbCrLf & _
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
        .Columns(DB_COL_PHASE & ":" & DB_COL_COMMENTS).AutoFit   ' User columns (K:N)
        .Columns("C").ColumnWidth = 25  ' Customer Name
        .Columns(DB_COL_COMMENTS).ColumnWidth = 40  ' Widen User Comments (N)
    End With

    ' 10. Restore user data from UserEdits to Dashboard using Dictionary
    Set userEditsDict = LoadUserEditsToDictionary(wsEdits)

    For i = 4 To lastRow
        docNum = Trim(CStr(ws.Cells(i, "A").Value))
        If docNum <> "" And docNum <> "Document Number" Then
            If userEditsDict.Exists(docNum) Then
                editRow = userEditsDict(docNum) ' Get the row number from the dictionary

                ' Map UserEdits data back to Dashboard using the direct column mapping:
                ws.Cells(i, DB_COL_PHASE).Value = wsEdits.Cells(editRow, UE_COL_PHASE).Value       ' Engagement Phase
                ws.Cells(i, DB_COL_LASTCONTACT).Value = wsEdits.Cells(editRow, UE_COL_LASTCONTACT).Value ' Last Contact Date
                ws.Cells(i, DB_COL_EMAIL).Value = wsEdits.Cells(editRow, UE_COL_EMAIL).Value         ' Email Contact
                ws.Cells(i, DB_COL_COMMENTS).Value = wsEdits.Cells(editRow, UE_COL_COMMENTS).Value     ' User Comments
            Else
                ' Clear existing user data if no match found in UserEdits (optional, depends on desired behavior)
                ws.Cells(i, DB_COL_PHASE).Value = ""
                ws.Cells(i, DB_COL_LASTCONTACT).Value = ""
                ws.Cells(i, DB_COL_EMAIL).Value = ""
                ws.Cells(i, DB_COL_COMMENTS).Value = ""
            End If
        End If
    Next i
    Set userEditsDict = Nothing ' Clean up dictionary

    ' 11. Freeze header rows
    FreezeDashboard ws

    ' 12. Apply color conditional formatting (optional)
    ApplyColorFormatting ws

    ' 13. Protect columns A-J, allow K-N (using UserInterfaceOnly)
    ProtectUserColumns ws ' Protection is now applied at the end of this sub

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
        msgText = DASHBOARD_SHEET_NAME & " refreshed successfully!" & vbCrLf & _
                  USEREDITS_SHEET_NAME & " were preserved and applied to the dashboard." & vbCrLf & _
                  "No changes from the dashboard were saved to " & USEREDITS_SHEET_NAME & "."
    Else
        msgText = DASHBOARD_SHEET_NAME & " refreshed successfully!" & vbCrLf & _
                  "Dashboard edits were saved to " & USEREDITS_SHEET_NAME & " before refresh." & vbCrLf & _
                  USEREDITS_SHEET_NAME & " were then restored to the dashboard."
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
            If InStr(1, sh.Name, USEREDITS_SHEET_NAME & "_Backup_") > 0 And sh.Name <> USEREDITS_SHEET_NAME & "_Backup_" & Format(Now, "yyyymmdd") Then ' Use Constant
                oldSheets.Add sh
            End If
        Next sh

        For i = 1 To oldSheets.Count
            oldSheets(i).Delete
        Next i
        Application.DisplayAlerts = True
        On Error GoTo ErrorHandler ' Restore error handling
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
' Refactored to use Dictionary lookup for performance.
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
        docNum = Trim(CStr(wsSrc.Cells(i, "A").Value))  ' col A = Document Number
        If docNum <> "" And docNum <> "Document Number" Then

            ' Check if this row has any user edits (K-N columns)
            hasUserEdits = False
            If wsSrc.Cells(i, DB_COL_PHASE).Value <> "" Or _
               wsSrc.Cells(i, DB_COL_LASTCONTACT).Value <> "" Or _
               wsSrc.Cells(i, DB_COL_EMAIL).Value <> "" Or _
               wsSrc.Cells(i, DB_COL_COMMENTS).Value <> "" Then
                hasUserEdits = True
            End If

            ' Find existing row using dictionary
            If userEditsDict.Exists(docNum) Then
                editRow = userEditsDict(docNum) ' Get existing row number
            Else
                editRow = 0 ' Flag as not found
            End If

            ' Process this document number if:
            ' 1. It has user edits in columns K-N, OR
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

                ' Get current values from dashboard
                Dim dbPhase, dbLastContact, dbEmail, dbComments
                dbPhase = wsSrc.Cells(i, DB_COL_PHASE).Value
                dbLastContact = wsSrc.Cells(i, DB_COL_LASTCONTACT).Value
                dbEmail = wsSrc.Cells(i, DB_COL_EMAIL).Value
                dbComments = wsSrc.Cells(i, DB_COL_COMMENTS).Value

                ' Only update UserEdits if either:
                ' 1. This is a new entry (editRow was 0 initially), or
                ' 2. The value in the dashboard is different from what's in UserEdits

                If editRow = 0 Or wsEdits.Cells(destRow, UE_COL_PHASE).Value <> dbPhase Then
                    wsEdits.Cells(destRow, UE_COL_PHASE).Value = dbPhase
                    wasChanged = True
                End If

                If editRow = 0 Or wsEdits.Cells(destRow, UE_COL_LASTCONTACT).Value <> dbLastContact Then
                    wsEdits.Cells(destRow, UE_COL_LASTCONTACT).Value = dbLastContact
                    wasChanged = True
                End If

                If editRow = 0 Or wsEdits.Cells(destRow, UE_COL_EMAIL).Value <> dbEmail Then
                    wsEdits.Cells(destRow, UE_COL_EMAIL).Value = dbEmail
                    wasChanged = True
                End If

                If editRow = 0 Or wsEdits.Cells(destRow, UE_COL_COMMENTS).Value <> dbComments Then
                    wsEdits.Cells(destRow, UE_COL_COMMENTS).Value = dbComments
                    wasChanged = True
                End If

                ' Set ChangeSource to workbook identity and update timestamp only if something changed
                If wasChanged Then ' Update timestamp if any field was modified or if it's a new entry with edits
                    wsEdits.Cells(destRow, UE_COL_SOURCE).Value = GetWorkbookIdentity()  ' Use workbook identity
                    wsEdits.Cells(destRow, UE_COL_TIMESTAMP).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")  ' Timestamp
                    LogUserEditsOperation "Updated UserEdits for DocNumber " & docNum & " with attribution " & GetWorkbookIdentity()
                End If
            End If
        End If
'NextIteration: ' Removed GoTo
    Next i
    Set userEditsDict = Nothing ' Clean up dictionary

    LogUserEditsOperation "Completed SaveUserEditsFromDashboard"
    Exit Sub

ErrorHandler:
    LogUserEditsOperation "ERROR in SaveUserEditsFromDashboard: " & Err.Description
    ' Resume NextIteration ' Removed GoTo - Let error propagate or handle differently
    Set userEditsDict = Nothing ' Clean up dictionary on error
    ' Consider re-enabling events if necessary before exiting on error
    Application.EnableEvents = True
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
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Use Constant
    On Error GoTo ErrorHandler

    ' If it exists, determine if we need to create a backup before modifying
    If Not wsEdits Is Nothing Then
        ' Check if we need to restructure (missing Timestamp column or wrong order)
        If wsEdits.Cells(1, UE_COL_TIMESTAMP).Value <> "Timestamp" Or _
           wsEdits.Cells(1, UE_COL_SOURCE).Value <> "ChangeSource" Or _
           wsEdits.Cells(1, UE_COL_PHASE).Value <> "Engagement Phase" Then ' Use Constants
            needsBackup = True
            LogUserEditsOperation USEREDITS_SHEET_NAME & " sheet structure needs update - will create backup"
        End If

        ' Create backup if needed
        If needsBackup Then
            On Error Resume Next
            Set wsBackup = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME & "_Backup") ' Use Constant
            If wsBackup Is Nothing Then
                Set wsBackup = ThisWorkbook.Sheets.Add(After:=wsEdits)
                wsBackup.Name = USEREDITS_SHEET_NAME & "_Backup" ' Use Constant
                LogUserEditsOperation "Created " & USEREDITS_SHEET_NAME & "_Backup sheet"
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
        wsEdits.Name = USEREDITS_SHEET_NAME ' Use Constant
        LogUserEditsOperation "Created new " & USEREDITS_SHEET_NAME & " sheet" ' Use Constant

        ' Set up headers with improved styling using constants
        With wsEdits.Range(wsEdits.Cells(1, UE_COL_DOCNUM), wsEdits.Cells(1, UE_COL_TIMESTAMP)) ' Use Constants for range
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

            ' Set up new headers using constants
            With wsEdits.Range(wsEdits.Cells(1, UE_COL_DOCNUM), wsEdits.Cells(1, UE_COL_TIMESTAMP)) ' Use Constants for range
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
                        ' Map old structure to new structure using constants
                        wsEdits.Cells(i, UE_COL_DOCNUM).Value = wsBackup.Cells(i, 1).Value  ' DocNumber

                        ' Map based on headers in backup (assuming old structure might vary)
                        Dim colPhase As String, colLastContact As String, colEmail As String, colComments As String
                        colPhase = ""
                        colLastContact = ""
                        colEmail = ""
                        colComments = ""

                        ' Find old columns by header text (more robust than fixed indices)
                        Dim h As Long, headerText As String
                        For h = 1 To wsBackup.UsedRange.Columns.Count
                            headerText = CStr(wsBackup.Cells(1, h).Value)
                            Select Case headerText
                                Case "UserStageOverride", "EngagementPhase", "Engagement Phase"
                                    colPhase = wsBackup.Cells(i, h).Value
                                Case "LastContactDate", "Last Contact Date"
                                    colLastContact = wsBackup.Cells(i, h).Value
                                Case "EmailContact", "Email Contact"
                                    colEmail = wsBackup.Cells(i, h).Value
                                Case "UserComments", "User Comments"
                                    colComments = wsBackup.Cells(i, h).Value
                            End Select
                        Next h

                        ' Set values in new structure using constants
                        wsEdits.Cells(i, UE_COL_PHASE).Value = colPhase
                        wsEdits.Cells(i, UE_COL_LASTCONTACT).Value = colLastContact
                        wsEdits.Cells(i, UE_COL_EMAIL).Value = colEmail
                        wsEdits.Cells(i, UE_COL_COMMENTS).Value = colComments
                        wsEdits.Cells(i, UE_COL_SOURCE).Value = GetWorkbookIdentity()  ' Use workbook identity
                        wsEdits.Cells(i, UE_COL_TIMESTAMP).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")  ' Current timestamp
                    End If
                Next i
            End If

            LogUserEditsOperation "Migrated " & USEREDITS_SHEET_NAME & " data to new structure"
        Else
            ' Just ensure headers are correct using constants
            With wsEdits.Range(wsEdits.Cells(1, UE_COL_DOCNUM), wsEdits.Cells(1, UE_COL_TIMESTAMP)) ' Use Constants for range
                .Value = Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                               "Email Contact", "User Comments", "ChangeSource", "Timestamp")
                .Font.Bold = True
                .Interior.Color = RGB(16, 107, 193)
                .Font.Color = RGB(255, 255, 255)
            End With
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
' Note: Using CodeName directly (e.g., Sheet12) is preferred if reliable.
' This function remains for cases where name-based lookup is needed.
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
'===============================================================================
Private Sub CleanupDashboardLayout(ws As Worksheet)
    Application.ScreenUpdating = False

    ' No need to Unprotect if UserInterfaceOnly:=True is used during Protect
    ' On Error Resume Next
    ' ws.Unprotect Password:="password" ' REMOVED
    ' On Error GoTo 0

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
        ' Headers correspond to columns:
        ' K - DB_COL_PHASE
        ' L - DB_COL_LASTCONTACT
        ' M - DB_COL_EMAIL
        ' N - DB_COL_COMMENTS
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
' INITIALIZEDASHBOARDLAYOUT: Clears rows 4+ in A-N, sets up header row in A3:N3
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
        ' Headers correspond to columns:
        ' K - DB_COL_PHASE
        ' L - DB_COL_LASTCONTACT
        ' M - DB_COL_EMAIL
        ' N - DB_COL_COMMENTS
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

        .Columns(DB_COL_PHASE).ColumnWidth = 20 ' K
        .Columns(DB_COL_LASTCONTACT).ColumnWidth = 15 ' L
        .Columns(DB_COL_EMAIL).ColumnWidth = 25 ' M
        .Columns(DB_COL_COMMENTS).ColumnWidth = 40 ' N
    End With
End Sub

'===============================================================================
' POPULATEMASTERQUOTESDATA: Pulls columns A-J from MasterQuotes_Final
'===============================================================================
Private Sub PopulateMasterQuotesData(ws As Worksheet)
    ' Use constant for the source name
    Dim sourceName As String
    sourceName = MASTER_QUOTES_FINAL_SOURCE

    With ws
        ' A: Document Number
        .Range("A4").Formula = _
            "=IF(ROWS($A$4:A4)<=ROWS(" & sourceName & "[Document Number])," & _
            "IFERROR(INDEX(" & sourceName & "[Document Number],ROWS($A$4:A4)),""""),"""")"

        ' B: Client ID -> from Customer Number
        .Range("B4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Customer Number],ROWS($A$4:A4)),""""),"""")"

        ' C: Customer Name
        .Range("C4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Customer Name],ROWS($A$4:A4)),""""),"""")"

        ' D: Document Amount
        .Range("D4").Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[Document Amount],ROWS($A$4:A4)),""""),"""")"

        ' E: Document Date
        .Range("E4").Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[Document Date],ROWS($A$4:A4)),""""),"""")"

        ' F: First Date Pulled
        .Range("F4").Formula = _
            "=IF(A4<>"""",IFERROR(--INDEX(" & sourceName & "[First Date Pulled],ROWS($A$4:A4)),""""),"""")"

        ' G: Salesperson ID
        .Range("G4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[Salesperson ID],ROWS($A$4:A4)),""""),"""")"

        ' H: Entered By (was User To Enter)
        .Range("H4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[User To Enter],ROWS($A$4:A4)),""""),"""")"

        ' I: Occurrence Counter (was Auto Stage)
        .Range("I4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[AutoStage],ROWS($A$4:A4)),""""),"""")"

        ' J: Missing Quote Alert (was Auto Note)
        .Range("J4").Formula = _
            "=IF(A4<>"""",IFERROR(INDEX(" & sourceName & "[AutoNote],ROWS($A$4:A4)),""""),"""")"

        ' Autofill down (adjust range if needed)
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
' FREEZEDASHBOARD: Freezes rows 1-3
'===============================================================================
Private Sub FreezeDashboard(ws As Worksheet)
    ws.Activate

    ' First unfreeze any existing splits
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 0
    ActiveWindow.SplitColumn = 0

    ' Freeze rows 1-3
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

    ' No need to Unprotect if UserInterfaceOnly:=True is used during Protect
    ' On Error Resume Next
    ' ws.Unprotect Password:="password" ' REMOVED
    ' On Error GoTo 0

    ' Create buttons with improved spacing and styling
    ModernButton ws, "C2", "Standard Refresh", "Button_RefreshDashboard_SaveAndRestoreEdits"
    ModernButton ws, "E2", "Preserve UserEdits", "Button_RefreshDashboard_PreserveUserEdits"

    ' Protection will be applied in ProtectUserColumns after setting Locked status
    ' ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, _
    '            Scenarios:=True, AllowFormattingCells:=False, _
    '            AllowFormattingColumns:=False, AllowFormattingRows:=False ' REMOVED FROM HERE
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
' PROTECTUSERCOLUMNS: Lock A-J, unlock K-N
'===============================================================================
Public Sub ProtectUserColumns(ws As Worksheet)
    On Error Resume Next ' Ignore errors if sheet is already unprotected
    ws.Unprotect ' Unprotect first
    On Error GoTo 0 ' Resume default error handling

    ws.Cells.Locked = True
    ' Use constants for columns
    ws.Range(DB_COL_PHASE & "4:" & DB_COL_COMMENTS & ws.Rows.Count).Locked = False

    ' Re-apply protection here after setting Locked status
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

'===============================================================================
' APPLYCOLORFORMATTING: For coloring columns I (Occurrence) and K (Engagement)
'===============================================================================
Public Sub ApplyColorFormatting(ws As Worksheet)
    On Error Resume Next ' Ignore errors if sheet is already unprotected
    ws.Unprotect ' Unprotect first
    On Error GoTo 0 ' Resume default error handling

    ' Clear existing rules in I4:I1000 and K4:K1000
    ws.Range("I4:I1000," & DB_COL_PHASE & "4:" & DB_COL_PHASE & "1000").FormatConditions.Delete ' Use Constant for K

    Dim rngOccur As Range, rngPhase As Range
    Set rngOccur = ws.Range("I4:I1000")
    Set rngPhase = ws.Range(DB_COL_PHASE & "4:" & DB_COL_PHASE & "1000") ' Use Constant for K

    ' Apply conditional formatting to both columns
    ApplyStageFormatting rngOccur
    ApplyStageFormatting rngPhase

    ' Re-protect, ensuring UserInterfaceOnly is True
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True
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

    ' 1) Power Query named MASTER_QUOTES_FINAL_SOURCE
    For Each queryObj In ActiveWorkbook.Queries
        If queryObj.Name = MASTER_QUOTES_FINAL_SOURCE Then ' Use Constant
            IsMasterQuotesFinalPresent = True
            Exit For
        End If
    Next queryObj

    ' 2) ListObject named MASTER_QUOTES_FINAL_SOURCE
    If Not IsMasterQuotesFinalPresent Then
        For Each lo In ActiveWorkbook.ListObjects
            If lo.Name = MASTER_QUOTES_FINAL_SOURCE Then ' Use Constant
                IsMasterQuotesFinalPresent = True
                Exit For
            End If
        Next lo
    End If

    ' 3) Named Range MASTER_QUOTES_FINAL_SOURCE
    If Not IsMasterQuotesFinalPresent Then
        For Each nm In ActiveWorkbook.Names
            If nm.Name = MASTER_QUOTES_FINAL_SOURCE Then ' Use Constant
                IsMasterQuotesFinalPresent = True
                Exit For
            End If
        Next nm
    End If

    On Error GoTo 0
End Function

'===============================================================================
' LOADUSEREDITSTODICTIONARY: Loads UserEdits DocNumbers and row numbers into a dictionary
'===============================================================================
Public Function LoadUserEditsToDictionary(wsEdits As Worksheet) As Object ' Changed Private to Public
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive keys

    If wsEdits Is Nothing Then
        Set LoadUserEditsToDictionary = dict
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row

    If lastRow > 1 Then
        Dim i As Long
        Dim docNum As String
        Dim dataRange As Variant
        ' Read all data in one go for performance
        dataRange = wsEdits.Range(wsEdits.Cells(2, UE_COL_DOCNUM), wsEdits.Cells(lastRow, UE_COL_DOCNUM)).Value

        ' Handle case where only one row of data exists (returns a single value, not 2D array)
        If Not IsArray(dataRange) Then
            If lastRow = 2 Then
                docNum = Trim(CStr(dataRange))
                If docNum <> "" Then
                    If Not dict.Exists(docNum) Then
                        dict.Add docNum, 2 ' Store the actual row number
                    End If
                End If
            End If
        Else
            ' Process the 2D array
            For i = 1 To UBound(dataRange, 1) ' Loop through the array
                docNum = Trim(CStr(dataRange(i, 1)))
                If docNum <> "" Then
                    If Not dict.Exists(docNum) Then
                        dict.Add docNum, i + 1 ' Store the actual row number (i + 1 because array is 1-based and starts from row 2)
                    Else
                        ' Handle potential duplicates within UserEdits if necessary (e.g., log warning)
                        ' For now, keep the first one found
                    End If
                End If
            Next i
        End If
    End If

    Set LoadUserEditsToDictionary = dict
End Function

</final_file_content>

IMPORTANT: For any future changes to this file, use the final_file_content shown above as your reference. This content reflects the current state of the file, including any auto-formatting (e.g., if you used single quotes but the formatter converted them to double quotes). Always base your SEARCH/REPLACE operations on this final version to ensure accuracy.

<environment_details>
# VSCode Visible Files
../../../AppData/Local/Programs/Microsoft VS Code/src/ally_workbook/Module_Dashboard.bas
../../../AppData/Local/Programs/Microsoft VS Code/src/ally_workbook/Module_Dashboard.bas
src/ryan_workbook/Module_Dashboard.bas

# VSCode Open Tabs
src/ryan_workbook/Module_Dashboard.bas

# Current Time
4/10/2025, 2:23:10 PM (America/Phoenix, UTC-7:00)

# Context Window Usage
674,075 / 1,048.576K tokens used (64%)

# Current Mode
ACT MODE
</environment_details>

Attribute VB_Name = "Module_SyncTool_Logger"
Option Explicit

'===============================================================================
' MODULE_SYNCTOOL_LOGGER
' Contains logging functions for the SyncTool.
' Manages log files (SyncLog), document history, and conflict display.
'
' New Enhancement:
' A dedicated ErrorLog sheet is added to capture only error messages.
' This allows errors to be reviewed separately from general log messages.
'===============================================================================

'===============================================================================
' INITIALIZE_SYNC_LOG - Creates and formats the sync log sheet if needed.
'===============================================================================
Public Sub InitializeSyncLog()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    On Error Resume Next  ' Use Resume Next to handle sheet not found
    Set ws = ThisWorkbook.Sheets(SYNCTOOL_LOG_SHEET)
    On Error GoTo ErrorHandler  ' Restore error handling
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SYNCTOOL_LOG_SHEET
        
        ' Set up header
        Dim headerRange As Range
        Set headerRange = ws.Range("A1:C1")
        FormatHeaders ws, headerRange, Array("Timestamp", "Status", "Message")
        
        ' Set up column widths
        ws.Columns("A").ColumnWidth = 20
        ws.Columns("B").ColumnWidth = 15
        ws.Columns("C").ColumnWidth = 120
    End If
    
    ' Make sure the log sheet is visible
    ws.Visible = xlSheetVisible
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in InitializeSyncLog: " & err.Description, vbExclamation, "Logger Error"
End Sub

'===============================================================================
' LOG_MESSAGE - Records an entry in the synchronization log.
'===============================================================================
Public Sub LogMessage(message As String, Optional status As String = "INFO")
    On Error GoTo ErrorHandler
    
    ' Ensure log is initialized
    InitializeSyncLog
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SYNCTOOL_LOG_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    If lastRow < 1 Then lastRow = 1
    
    ' Add new log entry
    ws.Cells(lastRow + 1, "A").value = Format$(Now(), FORMAT_TIMESTAMP)
    ws.Cells(lastRow + 1, "B").value = status
    ws.Cells(lastRow + 1, "C").value = message
    
    ' Format row based on status
    Select Case UCase(status)
        Case "ERROR"
            ws.Range(ws.Cells(lastRow + 1, "A"), ws.Cells(lastRow + 1, "C")).Interior.Color = COLOR_ERROR
        Case "WARNING"
            ws.Range(ws.Cells(lastRow + 1, "A"), ws.Cells(lastRow + 1, "C")).Interior.Color = COLOR_WARNING
        Case Else
            ' INFO messages use default formatting
    End Select
    
    ws.Columns("A:C").AutoFit
    
    Debug.Print Format$(Now(), FORMAT_TIMESTAMP) & " - " & status & " - " & message
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in LogMessage: " & err.Description, vbExclamation, "Logger Error"
End Sub

'===============================================================================
' INITIALIZE_ERROR_LOG - Creates and formats the dedicated error log sheet if needed.
'
' Purpose:
' Checks for an ErrorLog sheet (named ERROR_LOG_SHEET). If it does not exist,
' creates a new sheet with headers for Timestamp, Error Code, Error Description, and Module.
'
' Returns:
' Nothing.
'===============================================================================
Public Sub InitializeErrorLog()
    On Error Resume Next  ' Use Resume Next to handle sheet not found
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ERROR_LOG_SHEET)
    On Error GoTo ErrorHandler  ' Restore error handling
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = ERROR_LOG_SHEET
        
        ' Set up header row
        With ws.Range("A1:D1")
            .value = Array("Timestamp", "Error Code", "Error Description", "Module")
            .Font.Bold = True
            .Interior.Color = RGB(255, 199, 206) ' Light red for errors
            .Font.Color = RGB(255, 255, 255) ' White text
        End With
        
        ws.Columns("A:D").AutoFit
    End If
    
    ' Set the ErrorLog sheet to visible
    ws.Visible = xlSheetVisible
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error initializing ErrorLog: " & err.Description, vbCritical, "ErrorLog Initialization Error"
End Sub

'===============================================================================
' LOG_ERROR - Records an error entry in the dedicated error log sheet.
'
' Parameters:
' message - A string describing the error.
' Optional moduleName - A string identifying the module or procedure where the error occurred.
'
' Returns:
' Nothing.
'===============================================================================
Public Sub LogError(message As String, Optional moduleName As String = "Unknown")
    On Error GoTo ErrorHandler
    
    ' Initialize the error log
    InitializeErrorLog
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ERROR_LOG_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    If lastRow < 1 Then lastRow = 1
    
    ws.Cells(lastRow + 1, "A").value = Format$(Now(), FORMAT_TIMESTAMP)
    ws.Cells(lastRow + 1, "B").value = err.Number
    ws.Cells(lastRow + 1, "C").value = message
    ws.Cells(lastRow + 1, "D").value = moduleName
    
    ws.Columns("A:D").AutoFit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in LogError: " & err.Description, vbCritical, "ErrorLog Error"
End Sub

'===============================================================================
' GET_MERGE_DATA_SHEET - Returns a reference to the MergeData sheet, creating it if needed.
'===============================================================================
Public Function GetMergeDataSheet() As Worksheet
    On Error Resume Next
    Set GetMergeDataSheet = ThisWorkbook.Sheets(SYNCTOOL_MERGEDATA_SHEET)
    On Error GoTo 0
    
    If GetMergeDataSheet Is Nothing Then
        Set GetMergeDataSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetMergeDataSheet.Name = SYNCTOOL_MERGEDATA_SHEET
    End If
End Function

'===============================================================================
' GET_DOC_HISTORY_SHEET - Returns a reference to the DocChangeHistory sheet, creating it if needed.
'===============================================================================
Public Function GetDocHistorySheet() As Worksheet
    On Error Resume Next
    Set GetDocHistorySheet = ThisWorkbook.Sheets(SYNCTOOL_HISTORY_SHEET)
    On Error GoTo 0
    
    If GetDocHistorySheet Is Nothing Then
        Set GetDocHistorySheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetDocHistorySheet.Name = SYNCTOOL_HISTORY_SHEET
        
        Dim headerRange As Range
        Set headerRange = GetDocHistorySheet.Range("A1:H1")
        FormatHeaders GetDocHistorySheet, headerRange, Array("Document Number", "Last Sync Date", "Change Source", _
                                                           "Engagement Phase", "Last Contact Date", "Email Contact", _
                                                           "User Comments", "Conflict Resolved")
    End If
End Function


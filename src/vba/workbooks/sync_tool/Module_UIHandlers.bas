Attribute VB_Name = "Module_UIHandlers"
Option Explicit

'===============================================================================
' MODULE_UIHANDLERS
' Contains all public macros for button-click events on the SQRCT Sync Tool Dashboard.
' Each macro calls into other modules (File Processor, Logger, etc.) to do the work.
' This separation ensures the UI layer is distinct from business logic.
'
' Best Practice:
' 1. Keep these macros public so Excel can assign them to shapes or form controls.
' 2. Reference other modules (Module_File_Processor, Module_SyncTool_Logger, etc.)
'    for the actual logic to avoid code duplication.
'===============================================================================

'-------------------------------------------------------------------------------
' BROWSE_ALLY_FILE_CLICK
' Invoked by the "Browse" button next to "Ally's Working File."
' Opens a file dialog and stores the chosen path in the dashboard cell.
'-------------------------------------------------------------------------------
Public Sub BrowseAllyFile_Click()
    On Error GoTo ErrorHandler
    
    Dim wsDashboard As Worksheet
    Dim filePath As String
    Dim initialFolder As String
    
    ' 1. Reference the dashboard sheet
    Set wsDashboard = Module_SyncTool_UI.GetSyncToolDashboard()
    
    If wsDashboard Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Could not access dashboard in BrowseAllyFile_Click", "ERROR"
        Exit Sub
    End If
    
    ' 2. Attempt to derive an initial folder from any existing path
    initialFolder = ""
    
    If Not Module_Utilities.IsNullOrEmpty(wsDashboard.Range(CELL_ALLY_PATH).value) Then
        initialFolder = Left(wsDashboard.Range(CELL_ALLY_PATH).value, InStrRev(wsDashboard.Range(CELL_ALLY_PATH).value, "\"))
    End If
    
    ' 3. Open file dialog
    filePath = Module_StartUp.BrowseForFile(initialFolder, "Select Ally's SQRCT Excel File")
    
    ' 4. If user selected a file, store it in the dashboard cell
    If Not Module_Utilities.IsNullOrEmpty(filePath) Then
        wsDashboard.Range(CELL_ALLY_PATH).value = filePath
        Module_SyncTool_Logger.LogMessage "Ally's file set to: " & filePath
    End If
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "Error in BrowseAllyFile_Click: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "Error selecting Ally's file: " & err.Description, vbExclamation, "Browse Ally File Error"
End Sub

'-------------------------------------------------------------------------------
' BROWSE_RYAN_FILE_CLICK
' Similar logic for Ryan's Working File
'-------------------------------------------------------------------------------
Public Sub BrowseRyanFile_Click()
    On Error GoTo ErrorHandler
    
    Dim wsDashboard As Worksheet
    Dim filePath As String
    Dim initialFolder As String
    
    Set wsDashboard = Module_SyncTool_UI.GetSyncToolDashboard()
    
    If wsDashboard Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Could not access dashboard in BrowseRyanFile_Click", "ERROR"
        Exit Sub
    End If
    
    initialFolder = ""
    
    If Not Module_Utilities.IsNullOrEmpty(wsDashboard.Range(CELL_RYAN_PATH).value) Then
        initialFolder = Left(wsDashboard.Range(CELL_RYAN_PATH).value, InStrRev(wsDashboard.Range(CELL_RYAN_PATH).value, "\"))
    End If
    
    filePath = Module_StartUp.BrowseForFile(initialFolder, "Select Ryan's SQRCT Excel File")
    
    If Not Module_Utilities.IsNullOrEmpty(filePath) Then
        wsDashboard.Range(CELL_RYAN_PATH).value = filePath
        Module_SyncTool_Logger.LogMessage "Ryan's file set to: " & filePath
    End If
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "Error in BrowseRyanFile_Click: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "Error selecting Ryan's file: " & err.Description, vbExclamation, "Browse Ryan File Error"
End Sub

'-------------------------------------------------------------------------------
' BROWSE_MASTER_FILE_CLICK
' For the "Automated Master File" browse button
'-------------------------------------------------------------------------------
Public Sub BrowseMasterFile_Click()
    On Error GoTo ErrorHandler
    
    Dim wsDashboard As Worksheet
    Dim filePath As String
    Dim initialFolder As String
    
    Set wsDashboard = Module_SyncTool_UI.GetSyncToolDashboard()
    
    If wsDashboard Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Could not access dashboard in BrowseMasterFile_Click", "ERROR"
        Exit Sub
    End If
    
    initialFolder = ""
    
    If Not Module_Utilities.IsNullOrEmpty(wsDashboard.Range(CELL_MASTER_PATH).value) Then
        initialFolder = Left(wsDashboard.Range(CELL_MASTER_PATH).value, InStrRev(wsDashboard.Range(CELL_MASTER_PATH).value, "\"))
    End If
    
    filePath = Module_StartUp.BrowseForFile(initialFolder, "Select Master SQRCT Excel File")
    
    If Not Module_Utilities.IsNullOrEmpty(filePath) Then
        wsDashboard.Range(CELL_MASTER_PATH).value = filePath
        Module_SyncTool_Logger.LogMessage "Master file set to: " & filePath
    End If
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "Error in BrowseMasterFile_Click: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "Error selecting Master file: " & err.Description, vbExclamation, "Browse Master File Error"
End Sub

'-------------------------------------------------------------------------------
' DIAGNOSE_FILES_CLICK
' Invoked by the "Diagnose Files" button.
' Uses FileProcessor to standardize and check the user edits in all selected files.
'-------------------------------------------------------------------------------
Public Sub DiagnoseFiles_Click()
    On Error GoTo ErrorHandler
    
    Dim wsDashboard As Worksheet
    Set wsDashboard = Module_SyncTool_UI.GetSyncToolDashboard()
    
    If wsDashboard Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Could not access dashboard in DiagnoseFiles_Click", "ERROR"
        Exit Sub
    End If
    
    Dim allyFilePath As String, ryanFilePath As String, masterFilePath As String
    
    allyFilePath = wsDashboard.Range(CELL_ALLY_PATH).value
    ryanFilePath = wsDashboard.Range(CELL_RYAN_PATH).value
    masterFilePath = wsDashboard.Range(CELL_MASTER_PATH).value
    
    ' Validate file paths
    If Not Module_File_Processor.ValidateFilePaths(allyFilePath, ryanFilePath, masterFilePath) Then
        Exit Sub
    End If
    
    Module_SyncTool_Logger.LogMessage "===== Beginning File Diagnostics: " & Format$(Now(), FORMAT_TIMESTAMP) & " ====="
    Module_SyncTool_UI.UpdateStatusDisplay "Diagnosing files..."
    
    ' Standardize the user edits
    Module_File_Processor.StandardizeUserEditsSheet allyFilePath, ATTRIBUTION_ALLY
    Module_File_Processor.StandardizeUserEditsSheet ryanFilePath, ATTRIBUTION_RYAN
    Module_File_Processor.StandardizeUserEditsSheet masterFilePath, ATTRIBUTION_MASTER
    
    Module_SyncTool_UI.UpdateStatusDisplay "Diagnostics completed"
    Module_SyncTool_Logger.LogMessage "===== File Diagnostics Completed: " & Format$(Now(), FORMAT_TIMESTAMP) & " ====="
    
    MsgBox "Diagnostics complete! Please check the SyncLog for any issues found.", vbInformation, "Diagnostics"
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_UI.UpdateStatusDisplay "Error diagnosing files: " & err.Description
    Module_SyncTool_Logger.LogMessage "Error in DiagnoseFiles_Click: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "An error occurred during file diagnostics: " & err.Description, vbExclamation, "Diagnostics Error"
End Sub

'-------------------------------------------------------------------------------
' SHOW_CONFLICTS_CLICK
' Invoked by the "Show Conflicts" button to identify potential conflicts before sync.
'-------------------------------------------------------------------------------
Public Sub ShowConflicts_Click()
    On Error GoTo ErrorHandler
    
    Dim wsDashboard As Worksheet
    Set wsDashboard = Module_SyncTool_UI.GetSyncToolDashboard()
    
    If wsDashboard Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Could not access dashboard in ShowConflicts_Click", "ERROR"
        Exit Sub
    End If
    
    Dim allyFilePath As String, ryanFilePath As String, masterFilePath As String
    
    allyFilePath = wsDashboard.Range(CELL_ALLY_PATH).value
    ryanFilePath = wsDashboard.Range(CELL_RYAN_PATH).value
    masterFilePath = wsDashboard.Range(CELL_MASTER_PATH).value
    
    If Not Module_File_Processor.ValidateFilePaths(allyFilePath, ryanFilePath, masterFilePath) Then
        Exit Sub
    End If
    
    Module_SyncTool_UI.UpdateStatusDisplay "Looking for potential conflicts..."
    Module_SyncTool_Logger.LogMessage "===== Beginning Conflict Detection: " & Format$(Now(), FORMAT_TIMESTAMP) & " ====="
    
    Dim allyData As Object, ryanData As Object, masterData As Object
    
    Set allyData = Module_File_Processor.ExtractUserEdits(allyFilePath, ATTRIBUTION_ALLY)
    Set ryanData = Module_File_Processor.ExtractUserEdits(ryanFilePath, ATTRIBUTION_RYAN)
    Set masterData = Module_File_Processor.ExtractUserEdits(masterFilePath, ATTRIBUTION_MASTER)
    
    Dim dataMap As Object
    Set dataMap = CreateObject("Scripting.Dictionary")
    
    dataMap.Add ATTRIBUTION_ALLY, allyData
    dataMap.Add ATTRIBUTION_RYAN, ryanData
    dataMap.Add ATTRIBUTION_MASTER, masterData
    
    Dim conflicts As Object
    Set conflicts = Module_Conflict_Handler.DetectConflicts(dataMap)
    
    ' Show them in MergeData
    On Error Resume Next
    Dim wsMergeData As Worksheet
    Set wsMergeData = Module_SyncTool_Logger.GetMergeDataSheet()
    
    If Not wsMergeData Is Nothing Then
        wsMergeData.Activate
        Module_Conflict_Handler.DisplayConflicts conflicts
    End If
    On Error GoTo ErrorHandler
    
    Module_SyncTool_UI.UpdateStatusDisplay "Conflict detection completed"
    Module_SyncTool_Logger.LogMessage "===== Conflict Detection Completed: " & Format$(Now(), FORMAT_TIMESTAMP) & " ====="
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_UI.UpdateStatusDisplay "Error detecting conflicts: " & err.Description
    Module_SyncTool_Logger.LogMessage "Error in ShowConflicts_Click: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "An error occurred while detecting conflicts: " & err.Description, vbExclamation, "Conflict Detection Error"
End Sub

'-------------------------------------------------------------------------------
' VIEW_LOG_CLICK
' Invoked by the "View Log" button to jump to the SyncLog sheet.
'-------------------------------------------------------------------------------
Public Sub ViewLog_Click()
    On Error GoTo ErrorHandler
    
    Module_SyncTool_Logger.InitializeSyncLog
    
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SYNCTOOL_LOG_SHEET)
    
    If Not ws Is Nothing Then
        ws.Activate
    End If
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "Error in ViewLog_Click: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "Error displaying SyncLog: " & err.Description, vbExclamation, "View Log Error"
End Sub

'-------------------------------------------------------------------------------
' SYNC_CLICK
' Invoked by the "Sync" button to start the synchronization process.
' For a one-way sync, we only merge into the Master.
'-------------------------------------------------------------------------------
Public Sub Sync_Click()
    On Error GoTo ErrorHandler
    
    ' Call the main synchronization routine in Module_SyncTool_Manager
    ' Pass 'False' to skip writing back to AF/RZ (one-way sync)
    Module_SyncTool_Manager.StartSynchronization False
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_UI.UpdateStatusDisplay "Error during synchronization: " & err.Description
    Module_SyncTool_Logger.LogMessage "Error in Sync_Click: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "An error occurred during synchronization: " & err.Description, vbExclamation, "Sync Error"
End Sub

'-------------------------------------------------------------------------------
' CREATE_SYNC_TOOL_DASHBOARD
' If you also want a button to re-create or refresh the dashboard, you can place it here.
'-------------------------------------------------------------------------------
Public Sub CreateSyncToolDashboard()
    On Error GoTo ErrorHandler
    
    Module_SyncTool_UI.CreateSyncToolDashboard
    
    MsgBox "SyncTool Dashboard has been refreshed!", vbInformation, "Dashboard Refreshed"
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "Error in CreateSyncToolDashboard: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "Error creating dashboard: " & err.Description, vbExclamation, "Dashboard Error"
End Sub

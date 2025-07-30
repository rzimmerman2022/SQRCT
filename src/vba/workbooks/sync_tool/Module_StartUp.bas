Attribute VB_Name = "Module_StartUp"
Option Explicit

'===============================================================================
' MODULE_STARTUP
'-------------------------------------------------------------------------------
' Purpose:
' Provides safe application startup and initialization by handling workbook
' open events and deferring initialization until all modules are loaded.
'===============================================================================

'===============================================================================
' SAFE_STARTUP - Entry point for deferred initialization.
'
' Purpose:
' Initializes the logging systems, ensures the SyncTool dashboard exists,
' and clears the application status. This routine is scheduled to run shortly
' after the workbook opens.
'
' Returns:
' Nothing.
'===============================================================================
Public Sub SafeStartup()
    On Error GoTo ErrorHandler
    
    ' Initialize the SyncLog and, if desired, the ErrorLog.
    Module_SyncTool_Logger.InitializeSyncLog
    
    ' To automatically create an ErrorLog that is visible, call:
    ' Module_SyncTool_Logger.InitializeErrorLog
    ' (Uncomment the following line if you want the ErrorLog auto-created and visible)
    Module_SyncTool_Logger.InitializeErrorLog
    
    Module_SyncTool_Logger.LogMessage "===== Application started safely: " & Format$(Now(), FORMAT_TIMESTAMP) & " ====="
    
    ' Check if the SyncTool dashboard exists. If not, create it.
    Dim hasDashboard As Boolean
    On Error Resume Next
    hasDashboard = Not ThisWorkbook.Sheets(SYNCTOOL_DASHBOARD_SHEET) Is Nothing
    On Error GoTo ErrorHandler
    
    If Not hasDashboard Then
        Module_SyncTool_Logger.LogMessage "Creating SyncTool dashboard..."
        Module_SyncTool_UI.CreateSyncToolDashboard
    End If
    
    ' Clear the application status.
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    Module_SyncTool_Logger.LogMessage "Startup Error: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "An error occurred during startup: " & err.Description, vbCritical, "Startup Error"
End Sub

'===============================================================================
' SYNC_TOOL_AUTO_OPEN - Entry point for VBA Auto_Open.
'
' Purpose:
' Schedules deferred initialization by calling SafeStartup one second after the
' workbook opens, ensuring all modules are loaded.
'
' Returns:
' Nothing.
'===============================================================================
Public Sub SyncTool_Auto_Open()
    On Error GoTo ErrorHandler
    
    ' Schedule SafeStartup to run one second after workbook open.
    Application.OnTime Now + TimeValue("00:00:01"), "SafeStartup"
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    Module_SyncTool_Logger.LogMessage "Auto_Open Error: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "An error occurred during Auto_Open: " & err.Description, vbCritical, "Auto_Open Error"
End Sub

'===============================================================================
' BROWSE_FOR_FILE - Opens a file dialog to allow the user to select an Excel file.
'
' Parameters:
' Optional initialFolder - The folder to initially display in the dialog.
' Optional title - The title of the file dialog.
'
' Returns:
' A string representing the selected file path, or an empty string if cancelled
' or an error occurs.
'===============================================================================
Public Function BrowseForFile(Optional initialFolder As String = "", Optional title As String = "Select an SQRCT Excel File") As String
    Dim fd As Office.FileDialog
    
    On Error GoTo ErrorHandler
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' If the file dialog cannot be created, log the error and exit.
    If fd Is Nothing Then
        Module_SyncTool_Logger.LogMessage "Error initializing file dialog.", "ERROR"
        BrowseForFile = ""
        Exit Function
    End If
    
    With fd
        .title = title
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm", 1
        .AllowMultiSelect = False
        
        ' Set the initial folder if provided.
        If Not Module_Utilities.IsNullOrEmpty(initialFolder) Then
            .InitialFileName = initialFolder
        End If
        
        If .Show = True Then
            BrowseForFile = .SelectedItems(1)
        Else
            BrowseForFile = ""
        End If
    End With
    
    Exit Function
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "BrowseForFile Error: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "Error opening file dialog: " & err.Description, vbExclamation, "File Dialog Error"
    BrowseForFile = ""
End Function

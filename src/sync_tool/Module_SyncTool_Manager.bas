Attribute VB_Name = "Module_SyncTool_Manager"
Option Explicit

'===============================================================================
' MODULE_SYNCTOOL_MANAGER
'-------------------------------------------------------------------------------
' Purpose:
' Coordinates the high-level synchronization workflow. It extracts data from
' Ally's (AF) and Ryan's (RZ) workbooks as well as from the current Master,
' merges and resolves conflicts, and writes the consolidated, Master-approved
' data only to the Automated Master workbook.
'===============================================================================

Public Sub StartSynchronization(Optional ByVal allowWriteBackToUsers As Boolean = False)
    On Error GoTo ErrorHandler
    
    ' Initialize logging (both SyncLog and ErrorLog, if implemented)
    Module_SyncTool_Logger.InitializeSyncLog
    Module_SyncTool_Logger.InitializeErrorLog ' Creates ErrorLog (visible if desired)
    
    Module_SyncTool_Logger.LogMessage "===== Beginning Synchronization: " & Format$(Now(), FORMAT_TIMESTAMP) & " ====="
    
    ' Retrieve and validate file paths from the dashboard.
    Dim filePaths As Object
    Set filePaths = GetValidatedFilePaths()
    
    If filePaths Is Nothing Then
        Module_SyncTool_UI.UpdateStatusDisplay "Synchronization cancelled: Invalid file paths"
        Exit Sub
    End If
    
    ' Standardize the UserEdits sheets in all files (AF, RZ, Master)
    Module_SyncTool_Logger.LogMessage "Standardizing UserEdits sheets..."
    Module_SyncTool_UI.UpdateStatusDisplay "Standardizing UserEdits sheets..."
    
    Module_File_Processor.StandardizeUserEditsSheet filePaths("Ally"), ATTRIBUTION_ALLY
    Module_File_Processor.StandardizeUserEditsSheet filePaths("Ryan"), ATTRIBUTION_RYAN
    Module_File_Processor.StandardizeUserEditsSheet filePaths("Master"), ATTRIBUTION_MASTER
    
    ' Extract data from all three workbooks.
    Module_SyncTool_UI.UpdateStatusDisplay "Extracting UserEdits data..."
    
    Dim allyData As Object, ryanData As Object, masterData As Object, dataMap As Object
    
    Set allyData = Module_File_Processor.ExtractUserEdits(filePaths("Ally"), ATTRIBUTION_ALLY)
    Set ryanData = Module_File_Processor.ExtractUserEdits(filePaths("Ryan"), ATTRIBUTION_RYAN)
    Set masterData = Module_File_Processor.ExtractUserEdits(filePaths("Master"), ATTRIBUTION_MASTER)
    
    ' Create a unified data map.
    Set dataMap = CreateObject("Scripting.Dictionary")
    dataMap.Add ATTRIBUTION_ALLY, allyData
    dataMap.Add ATTRIBUTION_RYAN, ryanData
    dataMap.Add ATTRIBUTION_MASTER, masterData
    
    ' Detect conflicts between data sources.
    Module_SyncTool_UI.UpdateStatusDisplay "Detecting conflicts..."
    Module_SyncTool_Logger.LogMessage "Detecting conflicts across data sources..."
    
    Dim conflicts As Object
    Set conflicts = Module_Conflict_Handler.DetectConflicts(dataMap)
    
    Module_SyncTool_Logger.LogMessage "Found " & conflicts.Count & " conflicts"
    
    ' If conflicts exist, display them for review.
    If conflicts.Count > 0 Then
        Dim wsMergeData As Worksheet
        
        On Error Resume Next
        Set wsMergeData = Module_SyncTool_Logger.GetMergeDataSheet()
        On Error GoTo ErrorHandler
        
        If Not wsMergeData Is Nothing Then
            wsMergeData.Visible = xlSheetVisible
            wsMergeData.Activate
            
            Module_Conflict_Handler.DisplayConflicts conflicts
            
            Dim response As VbMsgBoxResult
            response = MsgBox("Conflicts have been detected and are displayed in the MergeData sheet." & vbCrLf & _
                             "Review the conflicts before continuing." & vbCrLf & vbCrLf & _
                             "Do you want to continue with the synchronization?", _
                             vbQuestion + vbYesNo, "Conflicts Detected")
            
            If response = vbNo Then
                Module_SyncTool_UI.UpdateStatusDisplay "Synchronization cancelled by user"
                Module_SyncTool_Logger.LogMessage "Synchronization cancelled by user after conflict review"
                Exit Sub
            End If
        End If
    End If
    
    ' Merge all data (applying conflict resolution) into a consolidated dataset.
    Module_SyncTool_UI.UpdateStatusDisplay "Merging data..."
    
    Dim mergedData As Object
    Set mergedData = Module_Conflict_Handler.MergeUserEdits(dataMap, conflicts)
    
    ' Write the merged data exclusively to the Master workbook.
    Module_SyncTool_UI.UpdateStatusDisplay "Writing merged data to Master..."
    Module_File_Processor.WriteUserEditsToWorkbook filePaths("Master"), mergedData, False
    
    ' Optionally, if you want to allow controlled write-back to AF/RZ in the future:
    If allowWriteBackToUsers Then
        Module_SyncTool_Logger.LogMessage "Writing updates to Ally and Ryan workbooks..."
        Module_File_Processor.WriteUserEditsToWorkbook filePaths("Ally"), mergedData, True
        Module_File_Processor.WriteUserEditsToWorkbook filePaths("Ryan"), mergedData, True
    Else
        Module_SyncTool_Logger.LogMessage "Skipping write-back to Ally/Ryan (raw user files remain unchanged)"
    End If
    
    ' Trigger refresh in the Master workbook only.
    Module_SyncTool_UI.UpdateStatusDisplay "Triggering refresh for Master..."
    Module_File_Processor.TriggerExternalWorkbookRefresh filePaths("Master")
    
    ' Update document history in the Master.
    UpdateDocumentHistory mergedData, conflicts
    
    ' Finalize synchronization.
    Module_SyncTool_UI.UpdateStatusDisplay "Synchronization complete!"
    Module_SyncTool_Logger.LogMessage "===== Synchronization completed: " & Format$(Now(), FORMAT_TIMESTAMP) & " ====="
    
    ' Update the Last Sync timestamp on the dashboard.
    On Error Resume Next
    Dim wsDashboard As Worksheet
    Set wsDashboard = Module_SyncTool_UI.GetSyncToolDashboard()
    
    If Not wsDashboard Is Nothing Then
        wsDashboard.Range(CELL_LAST_SYNC_TIME).value = Format(Now(), FORMAT_TIMESTAMP)
    End If
    On Error GoTo ErrorHandler
    
    MsgBox "Synchronization to Master is complete! AF and RZ files remain unaltered." & vbCrLf & vbCrLf & _
           conflicts.Count & " conflicts were resolved." & vbCrLf & _
           mergedData.Count & " documents were merged.", vbInformation, "Synchronization Complete"
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_UI.UpdateStatusDisplay "Error during synchronization: " & err.Description
    Module_SyncTool_Logger.LogMessage "ERROR in StartSynchronization: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    MsgBox "An error occurred during synchronization: " & err.Description, vbCritical, "Synchronization Error"
End Sub

'===============================================================================
' GET_VALIDATED_FILE_PATHS - Retrieves and validates file paths from the dashboard.
'
' Returns:
' A Dictionary containing validated file paths for "Ally", "Ryan", and "Master".
' Returns Nothing if validation fails.
'===============================================================================
Private Function GetValidatedFilePaths() As Object
    On Error GoTo ErrorHandler
    
    Dim filePaths As Object
    Set filePaths = CreateObject("Scripting.Dictionary")
    
    Dim wsDashboard As Worksheet
    Set wsDashboard = Module_SyncTool_UI.GetSyncToolDashboard()
    
    If wsDashboard Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Could not access dashboard sheet for file paths", "ERROR"
        Set GetValidatedFilePaths = Nothing
        Exit Function
    End If
    
    Dim allyPath As String, ryanPath As String, masterPath As String
    
    allyPath = Trim(wsDashboard.Range(CELL_ALLY_PATH).value)
    ryanPath = Trim(wsDashboard.Range(CELL_RYAN_PATH).value)
    masterPath = Trim(wsDashboard.Range(CELL_MASTER_PATH).value)
    
    ' Validate all paths
    If Not Module_File_Processor.ValidateFilePaths(allyPath, ryanPath, masterPath) Then
        Set GetValidatedFilePaths = Nothing
        Exit Function
    End If
    
    ' Store validated paths in the dictionary
    filePaths.Add "Ally", allyPath
    filePaths.Add "Ryan", ryanPath
    filePaths.Add "Master", masterPath
    
    Module_SyncTool_Logger.LogMessage "File paths validated successfully:"
    Module_SyncTool_Logger.LogMessage "- Ally: " & allyPath
    Module_SyncTool_Logger.LogMessage "- Ryan: " & ryanPath
    Module_SyncTool_Logger.LogMessage "- Master: " & masterPath
    
    Set GetValidatedFilePaths = filePaths
    
    Exit Function
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "ERROR in GetValidatedFilePaths: " & err.Description & " (Error " & err.Number & ")", "ERROR"
    Set GetValidatedFilePaths = Nothing
End Function

'===============================================================================
' UPDATE_DOCUMENT_HISTORY - Records sync history for documents.
'
' Parameters:
' mergedData - Dictionary of merged document data.
' conflicts - Dictionary of conflict information.
'
' Purpose:
' Records all synchronized documents in the history sheet, noting which ones
' had conflicts and their resolution strategies.
'===============================================================================
Private Sub UpdateDocumentHistory(mergedData As Object, conflicts As Object)
    On Error GoTo ErrorHandler
    
    If mergedData Is Nothing Then Exit Sub
    
    Dim wsHistory As Worksheet
    On Error Resume Next
    Set wsHistory = Module_SyncTool_Logger.GetDocHistorySheet()
    On Error GoTo ErrorHandler
    
    If wsHistory Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Could not access document history sheet", "ERROR"
        Exit Sub
    End If
    
    ' Clear existing data (apart from header row)
    Dim lastRow As Long
    lastRow = wsHistory.Cells(wsHistory.Rows.Count, "A").End(xlUp).row
    
    If lastRow > 1 Then
        wsHistory.Range("A2:H" & lastRow).Clear
    End If
    
    ' Write history entries
    Dim docNum As Variant
    Dim rowData As Object
    Dim destRow As Long
    
    destRow = 2
    
    For Each docNum In mergedData.keys
        Set rowData = mergedData(docNum)
        
        wsHistory.Cells(destRow, "A").value = docNum
        wsHistory.Cells(destRow, "B").value = Format(Now(), FORMAT_TIMESTAMP)
        wsHistory.Cells(destRow, "C").value = rowData("ChangeSource")
        wsHistory.Cells(destRow, "D").value = rowData("EngagementPhase")
        wsHistory.Cells(destRow, "E").value = rowData("LastContactDate")
        wsHistory.Cells(destRow, "F").value = rowData("EmailContact")
        wsHistory.Cells(destRow, "G").value = rowData("UserComments")
        
        ' Mark if this document had a conflict
        wsHistory.Cells(destRow, "H").value = conflicts.Exists(docNum)
        
        ' Apply conditional formatting
        If conflicts.Exists(docNum) Then
            wsHistory.Range(wsHistory.Cells(destRow, "A"), wsHistory.Cells(destRow, "H")).Interior.Color = COLOR_WARNING
        End If
        
        destRow = destRow + 1
    Next docNum
    
    ' Format the sheet
    wsHistory.Columns("A:H").AutoFit
    wsHistory.Activate
    
    Module_SyncTool_Logger.LogMessage "Document history updated with " & mergedData.Count & " records"
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "ERROR in UpdateDocumentHistory: " & err.Description & " (Error " & err.Number & ")", "ERROR"
End Sub


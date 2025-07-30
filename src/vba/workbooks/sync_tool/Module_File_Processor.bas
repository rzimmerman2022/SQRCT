Attribute VB_Name = "Module_File_Processor"
Option Explicit

'===============================================================================
' MODULE_FILE_PROCESSOR
'-------------------------------------------------------------------------------
' Purpose:
' Contains functions for file operations and data extraction.
' Handles opening/closing files and reading/writing data from external workbooks.
'===============================================================================

'===============================================================================
' FILE_EXISTS - Checks if a file exists at the specified path.
'
' Parameters:
' filePath - A string containing the full path of the file.
'
' Returns:
' Boolean - True if the file exists; False otherwise.
'===============================================================================
Public Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    
    ' Handle empty path.
    If Module_Utilities.IsNullOrEmpty(filePath) Then
        FileExists = False
        Module_SyncTool_Logger.LogMessage "WARNING: Empty file path provided to FileExists function", "WARNING"
        Exit Function
    End If
    
    ' Use the Dir function to check for file existence.
    FileExists = (Dir(filePath) <> "")
    
    ' If an error occurs (e.g., network path issues), log it.
    If err.Number <> 0 Then
        FileExists = False
        Module_SyncTool_Logger.LogMessage "WARNING: Error checking file existence: " & err.Description & " (Error " & err.Number & ")", "WARNING"
    End If
    
    On Error GoTo 0
End Function

'===============================================================================
' VALIDATE_FILE_PATHS - Verifies that all file paths are valid and accessible.
'
' Parameters:
' allyFilePath - Full path for Ally's workbook.
' ryanFilePath - Full path for Ryan's workbook.
' masterFilePath - Full path for the Master workbook.
'
' Returns:
' Boolean - True if all paths are valid; False otherwise.
'===============================================================================
Public Function ValidateFilePaths(allyFilePath As String, ryanFilePath As String, masterFilePath As String) As Boolean
    Dim errorMessages As New Collection
    Dim overallResult As Boolean
    
    overallResult = True
    
    ' Validate each file path and accumulate errors.
    If Module_Utilities.IsNullOrEmpty(allyFilePath) Then
        errorMessages.Add "- Ally's file path is empty. Please select a valid file."
        overallResult = False
    ElseIf Not FileExists(allyFilePath) Then
        errorMessages.Add "- Ally's file not found: " & allyFilePath
        overallResult = False
    End If
    
    If Module_Utilities.IsNullOrEmpty(ryanFilePath) Then
        errorMessages.Add "- Ryan's file path is empty. Please select a valid file."
        overallResult = False
    ElseIf Not FileExists(ryanFilePath) Then
        errorMessages.Add "- Ryan's file not found: " & ryanFilePath
        overallResult = False
    End If
    
    If Module_Utilities.IsNullOrEmpty(masterFilePath) Then
        errorMessages.Add "- Master file path is empty. Please select a valid file."
        overallResult = False
    ElseIf Not FileExists(masterFilePath) Then
        errorMessages.Add "- Master file not found: " & masterFilePath
        overallResult = False
    End If
    
    ' Check for duplicate file paths.
    If allyFilePath = ryanFilePath And Not Module_Utilities.IsNullOrEmpty(allyFilePath) Then
        errorMessages.Add Module_Utilities.FormatErrorMessage(ERR_DUPLICATE_PATH, "Ally's", "Ryan's")
        overallResult = False
    End If
    
    If allyFilePath = masterFilePath And Not Module_Utilities.IsNullOrEmpty(allyFilePath) Then
        errorMessages.Add Module_Utilities.FormatErrorMessage(ERR_DUPLICATE_PATH, "Ally's", "Master")
        overallResult = False
    End If
    
    If ryanFilePath = masterFilePath And Not Module_Utilities.IsNullOrEmpty(ryanFilePath) Then
        errorMessages.Add Module_Utilities.FormatErrorMessage(ERR_DUPLICATE_PATH, "Ryan's", "Master")
        overallResult = False
    End If
    
    ' If errors exist, display them and log each error.
    If Not overallResult Then
        Dim errorMsg As String
        errorMsg = "The following errors were detected:" & vbCrLf & vbCrLf
        
        Dim i As Long
        For i = 1 To errorMessages.Count
            errorMsg = errorMsg & errorMessages(i) & vbCrLf
        Next i
        
        MsgBox errorMsg, vbExclamation, "Invalid File Paths"
        
        For i = 1 To errorMessages.Count
            Module_SyncTool_Logger.LogMessage errorMessages(i), "ERROR"
        Next i
    End If
    
    ValidateFilePaths = overallResult
End Function

'===============================================================================
' OPEN_WORKBOOK_SAFELY - Opens a workbook with comprehensive error handling.
'
' Parameters:
' filePath - The full path to the workbook.
' readOnly - Optional Boolean indicating if the workbook should be opened read-only (default True).
'
' Returns:
' Workbook - The opened Workbook object, or Nothing if opening fails.
'===============================================================================
Public Function OpenWorkbookSafely(filePath As String, Optional readOnly As Boolean = True) As Workbook
    On Error GoTo ErrorHandler
    
    ' Check if the file exists.
    If Not FileExists(filePath) Then
        Module_SyncTool_Logger.LogMessage "ERROR: " & Module_Utilities.FormatErrorMessage(ERR_FILE_NOT_FOUND, filePath), "ERROR"
        Set OpenWorkbookSafely = Nothing
        Exit Function
    End If
    
    ' Attempt to open the workbook.
    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath, readOnly:=readOnly)
    
    ' Return the workbook.
    Set OpenWorkbookSafely = wb
    
    Exit Function
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "ERROR: Failed to open workbook '" & Module_Utilities.GetFileName(filePath) & "': " & _
                                     err.Description & " (Error " & err.Number & ")", "ERROR"
    Set OpenWorkbookSafely = Nothing
End Function

'===============================================================================
' STANDARDIZE_USER_EDITS_SHEET - Ensures the UserEdits sheet has the proper structure.
'
' Parameters:
' workbookPath - The full path to the workbook.
' sourceAttribution - The expected source code (e.g., "AF", "RZ", or "MASTER").
'
' Behavior:
' Opens the workbook, verifies or creates the "UserEdits" sheet (EXTERNAL_USEREDITS_SHEET),
' standardizes the headers, updates incorrect ChangeSource values, and sets timestamps.
'
' Returns:
' Nothing.
'===============================================================================
Public Sub StandardizeUserEditsSheet(workbookPath As String, sourceAttribution As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim backupCreated As Boolean
    Dim docNum As String
    
    On Error GoTo ErrorHandler
    
    Module_SyncTool_UI.UpdateStatusDisplay "Standardizing " & Module_Utilities.GetFileName(workbookPath) & "..."
    Module_SyncTool_Logger.LogMessage "Beginning standardization of " & Module_Utilities.GetFileName(workbookPath) & " with source: " & sourceAttribution
    
    ' Check if the file exists.
    If Not FileExists(workbookPath) Then
        Module_SyncTool_Logger.LogMessage "ERROR: " & Module_Utilities.FormatErrorMessage(ERR_FILE_NOT_FOUND, Module_Utilities.GetFileName(workbookPath)), "ERROR"
        Exit Sub
    End If
    
    ' Open the workbook.
    Set wb = OpenWorkbookSafely(workbookPath, False)
    
    If wb Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Could not open workbook to standardize: " & Module_Utilities.GetFileName(workbookPath), "ERROR"
        Exit Sub
    End If
    
    ' Attempt to get the UserEdits sheet.
    On Error Resume Next
    Set ws = wb.Sheets(EXTERNAL_USEREDITS_SHEET)
    On Error GoTo ErrorHandler
    
    backupCreated = False
    
    ' If the sheet does not exist, create it and set up headers.
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = EXTERNAL_USEREDITS_SHEET
        
        Dim headerRange As Range
        Set headerRange = ws.Range("A1:G1")
        
        Module_Utilities.FormatHeaders ws, headerRange, Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                                                           "Email Contact", "User Comments", "ChangeSource", "Timestamp")
        
        Module_SyncTool_Logger.LogMessage "Created new UserEdits sheet in " & Module_Utilities.GetFileName(workbookPath)
    Else
        ' If the sheet exists, check if its headers match the expected format.
        Dim headersCorrect As Boolean
        headersCorrect = True
        
        If ws.Cells(1, 1).value <> "DocNumber" Then headersCorrect = False
        If ws.Cells(1, 2).value <> "Engagement Phase" And _
           ws.Cells(1, 2).value <> "EngagementPhase" Then headersCorrect = False
        If ws.Cells(1, 6).value <> "ChangeSource" Then headersCorrect = False
        If ws.Cells(1, 7).value <> "Timestamp" Then headersCorrect = False
        
        If Not headersCorrect Then
            Dim backupName As String
            backupName = "UserEdits_Backup_" & Format(Now, "yyyymmdd_hhmmss")
            
            On Error Resume Next
            Dim wsBackup As Worksheet
            Set wsBackup = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            
            If err.Number <> 0 Then
                err.Clear
                backupName = "Backup_" & Format(Now, "yyyymmdd")
                Set wsBackup = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            End If
            
            If Not wsBackup Is Nothing Then
                wsBackup.Name = backupName
                ws.UsedRange.Copy wsBackup.Range("A1")
                wsBackup.Visible = xlSheetHidden
                
                backupCreated = True
                Module_SyncTool_Logger.LogMessage "Created backup of non-standard UserEdits as '" & backupName & "' in " & Module_Utilities.GetFileName(workbookPath)
            Else
                Module_SyncTool_Logger.LogMessage "WARNING: Could not create backup of UserEdits in " & Module_Utilities.GetFileName(workbookPath), "WARNING"
            End If
            
            On Error GoTo ErrorHandler
            
            Set headerRange = ws.Range("A1:G1")
            
            Module_Utilities.FormatHeaders ws, headerRange, Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                                                               "Email Contact", "User Comments", "ChangeSource", "Timestamp")
            
            Module_SyncTool_Logger.LogMessage "Standardized UserEdits headers in " & Module_Utilities.GetFileName(workbookPath) & _
                                             (IIf(backupCreated, ", backup created as " & backupName, ""))
        End If
    End If
    
    ' Update ChangeSource values and ensure timestamps exist.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    If lastRow > 1 Then
        Dim updatedCount As Long
        updatedCount = 0
        
        For i = 2 To lastRow
            docNum = Trim(ws.Cells(i, "A").value)
            
            If docNum <> "" Then
                Dim currentSource As String
                currentSource = Trim(ws.Cells(i, "F").value)
                
                If currentSource = "" Or (currentSource <> sourceAttribution And Not Module_Utilities.IsValidAttribution(currentSource)) Then
                    ws.Cells(i, "F").value = sourceAttribution
                    updatedCount = updatedCount + 1
                End If
                
                If Trim(ws.Cells(i, "G").value) = "" Then
                    ws.Cells(i, "G").value = Format$(Now(), FORMAT_TIMESTAMP)
                End If
            End If
        Next i
        
        If updatedCount > 0 Then
            Module_SyncTool_Logger.LogMessage "Updated " & updatedCount & " ChangeSource values to '" & sourceAttribution & "' in " & Module_Utilities.GetFileName(workbookPath)
        End If
    End If
    
    ' Hide the UserEdits sheet, save and close the workbook.
    ws.Visible = xlSheetHidden
    
    wb.Save
    wb.Close
    
    Module_SyncTool_Logger.LogMessage "Completed standardization of " & Module_Utilities.GetFileName(workbookPath)
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_Logger.LogMessage "ERROR standardizing " & Module_Utilities.GetFileName(workbookPath) & ": " & _
                                     err.Description & " (Error " & err.Number & ")", "ERROR"
    
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'===============================================================================
' EXTRACT_USER_EDITS - Extracts data from the UserEdits sheet in the specified workbook.
'
' Parameters:
' workbookPath - Full path of the workbook.
' sourceAttribution - The expected source attribution code (e.g., "AF", "RZ", "MASTER").
'
' Returns:
' A Dictionary where each key is a document number and each value is a Dictionary
' containing standardized row data.
'===============================================================================
Public Function ExtractUserEdits(workbookPath As String, sourceAttribution As String) As Object
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim i As Long
    Dim docNum As String
    Dim processedCount As Long
    Dim totalRows As Long
    
    ' Validate inputs.
    If Module_Utilities.IsNullOrEmpty(workbookPath) Then
        Module_SyncTool_Logger.LogMessage "ERROR: Empty workbook path provided to ExtractUserEdits", "ERROR"
        Set ExtractUserEdits = CreateObject("Scripting.Dictionary")
        Exit Function
    End If
    
    If Not FileExists(workbookPath) Then
        Module_SyncTool_Logger.LogMessage "ERROR: " & Module_Utilities.FormatErrorMessage(ERR_FILE_NOT_FOUND, workbookPath), "ERROR"
        Set ExtractUserEdits = CreateObject("Scripting.Dictionary")
        Exit Function
    End If
    
    If Module_Utilities.IsNullOrEmpty(sourceAttribution) Then
        Module_SyncTool_Logger.LogMessage "ERROR: Empty source attribution provided to ExtractUserEdits", "ERROR"
        Set ExtractUserEdits = CreateObject("Scripting.Dictionary")
        Exit Function
    End If
    
    Module_SyncTool_UI.UpdateStatusDisplay "Extracting from: " & Module_Utilities.GetFileName(workbookPath)
    Module_SyncTool_Logger.LogMessage "Beginning extraction from " & Module_Utilities.GetFileName(workbookPath)
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Open the workbook in read-only mode.
    Set wb = OpenWorkbookSafely(workbookPath, True)
    
    If wb Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Failed to open workbook for extraction: " & Module_Utilities.GetFileName(workbookPath), "ERROR"
        Set ExtractUserEdits = dict
        Exit Function
    End If
    
    On Error Resume Next
    Set ws = wb.Sheets(EXTERNAL_USEREDITS_SHEET)
    On Error GoTo ErrorHandler
    
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        
        totalRows = lastRow - 1 ' Skip header row.
        
        If totalRows > 0 Then
            Module_SyncTool_UI.UpdateStatusDisplay "Processing 0 of " & totalRows & " records from " & Module_Utilities.GetFileName(workbookPath)
            Module_SyncTool_Logger.LogMessage "Found " & totalRows & " records in " & Module_Utilities.GetFileName(workbookPath)
        Else
            Module_SyncTool_UI.UpdateStatusDisplay "No records found in " & Module_Utilities.GetFileName(workbookPath)
            Module_SyncTool_Logger.LogMessage "No records found in " & Module_Utilities.GetFileName(workbookPath)
            GoTo CleanupAndExit
        End If
        
        processedCount = 0
        
        For i = 2 To lastRow
            docNum = Trim(ws.Cells(i, COL_DOCNUMBER).value)
            
            processedCount = processedCount + 1
            
            If processedCount Mod 20 = 0 Or processedCount = 1 Or processedCount = totalRows Then
                Module_SyncTool_UI.UpdateStatusDisplay "Processing " & processedCount & " of " & totalRows & " records from " & Module_Utilities.GetFileName(workbookPath)
                DoEvents
            End If
            
            If Not Module_Utilities.IsNullOrEmpty(docNum) Then
                Dim rowData As Object
                Set rowData = CreateObject("Scripting.Dictionary")
                
                rowData.Add "DocNumber", docNum
                rowData.Add "EngagementPhase", ws.Cells(i, COL_ENGAGEMENTPHASE).value
                rowData.Add "LastContactDate", ws.Cells(i, COL_LASTCONTACTDATE).value
                rowData.Add "EmailContact", ws.Cells(i, COL_EMAILCONTACT).value
                rowData.Add "UserComments", ws.Cells(i, COL_USERCOMMENTS).value
                
                Dim changeSource As String
                changeSource = Trim(ws.Cells(i, COL_CHANGESOURCE).value)
                
                If Module_Utilities.IsNullOrEmpty(changeSource) Then
                    changeSource = sourceAttribution
                ElseIf Not Module_Utilities.IsValidAttribution(changeSource) Then
                    Module_SyncTool_Logger.LogMessage "WARNING: Invalid ChangeSource '" & changeSource & "' in " & Module_Utilities.GetFileName(workbookPath) & " for DocNumber " & docNum, "WARNING"
                End If
                
                rowData.Add "ChangeSource", changeSource
                
                Dim timestamp As Variant
                timestamp = ws.Cells(i, COL_TIMESTAMP).value
                
                If IsDate(timestamp) Then
                    rowData.Add "LastModified", CDate(timestamp)
                Else
                    Module_SyncTool_Logger.LogMessage "WARNING: Invalid timestamp in " & Module_Utilities.GetFileName(workbookPath) & " for DocNumber " & docNum & ", using current time", "WARNING"
                    rowData.Add "LastModified", Now()
                End If
                
                rowData.Add "Source", Module_Utilities.GetFileName(workbookPath)
                
                If Not dict.Exists(docNum) Then
                    dict.Add docNum, rowData
                Else
                    If rowData("LastModified") > dict(docNum)("LastModified") Then
                        dict.Remove docNum
                        dict.Add docNum, rowData
                        Module_SyncTool_Logger.LogMessage "WARNING: Duplicate entry for DocNumber " & docNum & " in " & Module_Utilities.GetFileName(workbookPath) & ", keeping most recent", "WARNING"
                    End If
                End If
            End If
            
            If i Mod 100 = 0 Then DoEvents
        Next i
        
        Module_SyncTool_UI.UpdateStatusDisplay "Extracted " & dict.Count & " unique records from " & Module_Utilities.GetFileName(workbookPath)
        Module_SyncTool_Logger.LogMessage "Extracted " & dict.Count & " unique records from " & Module_Utilities.GetFileName(workbookPath)
    Else
        Module_SyncTool_UI.UpdateStatusDisplay "Error: UserEdits sheet not found in " & Module_Utilities.GetFileName(workbookPath)
        Module_SyncTool_Logger.LogMessage "ERROR: UserEdits sheet not found in " & Module_Utilities.GetFileName(workbookPath), "ERROR"
        Set dict = CreateObject("Scripting.Dictionary")
    End If
    
CleanupAndExit:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.StatusBar = False
    
    Set ExtractUserEdits = dict
    
    Exit Function
    
ErrorHandler:
    Module_SyncTool_UI.UpdateStatusDisplay "Error extracting data: " & err.Description
    Module_SyncTool_Logger.LogMessage "ERROR during extraction from " & Module_Utilities.GetFileName(workbookPath) & ": " & err.Description & " (Error " & err.Number & ")", "ERROR"
    
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    
    On Error GoTo 0
    Set ExtractUserEdits = CreateObject("Scripting.Dictionary")
End Function

'===============================================================================
' WRITE_USER_EDITS_TO_WORKBOOK - Writes merged data to an external workbook.
'
' Parameters:
' workbookPath - Full path to the target workbook.
' mergedData - A Dictionary containing merged document data.
' preserveSourceFiltering - Boolean indicating if source filtering should be applied.
'
' Behavior:
' Opens the workbook, ensures the UserEdits sheet exists and is properly formatted,
' writes the filtered merged data into the sheet, sorts the data, and then saves and closes the workbook.
'
' Returns:
' Nothing.
'===============================================================================
Public Sub WriteUserEditsToWorkbook(workbookPath As String, mergedData As Object, preserveSourceFiltering As Boolean)
    On Error GoTo ErrorHandler
    
    ' Validate inputs.
    If Module_Utilities.IsNullOrEmpty(workbookPath) Then
        Module_SyncTool_Logger.LogMessage "ERROR: Empty workbook path provided to WriteUserEditsToWorkbook", "ERROR"
        Exit Sub
    End If
    
    If Not FileExists(workbookPath) Then
        Module_SyncTool_Logger.LogMessage "ERROR: " & Module_Utilities.FormatErrorMessage(ERR_FILE_NOT_FOUND, workbookPath), "ERROR"
        Exit Sub
    End If
    
    If mergedData Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Invalid mergedData provided to WriteUserEditsToWorkbook", "ERROR"
        Exit Sub
    End If
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim docNum As Variant
    Dim rowData As Object
    Dim lastRow As Long, destRow As Long
    
    Module_SyncTool_UI.UpdateStatusDisplay "Writing to: " & Module_Utilities.GetFileName(workbookPath)
    Module_SyncTool_Logger.LogMessage "Begin writing to: " & Module_Utilities.GetFileName(workbookPath)
    
    ' Open the workbook.
    Set wb = OpenWorkbookSafely(workbookPath, False)
    
    If wb Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Failed to open workbook for writing: " & Module_Utilities.GetFileName(workbookPath), "ERROR"
        Exit Sub
    End If
    
    ' Get or create the UserEdits sheet.
    On Error Resume Next
    Set ws = wb.Sheets(EXTERNAL_USEREDITS_SHEET)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = EXTERNAL_USEREDITS_SHEET
        Module_SyncTool_Logger.LogMessage "Created new UserEdits sheet in " & Module_Utilities.GetFileName(workbookPath)
    End If
    
    ' Ensure headers are correct.
    Dim headerRange As Range
    Set headerRange = ws.Range("A1:G1")
    
    Module_Utilities.FormatHeaders ws, headerRange, Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                                                       "Email Contact", "User Comments", "ChangeSource", "Timestamp")
    
    ' Clear existing data (keep header row).
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    If lastRow > 1 Then ws.Range("A2:G" & lastRow).Clear
    
    ' Determine workbook type for source attribution based on filePath.
    Dim workbookType As String
    
    If InStr(LCase(workbookPath), "ally") > 0 Then
        workbookType = ATTRIBUTION_ALLY
    ElseIf InStr(LCase(workbookPath), "ryan") > 0 Then
        workbookType = ATTRIBUTION_RYAN
    Else
        workbookType = ATTRIBUTION_MASTER
    End If
    
    ' Write filtered merged data.
    destRow = 2
    
    Dim counter As Long, total As Long
    total = mergedData.Count: counter = 0
    
    Dim filteredData As Object
    Set filteredData = CreateObject("Scripting.Dictionary")
    
    Dim changeSource As String
    
    For Each docNum In mergedData.keys
        Set rowData = mergedData(docNum)
        
        Dim includeRecord As Boolean: includeRecord = True
        
        If preserveSourceFiltering And workbookType <> ATTRIBUTION_MASTER Then
            changeSource = rowData("ChangeSource")
            
            If InStr(changeSource, workbookType) = 0 And changeSource <> ATTRIBUTION_MASTER Then
                includeRecord = False
            End If
        End If
        
        If includeRecord Then filteredData.Add docNum, rowData
    Next docNum
    
    For Each docNum In filteredData.keys
        counter = counter + 1
        Set rowData = filteredData(docNum)
        
        If counter Mod 50 = 0 Then
            Module_SyncTool_UI.UpdateStatusDisplay "Writing " & counter & " of " & filteredData.Count & " records to " & Module_Utilities.GetFileName(workbookPath)
            DoEvents
        End If
        
        ws.Cells(destRow, COL_DOCNUMBER).value = CStr(docNum)
        ws.Cells(destRow, COL_ENGAGEMENTPHASE).value = rowData("EngagementPhase")
        ws.Cells(destRow, COL_LASTCONTACTDATE).value = rowData("LastContactDate")
        ws.Cells(destRow, COL_EMAILCONTACT).value = rowData("EmailContact")
        ws.Cells(destRow, COL_USERCOMMENTS).value = rowData("UserComments")
        
        If preserveSourceFiltering Then
            ws.Cells(destRow, COL_CHANGESOURCE).value = ConvertAttributionForWorkbook(rowData("ChangeSource"), workbookType)
        Else
            ws.Cells(destRow, COL_CHANGESOURCE).value = rowData("ChangeSource")
        End If
        
        ws.Cells(destRow, COL_TIMESTAMP).value = Format$(Now(), FORMAT_TIMESTAMP)
        
        If IsDate(rowData("LastContactDate")) Then
            ws.Cells(destRow, COL_LASTCONTACTDATE).NumberFormat = FORMAT_DATE
        End If
        
        Module_Format_Helpers.ApplyAttributionFormatting ws.Cells(destRow, COL_CHANGESOURCE)
        
        destRow = destRow + 1
    Next docNum
    
    If destRow > 2 Then
        ws.Range("A2:G" & destRow - 1).Sort Key1:=ws.Range("A2"), Order1:=xlAscending, Header:=xlNo
    End If
    
    ws.Columns("A:G").AutoFit
    ws.Cells.WrapText = False
    
    ws.Visible = xlSheetHidden
    
    wb.Save
    wb.Close SaveChanges:=True
    
    Module_SyncTool_UI.UpdateStatusDisplay "Successfully wrote " & counter & " records to " & Module_Utilities.GetFileName(workbookPath)
    Module_SyncTool_Logger.LogMessage "Completed writing " & counter & " records to " & Module_Utilities.GetFileName(workbookPath)
    
    Exit Sub
    
ErrorHandler:
    Module_SyncTool_UI.UpdateStatusDisplay "Error writing to workbook: " & err.Description
    Module_SyncTool_Logger.LogMessage "ERROR writing to " & Module_Utilities.GetFileName(workbookPath) & ": " & err.Description & " (Error " & err.Number & ")", "ERROR"
    
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    
    On Error GoTo 0
End Sub

'===============================================================================
' CONVERT_ATTRIBUTION_FOR_WORKBOOK - Simplifies attribution for the target workbook.
'
' Parameters:
' attribution - The original attribution string.
' workbookType - The target workbook's source code.
'
' Returns:
' A simplified attribution string appropriate for the target workbook.
'===============================================================================
Private Function ConvertAttributionForWorkbook(attribution As String, workbookType As String) As String
    If InStr(attribution, "+") > 0 Then
        If InStr(attribution, workbookType) > 0 Then
            ConvertAttributionForWorkbook = workbookType
        Else
            ConvertAttributionForWorkbook = attribution
        End If
    Else
        ConvertAttributionForWorkbook = attribution
    End If
End Function

'===============================================================================
' TRIGGER_EXTERNAL_WORKBOOK_REFRESH - Triggers a refresh in an external workbook.
'
' Parameters:
' filePath - The full path to the external workbook.
'
' Behavior:
' Attempts to open the workbook, run a refresh macro, and then save and close it.
'
' Returns:
' Nothing.
'===============================================================================
Public Sub TriggerExternalWorkbookRefresh(filePath As String)
    Dim wb As Workbook
    Dim success As Boolean
    
    On Error Resume Next
    
    If Not FileExists(filePath) Then
        Module_SyncTool_Logger.LogMessage "ERROR: " & Module_Utilities.FormatErrorMessage(ERR_FILE_NOT_FOUND, Module_Utilities.GetFileName(filePath)), "ERROR"
        Exit Sub
    End If
    
    Module_SyncTool_UI.UpdateStatusDisplay "Opening " & Module_Utilities.GetFileName(filePath) & " for refresh..."
    
    Set wb = OpenWorkbookSafely(filePath, False)
    
    If wb Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Failed to open workbook for refresh: " & Module_Utilities.GetFileName(filePath), "ERROR"
        Exit Sub
    End If
    
    err.Clear
    
    Module_SyncTool_UI.UpdateStatusDisplay "Triggering table refresh in " & Module_Utilities.GetFileName(filePath) & "..."
    
    success = False
    
    Application.Run "'" & wb.Name & "'!RefreshDataTables_FromSync"
    
    If err.Number = 0 Then
        success = True
    Else
        Module_SyncTool_Logger.LogMessage "WARNING: RefreshDataTables_FromSync failed in " & Module_Utilities.GetFileName(filePath) & ": " & err.Description, "WARNING"
        err.Clear
        
        Application.Run "'" & wb.Name & "'!RefreshTables"
        
        If err.Number = 0 Then success = True
    End If
    
    If success Then
        Module_SyncTool_Logger.LogMessage "Successfully triggered data refresh in " & Module_Utilities.GetFileName(filePath)
    Else
        Module_SyncTool_Logger.LogMessage "WARNING: Could not trigger data refresh in " & Module_Utilities.GetFileName(filePath) & ". This won't affect data synchronization, just the data display.", "WARNING"
    End If
    
    On Error Resume Next
    
    If Not wb Is Nothing Then
        Module_SyncTool_UI.UpdateStatusDisplay "Saving " & Module_Utilities.GetFileName(filePath) & "..."
        wb.Save
        wb.Close
    End If
    
    On Error GoTo 0
End Sub


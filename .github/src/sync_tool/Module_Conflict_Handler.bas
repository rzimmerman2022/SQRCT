Attribute VB_Name = "Module_Conflict_Handler"
Option Explicit

'===============================================================================
' MODULE_CONFLICT_HANDLER
' Contains functions for detecting and resolving conflicts between data sources.
' Focuses specifically on conflict identification and resolution.
'===============================================================================

'===============================================================================
' DETECT_CONFLICTS - Identifies conflicts between multiple data sources.
'
' Parameters:
' dataMap - A Dictionary object where each key is a source identifier (e.g., "AF", "RZ", "MASTER")
'           and each value is a Dictionary of document data (keyed by document number).
'
' Returns:
' A Dictionary containing conflict information for each document that has conflicting data.
'===============================================================================
Public Function DetectConflicts(dataMap As Object) As Object
    ' Validate input.
    If dataMap Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Invalid dataMap provided to DetectConflicts", "ERROR"
        Set DetectConflicts = CreateObject("Scripting.Dictionary")
        Exit Function
    End If
    
    Dim conflicts As Object
    Set conflicts = CreateObject("Scripting.Dictionary")
    
    Dim docNum As Variant
    Dim sourceKeys As Variant
    Dim i As Long, j As Long
    Dim docSources As Object
    
    Module_SyncTool_Logger.LogMessage "Scanning for conflicts between data sources..."
    
    ' Get an array of source keys from the dataMap.
    sourceKeys = dataMap.keys
    
    ' Gather all unique document numbers from each data source.
    Dim allDocs As Object
    Set allDocs = CreateObject("Scripting.Dictionary")
    
    Dim sKey As Variant
    
    ' Use the first source as baseline.
    On Error Resume Next  ' Handle case where source might have no documents
    For Each docNum In dataMap(sourceKeys(0)).keys
        allDocs(docNum) = True
    Next docNum
    On Error GoTo 0  ' Restore normal error handling
    
    ' Add document numbers from remaining sources.
    For i = 1 To UBound(sourceKeys)
        On Error Resume Next  ' Handle case where source might have no documents
        For Each docNum In dataMap(sourceKeys(i)).keys
            allDocs(docNum) = True
        Next docNum
        On Error GoTo 0  ' Restore normal error handling
    Next i
    
    ' Check each unique document for conflicts.
    For Each docNum In allDocs.keys
        ' Build a dictionary of sources that contain this document.
        Set docSources = CreateObject("Scripting.Dictionary")
        
        For i = 0 To UBound(sourceKeys)
            If dataMap(sourceKeys(i)).Exists(docNum) Then
                docSources.Add sourceKeys(i), dataMap(sourceKeys(i))(docNum)
            End If
        Next i
        
        ' Only proceed if more than one source has the document.
        If docSources.Count > 1 Then
            Dim hasConflict As Boolean
            hasConflict = False
            
            ' Compare each pair of sources for field-level differences.
            Dim sourceArray As Variant
            sourceArray = docSources.keys
            
            For i = 0 To docSources.Count - 2
                For j = i + 1 To docSources.Count - 1
                    If HasFieldConflicts(docSources(sourceArray(i)), docSources(sourceArray(j))) Then
                        hasConflict = True
                        Exit For
                    End If
                Next j
                
                If hasConflict Then Exit For
            Next i
            
            ' If conflict is detected, build the conflict information.
            If hasConflict Then
                Dim conflictInfo As Object
                Set conflictInfo = CreateObject("Scripting.Dictionary")
                
                conflictInfo.Add "DocNumber", docNum
                
                ' Store each source's last modified timestamp and user comments.
                For i = 0 To docSources.Count - 1
                    Dim src As String
                    src = sourceArray(i)
                    
                    conflictInfo.Add src & "Date", docSources(src)("LastModified")
                    conflictInfo.Add src & "Comment", docSources(src)("UserComments")
                Next i
                
                ' Determine the type of conflict based on field differences.
                conflictInfo.Add "Type", DetermineConflictType(docSources, sourceArray)
                
                ' Store all source data for further resolution.
                conflictInfo.Add "Sources", docSources
                
                ' Add this conflict to the conflicts dictionary.
                conflicts.Add docNum, conflictInfo
            End If
        End If
    Next docNum
    
    Module_SyncTool_Logger.LogMessage "Found " & conflicts.Count & " potential conflicts"
    
    Set DetectConflicts = conflicts
End Function

'===============================================================================
' HAS_FIELD_CONFLICTS - Determines if two data records have conflicting values.
'
' Parameters:
' data1, data2 - Dictionary objects representing a document record from two sources.
'
' Returns:
' Boolean - True if any key field differs between the two records; False otherwise.
'===============================================================================
Private Function HasFieldConflicts(data1 As Object, data2 As Object) As Boolean
    ' Validate inputs.
    If data1 Is Nothing Or data2 Is Nothing Then
        HasFieldConflicts = False
        Exit Function
    End If
    
    Dim hasConflict As Boolean
    hasConflict = False
    
    ' Compare key fields: UserComments, EngagementPhase, LastContactDate, EmailContact.
    If Not Module_Utilities.IsNullOrEmpty(data1("UserComments")) And _
       Not Module_Utilities.IsNullOrEmpty(data2("UserComments")) And _
       Trim(data1("UserComments")) <> Trim(data2("UserComments")) Then
        hasConflict = True
    End If
    
    If Not Module_Utilities.IsNullOrEmpty(data1("EngagementPhase")) And _
       Not Module_Utilities.IsNullOrEmpty(data2("EngagementPhase")) And _
       Trim(data1("EngagementPhase")) <> Trim(data2("EngagementPhase")) Then
        hasConflict = True
    End If
    
    If Not Module_Utilities.IsNullOrEmpty(data1("LastContactDate")) And _
       Not Module_Utilities.IsNullOrEmpty(data2("LastContactDate")) And _
       Trim(CStr(data1("LastContactDate"))) <> Trim(CStr(data2("LastContactDate"))) Then
        hasConflict = True
    End If
    
    If Not Module_Utilities.IsNullOrEmpty(data1("EmailContact")) And _
       Not Module_Utilities.IsNullOrEmpty(data2("EmailContact")) And _
       Trim(data1("EmailContact")) <> Trim(data2("EmailContact")) Then
        hasConflict = True
    End If
    
    HasFieldConflicts = hasConflict
End Function

'===============================================================================
' DETERMINE_CONFLICT_TYPE - Determines the primary type of conflict for a document.
'
' Parameters:
' docSources - A Dictionary containing all source records for a document.
' sourceArray - An array of source keys present in docSources.
'
' Returns:
' String - The conflict type (e.g., "Comments", "EngagementPhase", "LastContactDate",
'          "EmailContact", or "Timestamps" as a default).
'===============================================================================
Private Function DetermineConflictType(docSources As Object, sourceArray As Variant) As String
    Dim i As Long, j As Long
    Dim src1 As String, src2 As String
    
    ' Highest priority: UserComments conflicts.
    For i = 0 To UBound(sourceArray) - 1
        For j = i + 1 To UBound(sourceArray)
            src1 = sourceArray(i)
            src2 = sourceArray(j)
            
            If Not Module_Utilities.IsNullOrEmpty(docSources(src1)("UserComments")) And _
               Not Module_Utilities.IsNullOrEmpty(docSources(src2)("UserComments")) And _
               Trim(docSources(src1)("UserComments")) <> Trim(docSources(src2)("UserComments")) Then
                DetermineConflictType = "Comments"
                Exit Function
            End If
        Next j
    Next i
    
    ' Next, check EngagementPhase.
    For i = 0 To UBound(sourceArray) - 1
        For j = i + 1 To UBound(sourceArray)
            src1 = sourceArray(i)
            src2 = sourceArray(j)
            
            If Not Module_Utilities.IsNullOrEmpty(docSources(src1)("EngagementPhase")) And _
               Not Module_Utilities.IsNullOrEmpty(docSources(src2)("EngagementPhase")) And _
               Trim(docSources(src1)("EngagementPhase")) <> Trim(docSources(src2)("EngagementPhase")) Then
                DetermineConflictType = "EngagementPhase"
                Exit Function
            End If
        Next j
    Next i
    
    ' Next, check LastContactDate.
    For i = 0 To UBound(sourceArray) - 1
        For j = i + 1 To UBound(sourceArray)
            src1 = sourceArray(i)
            src2 = sourceArray(j)
            
            If Not Module_Utilities.IsNullOrEmpty(docSources(src1)("LastContactDate")) And _
               Not Module_Utilities.IsNullOrEmpty(docSources(src2)("LastContactDate")) And _
               Trim(CStr(docSources(src1)("LastContactDate"))) <> Trim(CStr(docSources(src2)("LastContactDate"))) Then
                DetermineConflictType = "LastContactDate"
                Exit Function
            End If
        Next j
    Next i
    
    ' Finally, check EmailContact.
    For i = 0 To UBound(sourceArray) - 1
        For j = i + 1 To UBound(sourceArray)
            src1 = sourceArray(i)
            src2 = sourceArray(j)
            
            If Not Module_Utilities.IsNullOrEmpty(docSources(src1)("EmailContact")) And _
               Not Module_Utilities.IsNullOrEmpty(docSources(src2)("EmailContact")) And _
               Trim(docSources(src1)("EmailContact")) <> Trim(docSources(src2)("EmailContact")) Then
                DetermineConflictType = "EmailContact"
                Exit Function
            End If
        Next j
    Next i
    
    ' Default: If no specific field conflict found, return "Timestamps".
    DetermineConflictType = "Timestamps"
End Function

'===============================================================================
' DISPLAY_CONFLICTS - Displays detected conflicts in the MergeData sheet for review.
'
' Parameters:
' conflicts - A Dictionary containing conflict information keyed by document number.
'
' Returns:
' Nothing.
'
' Behavior:
' Clears the MergeData sheet, sets up header rows, and then populates each row
' with conflict details including source timestamps, field values, and resolution
' recommendations.
'===============================================================================
Public Sub DisplayConflicts(conflicts As Object)
    ' Validate input.
    If conflicts Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Invalid conflicts object provided to DisplayConflicts", "ERROR"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim row As Long
    Dim docNum As Variant
    Dim conflictInfo As Object
    Dim src As Variant
    Dim srcCol As Long
    Dim resolution As String
    Dim attribution As String
    Dim mostRecentSource As String
    Dim colOffset As Long
    Dim allSources As Object
    
    ' If no conflicts, notify the user.
    If conflicts.Count = 0 Then
        MsgBox "No conflicts detected.", vbInformation
        Exit Sub
    End If
    
    ' Retrieve and clear the MergeData sheet.
    On Error Resume Next
    Set ws = Module_SyncTool_Logger.GetMergeDataSheet()
    On Error GoTo 0
    
    If ws Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Could not access MergeData sheet in DisplayConflicts", "ERROR"
        Exit Sub
    End If
    
    ws.Cells.Clear
    
    ' Set up header rows.
    ws.Range("A1").value = "Document Number"
    ws.Range("B1").value = "Conflict Type"
    
    ' Starting from column C for source-specific data.
    colOffset = 3
    Set allSources = CreateObject("Scripting.Dictionary")
    
    ' Build headers for each source across all conflicts.
    For Each docNum In conflicts.keys
        Set conflictInfo = conflicts(docNum)
        
        For Each src In conflictInfo("Sources").keys
            If Not allSources.Exists(src) Then
                allSources.Add src, colOffset
                ws.Cells(1, colOffset).value = src & " Last Edit"
                ws.Cells(1, colOffset + 1).value = src & " Value"
                colOffset = colOffset + 2
            End If
        Next src
    Next docNum
    
    ' Add columns for resolution recommendations.
    ws.Cells(1, colOffset).value = "Resolution"
    ws.Cells(1, colOffset + 1).value = "Final Attribution"
    
    ' Apply standard header formatting.
    Dim headerRange As Range
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, colOffset + 1))
    Module_Utilities.FormatHeaders ws, headerRange
    
    ' Populate conflict rows.
    row = 2
    
    For Each docNum In conflicts.keys
        Set conflictInfo = conflicts(docNum)
        
        ws.Cells(row, "A").value = conflictInfo("DocNumber")
        ws.Cells(row, "B").value = conflictInfo("Type")
        
        ' For each source, display the last edit timestamp and corresponding field value.
        For Each src In conflictInfo("Sources").keys
            srcCol = allSources(src)
            
            ws.Cells(row, srcCol).value = Format(conflictInfo(src & "Date"), FORMAT_TIMESTAMP)
            
            Select Case conflictInfo("Type")
                Case "Comments"
                    ws.Cells(row, srcCol + 1).value = conflictInfo(src & "Comment")
                Case "EngagementPhase"
                    ws.Cells(row, srcCol + 1).value = conflictInfo("Sources")(src)("EngagementPhase")
                Case "LastContactDate"
                    ws.Cells(row, srcCol + 1).value = conflictInfo("Sources")(src)("LastContactDate")
                Case "EmailContact"
                    ws.Cells(row, srcCol + 1).value = conflictInfo("Sources")(src)("EmailContact")
                Case "Timestamps"
                    ws.Cells(row, srcCol + 1).value = "[Various Fields]"
            End Select
        Next src
        
        ' Determine resolution recommendation.
        Select Case conflictInfo("Type")
            Case "Comments"
                resolution = "Keeping all comments with combined attribution"
                attribution = JoinSourceKeys(conflictInfo("Sources").keys)
            Case Else
                mostRecentSource = GetMostRecentSource(conflictInfo)
                resolution = "Using " & mostRecentSource & "'s value (most recent)"
                attribution = mostRecentSource
        End Select
        
        ws.Cells(row, colOffset).value = resolution
        ws.Cells(row, colOffset + 1).value = attribution
        
        ' Highlight the row if there was a conflict.
        ws.Range(ws.Cells(row, 1), ws.Cells(row, colOffset + 1)).Interior.Color = COLOR_WARNING
        
        row = row + 1
    Next docNum
    
    ' Auto-fit columns and activate the sheet.
    ws.Columns("A:" & Module_Utilities.ColumnLetterFromNumber(colOffset + 1)).AutoFit
    
    ws.Visible = xlSheetVisible
    ws.Activate
    
    MsgBox conflicts.Count & " potential conflicts were detected." & vbCrLf & _
         "These have been displayed in the MergeData sheet for review.", vbInformation, "Conflict Detection"
End Sub

'===============================================================================
' JOIN_SOURCE_KEYS - Combines source keys using a "+" separator.
'
' Parameters:
' keys - An array or collection of source keys.
'
' Returns:
' A string representing the combined source keys.
'===============================================================================
Private Function JoinSourceKeys(keys As Variant) As String
    Dim result As String
    Dim i As Integer
    
    result = ""
    
    For i = 0 To UBound(keys)
        If i > 0 Then result = result & "+"
        result = result & keys(i)
    Next i
    
    JoinSourceKeys = result
End Function

'===============================================================================
' GET_MOST_RECENT_SOURCE - Determines which source has the most recent edit.
'
' Parameters:
' conflictInfo - A Dictionary containing conflict details for a document.
'
' Returns:
' A string representing the source key with the most recent timestamp.
'===============================================================================
Private Function GetMostRecentSource(conflictInfo As Object) As String
    Dim mostRecentSource As String
    Dim mostRecentDate As Date
    Dim src As Variant
    
    mostRecentDate = #1/1/1900# ' Initialize to a very early date.
    
    For Each src In conflictInfo("Sources").keys
        If conflictInfo(src & "Date") > mostRecentDate Then
            mostRecentDate = conflictInfo(src & "Date")
            mostRecentSource = src
        End If
    Next src
    
    GetMostRecentSource = mostRecentSource
End Function

'===============================================================================
' MERGE_USER_EDITS - Merges data from multiple sources, applying conflict resolution.
'
' Parameters:
' dataMap - A Dictionary with source keys mapping to document data dictionaries.
' conflicts - A Dictionary of conflict information for documents with conflicts.
'
' Returns:
' A Dictionary containing the merged document data.
'===============================================================================
Public Function MergeUserEdits(dataMap As Object, conflicts As Object) As Object
    Dim mergedData As Object
    Set mergedData = CreateObject("Scripting.Dictionary")
    
    ' Validate inputs.
    If dataMap Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Invalid dataMap provided to MergeUserEdits", "ERROR"
        Set MergeUserEdits = mergedData
        Exit Function
    End If
    
    If conflicts Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Invalid conflicts provided to MergeUserEdits", "ERROR"
        Set MergeUserEdits = mergedData
        Exit Function
    End If
    
    Dim sourceKeys As Variant
    sourceKeys = dataMap.keys
    
    Dim docNum As Variant
    Dim rowData As Object
    Dim i As Long
    
    Module_SyncTool_UI.UpdateStatusDisplay "Merging data from all sources..."
    Module_SyncTool_Logger.LogMessage "Beginning data merge from " & dataMap.Count & " sources"
    
    ' Add all documents from sources that do not have conflicts.
    For i = 0 To UBound(sourceKeys)
        Dim source As String
        source = sourceKeys(i)
        
        Module_SyncTool_Logger.LogMessage "Processing " & dataMap(source).Count & " records from source: " & source
        
        Dim key As Variant
        For Each docNum In dataMap(source).keys
            If Not conflicts.Exists(docNum) Then
                If Not mergedData.Exists(docNum) Then
                    mergedData.Add docNum, dataMap(source)(docNum)
                Else
                    If dataMap(source)(docNum)("LastModified") > mergedData(docNum)("LastModified") Then
                        mergedData.Remove docNum
                        mergedData.Add docNum, dataMap(source)(docNum)
                    End If
                End If
            End If
        Next docNum
    Next i
    
    ' Resolve conflicts and merge them.
    Dim resolvedData As Object
    Set resolvedData = ResolveConflicts(conflicts, dataMap)
    
    ApplyConflictResolutions mergedData, resolvedData
    
    Module_SyncTool_UI.UpdateStatusDisplay "Merged " & mergedData.Count & " total records"
    Module_SyncTool_Logger.LogMessage "Merged data complete: " & mergedData.Count & " total records (" & _
                                     conflicts.Count & " conflict resolutions applied)"
    
    Set MergeUserEdits = mergedData
End Function

'===============================================================================
' RESOLVE_CONFLICTS - Resolves conflicts and produces merged data for conflicting documents.
'
' Parameters:
' conflicts - A Dictionary of conflict information keyed by document number.
' dataMap - The original data map of all sources.
'
' Returns:
' A Dictionary containing resolved document data.
'===============================================================================
Private Function ResolveConflicts(conflicts As Object, dataMap As Object) As Object
    Dim resolvedData As Object
    Set resolvedData = CreateObject("Scripting.Dictionary")
    
    Dim docNum As Variant
    Dim conflictInfo As Object
    Dim mergedData As Object
    
    For Each docNum In conflicts.keys
        Set conflictInfo = conflicts(docNum)
        Set mergedData = CreateObject("Scripting.Dictionary")
        
        ' Copy basic document info.
        mergedData.Add "DocNumber", docNum
        
        Dim mostRecentSource As String
        
        ' Handle resolution based on conflict type.
        Select Case conflictInfo("Type")
            Case "Comments"
                ' Concatenate all comments with attribution.
                mergedData.Add "UserComments", ConcatenateComments(conflictInfo)
                
                mostRecentSource = GetMostRecentSource(conflictInfo)
                mergedData.Add "EngagementPhase", conflictInfo("Sources")(mostRecentSource)("EngagementPhase")
                mergedData.Add "LastContactDate", conflictInfo("Sources")(mostRecentSource)("LastContactDate")
                mergedData.Add "EmailContact", conflictInfo("Sources")(mostRecentSource)("EmailContact")
                
                ' Combine all attribution codes.
                mergedData.Add "ChangeSource", JoinSourceKeys(conflictInfo("Sources").keys)
                
            Case "EngagementPhase", "LastContactDate", "EmailContact", "Timestamps"
                mostRecentSource = GetMostRecentSource(conflictInfo)
                
                mergedData.Add "EngagementPhase", conflictInfo("Sources")(mostRecentSource)("EngagementPhase")
                mergedData.Add "LastContactDate", conflictInfo("Sources")(mostRecentSource)("LastContactDate")
                mergedData.Add "EmailContact", conflictInfo("Sources")(mostRecentSource)("EmailContact")
                mergedData.Add "UserComments", conflictInfo("Sources")(mostRecentSource)("UserComments")
                mergedData.Add "ChangeSource", mostRecentSource
        End Select
        
        ' Update the timestamp to the current time.
        mergedData.Add "LastModified", Now()
        
        ' Add resolved document to the resolved data dictionary.
        resolvedData.Add docNum, mergedData
    Next docNum
    
    Set ResolveConflicts = resolvedData
End Function

'===============================================================================
' CONCATENATE_COMMENTS - Concatenates user comments from all sources with attribution.
'
' Parameters:
' conflictInfo - A Dictionary containing conflict information for a document.
'
' Returns:
' A string containing all comments, separated by line breaks and prefixed by the source.
'===============================================================================
Private Function ConcatenateComments(conflictInfo As Object) As String
    Dim result As String
    Dim src As Variant
    
    result = ""
    
    For Each src In conflictInfo("Sources").keys
        Dim srcComment As String
        srcComment = Trim(conflictInfo(src & "Comment"))
        
        If srcComment <> "" Then
            If result <> "" Then result = result & vbCrLf & "---" & vbCrLf
            result = result & src & ": " & srcComment
        End If
    Next src
    
    ConcatenateComments = result
End Function

'===============================================================================
' APPLY_CONFLICT_RESOLUTIONS - Updates merged data with resolved conflict data.
'
' Parameters:
' mergedData - A Dictionary containing initial merged document data.
' resolvedData - A Dictionary of resolved conflict data.
'
' Returns:
' Nothing. The mergedData dictionary is updated in place.
'===============================================================================
Public Sub ApplyConflictResolutions(mergedData As Object, resolvedData As Object)
    ' Validate inputs.
    If mergedData Is Nothing Or resolvedData Is Nothing Then
        Module_SyncTool_Logger.LogMessage "ERROR: Invalid input to ApplyConflictResolutions", "ERROR"
        Exit Sub
    End If
    
    Dim docNum As Variant
    
    For Each docNum In resolvedData.keys
        If mergedData.Exists(docNum) Then
            mergedData.Remove docNum
        End If
        
        mergedData.Add docNum, resolvedData(docNum)
    Next docNum
End Sub


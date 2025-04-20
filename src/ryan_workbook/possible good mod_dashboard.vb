Option Explicit

'============================================================
'                 MODULE-LEVEL CONSTANTS
'============================================================

' --- Workbook Protection ---
' Make PUBLIC so modArchival can use it for sheet protection
Public Const PW_WORKBOOK As String = ""    ' <--- ADD PASSWORD HERE IF ANY (Keep Public)

' --- Sheet Names ---
' Make DASHBOARD_SHEET_NAME Public for modArchival navigation
Public Const DASHBOARD_SHEET_NAME As String = "SQRCT Dashboard"
Public Const USEREDITS_SHEET_NAME As String = "UserEdits" ' Keep Private? (Review if modUtilities needs it Public)
Private Const USEREDITSLOG_SHEET_NAME As String = "UserEditsLog" ' Keep Private
Private Const TEXT_ONLY_SHEET_NAME As String = "SQRCT Dashboard (Text-Only)" ' Keep Private

' --- Data Source Names ---
Private Const MASTER_QUOTES_FINAL_SOURCE As String = "MasterQuotes_Final" ' Keep Private
Private Const PQ_LATEST_LOCATION_SHEET As String = "DocNum_LatestLocation" ' Keep Private
Private Const PQ_LATEST_LOCATION_TABLE As String = "DocNum_LatestLocation" ' Keep Private
Private Const PQ_DOCNUM_COL_NAME As String = "PrimaryDocNumber" ' Keep Private
Private Const PQ_LOCATION_COL_NAME As String = "MostRecent_FolderLocation" ' Keep Private

' --- UserEdits Sheet Columns ---
Public Const UE_COL_DOCNUM      As String = "A"
Public Const UE_COL_PHASE       As String = "B"
Public Const UE_COL_LASTCONTACT As String = "C"
Public Const UE_COL_COMMENTS    As String = "D"
Public Const UE_COL_SOURCE      As String = "E"
Public Const UE_COL_TIMESTAMP   As String = "F"

' --- Dashboard Column Letters (Corrected & Public where needed) ---
' NOTE: J = Workflow, K = Missing Quote (User Confirmed Order)
' *** ADDED Missing Public Constants needed by modArchival ***
Public Const DB_COL_AMOUNT      As String = "D" ' PUBLIC needed for modArchival formatting call
Public Const DB_COL_DOC_DATE    As String = "E" ' PUBLIC needed for modArchival formatting call
Public Const DB_COL_FIRST_PULL  As String = "F" ' PUBLIC needed for modArchival formatting call
Public Const DB_COL_PULL_COUNT  As String = "I" ' PUBLIC needed for modArchival formatting call
' *** End Added Constants ***
Public Const DB_COL_WORKFLOW_LOCATION As String = "J" ' PUBLIC needed for modArchival formatting call
Public Const DB_COL_MISSING_QUOTE As String = "K" ' Keep Public for consistency or future use
Public Const DB_COL_PHASE As String = "L" ' PUBLIC needed for modArchival filtering/formatting
Public Const DB_COL_LASTCONTACT As String = "M" ' PUBLIC needed for modArchival formatting/UserEdits
Public Const DB_COL_COMMENTS As String = "N" ' PUBLIC needed for modArchival range definitions/UserEdits

' --- Other Settings ---
Public Const DEBUG_LOGGING As Boolean = True ' Master switch for DebugLog output
Public Const PHASE_LIST_NAMED_RANGE As String = "PHASE_LIST" ' Name of the range holding valid phases


'===============================================================================
'                         0. CORE HELPER ROUTINES
'===============================================================================
'------------------------------------------------------------------------------
' BuildRowIndexDict - Builds CleanDocNum -> Sheet Row# dictionary for UserEdits
' Specifically for Worksheet_Change event to find rows quickly.
'------------------------------------------------------------------------------
Public Function BuildRowIndexDict(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive keys

    If ws Is Nothing Then
        LogUserEditsOperation "BuildRowIndexDict: Worksheet object not provided."
        Set BuildRowIndexDict = dict ' Return empty
        Exit Function
    End If

    Dim lastR As Long, r As Long
    Dim key As String

    On Error Resume Next ' Handle errors getting last row or reading cells
    lastR = ws.Cells(ws.rows.Count, UE_COL_DOCNUM).End(xlUp).Row
    If Err.Number <> 0 Then
        LogUserEditsOperation "BuildRowIndexDict: Error getting last row from '" & ws.Name & "'. " & Err.Description
        Set BuildRowIndexDict = dict ' Return empty
        Exit Function
    End If
    On Error GoTo 0 ' Restore default error handling for loop

    If lastR >= 2 Then ' Headers are row 1
        For r = 2 To lastR
            key = CleanDocumentNumber(CStr(ws.Cells(r, UE_COL_DOCNUM).value)) ' Use Public CleanDocumentNumber
            If Len(key) > 0 Then
                If Not dict.Exists(key) Then
                     dict.Add key, r ' Key = Cleaned DocNum, Item = Row Number
                ' Else: Keep first row found if duplicates exist
                End If
            End If
        Next r
    End If

    LogUserEditsOperation "BuildRowIndexDict: Built dictionary with " & dict.Count & " row index entries for sheet '" & ws.Name & "'."
    Set BuildRowIndexDict = dict
End Function

Sub ListAllTables()
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            Debug.Print ws.Name, lo.Name
        Next lo
    Next ws
End Sub

'------------------------------------------------------------------------------
' DebugLog - Wrapper for detailed logging during debugging
' Writes to UserEditsLog sheet via LogUserEditsOperation AND prints to Immediate Window.
' Controlled by DEBUG_LOGGING constant.
'------------------------------------------------------------------------------
Public Sub DebugLog(procName As String, msg As String)
    If Not DEBUG_LOGGING Then Exit Sub ' Exit if detailed logging is turned off

    Dim logMsg As String
    logMsg = "[" & procName & "] " & msg ' Format the message

    ' Also print to Immediate Window (Ctrl+G) as a reliable fallback/real-time view
    Debug.Print Now(); " "; logMsg

    ' Attempt to write to the log sheet using the original logger
    On Error Resume Next ' Temporarily ignore errors from the logger itself
    LogUserEditsOperation logMsg ' Call the original logger
    If Err.Number <> 0 Then
        ' If main logger failed, print a warning to Immediate Window
        Debug.Print Now(); " [DebugLog] WARNING: Call to LogUserEditsOperation failed. Error: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0 ' Restore default error handling
End Sub

'------------------------------------------------------------------------------
' ResolvePhase - Determines the correct Engagement Phase to display on dashboard
'------------------------------------------------------------------------------
Public Function ResolvePhase(ByVal historicStage As String, _
                             ByVal autoStage As String, _
                             ByVal userPhase As String, _
                             ByVal dataSource As String) As String
    ' Purpose: Prioritize UserEdits unless it's the specific "Legacy Process" placeholder
    '          For Legacy source, fall back to HistoricStage if no valid UserEdit
    '          For non-Legacy source, fall back to AutoStage if no valid UserEdit

    Dim cleanUserPhase As String
    cleanUserPhase = Trim$(userPhase)

    ' Treat the specific placeholder text literally as "nothing entered by user"
    If LCase$(cleanUserPhase) = "legacy process" Then cleanUserPhase = ""

    ' Determine final phase based on DataSource
    If LCase$(dataSource) = "legacy" Then
        ' For legacy data source records:
        If cleanUserPhase <> "" Then
            ResolvePhase = cleanUserPhase    ' User entered a real override
        Else
            ResolvePhase = historicStage     ' No valid user edit, use the original Historic Stage
        End If
    Else
        ' For non-legacy data source records (e.g., "CSV"):
        If cleanUserPhase <> "" Then
            ResolvePhase = cleanUserPhase    ' User entered an override
        Else
            ResolvePhase = autoStage         ' No valid user edit, use the calculated AutoStage
        End If
    End If

End Function

'------------------------------------------------------------------------------
' RowIndexDictAdd - Helper to update Row Index Dictionary after adding new row
'------------------------------------------------------------------------------
Public Sub RowIndexDictAdd(ByRef dict As Object, ByVal key As String, ByVal rowNum As Long)
    ' Adds item to dictionary only if key doesn't already exist
    If dict Is Nothing Then Exit Sub
    If Len(key) > 0 Then
        If Not dict.Exists(key) Then
            dict.Add key, rowNum
        End If
    End If
End Sub
'-------------------------------------------------------------------
' Returns a Dictionary:  hdr("Document Amount") = 7   (etc.)
' It tries the table first, then a named range.
'-------------------------------------------------------------------
Private Function GetMQF_HeaderMap() As Object
    Dim d       As Object:    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Dim lo      As ListObject
    Dim ws      As Worksheet
    Dim hdrRow  As Variant
    Dim c       As Long

    ' 1) Try to find the table by iterating every ListObject on every sheet
    On Error Resume Next ' Added safety in case of issues accessing ListObjects collection itself
    For Each ws In ThisWorkbook.Worksheets
        If ws.ListObjects.Count > 0 Then ' Only check if the sheet HAS listobjects
            For Each lo In ws.ListObjects
                If lo.Name = MASTER_QUOTES_FINAL_SOURCE Then
                    hdrRow = lo.HeaderRowRange.value
                    Err.Clear ' Clear any error from HeaderRowRange access
                    Exit For ' Exit inner loop once found
                End If
            Next lo
        End If
        If Not IsEmpty(hdrRow) Then Exit For ' Exit outer loop if found
    Next ws
    On Error GoTo 0 ' Restore default error handling

    ' 2) If we still have nothing, fall back to a Named Range
    If IsEmpty(hdrRow) Then
        Dim rng As Range ' Declare here
        On Error Resume Next ' Check for named range existence
        Set rng = Nothing ' Reset rng
        Set rng = ThisWorkbook.Names(MASTER_QUOTES_FINAL_SOURCE).RefersToRange
        On Error GoTo 0 ' Restore default error handling

        If Not rng Is Nothing Then
             On Error Resume Next ' Handle error reading row 1 value
             hdrRow = rng.rows(1).value
             If Err.Number <> 0 Then
                 LogUserEditsOperation "GetMQF_HeaderMap: Found Named Range '" & MASTER_QUOTES_FINAL_SOURCE & "' but failed to read header row: " & Err.Description
                 hdrRow = Empty ' Ensure hdrRow is empty if read failed
                 Err.Clear
             End If
             On Error GoTo 0
        Else
            ' Neither Table nor Named Range found - log and bail out
            LogUserEditsOperation "GetMQF_HeaderMap: CRITICAL - Could not find source '" & MASTER_QUOTES_FINAL_SOURCE & "' as Table or Named Range."
            Set GetMQF_HeaderMap = d ' Return empty dictionary
            Exit Function
        End If
    End If

    ' 3) Build the dictionary from the header row array
    If IsArray(hdrRow) Then
         On Error Resume Next ' Handle potential errors accessing array bounds/values
         For c = 1 To UBound(hdrRow, 2)
             Dim headerName As String
             headerName = Trim$(hdrRow(1, c) & "") ' Ensure it's a string and trimmed
             If Len(headerName) > 0 Then
                 If Not d.Exists(headerName) Then
                     d.Add headerName, c
                 Else
                     LogUserEditsOperation "GetMQF_HeaderMap: Warning - Duplicate header '" & headerName & "' found. Using first instance."
                 End If
             End If
         Next c
         If Err.Number <> 0 Then
              LogUserEditsOperation "GetMQF_HeaderMap: Error building dictionary from header row array: " & Err.Description
              Err.Clear
         End If
         On Error GoTo 0
    ElseIf Not IsEmpty(hdrRow) Then ' Handle case where header range might be single cell
         Dim headerNameSingle As String
         headerNameSingle = Trim$(hdrRow & "")
         If Len(headerNameSingle) > 0 Then d(headerNameSingle) = 1
    Else
         LogUserEditsOperation "GetMQF_HeaderMap: Header row data (hdrRow) was empty or invalid after source lookup."
    End If

    Set GetMQF_HeaderMap = d ' Return the dictionary (might be empty if errors occurred)
End Function

' Quick wrapper to fetch an index or raise a clear log + error
Private Function MQFIdx(hdrMap As Object, hdrName As String, proc As String) As Long
    If hdrMap.Exists(hdrName) Then
        MQFIdx = hdrMap(hdrName)
    Else
        LogUserEditsOperation proc & ": REQUIRED column """ & hdrName & """ not found."
        Err.Raise vbObjectError + 513, , proc & ": Missing column """ & hdrName & """"
    End If
End Function

    '-------------------------------------------------------------------------------
    ' GetTableOrRangeData
    ' Reads data from an Excel Table (ListObject) or a Named Range into a 2D Variant Array.
    ' Returns an empty array if source not found or empty.
    '-------------------------------------------------------------------------------
    Private Function GetTableOrRangeData(SourceName As String) As Variant
        Dim lo As ListObject
        Dim rng As Range
        Dim ws As Worksheet
        Dim dataArray As Variant
        Dim sourceFound As Boolean

        ' Default to an empty array (specifically, Variant/Empty)
        GetTableOrRangeData = VBA.Array() ' Use VBA.Array() to return an empty variant

        On Error Resume Next ' Allow checking for objects without halting

        ' Check for Table (ListObject) on any sheet
        Set lo = Nothing: Err.Clear
        For Each ws In ThisWorkbook.Worksheets
            Set lo = ws.ListObjects(SourceName)
            If Err.Number = 0 And Not lo Is Nothing Then Exit For ' Found it
        Next ws

        If Not lo Is Nothing Then
            ' Source is a Table
            If Not lo.DataBodyRange Is Nothing Then
                If lo.ListRows.Count > 0 Then
                    dataArray = lo.DataBodyRange.Value2 ' Use Value2 for performance
                    sourceFound = True
                End If
            End If
            LogUserEditsOperation "GetTableOrRangeData: Read " & lo.ListRows.Count & " rows from Table '" & SourceName & "'."
        Else
            ' Check for Named Range if not found as Table
            Set rng = Nothing: Err.Clear
            Set rng = ThisWorkbook.Names(SourceName).RefersToRange
            If Err.Number = 0 And Not rng Is Nothing Then
                 ' Source is a Named Range - assume data starts from first row
                 If Application.WorksheetFunction.CountA(rng) > 0 Then ' Basic check if range has any data
                     If rng.rows.Count > 0 And rng.Columns.Count > 0 Then
                         dataArray = rng.Value2 ' Use Value2
                         sourceFound = True
                     End If
                 End If
                 LogUserEditsOperation "GetTableOrRangeData: Read data from Named Range '" & SourceName & "'."
            End If
        End If

        On Error GoTo 0 ' Restore default error handling

        If sourceFound Then
            ' Ensure result is always a 2D array, even if only one row/column was read
            If Not IsArray(dataArray) Then
                 ' Handle single cell case
                 Dim tempData(1 To 1, 1 To 1) As Variant
                 tempData(1, 1) = dataArray
                 GetTableOrRangeData = tempData
            ElseIf LBound(dataArray, 1) > UBound(dataArray, 1) Then
                 ' Handle empty array from range read (rare case)
                 GetTableOrRangeData = VBA.Array()
            ElseIf UBound(dataArray, 1) >= LBound(dataArray, 1) And UBound(dataArray, 2) = 1 Then
                 ' Handle single column case - already 2D by default when read from range
                 GetTableOrRangeData = dataArray
            ElseIf UBound(dataArray, 1) = 1 And UBound(dataArray, 2) > 1 Then
                 ' Handle single row case - Reading Range.Value2 usually returns 2D (1 to 1, 1 to N)
                 GetTableOrRangeData = dataArray
            Else
                 ' Standard 2D array
                 GetTableOrRangeData = dataArray
            End If
        Else
            LogUserEditsOperation "GetTableOrRangeData: Source '" & SourceName & "' not found or empty."
            ' Return the empty array initialized at the start
        End If

    End Function

    '-------------------------------------------------------------------------------
    ' BuildDictionaryFromArray
    ' Creates a Scripting Dictionary from a 2D array.
    ' KeyColIndex: 1-based index of the column in the array to use for dictionary Keys.
    ' ValueColIndex: 1-based index of the column for dictionary Items.
    ' Uses CleanDocumentNumber on keys if useCleanKey is True.
    ' Handles potential errors gracefully.
    '-------------------------------------------------------------------------------
    Private Function BuildDictionaryFromArray(InputArray As Variant, KeyColIndex As Long, ValueColIndex As Long, Optional UseCleanKey As Boolean = False) As Object
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = vbTextCompare ' Default to case-insensitive keys

        Dim r As Long
        Dim keyVal As String
        Dim itemVal As Variant

        If Not IsArray(InputArray) Then
            LogUserEditsOperation "BuildDictionaryFromArray: Input is not a valid array."
            Set BuildDictionaryFromArray = dict ' Return empty dictionary
            Exit Function
        End If
        ' Check if array actually has data rows
        If LBound(InputArray, 1) > UBound(InputArray, 1) Then
             LogUserEditsOperation "BuildDictionaryFromArray: Input array is empty."
             Set BuildDictionaryFromArray = dict ' Return empty dictionary
             Exit Function
        End If


        On Error Resume Next ' Handle potential errors like array bounds

        For r = LBound(InputArray, 1) To UBound(InputArray, 1) ' Use LBound/UBound for safety
            If KeyColIndex > UBound(InputArray, 2) Or ValueColIndex > UBound(InputArray, 2) Then
                LogUserEditsOperation "BuildDictionaryFromArray: Key or Value column index out of bounds for array size (" & UBound(InputArray, 2) & ")."
                Set BuildDictionaryFromArray = dict ' Return potentially partially built dict
                Exit Function ' Stop processing if indices are invalid
            End If

            If UseCleanKey Then
                keyVal = CleanDocumentNumber(CStr(InputArray(r, KeyColIndex)))
            Else
                keyVal = Trim$(CStr(InputArray(r, KeyColIndex)))
            End If

            itemVal = InputArray(r, ValueColIndex)

            If keyVal <> "" Then
                If Not dict.Exists(keyVal) Then
                    dict.Add keyVal, itemVal
                Else
                    ' Handle duplicate keys if necessary - default keeps the first found
                    ' LogUserEditsOperation "BuildDictionaryFromArray: Duplicate key '" & keyVal & "' found at array row " & r & ". Keeping first value."
                End If
            End If
        Next r

        If Err.Number <> 0 Then
            LogUserEditsOperation "BuildDictionaryFromArray: Error during dictionary build: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0 ' Restore default

        Set BuildDictionaryFromArray = dict
    End Function

'------------------------------------------------------------------------------
' CleanDocumentNumber - Standardizes DocNum for matching.
' Mirrors Power Query Step Q logic for consistency.
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Function CleanDocumentNumber(ByVal raw As String) As String
    Dim cleaned As String
    Dim prefix As String, tail As String
    Dim num As Long, tailLen As Long

    ' 1. Trim and Uppercase
    cleaned = UCase$(Trim$(raw))

    ' 2. Handle empty or short strings
    If Len(cleaned) = 0 Then CleanDocumentNumber = "": Exit Function
    If Len(cleaned) <= 5 Then CleanDocumentNumber = cleaned: Exit Function ' Return as-is if 5 chars or less

    ' 3. Split into Prefix (5 chars) and Tail
    prefix = Left$(cleaned, 5)
    tail = Mid$(cleaned, 6)

    ' 4. Handle BSMOQ prefix exception (return as-is)
    If prefix = "BSMOQ" Then CleanDocumentNumber = cleaned: Exit Function

    ' 5. Pad numeric tail to 5 digits (or original length if longer)
    If VBA.isNumeric(tail) Then ' Use VBA.IsNumeric for clarity
        On Error Resume Next ' Handle potential overflow if tail is huge number
        num = Val(tail)
        If Err.Number <> 0 Then ' Val failed (e.g., too large)
            Err.Clear
            ' Keep original tail if Val fails
        Else
            On Error GoTo 0 ' Restore error handling
            tailLen = Len(tail)
            ' Ensure padding doesn't truncate if original tail was > 5 digits and numeric
            tail = Right$("00000" & CStr(num), IIf(tailLen < 5, 5, tailLen))
        End If
        On Error GoTo 0
    Else
        ' Keep original tail if it's not purely numeric
    End If

    ' 6. Combine and return
    CleanDocumentNumber = prefix & tail
End Function

'------------------------------------------------------------------------------
' LogUserEditsOperation – Timestamped entry to hidden log sheet
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Sub LogUserEditsOperation(msg As String)
    Dim wsLog As Worksheet, r As Long
    On Error Resume Next ' Prevent error if sheet exists but is protected differently
    Set wsLog = ThisWorkbook.Sheets(USEREDITSLOG_SHEET_NAME)
    On Error GoTo 0 ' Restore error handling

    If wsLog Is Nothing Then
        On Error Resume Next ' Prevent error if workbook is protected
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        If wsLog Is Nothing Then Exit Sub ' Failed to add sheet
        On Error GoTo 0
        wsLog.Name = USEREDITSLOG_SHEET_NAME
        wsLog.Range("A1:C1").value = Array("Timestamp", "Workbook", "Operation")
        wsLog.Range("A1:C1").Font.Bold = True
        wsLog.Visible = xlSheetHidden ' Hide it by default
    End If

    On Error Resume Next ' Avoid error if log sheet is protected
    r = wsLog.Cells(wsLog.rows.Count, "A").End(xlUp).Row + 1
    If r < 2 Then r = 2 ' Ensure we start writing at row 2 if sheet was empty
    wsLog.Cells(r, "A").value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    wsLog.Cells(r, "B").value = Module_Identity.GetWorkbookIdentity() ' Use Identity Module
    wsLog.Cells(r, "C").value = msg
    If Err.Number <> 0 Then Debug.Print "Error writing to UserEditsLog: " & Err.Description: Err.Clear
    On Error GoTo 0
End Sub

    '------------------------------------------------------------------------------
    ' LoadUserEditsToDictionary (MODIFIED)
    ' Reads UserEdits sheet, builds Dictionary mapping CleanDocNum -> Array(Phase, Contact, Comments).
    '------------------------------------------------------------------------------
    Public Function LoadUserEditsToDictionary() As Object ' Removed wsEdits parameter, uses constant
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = vbTextCompare ' Use case-insensitive comparison

        Dim wsEdits As Worksheet
        Dim lastRow As Long
        Dim dataArray As Variant
        Dim r As Long
        Dim cleanedDocNum As String
        Dim userEditValues(0 To 2) As Variant ' Array to hold Phase, Contact, Comments

        ' Get UserEdits sheet reference
        On Error Resume Next
        Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
        On Error GoTo 0 ' Restore error handling for main sub logic
        If wsEdits Is Nothing Then
            LogUserEditsOperation "LoadUserEditsToDictionary: UserEdits sheet '" & USEREDITS_SHEET_NAME & "' not found."
            Set LoadUserEditsToDictionary = dict ' Return empty dictionary
            Exit Function
        End If

        ' Find last row with data in DocNum column
        lastRow = wsEdits.Cells(wsEdits.rows.Count, UE_COL_DOCNUM).End(xlUp).Row

        If lastRow <= 1 Then
            LogUserEditsOperation "LoadUserEditsToDictionary: No data found on UserEdits sheet."
            Set LoadUserEditsToDictionary = dict ' Return empty dictionary
            Exit Function
        End If

        ' Read relevant columns (A to D) into an array for speed
        On Error Resume Next
        dataArray = wsEdits.Range(UE_COL_DOCNUM & "2:" & UE_COL_COMMENTS & lastRow).Value2 ' Read A:D
        If Err.Number <> 0 Then
             LogUserEditsOperation "LoadUserEditsToDictionary: Error reading data from UserEdits sheet: " & Err.Description
             Set LoadUserEditsToDictionary = dict ' Return empty dictionary
             Exit Function
        End If
        On Error GoTo 0

        ' Process the array
        If IsArray(dataArray) Then
             ' Check if it's a 2D array (multiple rows) or 1D (single row)
             Dim is2D As Boolean
             On Error Resume Next
             Dim checkBound As Long: checkBound = UBound(dataArray, 2)
             is2D = (Err.Number = 0)
             On Error GoTo 0

             If is2D Then ' Multiple rows read
                 For r = 1 To UBound(dataArray, 1) ' Loop through rows of the array
                     cleanedDocNum = CleanDocumentNumber(CStr(dataArray(r, 1))) ' Col 1 = DocNum (from Col A)

                     If cleanedDocNum <> "" Then
                         If Not dict.Exists(cleanedDocNum) Then
                            ' Store Phase (Col B -> Array Index 2), LastContact (Col C -> 3), Comments (Col D -> 4)
                            userEditValues(0) = dataArray(r, 2) ' Phase
                            userEditValues(1) = dataArray(r, 3) ' Last Contact
                            userEditValues(2) = dataArray(r, 4) ' Comments
                            dict.Add key:=cleanedDocNum, Item:=userEditValues ' Store the array as the item
                         Else
                            ' Log duplicate cleaned key if needed
                         End If
                     End If
                 Next r
                 LogUserEditsOperation "LoadUserEditsToDictionary: Processed " & UBound(dataArray, 1) & " rows from UserEdits sheet."
             Else ' Single row of data read (returned as 1D array)
                 cleanedDocNum = CleanDocumentNumber(CStr(dataArray(1))) ' Index 1 = DocNum
                 If cleanedDocNum <> "" Then
                     If Not dict.Exists(cleanedDocNum) Then
                         userEditValues(0) = dataArray(2) ' Phase
                         userEditValues(1) = dataArray(3) ' Last Contact
                         userEditValues(2) = dataArray(4) ' Comments
                         dict.Add key:=cleanedDocNum, Item:=userEditValues
                         LogUserEditsOperation "LoadUserEditsToDictionary: Processed 1 row from UserEdits sheet."
                     End If
                 End If
             End If
        Else
            ' Handle case where only one cell was read (unlikely for A:D but safe)
            LogUserEditsOperation "LoadUserEditsToDictionary: Read only single value, expected array."
            ' Cannot process further
        End If

        Set LoadUserEditsToDictionary = dict ' Return the dictionary
    End Function

'------------------------------------------------------------------------------
' IsMasterQuotesFinalPresent - Checks if the source exists (PQ, Table, or NR)
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Function IsMasterQuotesFinalPresent() As Boolean
    Dim lo As ListObject
    Dim nm As Name
    Dim queryObj As Object
    Dim ws As Worksheet

    IsMasterQuotesFinalPresent = False
    On Error Resume Next ' Ignore errors during checks

    ' 1. Check for Power Query
    Set queryObj = Nothing: Err.Clear
    Set queryObj = ThisWorkbook.Queries(MASTER_QUOTES_FINAL_SOURCE)
    If Err.Number = 0 And Not queryObj Is Nothing Then
        IsMasterQuotesFinalPresent = True
        GoTo ExitPoint_IsPresent ' Found it
    End If

    ' 2. Check for Table (ListObject) on any sheet
    Set lo = Nothing: Err.Clear
    For Each ws In ThisWorkbook.Worksheets
        Set lo = ws.ListObjects(MASTER_QUOTES_FINAL_SOURCE)
        If Err.Number = 0 And Not lo Is Nothing Then
            IsMasterQuotesFinalPresent = True
            GoTo ExitPoint_IsPresent ' Found it
        End If
    Next ws

    ' 3. Check for Named Range
    Set nm = Nothing: Err.Clear
    Set nm = ThisWorkbook.Names(MASTER_QUOTES_FINAL_SOURCE)
    If Err.Number = 0 And Not nm Is Nothing Then
        ' Further check if the named range actually refers to a valid range
        Dim tempRange As Range
        Set tempRange = Nothing
        Set tempRange = nm.RefersToRange
        If Err.Number = 0 And Not tempRange Is Nothing Then
             IsMasterQuotesFinalPresent = True
             GoTo ExitPoint_IsPresent ' Found it
        End If
    End If

ExitPoint_IsPresent:
    On Error GoTo 0 ' Restore default error handling
    Set lo = Nothing: Set nm = Nothing: Set queryObj = Nothing: Set ws = Nothing: Set tempRange = Nothing
End Function

'-------------------------------------------------------------------------------
' BuildDashboardDataArray (REFACTORED v3 - Using Header Map & ResolvePhase)
' Loads data from PQ sources and UserEdits into memory, merges them,
' and returns a 1-based 2D Variant array formatted for SQRCT Dashboard (A:N).
' Uses GetMQF_HeaderMap/MQFIdx helpers for robustness against column reordering.
' Uses ResolvePhase helper to handle Engagement Phase logic correctly.
' *** ORDER: J=Workflow Location, K=Missing Quote Alert ***
'-------------------------------------------------------------------------------
Private Function BuildDashboardDataArray() As Variant
    ' --- Declare Variables ---
    Dim arrQuotes As Variant        ' From MasterQuotes_Final source
    Dim arrLoc As Variant           ' From DocNum_LatestLocation source
    Dim dictLoc As Object           ' Lookup: CleanDocNum -> Location
    Dim dictUserEdits As Object     ' Lookup: CleanDocNum -> Array(Phase, Contact, Comments)
    Dim arrOutput() As Variant      ' Final output array (1-based, 14 columns)
    Dim hdrMap As Object            ' Dictionary mapping Header Name -> Source Column Index
    Dim r As Long, numRows As Long  ' Loop counter, row count
    Dim cleanedDocNum As String, rawDocNum As String
    Dim editData As Variant         ' Temp array for user edits
    Dim startTime As Double: startTime = Timer
    Dim i As Long                   ' Generic loop counter

    '--- Column Indices for Location data (Still using fixed indices for this source) ---
    Const LOC_IDX_DOCNUM As Long = 1
    Const LOC_IDX_LOCATION As Long = 2

    LogUserEditsOperation "BuildDashboardDataArray (Header Map + ResolvePhase): Starting..."

    '--- STEP 1: Get Header Map for MasterQuotes_Final source ---
    Set hdrMap = GetMQF_HeaderMap() ' Calls helper function
    If hdrMap Is Nothing Or hdrMap.Count = 0 Then
        LogUserEditsOperation "BuildDashboardDataArray: CRITICAL - Could not build header map for MasterQuotes_Final. Aborting."
        GoTo Fail_Build ' Use GoTo for failure exit
    End If

    '--- STEP 2: Dynamically Get Column Indices using Header Names ---
    '    MQFIdx helper function returns 0 and logs error if header not found.
    Dim IDX_DOCNUM As Long:      IDX_DOCNUM = MQFIdx(hdrMap, "Document Number", "BuildDashboardDataArray")
    Dim IDX_CUSTNUM As Long:     IDX_CUSTNUM = MQFIdx(hdrMap, "Customer Number", "BuildDashboardDataArray")
    Dim IDX_CUSTNAME As Long:    IDX_CUSTNAME = MQFIdx(hdrMap, "Customer Name", "BuildDashboardDataArray")
    Dim IDX_DOCAMT As Long:      IDX_DOCAMT = MQFIdx(hdrMap, "Document Amount", "BuildDashboardDataArray")
    Dim IDX_DOCDATE As Long:     IDX_DOCDATE = MQFIdx(hdrMap, "Document Date", "BuildDashboardDataArray")
    Dim IDX_FIRSTPULL As Long:   IDX_FIRSTPULL = MQFIdx(hdrMap, "First Date Pulled", "BuildDashboardDataArray")
    Dim IDX_SPID As Long:        IDX_SPID = MQFIdx(hdrMap, "Salesperson ID", "BuildDashboardDataArray")
    Dim IDX_USERENTER As Long:   IDX_USERENTER = MQFIdx(hdrMap, "User To Enter", "BuildDashboardDataArray")
    Dim IDX_PULLCOUNT As Long:   IDX_PULLCOUNT = MQFIdx(hdrMap, "Pull Count", "BuildDashboardDataArray")
    Dim IDX_AUTONOTE As Long:    IDX_AUTONOTE = MQFIdx(hdrMap, "AutoNote", "BuildDashboardDataArray")      ' Source for K (Missing Quote)
    Dim IDX_AUTOSTAGE As Long:   IDX_AUTOSTAGE = MQFIdx(hdrMap, "AutoStage", "BuildDashboardDataArray")    ' Default for L (Phase)
    Dim IDX_DATASOURCE As Long:  IDX_DATASOURCE = MQFIdx(hdrMap, "DataSource", "BuildDashboardDataArray") ' Needed for ResolvePhase
    Dim IDX_HISTORICSTAGE As Long: IDX_HISTORICSTAGE = MQFIdx(hdrMap, "Historic Stage", "BuildDashboardDataArray") ' Needed for ResolvePhase

    '--- Check if any essential index lookup failed ---
    If IDX_DOCNUM = 0 Or IDX_CUSTNUM = 0 Or IDX_CUSTNAME = 0 Or IDX_DOCAMT = 0 Or _
       IDX_DOCDATE = 0 Or IDX_FIRSTPULL = 0 Or IDX_SPID = 0 Or IDX_USERENTER = 0 Or _
       IDX_PULLCOUNT = 0 Or IDX_AUTONOTE = 0 Or IDX_AUTOSTAGE = 0 Or _
       IDX_DATASOURCE = 0 Or IDX_HISTORICSTAGE = 0 Then
            LogUserEditsOperation "BuildDashboardDataArray: CRITICAL - One or more required headers not found in MasterQuotes_Final source via header map. Aborting."
            GoTo Fail_Build
    End If
    '--- End Header Map Lookup ---

    ' --- STEP 3: Load Data ---
    arrQuotes = GetTableOrRangeData(MASTER_QUOTES_FINAL_SOURCE)
    arrLoc = GetTableOrRangeData(PQ_LATEST_LOCATION_TABLE)
    Set dictUserEdits = LoadUserEditsToDictionary() ' Uses refactored version returning arrays

    ' --- STEP 4: Validate arrQuotes Array State ---
    If Not IsArray(arrQuotes) Then GoTo Fail_Build ' Array not returned
    If LBound(arrQuotes, 1) > UBound(arrQuotes, 1) Then GoTo Fail_Build ' Array has no rows
    ' No column count check needed here - handled by MQFIdx checks above
    numRows = UBound(arrQuotes, 1)
    LogUserEditsOperation "BuildDashboardDataArray: Loaded " & numRows & " rows from MasterQuotes_Final."

    ' --- STEP 5: Build Location Dictionary ---
    Set dictLoc = BuildDictionaryFromArray(arrLoc, LOC_IDX_DOCNUM, LOC_IDX_LOCATION, True) ' Uses indices
    LogUserEditsOperation "BuildDashboardDataArray: Built Location Dictionary with " & dictLoc.Count & " entries."
    LogUserEditsOperation "BuildDashboardDataArray: Loaded User Edits Dictionary with " & dictUserEdits.Count & " entries."

    ' --- STEP 6: Prepare Output Array (1 to numRows, 1 to 14 corresponding to A:N) ---
    ReDim arrOutput(1 To numRows, 1 To 14)

    ' --- STEP 7: Loop through Source Data and Merge ---
    Dim currentHistoricStage As String
    Dim currentAutoStage As String
    Dim currentUserPhase As String
    Dim currentDataSource As String

    For r = 1 To numRows
        ' Get Cleaned Document Number Key using dynamic index
        rawDocNum = CStr(arrQuotes(r, IDX_DOCNUM))
        cleanedDocNum = CleanDocumentNumber(rawDocNum)

        ' Populate A-I directly using dynamic indices
        arrOutput(r, 1) = rawDocNum
        arrOutput(r, 2) = arrQuotes(r, IDX_CUSTNUM)
        arrOutput(r, 3) = arrQuotes(r, IDX_CUSTNAME)
        arrOutput(r, 4) = arrQuotes(r, IDX_DOCAMT)
        arrOutput(r, 5) = arrQuotes(r, IDX_DOCDATE)
        arrOutput(r, 6) = arrQuotes(r, IDX_FIRSTPULL)
        arrOutput(r, 7) = arrQuotes(r, IDX_SPID)
        arrOutput(r, 8) = arrQuotes(r, IDX_USERENTER)
        arrOutput(r, 9) = arrQuotes(r, IDX_PULLCOUNT)

        ' Populate J (Workflow Location) using lookup (Index 10)
        If dictLoc.Exists(cleanedDocNum) Then
            arrOutput(r, 10) = dictLoc(cleanedDocNum)
            If Len(Trim(CStr(arrOutput(r, 10)))) = 0 Then arrOutput(r, 10) = "Quote Only"
        Else
            arrOutput(r, 10) = "Quote Only"
        End If

        ' Populate K (Missing Quote Alert) from AutoNote using dynamic index (Index 11)
        arrOutput(r, 11) = arrQuotes(r, IDX_AUTONOTE)

        ' Populate L-N using ResolvePhase logic
        ' Get values needed for ResolvePhase from the source array
        currentHistoricStage = CStr(arrQuotes(r, IDX_HISTORICSTAGE))
        currentAutoStage = CStr(arrQuotes(r, IDX_AUTOSTAGE))
        currentDataSource = CStr(arrQuotes(r, IDX_DATASOURCE))

        ' Get User Edit Phase if it exists, otherwise pass blank
        If dictUserEdits.Exists(cleanedDocNum) Then
            editData = dictUserEdits(cleanedDocNum) ' Retrieve the array(Phase, Contact, Comments)
            currentUserPhase = CStr(editData(0))
            ' Populate M and N directly from UserEdits if present
            arrOutput(r, 13) = editData(1) ' M: UserEdit LastContact
            arrOutput(r, 14) = editData(2) ' N: UserEdit Comments
        Else
            currentUserPhase = "" ' No user edit exists for Phase
            ' Set default blanks for M and N if no user edit found
            arrOutput(r, 13) = vbNullString ' M: Default LastContact (Blank)
            arrOutput(r, 14) = vbNullString ' N: Default Comments (Blank)
        End If

        ' Call helper function to determine final Phase for Column L (Index 12)
        arrOutput(r, 12) = ResolvePhase(currentHistoricStage, currentAutoStage, currentUserPhase, currentDataSource)

    Next r

    LogUserEditsOperation "BuildDashboardDataArray: Merge complete. Time: " & Format(Timer - startTime, "0.00") & "s"
    BuildDashboardDataArray = arrOutput ' Return the final array
    Exit Function ' Successful exit

Fail_Build: ' Label for failure exits
    LogUserEditsOperation "BuildDashboardDataArray: ERROR - Failed to load or validate source data. Returning False."
    BuildDashboardDataArray = False ' Indicate failure by returning False (Variant/Boolean)
    ' Ensure objects are cleaned up even on failure
    Set hdrMap = Nothing
    Set dictLoc = Nothing
    Set dictUserEdits = Nothing
End Function

    

'===============================================================================
'           1. DASHBOARD CREATION / REFRESH MASTER ROUTINE
'===============================================================================

'------------------------------------------------------------------------------
' Button macros (Assign these to shapes/buttons on the dashboard)
' *** Kept from User's Provided Code ***
'------------------------------------------------------------------------------
Public Sub Button_RefreshDashboard_SaveAndRestoreEdits()
    ' Standard workflow: Saves dashboard edits (L-N) -> Refreshes A-L -> Restores all UserEdits (L-N)
    RefreshDashboard PreserveUserEdits:=False
End Sub

Public Sub Button_RefreshDashboard_PreserveUserEdits() ' Added based on button definition in modArchival
    ' Preserve workflow: Refreshes A-L using source + UserEdits, dashboard edits (L-N) are ignored/overwritten
    RefreshDashboard PreserveUserEdits:=True
End Sub

'------------------------------------------------------------------------------
' RefreshDashboard — Master routine orchestrating the entire refresh
' REVISED: Includes Number Formatting Fix, Corrected Step 9 for UI Consistency
'          FIXED: Call to ApplyPhaseValidationToListColumn now references modUtilities
'------------------------------------------------------------------------------
Public Sub RefreshDashboard(Optional PreserveUserEdits As Boolean = False)

    Dim ws As Worksheet, wsEdits As Worksheet
    Dim lastRow As Long ' Last row *on the dashboard sheet* after writing data
    Dim t_Start As Double, t_Build As Double, t_Write As Double, t_Sort As Double
    Dim t_Format As Double, t_Protect As Double, t_TextOnly As Double, t_Archival As Double ' Timers
    Dim backupCreated As Boolean
    Dim calcState As XlCalculation: calcState = Application.Calculation ' Store current calculation state
    Dim eventsState As Boolean: eventsState = Application.EnableEvents ' Store current event state
    Dim screenState As Boolean: screenState = Application.ScreenUpdating ' Store current screen state
    Dim currentSheet As Worksheet: Set currentSheet = ActiveSheet ' Remember active sheet
    Dim arrFinalData As Variant ' Array to hold the merged data for the dashboard
    Dim wbWasLocked As Boolean
    Dim arrIsEmpty As Boolean ' Declare here
    Dim errNum As Long, errDesc As String, errLine As Long ' For Error Handler

    t_Start = Timer ' Start overall timer
    DebugLog "RefreshDashboard", "ENTER. Mode: " & IIf(PreserveUserEdits, "PreserveUserEdits", "SaveAndRestore")

    ' --- Error Handling & Application Settings ---
    On Error GoTo ErrorHandler ' Master error handler for the refresh process
    DebugLog "RefreshDashboard", "Applying initial application settings (ScreenUpdating=False, EnableEvents=False, Calc=Manual, Alerts=False)..."
    Application.ScreenUpdating = False
    Application.EnableEvents = False ' Turn off events during manipulation
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' --- Check & Handle Workbook Structure Protection ---
    DebugLog "RefreshDashboard", "Checking Workbook Structure Protection..."
    wbWasLocked = ThisWorkbook.ProtectStructure
    DebugLog "RefreshDashboard", "Workbook Structure Locked = " & wbWasLocked
    If wbWasLocked Then
        DebugLog "RefreshDashboard", "Workbook structure is protected. Attempting temporary unlock..."
        If Not ToggleWorkbookStructure(False) Then ' Call helper to Unlock Structure
            MsgBox "Failed to unprotect workbook structure. Required for creating backup/log/setup sheets. Aborting.", vbCritical, "Structure Protection Error"
            GoTo Cleanup ' Abort if unlock fails
        End If
         DebugLog "RefreshDashboard", "Workbook structure unlocked."
    End If

    ' --- Attempt pre-emptive backup ---
    DebugLog "RefreshDashboard", "Attempting pre-refresh backup..."
    backupCreated = CreateUserEditsBackup(Format(Now, "yyyymmdd_hhmmss"))
    DebugLog "RefreshDashboard", "Pre-refresh UserEdits backup created: " & backupCreated

    '--- STEP 1: Ensure UserEdits Sheet Exists and Get Reference ---
    DebugLog "RefreshDashboard", "Step 1: SetupUserEditsSheet..."
    SetupUserEditsSheet ' Creates or verifies the UserEdits sheet structure
    Set wsEdits = Nothing
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    On Error GoTo ErrorHandler
    If wsEdits Is Nothing Then
        DebugLog "RefreshDashboard", "CRITICAL ERROR - Could not find or create '" & USEREDITS_SHEET_NAME & "'. Aborting."
        MsgBox "CRITICAL ERROR: Could not find or create the '" & USEREDITS_SHEET_NAME & "' sheet. Aborting refresh.", vbCritical, "Refresh Aborted"
        GoTo Cleanup
    End If
    DebugLog "RefreshDashboard", "Step 1: UserEdits sheet object obtained."

    '--- STEP 2: Save Current Dashboard Edits (L-N) to UserEdits (if not preserving) ---
    If Not PreserveUserEdits Then
        DebugLog "RefreshDashboard", "Step 2: SaveAndRestore Mode - Calling SaveUserEditsFromDashboard..."
        SaveUserEditsFromDashboard
        DebugLog "RefreshDashboard", "Step 2: Finished SaveUserEditsFromDashboard."
    Else
        DebugLog "RefreshDashboard", "Step 2: PreserveUserEdits Mode - Skipping save of dashboard edits."
    End If

    '--- STEP 3: Get or Create Dashboard Sheet & Prepare Layout ---
    DebugLog "RefreshDashboard", "Step 3: GetOrCreateDashboardSheet..."
    Set ws = GetOrCreateDashboardSheet(DASHBOARD_SHEET_NAME)
    If ws Is Nothing Then
         DebugLog "RefreshDashboard", "CRITICAL ERROR - Could not find or create '" & DASHBOARD_SHEET_NAME & "'. Aborting."
         MsgBox "CRITICAL ERROR: Could not find or create the '" & DASHBOARD_SHEET_NAME & "' sheet. Aborting refresh.", vbCritical, "Refresh Aborted"
         GoTo Cleanup
    End If
    DebugLog "RefreshDashboard", "Step 3: Dashboard sheet object obtained: '" & ws.Name & "'"

    On Error Resume Next
    ws.Unprotect Password:=PW_WORKBOOK
    If Err.Number <> 0 Then DebugLog "RefreshDashboard", "Warning: Failed to unprotect dashboard sheet. Err=" & Err.Number: Err.Clear
    On Error GoTo ErrorHandler

    DebugLog "RefreshDashboard", "Step 3: CleanupDashboardLayout..."
    CleanupDashboardLayout ws
    DebugLog "RefreshDashboard", "Step 3: InitializeDashboardLayout..."
    InitializeDashboardLayout ws

    '--- STEP 4: Build Final Data Array In Memory ---
    DebugLog "RefreshDashboard", "Step 4: BuildDashboardDataArray..."
    t_Build = Timer
    arrFinalData = BuildDashboardDataArray()
    DebugLog "RefreshDashboard", "Step 4: Finished BuildDashboardDataArray. Time: " & Format(Timer - t_Build, "0.00") & "s"

    If Not IsArray(arrFinalData) Then
         DebugLog "RefreshDashboard", "CRITICAL ERROR - BuildDashboardDataArray failed to return valid data. Aborting."
         MsgBox "Critical error building dashboard data array. Please check logs or source data.", vbCritical, "Refresh Aborted"
         GoTo Cleanup
    End If
    arrIsEmpty = False
    On Error Resume Next
    arrIsEmpty = (LBound(arrFinalData, 1) > UBound(arrFinalData, 1))
    On Error GoTo ErrorHandler
    DebugLog "RefreshDashboard", "Step 4: Data Array IsArray=" & IsArray(arrFinalData) & ", IsEmpty=" & arrIsEmpty
    If arrIsEmpty Then
         DebugLog "RefreshDashboard", "BuildDashboardDataArray returned empty data array. Dashboard will be empty."
    End If

    '--- STEP 5: Write Data Array to Dashboard ---
    DebugLog "RefreshDashboard", "Step 5: Writing data array to dashboard sheet..."
    t_Write = Timer
    ws.Range("A4:" & DB_COL_COMMENTS & ws.rows.Count).ClearContents
    If Not arrIsEmpty Then
        DebugLog "RefreshDashboard", "Step 5: Resizing target range A4 to " & UBound(arrFinalData, 1) & " rows, " & UBound(arrFinalData, 2) & " cols."
        ws.Range("A4").Resize(UBound(arrFinalData, 1), UBound(arrFinalData, 2)).value = arrFinalData
        DebugLog "RefreshDashboard", "Step 5: Finished writing " & UBound(arrFinalData, 1) & " rows to dashboard."
    Else
         DebugLog "RefreshDashboard", "Step 5: Skipping write to dashboard (data array was empty)."
    End If
    DebugLog "RefreshDashboard", "Step 5: Finished writing data. Time: " & Format(Timer - t_Write, "0.00") & "s"

    ' --- STEP 6: Calculate Last Row and Sort Dashboard Data ---
    lastRow = 0
    On Error Resume Next
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    On Error GoTo ErrorHandler
    DebugLog "RefreshDashboard", "Step 6: Calculated lastRow = " & lastRow & ". Dashboard now has " & IIf(lastRow < 4, 0, lastRow - 3) & " data rows."

    If lastRow >= 5 Then
        DebugLog "RefreshDashboard", "Step 6: Sorting dashboard rows 4:" & lastRow & "..."
        t_Sort = Timer
        SortDashboardData ws, lastRow
        DebugLog "RefreshDashboard", "Step 6: Finished sorting. Time: " & Format(Timer - t_Sort, "0.00") & "s"
    Else
        DebugLog "RefreshDashboard", "Step 6: Skipping sort (less than 2 data rows)."
    End If
    On Error Resume Next
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row ' Recalculate after sort
    On Error GoTo ErrorHandler
    DebugLog "RefreshDashboard", "Step 6: lastRow after sort = " & lastRow

    '--- STEP 7: Apply Final Column Widths, Row Height, AND Number Formats ---
    DebugLog "RefreshDashboard", "Step 7: Applying final column widths/row height/number formats..."
    t_Format = Timer
    Application.ScreenUpdating = False
    On Error Resume Next ' Local error handling for formatting section

    ' Column Widths
    ws.Columns("A:" & DB_COL_LASTCONTACT).AutoFit ' AutoFit A through M
    ' ws.Columns("D").AutoFit ' Let specific width apply below
    ' ws.Columns("E").AutoFit ' Let specific width apply below
    ' ws.Columns("F").AutoFit ' Let specific width apply below
    Const MIN_BUTTON_WIDTH As Double = 20 ' Min width for button columns C, D
    Const MAX_BUTTON_WIDTH As Double = 30 ' Max width for button columns C, D
    If ws.Columns("C").ColumnWidth < MIN_BUTTON_WIDTH Then ws.Columns("C").ColumnWidth = MIN_BUTTON_WIDTH
    If ws.Columns("C").ColumnWidth > MAX_BUTTON_WIDTH Then ws.Columns("C").ColumnWidth = MAX_BUTTON_WIDTH
    If ws.Columns("D").ColumnWidth < MIN_BUTTON_WIDTH Then ws.Columns("D").ColumnWidth = MIN_BUTTON_WIDTH
    If ws.Columns("D").ColumnWidth > MAX_BUTTON_WIDTH Then ws.Columns("D").ColumnWidth = MAX_BUTTON_WIDTH
    ws.Columns("E").ColumnWidth = 10.5  ' Doc Date
    ws.Columns("F").ColumnWidth = 10.5  ' First Date Pulled
    ws.Columns(DB_COL_COMMENTS).ColumnWidth = 45 ' Fixed width for Comments (N)

    ' Header Row 3 Format
    With ws.rows(3)
        .RowHeight = 15
        .WrapText = False
    End With
    ws.Range("A3:" & DB_COL_COMMENTS & "3").ShrinkToFit = False

    ' *** ADDED/CONFIRMED: Apply Number Formats to Dashboard ***
    DebugLog "RefreshDashboard", "Step 7a: Applying specific number formats to dashboard..."
    If lastRow >= 4 Then ' Only format if data rows exist
        ws.Range(DB_COL_AMOUNT & "4:" & DB_COL_AMOUNT & lastRow).NumberFormat = "$#,##0.00"      ' Amount (D)
        ws.Range(DB_COL_DOC_DATE & "4:" & DB_COL_DOC_DATE & lastRow).NumberFormat = "mm/dd/yyyy"   ' Document Date (E)
        ws.Range(DB_COL_FIRST_PULL & "4:" & DB_COL_FIRST_PULL & lastRow).NumberFormat = "mm/dd/yyyy" ' First Date Pulled (F)
        ws.Range(DB_COL_PULL_COUNT & "4:" & DB_COL_PULL_COUNT & lastRow).NumberFormat = "0"         ' Pull Count (I)
        ws.Range(DB_COL_LASTCONTACT & "4:" & DB_COL_LASTCONTACT & lastRow).NumberFormat = "mm/dd/yyyy" ' Last Contact Date (M)
    End If
    ' *** END NUMBER FORMATS ***

    If Err.Number <> 0 Then
        DebugLog "RefreshDashboard", "WARNING: Error during column/row/number formatting adjustments: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler ' Restore main error handler

    ' *** Re-apply Validation AFTER data write & format ***
    DebugLog "RefreshDashboard", "Step 7b: Re-applying phase validation..."
    ' *** FIXED: Call the sub from modUtilities where it now resides ***
    Call modUtilities.ApplyPhaseValidationToListColumn(ws, DB_COL_PHASE, 4) ' Apply to Dashboard Col L, starting Row 4
    DebugLog "RefreshDashboard", "Step 7b: Finished re-applying validation."

    Application.ScreenUpdating = True ' Turn back on before CF/Protect which might be slow
    DebugLog "RefreshDashboard", "Step 7: Finished final column/row/number formatting and validation. Time: " & Format(Timer - t_Format, "0.00") & "s"

    '--- STEP 8: Apply Conditional Formatting, Protection and Freeze Panes ---
    DebugLog "RefreshDashboard", "Step 8: Applying conditional formatting and protection..."
    t_Protect = Timer
    If lastRow >= 4 Then
        DebugLog "RefreshDashboard", "Step 8: Calling ApplyColorFormatting..."
        ApplyColorFormatting ws, 4
        DebugLog "RefreshDashboard", "Step 8: Calling ApplyWorkflowLocationFormatting..."
        ApplyWorkflowLocationFormatting ws, 4
    Else
         DebugLog "RefreshDashboard", "Step 8: Skipping CF (lastRow < 4)."
    End If
    DebugLog "RefreshDashboard", "Step 8: Calling ProtectUserColumns..."
    ProtectUserColumns ws ' Unlocks L:N, Protects sheet
    DebugLog "RefreshDashboard", "Step 8: Calling FreezeDashboard..."
    FreezeDashboard ws
    DebugLog "RefreshDashboard", "Step 8: Finished formatting/protection. Time: " & Format(Timer - t_Protect, "0.00") & "s"

    '--- Call Archival ---
    DebugLog "RefreshDashboard", "Calling modArchival.RefreshAllViews..."
    t_Archival = Timer
    modArchival.RefreshAllViews ' Creates/Refreshes Active/Archive views
    DebugLog "RefreshDashboard", "Returned from modArchival.RefreshAllViews. Time: " & Format(Timer - t_Archival, "0.00") & "s"

    '--- STEP 9 (REVISED): Apply Consistent UI Elements to Main Dashboard ---
    DebugLog "RefreshDashboard", "Step 9 (REVISED): Applying consistent UI elements (Row Heights & Nav Buttons) to main dashboard..."
    On Error Resume Next ' Local handling for UI elements

    ' Apply Row Heights to match Active/Archive views
    ws.rows(1).RowHeight = 32 ' Title Banner
    ws.rows(2).RowHeight = 28 ' Control Panel Row
    If lastRow >= 4 Then ws.rows("4:" & lastRow).RowHeight = 18 ' Data Rows

    ' Add the same Navigation Buttons as Active/Archive views
    ' Ensure modArchival.AddNavigationButtons is Public or move it
    modArchival.AddNavigationButtons ws ' Call the button routine from modArchival

    If Err.Number <> 0 Then
        DebugLog "RefreshDashboard", "WARNING: Error applying consistent UI elements to main dashboard. Err=" & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler ' Restore main error handler for RefreshDashboard
    DebugLog "RefreshDashboard", "Step 9 (REVISED): Finished applying consistent UI elements."
    ' --- End Revised Step 9 ---

    '--- STEP 10: Create/Update Text-Only Copy ---
    DebugLog "RefreshDashboard", "Step 10: CreateOrUpdateTextOnlySheet..."
    t_TextOnly = Timer
    CreateOrUpdateTextOnlySheet ws
    DebugLog "RefreshDashboard", "Step 10: Finished Text-Only copy. Time: " & Format(Timer - t_TextOnly, "0.00") & "s"

    '--- STEP 11: Completion Message & Cleanup ---
    Dim msgText As String
    If PreserveUserEdits Then
        msgText = DASHBOARD_SHEET_NAME & " refreshed successfully!" & vbCrLf & vbCrLf & _
                  "User edits from the '" & USEREDITS_SHEET_NAME & "' sheet were applied." & vbCrLf & _
                  "(Dashboard edits were NOT saved back during this refresh.)"
    Else
        msgText = DASHBOARD_SHEET_NAME & " refreshed successfully!" & vbCrLf & vbCrLf & _
                  "Edits made on the dashboard were saved to '" & USEREDITS_SHEET_NAME & "'." & vbCrLf & _
                  "All data (including User Edits) was merged and displayed."
    End If
    Application.DisplayAlerts = True
    MsgBox msgText, vbInformation, "Dashboard Refresh Complete"
    Application.DisplayAlerts = False

    DebugLog "RefreshDashboard", "Dashboard refresh process completed successfully."

    If backupCreated Then
         DebugLog "RefreshDashboard", "Calling CleanupOldBackups..."
         CleanupOldBackups
    End If

Cleanup:
    DebugLog "RefreshDashboard", "Cleanup Label Reached..."
    On Error Resume Next

    If wbWasLocked Then
        DebugLog "RefreshDashboard", "Cleanup: Re-locking workbook structure..."
        Call ToggleWorkbookStructure(True)
    End If

    DebugLog "RefreshDashboard", "Cleanup: Restoring Application Settings..."
    Application.ScreenUpdating = screenState
    Application.Calculation = calcState
    Application.DisplayAlerts = True
    Application.EnableEvents = eventsState
    DebugLog "RefreshDashboard", "Cleanup: Application Settings Restored (Events=" & eventsState & ")."

    Set ws = Nothing
    Set wsEdits = Nothing
    If IsArray(arrFinalData) Then Erase arrFinalData

    If Not currentSheet Is Nothing Then
        If ActiveSheet.Name <> currentSheet.Name Then currentSheet.Activate
    End If
    Set currentSheet = Nothing

    DebugLog "RefreshDashboard", "EXIT (Cleanup Complete). Total time: " & Format(Timer - t_Start, "0.00") & "s"
    Exit Sub

ErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    errLine = Erl

    DebugLog "RefreshDashboard", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    DebugLog "RefreshDashboard", "ERROR Handler! Err=" & errNum & ": " & errDesc & " near line " & errLine
    DebugLog "RefreshDashboard", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"

    If backupCreated Then
        DebugLog "RefreshDashboard", "Attempting to restore UserEdits from pre-refresh backup..."
        If RestoreUserEditsFromBackup() Then
            DebugLog "RefreshDashboard", "UserEdits restore from backup SUCCEEDED."
            MsgBox "An error occurred during the refresh." & vbCrLf & vbCrLf & _
                   "Error: " & errDesc & vbCrLf & "(Error Code: " & errNum & ")" & vbCrLf & vbCrLf & _
                   "Your UserEdits sheet has been restored from the backup created before this refresh.", vbCritical, "Dashboard Refresh Error"
        Else
            DebugLog "RefreshDashboard", "UserEdits restore from backup FAILED."
             MsgBox "An error occurred during the refresh." & vbCrLf & vbCrLf & _
                    "Error: " & errDesc & vbCrLf & "(Error Code: " & errNum & ")" & vbCrLf & vbCrLf & _
                    "ATTEMPT TO RESTORE USEREDITS FROM BACKUP FAILED. Please check manually for backup sheets ('" & USEREDITS_SHEET_NAME & "_Backup...').", vbCritical, "Dashboard Refresh Error"
        End If
    Else
         DebugLog "RefreshDashboard", "No pre-refresh backup was successfully created."
         MsgBox "An error occurred during the refresh." & vbCrLf & vbCrLf & _
                "Error: " & errDesc & vbCrLf & "(Error Code: " & errNum & ")" & vbCrLf & vbCrLf & _
                "No pre-refresh backup was successfully created.", vbCritical, "Dashboard Refresh Error"
    End If

    Resume Cleanup

End Sub




'================================================================================
'              2. CORE DATA POPULATION & RESTORATION SUB-ROUTINES
'================================================================================


'================================================================================
'              3. LAYOUT, SETUP, and UI ROUTINES
'              *** Using User's Preferred Versions ***
'================================================================================

'------------------------------------------------------------------------------
' GetOrCreateDashboardSheet - Returns the main dashboard Worksheet object
' *** Uses user's preferred version which calls SetupDashboard ***
'------------------------------------------------------------------------------
Private Function GetOrCreateDashboardSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        LogUserEditsOperation "Dashboard sheet '" & sheetName & "' not found. Creating new sheet." ' Added log
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
        ' Call user's SetupDashboard to establish initial layout for new sheet
        SetupDashboard ws ' Call the user's specific SetupDashboard for layout
        LogUserEditsOperation "Called SetupDashboard for new sheet." ' Added log
    End If

    Set GetOrCreateDashboardSheet = ws
End Function


'------------------------------------------------------------------------------
' CleanupDashboardLayout - Clears old data rows (4+), preserves rows 1-3
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Private Sub CleanupDashboardLayout(ws As Worksheet)
    Application.ScreenUpdating = False
    LogUserEditsOperation "CleanupDashboardLayout: Clearing data rows 4+ on sheet '" & ws.Name & "'." ' Added Log

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    Dim dataRange As Range
    Dim tempData As Variant

    ' This section seems intended to preserve data during cleanup,
    ' but the actual .Clear command later removes it anyway.
    ' Keeping the logic as provided by user, but noting it might be redundant.
    ' If lastRow >= 4 Then
    '     Set dataRange = ws.Range("A4:" & DB_COL_COMMENTS & lastRow)
    '     tempData = dataRange.Value
    ' End If

    ' This check seems unnecessary if InitializeDashboardLayout sets headers anyway
    ' Dim hasTitle As Boolean
    ' hasTitle = False
    ' Dim cell As Range
    ' For Each cell In ws.Range("A1:" & DB_COL_COMMENTS & "1").Cells
    '     If InStr(1, CStr(cell.Value), "STRATEGIC QUOTE RECOVERY", vbTextCompare) > 0 Then
    '         hasTitle = True
    '         Exit For
    '     End If
    ' Next cell

    ' --- The Actual Clearing Action ---
    On Error Resume Next ' Ignore error if already clear or protected
    ws.Range("A4:" & DB_COL_COMMENTS & ws.rows.Count).ClearContents ' Clear Data rows only
    If Err.Number <> 0 Then LogUserEditsOperation "CleanupDashboardLayout: Note - Error during clear contents.": Err.Clear
    On Error GoTo 0
    ' --- End Clearing ---

    ' This section recreating rows 1-3 seems redundant if InitializeDashboardLayout does it.
    ' Commenting out to avoid potential conflicts, relying on InitializeDashboardLayout.
    ' If Not hasTitle Then
    '      With ws.Range("A1:" & DB_COL_COMMENTS & "1")
    '          ... (Title formatting) ...
    '      End With
    ' End If
    ' With ws.Range("A2:" & DB_COL_COMMENTS & "2")
    '      ... (Row 2 formatting) ...
    ' End With
    ' With ws.Range("A2")
    '      ... (Control panel label formatting) ...
    ' End With
    ' With ws.Range(DB_COL_COMMENTS & "2")
    '      ... (? formatting) ...
    ' End With
    ' With ws.Range("A3:" & DB_COL_COMMENTS & "3")
    '      ... (Header formatting) ...
    ' End With

    ' This restore logic also seems misplaced here if RefreshDashboard handles restore
    ' If Not IsEmpty(tempData) Then
    '      ... (Restore tempData) ...
    ' End If

    Application.ScreenUpdating = True
End Sub

'------------------------------------------------------------------------------
' InitializeDashboardLayout - Sets headers A-N in Row 3, disables Col N wrap
' *** Corrected Header Order: J=Workflow Location, K=Missing Quote Alert ***
'------------------------------------------------------------------------------
Private Sub InitializeDashboardLayout(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    LogUserEditsOperation "InitializeDashboardLayout: Setting headers A-N, disabling Col N wrap."

    ' --- Clear Data Rows & Extra Columns ---
    ws.Range("A4:" & DB_COL_COMMENTS & ws.rows.Count).ClearContents
    On Error Resume Next ' Ignore error deleting columns
    ' ws.Range("O:" & ws.Columns.Count).Delete Shift:=xlToLeft
    On Error GoTo 0 ' Restore default error handling for this sub

    ' --- Set Headers in Row 3 ---
    With ws.Range("A3:" & DB_COL_COMMENTS & "3") ' Range A3:N3
        .ClearContents
            ' *** CORRECTED Header Order: J=Workflow Location, K=Missing Quote Alert ***
            ' *** Cleaned Syntax - No Comments Inside Array ***
            .value = Array( _
                "Document Number", "Client ID", "Customer Name", "Document Amount", "Document Date", _
                "First Date Pulled", "Salesperson ID", "Entered By", "Pull Count", _
                "Workflow Location", _
                "Missing Quote Alert", _
                "Engagement Phase", "Last Contact Date", "User Comments")
            ' *** End Corrected Array ***

        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False ' Headers should NOT wrap
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlHairline
        .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
    End With

    ' --- Set Row 3 Height (Done later in RefreshDashboard Step 9 after AutoFit) ---
    ' On Error Resume Next
    ' ws.rows(3).AutoFit ' Initial AutoFit can cause issues before final widths set
    ' On Error GoTo 0

    ' --- Disable Text Wrapping for the data area of Column N ---
    On Error Resume Next ' Ignore error if sheet is protected
    Dim wrapRange As Range
    Set wrapRange = ws.Range(DB_COL_COMMENTS & "4:" & DB_COL_COMMENTS & ws.rows.Count) ' Select Col N data area
    wrapRange.WrapText = False ' Ensure comments overflow/truncate, not wrap
    Set wrapRange = Nothing
    If Err.Number <> 0 Then LogUserEditsOperation "InitializeDashboardLayout: Warning - could not set WrapText for Column N data area.": Err.Clear
    On Error GoTo 0

    LogUserEditsOperation "InitializeDashboardLayout: Headers set, Col N data wrap DISABLED."
End Sub

'------------------------------------------------------------------------------
' SetupDashboard - Sets up static Rows 1 (Title) and 2 (Control Panel)
' *** This appears to be the user's preferred setup for Rows 1 & 2 ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' SetupDashboard - Sets up static Rows 1 (Title) and 2 (Control Panel)
' REVISED: Removed Row 2 merges, Added optional column widths
'------------------------------------------------------------------------------
Public Sub SetupDashboard(ws As Worksheet)
     LogUserEditsOperation "SetupDashboard: Setting up Title (Row 1) and Control Panel (Row 2)."
     Application.ScreenUpdating = False
     On Error Resume Next ' Ignore errors if sheet is protected

     ' --- Row 1: Title Bar ---
     With ws.Range("A1:" & DB_COL_COMMENTS & "1") ' A1:N1
         If .MergeCells Then .UnMerge ' Ensure unmerged before merging again
         .ClearContents
         .Merge
         .value = "STRATEGIC QUOTE RECOVERY & CONVERSION TRACKER"
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .Font.Size = 18
         .Font.Bold = True
         .Interior.Color = RGB(16, 107, 193) ' Blue background
         .Font.Color = RGB(255, 255, 255) ' White text
         .RowHeight = 32 ' Use constant later? ROW_HGT_BANNER
     End With

     ' --- Row 2: Control Panel Area ---
     With ws.Range("A2:" & DB_COL_COMMENTS & "2") ' A2:N2
         .ClearContents
         '--- REMOVED MERGE LINES for C2:D2, E2:F2 ---
         .UnMerge ' Ensure entire row is unmerged initially
         .Interior.Color = RGB(245, 245, 245) ' Light grey background
         .Borders(xlEdgeTop).LineStyle = xlContinuous
         .Borders(xlEdgeTop).Weight = xlThin
         .Borders(xlEdgeTop).Color = RGB(200, 200, 200)
         .Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Borders(xlEdgeBottom).Weight = xlThin
         .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
         .RowHeight = 28 ' Use constant later? ROW_HGT_CONTROLS
         .VerticalAlignment = xlCenter
     End With

     ' --- Row 2: "CONTROL PANEL" Label (A2) ---
     With ws.Range("A2")
         .value = "CONTROL PANEL"
         .Font.Bold = True
         .Font.Size = 10
         .Font.Name = "Segoe UI"
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .Interior.Color = RGB(70, 130, 180) ' Steel blue
         .Font.Color = RGB(255, 255, 255)
         ' .ColumnWidth = 16 ' Set width below if using optional block
         .Borders(xlEdgeRight).LineStyle = xlContinuous
         .Borders(xlEdgeRight).Weight = xlThin
         .Borders(xlEdgeRight).Color = RGB(200, 200, 200)
     End With
     ' Clear B2 for spacing
      ws.Range("B2").ClearContents

     ' --- Row 2: Help (?) Icon (N2) ---
      With ws.Range(DB_COL_COMMENTS & "2") ' N2 (Timestamp now goes here via AddNavigationButtons)
          ' Clear any old ? icon - timestamp will overwrite
          .ClearContents
          ' Optional: Keep ? if desired, place timestamp elsewhere (like M2)
          ' .Value = "?"
          ' .Font.Bold = True
          ' ... etc
      End With

     ' --- Optional: Set standard widths for Row 2 elements ---
     DebugLog "SetupDashboard", "Setting standard column widths for Row 2 elements..."
     With ws.Columns
        .Item("C").ColumnWidth = 15   ' Standard Refresh button width approx match
        .Item("D").ColumnWidth = 15   ' Preserve UserEdits button width approx match
        .Item("F").ColumnWidth = 11   ' All Items button width approx match
        .Item("G").ColumnWidth = 11   ' Active button width approx match
        .Item("H").ColumnWidth = 11   ' Archive button width approx match
        .Item("J").ColumnWidth = 14   ' All Count label width
        .Item("K").ColumnWidth = 14   ' Active Count label width
        .Item("L").ColumnWidth = 14   ' Archive Count label width
        .Item("N").ColumnWidth = 20   ' Timestamp label width
     End With
     ' --- End Optional Widths ---

     If Err.Number <> 0 Then LogUserEditsOperation "SetupDashboard: Note - Error setting up rows 1-2.": Err.Clear
     On Error GoTo 0
     Application.ScreenUpdating = True
 End Sub

'------------------------------------------------------------------------------
' ModernButton - Creates styled buttons
' REVISED: Function returning Shape, accepts Range & Width, includes ByVal & visibility fixes
'------------------------------------------------------------------------------
Public Function ModernButton(ws As Worksheet, targetCell As Range, ByVal buttonText As String, ByVal macroName As String, ByVal buttonWidth As Double) As Shape
    Dim btn As Shape
    On Error GoTo ModernButtonErrorHandler

    ' --- Calculate Position and Size ---
    Dim pad As Double: pad = 2 ' Padding inside target cell
    Dim btnLeft As Double: btnLeft = targetCell.Left + pad ' Position relative to target cell
    Dim btnTop As Double: btnTop = targetCell.Top + pad   ' Position relative to target cell
    Const BTN_HEIGHT As Double = 24 ' Fixed height for consistency

    ' --- Add the shape using PASSED IN width ---
    Set btn = Nothing
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, buttonWidth, BTN_HEIGHT) ' Use passed buttonWidth
    If btn Is Nothing Then
        DebugLog "ModernButton", "ERROR: Failed to add shape for '" & buttonText & "'."
        Set ModernButton = Nothing
        Exit Function
    End If

    ' --- Style the button ---
    With btn
        ' Size
         .Width = buttonWidth ' Set width definitively
         .Height = BTN_HEIGHT
         .LockAspectRatio = msoFalse

        ' Force Visible Fill
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(0, 112, 192)   ' Example blue fill
        If .Fill.Visible = msoFalse Then .Fill.Visible = msoTrue ' Double-check

        ' Force No Outline
        .Line.Visible = msoFalse

        ' Rounded Corners
        .Adjustments(1) = 0.25

        ' Text Formatting
        On Error Resume Next
        With .TextFrame2 ' Try modern text frame
            .TextRange.Text = buttonText
            .TextRange.Font.Fill.Visible = msoTrue ' Force text visible
            .TextRange.Font.Fill.Solid
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite ' Example white text
             If .TextRange.Font.Fill.Visible = msoFalse Then .TextRange.Font.Fill.Visible = msoTrue ' Double-check
            .TextRange.Font.Size = 9
            .HorizontalAnchor = msoAnchorCenter
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Font.Name = "Segoe UI"
            .WordWrap = msoFalse
        End With
        If Err.Number <> 0 Then ' Fallback to older text frame
            Err.Clear
            DebugLog "ModernButton", "Note: TextFrame2 failed for '" & buttonText & "', using TextFrame fallback."
            With .TextFrame
                .Characters.Text = buttonText
                .Characters.Font.Color = vbWhite ' Fallback color
                .Characters.Font.Size = 9
                .Characters.Font.Name = "Segoe UI"
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
                .AutoSize = False
                .WrapText = False
            End With
        End If
        Err.Clear
        On Error GoTo ModernButtonErrorHandler ' Restore error handler for rest of With block

        ' Placement & Action & Temp Name
        .Placement = xlMoveAndSize
        .Name = "Temp_" & Replace(buttonText, " ", "_") ' Temporary - AddNavigationButtons will rename
        .OnAction = macroName

        ' Ensure it's on top
         .ZOrder msoBringToFront
    End With

    ' --- Return the created shape ---
    Set ModernButton = btn
    Exit Function ' Normal Exit

ModernButtonErrorHandler:
    DebugLog "ModernButton", "ERROR [" & Err.Number & "] " & Err.Description & " creating button '" & buttonText & "'"
    Set ModernButton = Nothing ' Return Nothing on error
End Function



'------------------------------------------------------------------------------
' FreezeDashboard - Freezes rows 1-3
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Private Sub FreezeDashboard(ws As Worksheet)
    If ws Is Nothing Then Exit Sub ' Added safety check
    LogUserEditsOperation "FreezeDashboard: Freezing panes at row 4." ' Added Log
    ws.Activate ' Ensure sheet is active
    On Error Resume Next ' Ignore error if already frozen/unfrozen
    ActiveWindow.FreezePanes = False ' Unfreeze first
    ws.Range("A4").Select          ' Select cell below freeze row
    ActiveWindow.FreezePanes = True  ' Freeze above selected cell
    ws.Range("A1").Select ' Select A1 after freezing
    If Err.Number <> 0 Then LogUserEditsOperation "FreezeDashboard: Note - Error during freeze panes operation.": Err.Clear
    On Error GoTo 0
End Sub

'================================================================================
'              4. UserEdits SHEET MANAGEMENT (Save, Setup, Backup)
'              *** Using Reviewed/Robust Versions ***
'================================================================================

'------------------------------------------------------------------------------
' SaveUserEditsFromDashboard - Captures L-N from Dashboard -> UserEdits A-F
' Uses dictionary lookup for row numbers and cleaned keys for reliable matching.
' *** CORRECTED to handle dictionary changes from refactoring ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' SaveUserEditsFromDashboard - Captures L-N from Dashboard -> UserEdits A-F
' Uses dictionary lookup for row numbers and cleaned keys for reliable matching.
' *** UPDATED: Prevents saving literal "Legacy Process" for Phase ***
'------------------------------------------------------------------------------
Public Sub SaveUserEditsFromDashboard()
    Dim wsDash As Worksheet, wsEdits As Worksheet
    Dim lastRowDash As Long, lastRowEdits As Long, destRow As Long
    Dim i As Long, r As Long ' Loop counters
    Dim dashDocNum As String, cleanedDashDocNum As String, cleanedDocNum As String
    Dim editRow As Variant      ' Stores existing SHEET row number from rowIndexDict
    Dim hasEdits As Boolean, wasChanged As Boolean
    Dim editsSavedCount As Long, editsUpdatedCount As Long
    Dim rowIndexDict As Object   ' Dictionary mapping CleanDocNum -> Sheet Row Number

    On Error GoTo ErrorHandler_SaveUserEdits ' Use a specific error handler for this sub

    LogUserEditsOperation "SaveUserEditsFromDashboard: Starting process..."

    ' --- Get Worksheet Objects ---
    Set wsDash = GetOrCreateDashboardSheet(DASHBOARD_SHEET_NAME) ' Use helper
    If wsDash Is Nothing Then
        LogUserEditsOperation "SaveUserEditsFromDashboard: ERROR - Dashboard sheet '" & DASHBOARD_SHEET_NAME & "' not found."
        Exit Sub
    End If
    SetupUserEditsSheet ' Ensure UserEdits sheet exists and is structured correctly
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    If wsEdits Is Nothing Then
        LogUserEditsOperation "SaveUserEditsFromDashboard: ERROR - UserEdits sheet '" & USEREDITS_SHEET_NAME & "' could not be accessed."
        Exit Sub
    End If

    ' --- Build Row Index Dictionary (CleanDocNum -> Sheet Row Number) ---
    LogUserEditsOperation "SaveUserEditsFromDashboard: Building Row Index Dictionary..."
    Set rowIndexDict = BuildRowIndexDict(wsEdits) ' Use Helper Function
    LogUserEditsOperation "SaveUserEditsFromDashboard: Built Row Index Dict with " & rowIndexDict.Count & " entries."
    ' --- End Build Row Index Dictionary ---

    ' --- Prepare for Dashboard Loop ---
    lastRowDash = wsDash.Cells(wsDash.rows.Count, "A").End(xlUp).Row
    LogUserEditsOperation "SaveUserEditsFromDashboard: Checking dashboard rows 4 to " & lastRowDash & " for edits..."

    ' --- Loop Through Dashboard Rows ---
    If lastRowDash >= 4 Then
        For i = 4 To lastRowDash
            dashDocNum = CStr(wsDash.Cells(i, "A").value)         ' Get raw DocNum from Dashboard
            cleanedDashDocNum = CleanDocumentNumber(dashDocNum) ' Clean it for lookup
            wasChanged = False ' Reset change tracker for each row

            If cleanedDashDocNum <> "" Then ' Only process if there's a valid DocNum

                ' Check if dashboard row L, M, or N has any data (indicates potential user edit)
                hasEdits = False
                If wsDash.Cells(i, DB_COL_PHASE).value <> "" Or _
                   wsDash.Cells(i, DB_COL_LASTCONTACT).value <> "" Or _
                   wsDash.Cells(i, DB_COL_COMMENTS).value <> "" Then
                    hasEdits = True
                End If

                ' Find if this cleaned DocNum already exists in UserEdits using the Row Index Dict
                editRow = 0 ' Reset flag
                If rowIndexDict.Exists(cleanedDashDocNum) Then
                    editRow = rowIndexDict(cleanedDashDocNum) ' Gets Row Number (Long)
                End If

                ' --- Process If Dashboard Has Edits OR If Entry Exists in UserEdits ---
                ' (We process existing entries even if dashboard L-N are blank now,
                '  to ensure UserEdits reflects the latest state from the dashboard)
                If hasEdits Or editRow > 0 Then

                    ' Determine destination row in UserEdits sheet
                    If editRow > 0 Then
                        destRow = CLng(editRow) ' Update existing row found via rowIndexDict
                    Else
                        ' Add new row to UserEdits sheet
                        lastRowEdits = wsEdits.Cells(wsEdits.rows.Count, UE_COL_DOCNUM).End(xlUp).Row + 1
                        If lastRowEdits < 2 Then lastRowEdits = 2 ' Ensure starting at row 2
                        destRow = lastRowEdits
                        wsEdits.Cells(destRow, UE_COL_DOCNUM).value = cleanedDashDocNum ' Write the CLEANED document number
                        RowIndexDictAdd rowIndexDict, cleanedDashDocNum, destRow ' Keep rowIndexDict in sync
                        wasChanged = True ' New row always counts as changed
                        editsSavedCount = editsSavedCount + 1
                    End If

                    ' Get current values from Dashboard L, M, N
                    Dim dbPhaseVal As Variant ' Use Variant to handle potential blanks/errors
                    Dim dbLastContactVal As Variant
                    Dim dbCommentsVal As Variant
                    dbPhaseVal = wsDash.Cells(i, DB_COL_PHASE).value
                    dbLastContactVal = wsDash.Cells(i, DB_COL_LASTCONTACT).value
                    dbCommentsVal = wsDash.Cells(i, DB_COL_COMMENTS).value

                    ' --- Prevent saving "Legacy Process" placeholder --- <<< UPDATE IS HERE
                    If LCase$(Trim$(CStr(dbPhaseVal))) = "legacy process" Then
                        dbPhaseVal = vbNullString ' Treat as blank before comparing/saving
                    End If
                    ' --- End Update ---

                    ' Compare with UserEdits values on sheet and update if different
                    ' Use CStr for safe comparison of variants/values
                    If editRow = 0 Or CStr(wsEdits.Cells(destRow, UE_COL_PHASE).value) <> CStr(dbPhaseVal) Then
                        wsEdits.Cells(destRow, UE_COL_PHASE).value = dbPhaseVal ' Write possibly modified dbPhaseVal to UserEdits B
                        If editRow > 0 Then wasChanged = True
                    End If
                    If editRow = 0 Or CStr(wsEdits.Cells(destRow, UE_COL_LASTCONTACT).value) <> CStr(dbLastContactVal) Then
                         wsEdits.Cells(destRow, UE_COL_LASTCONTACT).value = dbLastContactVal ' Write to UserEdits C
                         If editRow > 0 Then wasChanged = True
                    End If
                    If editRow = 0 Or CStr(wsEdits.Cells(destRow, UE_COL_COMMENTS).value) <> CStr(dbCommentsVal) Then
                         wsEdits.Cells(destRow, UE_COL_COMMENTS).value = dbCommentsVal ' Write to UserEdits D
                         If editRow > 0 Then wasChanged = True
                    End If

                    ' If any change was made, update Source/Timestamp
                    If wasChanged Then
                        wsEdits.Cells(destRow, UE_COL_SOURCE).value = Module_Identity.GetWorkbookIdentity() ' Write to UserEdits E
                        wsEdits.Cells(destRow, UE_COL_TIMESTAMP).value = Format$(Now(), "yyyy-mm-dd hh:nn:ss") ' Write to UserEdits F
                        If editRow > 0 Then editsUpdatedCount = editsUpdatedCount + 1
                    End If
                End If ' End If hasEdits Or editRow > 0
            End If ' End If cleanedDashDocNum <> ""
        Next i ' Next dashboard row
    End If ' End If lastRowDash >= 4

    LogUserEditsOperation "SaveUserEditsFromDashboard: Finished. New Edits Saved: " & editsSavedCount & ". Existing Edits Updated: " & editsUpdatedCount & "."

    ' --- Cleanup for this specific Sub ---
    Set rowIndexDict = Nothing
    Set wsDash = Nothing
    Set wsEdits = Nothing
    Exit Sub ' Normal Exit

ErrorHandler_SaveUserEdits:
    LogUserEditsOperation "ERROR in SaveUserEditsFromDashboard: [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")"
    MsgBox "An error occurred while saving user edits: " & vbCrLf & Err.Description, vbCritical, "Save Edits Error"
    ' Ensure objects are released even on error
    Set rowIndexDict = Nothing
    Set wsDash = Nothing
    Set wsEdits = Nothing
    ' Note: We don't re-enable events here; RefreshDashboard's main handler does that.
End Sub




'------------------------------------------------------------------------------
' SetupUserEditsSheet - Creates or Verifies the UserEdits sheet structure (A-F)
' Includes safety checks and backup before modifying structure.
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Sub SetupUserEditsSheet()
    Dim wsEdits As Worksheet
    Dim needsBackup As Boolean
    Dim wsBackup As Worksheet
    Dim currentHeaders As Variant
    Dim expectedHeaders As Variant
    Dim structureCorrect As Boolean
    Dim backupSuccess As Boolean
    Dim i As Long
    Dim emergencyFlag As Boolean

    expectedHeaders = Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                            "User Comments", "ChangeSource", "Timestamp") ' A-F

    LogUserEditsOperation "SetupUserEditsSheet: Checking structure..."

    ' --- Check if Sheet Exists ---
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    If Err.Number <> 0 Then Err.Clear ' Ignore error if sheet doesn't exist
    On Error GoTo ErrorHandler_SetupUserEdits ' Use specific handler for this sub

    If wsEdits Is Nothing Then
        ' --- Create New Sheet ---
        LogUserEditsOperation "SetupUserEditsSheet: Sheet doesn't exist. Creating new."
        Set wsEdits = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsEdits.Name = USEREDITS_SHEET_NAME
        With wsEdits.Range(UE_COL_DOCNUM & "1:" & UE_COL_TIMESTAMP & "1") ' A1:F1
            .value = expectedHeaders
            .Font.Bold = True
            .Interior.Color = RGB(16, 107, 193)
            .Font.Color = RGB(255, 255, 255)
        End With
        wsEdits.Visible = xlSheetHidden
        LogUserEditsOperation "SetupUserEditsSheet: Created new sheet with standard headers."
        Exit Sub ' Done
    End If

    ' --- Sheet Exists: Verify Header Structure (A-F) ---
    structureCorrect = False
    On Error Resume Next ' Handle error reading headers (e.g., protected sheet)
    currentHeaders = wsEdits.Range(UE_COL_DOCNUM & "1:" & UE_COL_TIMESTAMP & "1").value ' A1:F1
    If Err.Number <> 0 Then
        LogUserEditsOperation "SetupUserEditsSheet: WARNING - Error reading headers from existing sheet: " & Err.Description
        needsBackup = True ' Assume structure is bad if headers can't be read
        Err.Clear
    Else
        If IsArray(currentHeaders) Then
             If UBound(currentHeaders, 2) = UBound(expectedHeaders) + 1 Then ' Check if it has 6 columns
                 structureCorrect = True ' Assume correct until mismatch found
                 For i = LBound(expectedHeaders) To UBound(expectedHeaders)
                     If CStr(currentHeaders(1, i + 1)) <> expectedHeaders(i) Then
                         LogUserEditsOperation "SetupUserEditsSheet: Header mismatch found at column " & i + 1 & ". Expected: '" & expectedHeaders(i) & "', Found: '" & CStr(currentHeaders(1, i + 1)) & "'."
                         structureCorrect = False
                         Exit For
                     End If
                 Next i
             Else
                 LogUserEditsOperation "SetupUserEditsSheet: Incorrect number of header columns found (" & UBound(currentHeaders, 2) & "). Expected 6."
                 structureCorrect = False ' Mark as incorrect if column count wrong
             End If
        Else ' currentHeaders is not an array (e.g., single cell A1 is empty or sheet is weird)
             LogUserEditsOperation "SetupUserEditsSheet: Could not read headers as an array. Assuming structure is incorrect."
             structureCorrect = False
        End If
        needsBackup = Not structureCorrect
    End If
    On Error GoTo ErrorHandler_SetupUserEdits

    If structureCorrect Then
        LogUserEditsOperation "SetupUserEditsSheet: Structure verified as correct."
        Exit Sub ' Structure is fine, nothing more to do
    End If

    ' --- Structure Incorrect: Backup and Rebuild ---
    LogUserEditsOperation "SetupUserEditsSheet: Structure incorrect or unreadable. Attempting backup and rebuild."
    backupSuccess = CreateUserEditsBackup("StructureUpdate_" & Format(Now, "yyyymmdd_hhmmss"))

    If Not backupSuccess Then
        LogUserEditsOperation "SetupUserEditsSheet: CRITICAL - Backup failed. Aborting structure update to prevent data loss."
        MsgBox "WARNING: Could not create a backup of the '" & USEREDITS_SHEET_NAME & "' sheet." & vbCrLf & vbCrLf & _
               "The sheet structure needs updating, but no changes will be made to avoid potential data loss." & vbCrLf & _
               "Please check sheet protection or other issues preventing backup.", vbCritical, "UserEdits Structure Update Failed"
        Exit Sub ' Abort if backup failed
    End If

    LogUserEditsOperation "SetupUserEditsSheet: Backup successful. Clearing and rebuilding sheet '" & wsEdits.Name & "'."

    ' --- Clear and Recreate Headers ---
    On Error Resume Next ' Handle error during clear (e.g., protection)
    wsEdits.Cells.Clear
    If Err.Number <> 0 Then
        LogUserEditsOperation "SetupUserEditsSheet: ERROR clearing sheet after backup. Manual intervention may be required. Error: " & Err.Description
        MsgBox "ERROR: Could not clear the '" & USEREDITS_SHEET_NAME & "' sheet after creating a backup." & vbCrLf & _
               "Please check sheet protection. Manual cleanup might be needed.", vbExclamation, "Clear Error"
        Exit Sub ' Stop if clear fails
    End If
    On Error GoTo ErrorHandler_SetupUserEdits ' Restore handler

    With wsEdits.Range(UE_COL_DOCNUM & "1:" & UE_COL_TIMESTAMP & "1") ' A1:F1
        .value = expectedHeaders
        .Font.Bold = True
        .Interior.Color = RGB(16, 107, 193)
        .Font.Color = RGB(255, 255, 255)
    End With

    LogUserEditsOperation "SetupUserEditsSheet: Rebuilt headers. Attempting to restore data from most recent backup..."

    ' --- Attempt to Restore from Backup ---
    If RestoreUserEditsFromBackup() Then
        LogUserEditsOperation "SetupUserEditsSheet: Data restored from backup after structure update."
    Else
        LogUserEditsOperation "SetupUserEditsSheet: WARNING - Could not automatically restore data from backup after structure update. Manual restore may be needed from backup sheet."
    End If

    LogUserEditsOperation "SetupUserEditsSheet: Structure update process complete."
    Exit Sub

ErrorHandler_SetupUserEdits:
    LogUserEditsOperation "SetupUserEditsSheet: CRITICAL ERROR [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")"
    MsgBox "An critical error occurred during UserEdits Sheet Setup: " & Err.Description & vbCrLf & _
           "Please check the '" & USEREDITSLOG_SHEET_NAME & "' sheet for details.", vbCritical, "UserEdits Setup Error"
End Sub

'------------------------------------------------------------------------------
' CreateUserEditsBackup - Creates a timestamped backup copy of UserEdits
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Function CreateUserEditsBackup(Optional backupSuffix As String = "") As Boolean
    Dim wsEdits As Worksheet, wsBackup As Worksheet
    Dim backupName As String
    Dim backupTimestamp As String: backupTimestamp = Format(Now, "yyyymmdd_hhmmss")

    On Error GoTo BackupErrorHandler

    ' --- Determine Backup Name ---
    If backupSuffix = "" Then
        backupName = USEREDITS_SHEET_NAME & "_Backup_" & backupTimestamp
    Else
        ' Ensure suffix doesn't create overly long name (Excel limit is 31 chars)
        If Len(USEREDITS_SHEET_NAME & "_Backup_" & backupSuffix) > 31 Then
            backupName = Left$(USEREDITS_SHEET_NAME & "_Backup_" & backupSuffix, 31)
        Else
            backupName = USEREDITS_SHEET_NAME & "_Backup_" & backupSuffix
        End If
    End If

    ' --- Check Source Sheet ---
    On Error Resume Next
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    If Err.Number <> 0 Then
        LogUserEditsOperation "CreateUserEditsBackup: Source sheet '" & USEREDITS_SHEET_NAME & "' not found. Cannot create backup."
        CreateUserEditsBackup = False
        Exit Function
    End If
    On Error GoTo BackupErrorHandler ' Restore handler

    ' --- Handle Existing/New Backup Sheet ---
    Application.DisplayAlerts = False ' Suppress sheet delete prompt if overwriting
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets(backupName)
    If Err.Number = 0 Then ' Backup with this exact name already exists
         wsBackup.Delete ' Delete existing sheet cleanly
         If Err.Number <> 0 Then GoTo BackupFailed ' Error deleting existing sheet
         LogUserEditsOperation "CreateUserEditsBackup: Deleted existing backup sheet named '" & backupName & "'."
    End If
    Err.Clear ' Clear potential error from checking sheet existence

    ' Create new backup sheet
    Set wsBackup = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    If wsBackup Is Nothing Then GoTo BackupFailed ' Could not add sheet
    wsBackup.Name = backupName
    If Err.Number <> 0 Then ' Failed to name the sheet (e.g., invalid chars, length)
         LogUserEditsOperation "CreateUserEditsBackup: ERROR - Failed to name backup sheet '" & backupName & "'. Using default name '" & wsBackup.Name & "'."
         backupName = wsBackup.Name ' Use the default name assigned by Excel
         Err.Clear
    End If
    Application.DisplayAlerts = True ' Restore alerts
    On Error GoTo BackupErrorHandler ' Restore handler

    ' --- Copy Data ---
    wsEdits.UsedRange.Copy wsBackup.Range("A1")
    If Err.Number <> 0 Then GoTo BackupFailed ' Error during copy

    ' --- Hide Backup ---
    wsBackup.Visible = xlSheetHidden
    LogUserEditsOperation "CreateUserEditsBackup: Successfully created backup: '" & backupName & "'."
    CreateUserEditsBackup = True
    Set wsEdits = Nothing: Set wsBackup = Nothing
    Exit Function

BackupFailed:
    LogUserEditsOperation "CreateUserEditsBackup: ERROR creating backup '" & backupName & "'. Error: " & Err.Description
    ' Attempt to delete partially created backup sheet if it exists
    If Not wsBackup Is Nothing Then
        Application.DisplayAlerts = False
        On Error Resume Next ' Ignore error if sheet cannot be deleted
        wsBackup.Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
    End If
BackupErrorHandler:
    If Err.Number <> 0 And Err.Number <> 9 Then ' Log error if not already logged by BackupFailed (ignore Subscript out of range if sheet check fails)
         LogUserEditsOperation "CreateUserEditsBackup: ERROR [" & Err.Number & "] " & Err.Description
    End If
    Application.DisplayAlerts = True ' Ensure alerts are back on
    CreateUserEditsBackup = False
    Set wsEdits = Nothing: Set wsBackup = Nothing
End Function


'------------------------------------------------------------------------------
' RestoreUserEditsFromBackup - Restores UserEdits from the MOST RECENT backup
' Used primarily by the ErrorHandler in RefreshDashboard.
' *** Using robust version ***
'------------------------------------------------------------------------------
Public Function RestoreUserEditsFromBackup(Optional specificBackupName As String = "") As Boolean
    Dim wsEdits As Worksheet
    Dim wsBackup As Worksheet, candidateSheet As Worksheet
    Dim backupNameToRestore As String
    Dim latestBackupTimestamp As Double: latestBackupTimestamp = 0 ' Use timestamp for more precision
    Dim datePart As String, timePart As String, backupTimestampCurrent As Double

    On Error GoTo RestoreErrorHandler

    ' --- Find the Backup Sheet to Restore ---
    If specificBackupName <> "" Then
        backupNameToRestore = specificBackupName ' Use specified name
    Else
        ' Find the most recent backup based on timestamp in name yyyymmdd_hhmmss
        For Each candidateSheet In ThisWorkbook.Sheets
            If candidateSheet.Name Like USEREDITS_SHEET_NAME & "_Backup_????????_??????*" Then ' Match pattern yyyymmdd_hhmmss
                On Error Resume Next ' Ignore parse errors
                backupTimestampCurrent = 0
                datePart = Mid$(candidateSheet.Name, Len(USEREDITS_SHEET_NAME & "_Backup_") + 1, 8) ' yyyymmdd
                timePart = Mid$(candidateSheet.Name, Len(USEREDITS_SHEET_NAME & "_Backup_") + 10, 6) ' hhmmss
                ' Combine date and time into a sortable numeric value (YYYYMMDDHHMMSS)
                backupTimestampCurrent = Val(datePart & timePart)
                If Err.Number = 0 And backupTimestampCurrent > 0 Then ' Successfully parsed timestamp
                    If backupTimestampCurrent > latestBackupTimestamp Then
                         latestBackupTimestamp = backupTimestampCurrent
                         backupNameToRestore = candidateSheet.Name
                    End If
                End If
                Err.Clear
                On Error GoTo RestoreErrorHandler ' Restore handler
            End If
        Next candidateSheet
    End If

    If backupNameToRestore = "" Then
        LogUserEditsOperation "RestoreUserEditsFromBackup: No suitable backup sheet found to restore from."
        RestoreUserEditsFromBackup = False
        Exit Function
    End If

    ' --- Get Backup and Target Sheets ---
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets(backupNameToRestore)
    If wsBackup Is Nothing Then Err.Raise 9, , "Backup sheet '" & backupNameToRestore & "' not found." ' Raise error if specific sheet missing
    On Error GoTo RestoreErrorHandler

    ' Ensure UserEdits sheet exists (might have been deleted during failed process)
    SetupUserEditsSheet ' This will create it if missing
    Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
    If wsEdits Is Nothing Then Err.Raise 9, , "Target sheet '" & USEREDITS_SHEET_NAME & "' could not be accessed."

    ' --- Perform Restore ---
    LogUserEditsOperation "RestoreUserEditsFromBackup: Restoring data from '" & backupNameToRestore & "' to '" & wsEdits.Name & "'."
    wsEdits.Cells.Clear ' Clear target sheet first
    wsBackup.UsedRange.Copy wsEdits.Range("A1")

    LogUserEditsOperation "RestoreUserEditsFromBackup: Restore complete."
    RestoreUserEditsFromBackup = True
    Set wsEdits = Nothing: Set wsBackup = Nothing: Set candidateSheet = Nothing
    Exit Function

RestoreErrorHandler:
    LogUserEditsOperation "RestoreUserEditsFromBackup: ERROR [" & Err.Number & "] " & Err.Description
    RestoreUserEditsFromBackup = False
    Set wsEdits = Nothing: Set wsBackup = Nothing: Set candidateSheet = Nothing
End Function

'------------------------------------------------------------------------------
' CleanupOldBackups - Deletes backup sheets older than X days
' *** Using robust version ***
'------------------------------------------------------------------------------
Private Sub CleanupOldBackups()
    Const DAYS_TO_KEEP As Long = 7 ' Keep backups for 7 days
    Dim cutoffDate As Date: cutoffDate = Date - DAYS_TO_KEEP
    Dim backupBaseName As String: backupBaseName = USEREDITS_SHEET_NAME & "_Backup_"
    Dim oldSheets As New Collection ' Use Collection to store sheets for deletion
    Dim Sh As Worksheet
    Dim datePart As String, backupDate As Date, deleteCount As Long

    LogUserEditsOperation "CleanupOldBackups: Checking for backups older than " & Format(cutoffDate, "yyyy-mm-dd") & "..."

    On Error Resume Next ' Ignore errors iterating sheets

    For Each Sh In ThisWorkbook.Sheets
        If Sh.Visible = xlSheetHidden And Sh.Name Like backupBaseName & "????????_??????*" Then ' Match yyyymmdd_hhmmss pattern
            datePart = Mid$(Sh.Name, Len(backupBaseName) + 1, 8) ' yyyymmdd
            backupDate = DateSerial(1900, 1, 1) ' Default if parse fails
            Err.Clear
            backupDate = CDate(Format(datePart, "@@@@-@@-@@")) ' Parse only date part

             If Err.Number = 0 Then ' Successfully parsed date
                 If backupDate < cutoffDate Then
                     oldSheets.Add Sh ' Add sheet object to collection
                 End If
             Else
                  Err.Clear
             End If
        End If
    Next Sh

    If oldSheets.Count > 0 Then
        Application.DisplayAlerts = False ' Suppress delete confirmation prompts
        For Each Sh In oldSheets
            On Error Resume Next ' Ignore error deleting single sheet
            Sh.Delete
            If Err.Number = 0 Then deleteCount = deleteCount + 1 Else Err.Clear
        Next Sh
        Application.DisplayAlerts = True
        LogUserEditsOperation "CleanupOldBackups: Deleted " & deleteCount & " old backup sheets (older than " & DAYS_TO_KEEP & " days)."
    Else
         LogUserEditsOperation "CleanupOldBackups: No old backup sheets found for deletion."
    End If

    On Error GoTo 0 ' Restore default error handling
    Set oldSheets = Nothing: Set Sh = Nothing
End Sub

'================================================================================
'              5. FORMATTING & SORTING ROUTINES
'              *** Using User's Preferred Versions ***
'================================================================================

'------------------------------------------------------------------------------
' SortDashboardData - Sorts A3:N<lastRow> by F (asc), then D (desc)
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Private Sub SortDashboardData(ws As Worksheet, lastRow As Long)
    If ws Is Nothing Or lastRow < 5 Then Exit Sub ' Need header + at least 2 data rows

    LogUserEditsOperation "SortDashboardData: Sorting range A3:" & DB_COL_COMMENTS & lastRow & "..." ' Added Log
    Dim sortRange As Range
    Set sortRange = ws.Range("A3:" & DB_COL_COMMENTS & lastRow) ' A3:N<lastRow>

    With ws.Sort
        .SortFields.Clear
        ' Key 1: First Date Pulled (Column F), Ascending
        .SortFields.Add key:=sortRange.Columns(6), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ' Key 2: Document Amount (Column D), Descending
        .SortFields.Add key:=sortRange.Columns(4), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

        .SetRange sortRange
        .Header = xlYes ' Row 3 contains headers
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        On Error Resume Next ' Handle error if sort fails (e.g., protected sheet, merged cells)
        .Apply
        If Err.Number <> 0 Then LogUserEditsOperation "SortDashboardData: ERROR applying sort. Error: " & Err.Description: Err.Clear
        On Error GoTo 0 ' Restore default error handling
    End With
    Set sortRange = Nothing
    LogUserEditsOperation "SortDashboardData: Sort applied." ' Added Log
End Sub

'------------------------------------------------------------------------------
' ApplyColorFormatting - Applies conditional formatting to Col L (Phase)
' *** Using user's provided version ***
'------------------------------------------------------------------------------
Public Sub ApplyColorFormatting(ws As Worksheet, Optional startDataRow As Long = 4)
    On Error GoTo ErrorHandler_ColorFormat ' Use specific handler
    If ws Is Nothing Then Exit Sub ' Added check

    Dim endRow As Long
    endRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    If endRow < startDataRow Then Exit Sub ' No data rows to format

    LogUserEditsOperation "ApplyColorFormatting: Applying CF to " & DB_COL_PHASE & startDataRow & ":" & DB_COL_PHASE & endRow & "." ' Added Log

    Dim rngPhase As Range
    Set rngPhase = ws.Range(DB_COL_PHASE & startDataRow & ":" & DB_COL_PHASE & endRow) ' Column L data range

    ' No need to unprotect/reprotect here if called correctly from RefreshDashboard
    ApplyStageFormatting rngPhase ' Call helper sub with specific rules

    Set rngPhase = Nothing
    Exit Sub

ErrorHandler_ColorFormat:
    LogUserEditsOperation "Error in ApplyColorFormatting: " & Err.Description
    Set rngPhase = Nothing
End Sub

'------------------------------------------------------------------------------
' ApplyStageFormatting - Helper containing specific conditional format rules for Phase (L)
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Private Sub ApplyStageFormatting(targetRng As Range)
    If targetRng Is Nothing Then Exit Sub
    If targetRng.Cells.CountLarge = 0 Then Exit Sub

    On Error Resume Next ' Handle errors applying formats
    targetRng.FormatConditions.Delete ' Clear existing rules first
    If Err.Number <> 0 Then LogUserEditsOperation "ApplyStageFormatting: Warning - Could not delete existing format conditions.": Err.Clear

    Dim fc As FormatCondition
    Dim formulaBase As String
    Dim firstCellAddress As String

    ' Use address of the first cell in the range (e.g., L4) for relative formulas
    firstCellAddress = targetRng.Cells(1).Address(RowAbsolute:=False, ColumnAbsolute:=False) ' Use relative row/col

    ' Using xlExpression with case-sensitive EXACT match as in user's code
    formulaBase = "=EXACT(" & firstCellAddress & ",""{PHASE}"")"

    ' --- Define Standard Phase Rules ---
    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "First F/U"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(208, 230, 245)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Second F/U"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(146, 198, 237)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Third F/U"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(245, 225, 113)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Long-Term F/U"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 150, 54)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Requoting"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(227, 215, 232)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Pending"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 247, 209)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "No Response"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(245, 238, 224)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Texas (No F/U)"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(230, 217, 204)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "AF"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(162, 217, 210)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "RZ"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(138, 155, 212)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "KMH"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(247, 196, 175)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "RI"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(191, 225, 243)

    Set fc = targetRng.FormatConditions.Add( _
            Type:=xlExpression, _
            Formula1:="=OR(EXACT(" & firstCellAddress & ",""OM"")," & _
                       "EXACT(" & firstCellAddress & ",""WW/OM""))")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 204, 229)   'light pink

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Converted"))
    If Not fc Is Nothing Then With fc: .Interior.Color = RGB(120, 235, 120): .Font.Bold = True: End With

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Declined"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(209, 47, 47)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Closed (Extra Order)"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(184, 39, 39)

    Set fc = targetRng.FormatConditions.Add(Type:=xlExpression, Formula1:=Replace(formulaBase, "{PHASE}", "Closed"))
    If Not fc Is Nothing Then fc.Interior.Color = RGB(166, 28, 28)
    
        ' --- START: ADD RULES FOR "Other (...)" ---
    ' Rule for "Other (Active)" - Subtle highlight, treated as active
    Set fc = targetRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Other (Active)""")
    If Not fc Is Nothing Then
        fc.Interior.Color = RGB(235, 245, 255) ' Very light blue tint
        fc.Font.Italic = True
    End If

    ' Rule for "Other (Archive)" - Different subtle highlight, treated as archived
    Set fc = targetRng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Other (Archive)""")
    If Not fc Is Nothing Then
        fc.Interior.Color = RGB(255, 245, 235) ' Very light orange/tan tint
        fc.Font.Italic = True
    End If

    If Err.Number <> 0 Then LogUserEditsOperation "ApplyStageFormatting: ERROR applying one or more format conditions. Error: " & Err.Description: Err.Clear
    On Error GoTo 0
    Set fc = Nothing
End Sub


'------------------------------------------------------------------------------
' ApplyWorkflowLocationFormatting - Applies CF to Col J (Workflow)
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Public Sub ApplyWorkflowLocationFormatting(ws As Worksheet, Optional startDataRow As Long = 4)
    On Error GoTo ErrorHandler_WorkflowFormat ' Use specific handler

    If ws Is Nothing Then Exit Sub ' Added check
    Dim endRow As Long
    endRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    If endRow < startDataRow Then Exit Sub ' No data rows to format

    LogUserEditsOperation "ApplyWorkflowLocationFormatting: Applying CF to " & DB_COL_WORKFLOW_LOCATION & startDataRow & ":" & DB_COL_WORKFLOW_LOCATION & endRow & "." ' Added Log

    Dim rngLocation As Range
    Set rngLocation = ws.Range(DB_COL_WORKFLOW_LOCATION & startDataRow & ":" & DB_COL_WORKFLOW_LOCATION & endRow) ' Column J data range

    ' No need to unprotect/reprotect here if called correctly from RefreshDashboard
    On Error Resume Next ' Handle errors applying formats
    rngLocation.FormatConditions.Delete ' Clear existing rules first
    If Err.Number <> 0 Then LogUserEditsOperation "ApplyWorkflowLocationFormatting: Warning - Could not delete existing conditions.": Err.Clear

    Dim fc As FormatCondition
    ' Add rules based on exact text values (Copied from user's code)
    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Quote Only""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(230, 240, 248)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""1. New Orders""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(208, 230, 245)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""2. Open Orders""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 247, 209)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""3. As Available Orders""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(227, 215, 232)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""4. Hold Orders""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 150, 54)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""5. Closed Files""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(230, 230, 230)

    Set fc = rngLocation.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""6. Declined Orders""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(209, 47, 47)

    If Err.Number <> 0 Then LogUserEditsOperation "ApplyWorkflowLocationFormatting: ERROR applying conditions. Error: " & Err.Description: Err.Clear
    On Error GoTo 0 ' Restore default handler

    Set rngLocation = Nothing: Set fc = Nothing
    Exit Sub

ErrorHandler_WorkflowFormat:
    LogUserEditsOperation "Error in ApplyWorkflowLocationFormatting: " & Err.Description
    Set rngLocation = Nothing: Set fc = Nothing
End Sub


'------------------------------------------------------------------------------
' ProtectUserColumns - Locks columns A-K, Unlocks L-N, Protects sheet
' *** Uses user's provided version ***
'------------------------------------------------------------------------------
Public Sub ProtectUserColumns(ws As Worksheet)
    If ws Is Nothing Then Exit Sub ' Added check
    LogUserEditsOperation "ProtectUserColumns: Locking A-K, Unlocking L-N, Protecting sheet '" & ws.Name & "'." ' Added Log
    On Error Resume Next ' Ignore errors if already protected/unprotected

    ws.Unprotect ' Ensure unprotected first

    ws.Cells.Locked = True ' Lock all

    Dim lastDataRowProtected As Long ' Use different variable name
    lastDataRowProtected = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    If lastDataRowProtected >= 4 Then
        ws.Range(DB_COL_PHASE & "4:" & DB_COL_COMMENTS & lastDataRowProtected).Locked = False ' Unlock L:N
    End If

    ' Protect sheet allowing selection of locked/unlocked cells
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True

    If Err.Number <> 0 Then LogUserEditsOperation "ProtectUserColumns: ERROR applying protection. Error: " & Err.Description: Err.Clear
    On Error GoTo 0
End Sub


'================================================================================
'              6. TEXT-ONLY DASHBOARD COPY ROUTINE
'              *** Using Reviewed Version ***
'================================================================================

'------------------------------------------------------------------------------
' CreateOrUpdateTextOnlySheet - Creates a values-only copy of the dashboard
'------------------------------------------------------------------------------
Private Sub CreateOrUpdateTextOnlySheet(wsSource As Worksheet)
    If wsSource Is Nothing Then Exit Sub ' Exit if source worksheet is invalid
    LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Starting process..."

    Dim wsValues As Worksheet, currentSheet As Worksheet
    Dim lastRowSource As Long, lastRowValues As Long
    Dim srcRange As Range

    On Error GoTo TextOnlyErrorHandler ' Enable error handling for this sub
    Set currentSheet = ActiveSheet ' Remember active sheet

    ' --- Get or Create Text-Only Sheet ---
    On Error Resume Next ' Temporarily ignore errors for sheet check/add
    Set wsValues = ThisWorkbook.Sheets(TEXT_ONLY_SHEET_NAME)
    If wsValues Is Nothing Then ' Sheet doesn't exist, create it
        Set wsValues = ThisWorkbook.Sheets.Add(After:=wsSource)
        If wsValues Is Nothing Then GoTo TextOnlyFatalError ' Could not add sheet
        wsValues.Name = TEXT_ONLY_SHEET_NAME
        If Err.Number <> 0 Then ' Naming failed (e.g., conflict)
            LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Warning - Failed to name Text-Only sheet '" & TEXT_ONLY_SHEET_NAME & "'. Using default name '" & wsValues.Name & "'."
            Err.Clear
        Else
             LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Created new sheet: '" & TEXT_ONLY_SHEET_NAME & "'."
        End If
    End If
    On Error GoTo TextOnlyErrorHandler ' Restore main error handling

    ' --- Ensure sheet is valid and prepare ---
    If wsValues Is Nothing Then GoTo TextOnlyFatalError ' Should not happen, but safety check
    wsValues.Visible = xlSheetVisible ' Ensure visible for operations
    wsValues.Cells.Clear ' Clear existing content and formats

    ' --- Copy Data (Values and Number Formats) ---
    lastRowSource = wsSource.Cells(wsSource.rows.Count, "A").End(xlUp).Row
    If lastRowSource >= 3 Then ' Ensure there's at least header data on source
        Set srcRange = wsSource.Range("A3:" & DB_COL_COMMENTS & lastRowSource) ' A3:N<lastRow> (Headers + Data)
        srcRange.Copy
        wsValues.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats ' Paste Headers+Data starting at A1
        wsValues.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths        ' Paste Column Widths
        Application.CutCopyMode = False
        LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Pasted headers & data (Rows " & 3 & "-" & lastRowSource & " from source) to '" & wsValues.Name & "'."
    Else
        LogUserEditsOperation "CreateOrUpdateTextOnlySheet: No source data found on '" & wsSource.Name & "' (A3:N) to copy. Only setting up headers."
        ' Setup headers manually if no data copied
        With wsValues.Range("A1:" & DB_COL_COMMENTS & "1") ' A1:N1
            .value = wsSource.Range("A3:" & DB_COL_COMMENTS & "3").value ' Copy header text
            .Font.Bold = wsSource.Range("A3").Font.Bold
            .Interior.Color = wsSource.Range("A3").Interior.Color
            .Font.Color = wsSource.Range("A3").Font.Color
            .HorizontalAlignment = xlCenter
        End With
    End If

    ' --- Re-apply Conditional Formatting (Data starts at row 2 on wsValues) ---
    lastRowValues = wsValues.Cells(wsValues.rows.Count, "A").End(xlUp).Row
    If lastRowValues >= 2 Then ' Headers are row 1, data starts row 2
        LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Applying conditional formatting..."
        ' Pass the specific range to ApplyStageFormatting
        ApplyStageFormatting wsValues.Range(DB_COL_PHASE & "2:" & DB_COL_PHASE & lastRowValues)
        ' Pass sheet and START row to ApplyWorkflowLocationFormatting
        ApplyWorkflowLocationFormatting wsValues, 2 ' Start formatting from row 2
        LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Applied conditional formatting to rows 2:" & lastRowValues & "."
    End If

    ' --- Final Formatting for Text-Only sheet ---
    wsValues.Columns("A:" & DB_COL_COMMENTS).VerticalAlignment = xlTop ' Align top
    wsValues.rows(1).HorizontalAlignment = xlCenter ' Center Headers

    ' --- Ensure Unprotected and Unfrozen ---
    On Error Resume Next ' Temporarily ignore errors for Unprotect/Freeze
    wsValues.Unprotect
    If ActiveSheet.Name = wsValues.Name Then ActiveWindow.FreezePanes = False
    On Error GoTo TextOnlyErrorHandler ' Restore main error handling

    ' --- Activate original sheet if needed ---
    If Not currentSheet Is Nothing Then
        If ActiveSheet.Name <> currentSheet.Name Then currentSheet.Activate
    End If
    LogUserEditsOperation "CreateOrUpdateTextOnlySheet: Finalized sheet '" & wsValues.Name & "'."

    ' --- Cleanup ---
    Set currentSheet = Nothing: Set wsValues = Nothing: Set srcRange = Nothing
    Exit Sub ' Normal exit

TextOnlyFatalError:
    LogUserEditsOperation "CreateOrUpdateTextOnlySheet: CRITICAL ERROR - Could not get or create the Text-Only sheet. Aborting Text-Only update."
    GoTo TextOnlyCleanup

TextOnlyErrorHandler:
    LogUserEditsOperation "CreateOrUpdateTextOnlySheet: ERROR [" & Err.Number & "] " & Err.Description & " (Line: " & Erl & ")"
TextOnlyCleanup: ' Label used by GoTo if sheet add failed or other fatal error
    Application.CutCopyMode = False ' Ensure copy mode is off
    On Error Resume Next ' Use Resume Next for final cleanup attempts
    If Not currentSheet Is Nothing Then
        If ActiveSheet.Name <> currentSheet.Name Then currentSheet.Activate
    End If
    Set currentSheet = Nothing: Set wsValues = Nothing: Set srcRange = Nothing
End Sub


'================================================================================
'              7. Legacy/Compatibility (Optional)
'================================================================================

' Keep this if older buttons/macros might still call the old name
Public Sub CreateSQRCTDashboard()
    LogUserEditsOperation "Legacy 'CreateSQRCTDashboard' called. Redirecting to 'RefreshDashboard_SaveAndRestoreEdits'."
    RefreshDashboard PreserveUserEdits:=False
End Sub

'------------------------------------------------------------------------------
' Module_Identity Placeholder (Ensure this module exists in your project)
'------------------------------------------------------------------------------
' If you don't have Module_Identity, create it and add this code:
'
' Option Explicit
' Public Const WORKBOOK_IDENTITY As String = "UNIQUE_ID" ' e.g., "RZ", "AF", "MASTER"
' Public Function GetWorkbookIdentity() As String
'    GetWorkbookIdentity = WORKBOOK_IDENTITY
' End Function
'------------------------------------------------------------------------------
Public Sub ResetDashboardLayout()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SQRCT Dashboard")
    ' Reset critical widths
    ws.Columns("C").ColumnWidth = 25  ' Customer Name
    ws.Columns("N").ColumnWidth = 45  ' User Comments
    ' Fix button areas
    ws.Range("C2:D2").Merge
    ws.Range("E2:F2").Merge
    ' Rebuild buttons
    modArchival.AddNavigationButtons ws
    MsgBox "Dashboard layout has been reset", vbInformation
End Sub
'============================================================
'  Workbook-level structure protection helper
'============================================================
Public Function ToggleWorkbookStructure(lockIt As Boolean) As Boolean
    ' Purpose: Protects or unprotects the Workbook Structure
    '          Used by modArchival when adding/deleting sheets.
    ' Returns: True if successful, False if error occurred.

    On Error GoTo ErrHandler
    
    LogUserEditsOperation "ToggleWorkbookStructure: Setting Lock to " & lockIt & "..." ' Added log
    
    If lockIt Then
        ThisWorkbook.Protect Structure:=True, Password:=PW_WORKBOOK ' Use PW_WORKBOOK constant
    Else
        ThisWorkbook.Unprotect Password:=PW_WORKBOOK ' Use PW_WORKBOOK constant
    End If
    
    ToggleWorkbookStructure = True ' Assume success if no error occurred
    LogUserEditsOperation "ToggleWorkbookStructure: Success." ' Added log
    Exit Function

ErrHandler:
    LogUserEditsOperation "ToggleWorkbookStructure ERR " & Err.Number & ": " & Err.Description
    ToggleWorkbookStructure = False ' Return False on error
End Function





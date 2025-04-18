Option Explicit

' --- Constants (REMOVED - Use Public Const from Module_Dashboard) ---
' If you had Private Const definitions here for sheet names or columns,
' delete them to avoid conflicts with the Public ones in Module_Dashboard.

' --- Worksheet Event Handler ---

' This code runs automatically whenever a cell value is changed ON THIS SHEET.
Private Sub Worksheet_Change(ByVal Target As Range)
    ' --- Pre-Checks ---
    If Target.Cells.CountLarge > 1 Then Exit Sub ' Exit if multiple cells changed
    ' Use Public Const from Module_Dashboard
    If Intersect(Target, Me.Range(Module_Dashboard.DB_COL_PHASE & ":" & Module_Dashboard.DB_COL_COMMENTS)) Is Nothing Then Exit Sub ' Exit if outside L:N
    If Target.Row < 4 Then Exit Sub ' Exit if header rows

    ' --- Prevent Event Recursion & Set Up Error Handling ---
    Application.EnableEvents = False
    On Error GoTo SafeExit_WorksheetChange

    ' --- Log the Initial Change ---
    Module_Dashboard.LogUserEditsOperation "Worksheet_Change Triggered: Cell " & Target.Address(False, False) & " changed." ' Use Public Logger

    ' --- Get the Document Number for the Changed Row ---
    Dim docNum As String
    Dim cleanedDocNum As String
    docNum = CStr(Me.Cells(Target.Row, "A").Value2) ' Get DocNum from Col A of changed row
    cleanedDocNum = Module_Dashboard.CleanDocumentNumber(docNum) ' Use Public Cleaner

    ' --- Process only if a valid Document Number exists ---
    If cleanedDocNum <> "" And cleanedDocNum <> Module_Dashboard.CleanDocumentNumber("Document Number") Then

        Dim wsEdits As Worksheet
        ' Dim userEditsDict As Object ' <<< No longer needed here
        Dim rowDict As Object     ' <<< NEW: Dictionary for Row Numbers
        Dim editRow As Variant    ' To store row number
        Dim destRow As Long
        Dim lastRowEdits As Long

        ' --- Get or Create UserEdits Sheet ---
        On Error Resume Next ' Check if sheet exists
        ' Use Public Const from Module_Dashboard
        Set wsEdits = ThisWorkbook.Sheets(Module_Dashboard.USEREDITS_SHEET_NAME)
        On Error GoTo SafeExit_WorksheetChange ' Restore main handler

        If wsEdits Is Nothing Then
            Module_Dashboard.LogUserEditsOperation "Worksheet_Change: '" & Module_Dashboard.USEREDITS_SHEET_NAME & "' not found. Attempting to create..."
            Module_Dashboard.SetupUserEditsSheet ' Call Public setup routine
            Set wsEdits = ThisWorkbook.Sheets(Module_Dashboard.USEREDITS_SHEET_NAME) ' Re-set reference
            If wsEdits Is Nothing Then
                Module_Dashboard.LogUserEditsOperation "Worksheet_Change: CRITICAL ERROR - Failed to create UserEdits sheet. Cannot save change."
                GoTo SafeExit_WorksheetChange ' Exit if creation failed
            End If
        End If

        ' --- Find or Create Entry in UserEdits using Row Index Dict ---
        ' Set userEditsDict = LoadUserEditsToDictionary(wsEdits) ' <<< DELETE/COMMENT OUT THIS OLD LINE

        Set rowDict = Module_Dashboard.BuildRowIndexDict(wsEdits) ' <<< NEW: Call helper to get Row# Dict

        ' Determine destination row
        If rowDict.Exists(cleanedDocNum) Then
            editRow = rowDict(cleanedDocNum) ' <<< CORRECTED: Get row number from rowDict
            destRow = CLng(editRow) ' Ensure it's Long
            Module_Dashboard.LogUserEditsOperation "Worksheet_Change: Found existing UserEdits entry for DocNum '" & cleanedDocNum & "' at row " & destRow & "."
        Else
            editRow = 0 ' Flag as not found
            lastRowEdits = wsEdits.Cells(wsEdits.rows.Count, Module_Dashboard.UE_COL_DOCNUM).End(xlUp).Row + 1
            If lastRowEdits < 2 Then lastRowEdits = 2 ' Ensure starting at row 2
            destRow = lastRowEdits
            ' Write the CLEANED document number to the new row
            wsEdits.Cells(destRow, Module_Dashboard.UE_COL_DOCNUM).Value = cleanedDocNum
            ' --- Keep rowDict synchronized ---
             Module_Dashboard.RowIndexDictAdd rowDict, cleanedDocNum, destRow ' <<< NEW: Update Row# Dict

            Module_Dashboard.LogUserEditsOperation "Worksheet_Change: Creating new UserEdits entry for DocNum '" & cleanedDocNum & "' at row " & destRow & "."
        End If

        ' --- Update UserEdits Row with Data from Changed Dashboard Cell(s) ---
        Module_Dashboard.LogUserEditsOperation "Worksheet_Change: Updating UserEdits Row " & destRow & " from Dashboard Row " & Target.Row & "."
        On Error Resume Next ' Handle potential errors writing to UserEdits
        ' Use Public Constants from Module_Dashboard
        wsEdits.Cells(destRow, Module_Dashboard.UE_COL_PHASE).Value = Me.Cells(Target.Row, Module_Dashboard.DB_COL_PHASE).Value           ' Update Phase (B <- L)
        wsEdits.Cells(destRow, Module_Dashboard.UE_COL_LASTCONTACT).Value = Me.Cells(Target.Row, Module_Dashboard.DB_COL_LASTCONTACT).Value ' Update LastContact (C <- M)
        wsEdits.Cells(destRow, Module_Dashboard.UE_COL_COMMENTS).Value = Me.Cells(Target.Row, Module_Dashboard.DB_COL_COMMENTS).Value      ' Update Comments (D <- N)
        wsEdits.Cells(destRow, Module_Dashboard.UE_COL_SOURCE).Value = Module_Identity.GetWorkbookIdentity()                 ' Update Source (E)
        wsEdits.Cells(destRow, Module_Dashboard.UE_COL_TIMESTAMP).Value = Format$(Now(), "yyyy-mm-dd hh:nn:ss") ' Update Timestamp (F)
        If Err.Number <> 0 Then
             Module_Dashboard.LogUserEditsOperation "Worksheet_Change: ERROR writing data to UserEdits row " & destRow & ". Error: " & Err.Description
             Err.Clear
        End If
        On Error GoTo SafeExit_WorksheetChange ' Restore main handler

    Else
        Module_Dashboard.LogUserEditsOperation "Worksheet_Change: Change ignored - Blank or invalid Document Number in Row " & Target.Row & ", Column A ('" & docNum & "')."
    End If ' End If cleanedDocNum <> ""

SafeExit_WorksheetChange:
    If Err.Number <> 0 Then
        MsgBox "An error occurred while trying to save your change to the UserEdits sheet:" & vbCrLf & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Worksheet Change Error"
        Module_Dashboard.LogUserEditsOperation "ERROR in Worksheet_Change: [" & Err.Number & "] " & Err.Description & " for Target " & Target.Address
    End If
    ' --- Cleanup & Re-enable Events ---
    Set rowDict = Nothing ' Release new dictionary
    Set wsEdits = Nothing
    Application.EnableEvents = True
End Sub


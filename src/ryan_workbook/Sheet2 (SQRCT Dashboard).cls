Option Explicit

' --- Constants ---
' Defines sheets and columns relevant to this specific sheet's event handler
Private Const USEREDITS_SHEET_NAME As String = "UserEdits"
Private Const UE_COL_DOCNUM As String = "A"
Private Const UE_COL_PHASE As String = "B"
Private Const UE_COL_LASTCONTACT As String = "C"
Private Const UE_COL_COMMENTS As String = "D"
Private Const UE_COL_SOURCE As String = "E"
Private Const UE_COL_TIMESTAMP As String = "F"

' Dashboard columns this sheet cares about (Editable columns L-N)
Private Const DB_COL_PHASE As String = "L"       ' Editable Engagement Phase
Private Const DB_COL_LASTCONTACT As String = "M" ' Editable Last Contact Date
Private Const DB_COL_COMMENTS As String = "N"    ' Editable User Comments
' --- End Constants ---


' --- Worksheet Event Handler ---

' This code runs automatically whenever a cell value is changed ON THIS SHEET.
Private Sub Worksheet_Change(ByVal Target As Range)
    ' --- Pre-Checks ---
    ' 1. Exit if multiple cells were changed at once (e.g., paste, fill down)
    If Target.Cells.CountLarge > 1 Then Exit Sub
    ' 2. Exit if the change was outside the user-editable columns (L:N)
    If Intersect(Target, Me.Range(DB_COL_PHASE & ":" & DB_COL_COMMENTS)) Is Nothing Then Exit Sub
    ' 3. Exit if the change was in the header rows (Rows 1-3)
    If Target.Row < 4 Then Exit Sub

    ' --- Prevent Event Recursion & Set Up Error Handling ---
    Application.EnableEvents = False
    On Error GoTo SafeExit_WorksheetChange

    ' --- Log the Initial Change ---
    ' Assumes LogUserEditsOperation is Public in Module_Dashboard
    LogUserEditsOperation "Worksheet_Change Triggered: Cell " & Target.Address(False, False) & " changed."

    ' --- Get the Document Number for the Changed Row ---
    Dim docNum As String
    Dim cleanedDocNum As String
    ' Use .Value2 for potentially faster access to raw value, CStr ensures string conversion
    docNum = CStr(Me.Cells(Target.Row, "A").Value2)
    ' *** Use the SAME cleaning function as the main module ***
    ' Assumes CleanDocumentNumber is accessible (e.g., Public in Module_Dashboard)
    cleanedDocNum = CleanDocumentNumber(docNum)

    ' --- Process only if a valid Document Number exists ---
    If cleanedDocNum <> "" And cleanedDocNum <> CleanDocumentNumber("Document Number") Then

        Dim wsEdits As Worksheet
        Dim userEditsDict As Object ' Dictionary for lookup
        Dim editRow As Variant      ' To store row number
        Dim destRow As Long
        Dim lastRowEdits As Long

        ' --- Get or Create UserEdits Sheet ---
        On Error Resume Next ' Check if sheet exists
        Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME)
        On Error GoTo SafeExit_WorksheetChange ' Restore main handler

        If wsEdits Is Nothing Then
            LogUserEditsOperation "Worksheet_Change: '" & USEREDITS_SHEET_NAME & "' not found. Attempting to create..."
            ' Call setup routine from the standard module
            ' Assumes SetupUserEditsSheet is Public in Module_Dashboard
            SetupUserEditsSheet
            Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Re-set reference
            If wsEdits Is Nothing Then
                 LogUserEditsOperation "Worksheet_Change: CRITICAL ERROR - Failed to create UserEdits sheet. Cannot save change."
                 GoTo SafeExit_WorksheetChange ' Exit if creation failed
            End If
        End If

        ' --- Find or Create Entry in UserEdits ---
        ' Load dictionary to find existing entry efficiently
        ' Assumes LoadUserEditsToDictionary is Public in Module_Dashboard
        Set userEditsDict = LoadUserEditsToDictionary(wsEdits)

        ' Determine destination row
        If userEditsDict.Exists(cleanedDocNum) Then
            editRow = userEditsDict(cleanedDocNum) ' Get existing SHEET row number
            destRow = editRow
            LogUserEditsOperation "Worksheet_Change: Found existing UserEdits entry for DocNum '" & cleanedDocNum & "' at row " & destRow & "."
        Else
            editRow = 0 ' Flag as not found
            lastRowEdits = wsEdits.Cells(wsEdits.rows.Count, UE_COL_DOCNUM).End(xlUp).Row + 1
            If lastRowEdits < 2 Then lastRowEdits = 2 ' Ensure starting at row 2
            destRow = lastRowEdits
            ' Write the CLEANED document number to the new row
            wsEdits.Cells(destRow, UE_COL_DOCNUM).Value = cleanedDocNum
            LogUserEditsOperation "Worksheet_Change: Creating new UserEdits entry for DocNum '" & cleanedDocNum & "' at row " & destRow & "."
        End If

        ' --- Update UserEdits Row with Data from Changed Dashboard Cell(s) ---
        ' Note: Target refers *only* to the cell(s) that triggered the event
        ' We copy all L, M, N values from the dashboard row to ensure UserEdits is fully synced for that DocNum
        LogUserEditsOperation "Worksheet_Change: Updating UserEdits Row " & destRow & " from Dashboard Row " & Target.Row & "."
        On Error Resume Next ' Handle potential errors writing to UserEdits
        wsEdits.Cells(destRow, UE_COL_PHASE).Value = Me.Cells(Target.Row, DB_COL_PHASE).Value       ' Update Phase (B <- L)
        wsEdits.Cells(destRow, UE_COL_LASTCONTACT).Value = Me.Cells(Target.Row, DB_COL_LASTCONTACT).Value ' Update LastContact (C <- M)
        wsEdits.Cells(destRow, UE_COL_COMMENTS).Value = Me.Cells(Target.Row, DB_COL_COMMENTS).Value     ' Update Comments (D <- N)
        wsEdits.Cells(destRow, UE_COL_SOURCE).Value = Module_Identity.GetWorkbookIdentity()        ' Update Source (E)
        wsEdits.Cells(destRow, UE_COL_TIMESTAMP).Value = Format$(Now(), "yyyy-mm-dd hh:nn:ss") ' Update Timestamp (F)
        If Err.Number <> 0 Then
             LogUserEditsOperation "Worksheet_Change: ERROR writing data to UserEdits row " & destRow & ". Error: " & Err.Description
             Err.Clear
        End If
        On Error GoTo SafeExit_WorksheetChange ' Restore main handler

    Else
        LogUserEditsOperation "Worksheet_Change: Change ignored - Blank or invalid Document Number in Row " & Target.Row & ", Column A ('" & docNum & "')."
    End If ' End If cleanedDocNum <> ""

SafeExit_WorksheetChange:
    If Err.Number <> 0 Then
        MsgBox "An error occurred while trying to save your change to the UserEdits sheet:" & vbCrLf & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Worksheet Change Error"
        ' Log the error using the helper function
        LogUserEditsOperation "ERROR in Worksheet_Change: [" & Err.Number & "] " & Err.Description & " for Target " & Target.Address
    End If
    ' --- Cleanup & Re-enable Events ---
    Set userEditsDict = Nothing
    Set wsEdits = Nothing
    Application.EnableEvents = True
End Sub


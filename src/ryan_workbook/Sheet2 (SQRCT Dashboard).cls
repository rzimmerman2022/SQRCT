VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet69"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' --- Constants ---
' Note: These could be moved to a central module if preferred
Private Const USEREDITS_SHEET_NAME As String = "UserEdits"
Private Const UE_COL_DOCNUM As String = "A"
Private Const UE_COL_PHASE As String = "B"
Private Const UE_COL_LASTCONTACT As String = "C"
' Private Const UE_COL_EMAIL As String = "D" ' REMOVED - Column Shift Required Below
Private Const UE_COL_COMMENTS As String = "D" ' Shifted from E
Private Const UE_COL_SOURCE As String = "E"   ' Shifted from F
Private Const UE_COL_TIMESTAMP As String = "F" ' Shifted from G

' UPDATED: Dashboard column constants to match Module_Dashboard
Private Const DB_COL_PHASE As String = "L"         ' Shifted from K
Private Const DB_COL_LASTCONTACT As String = "M"    ' Shifted from L
Private Const DB_COL_COMMENTS As String = "N"       ' Shifted from M
' --- End Constants ---


Private Sub Worksheet_Change(ByVal Target As Range)
    ' Only run if changes are in columns L:N (user-editable columns) - Use Constants (Adjusted range)
    If Intersect(Target, Me.Range(DB_COL_PHASE & ":" & DB_COL_COMMENTS)) Is Nothing Then Exit Sub
    ' Skip rows above row 4 (header rows)
    If Target.Row < 4 Then Exit Sub
    Application.EnableEvents = False
    On Error GoTo SafeExit

    ' Log the change operation (using Module_Dashboard_UserEdits' logger)
    Module_Dashboard_UserEdits.LogUserEditsOperation "Worksheet change in row " & Target.Row & ", column " & Target.Column

    ' Grab the Document Number from column A (unique key)
    Dim docNum As String
    docNum = Trim(CStr(Me.Cells(Target.Row, "A").Value))
    If docNum <> "" And docNum <> "Document Number" Then
        Dim wsEdits As Worksheet
        Dim userEditsDict As Object ' Dictionary for lookup
        Dim editRow As Variant      ' To store row number

        On Error Resume Next
        Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Use Constant
        On Error GoTo SafeExit

        If wsEdits Is Nothing Then
            ' Create UserEdits sheet directly with standardized structure
            ' Call SetupUserEditsSheet from Module_Dashboard_UserEdits to ensure consistency
            Module_Dashboard_UserEdits.SetupUserEditsSheet
            Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Re-set reference after creation
            If wsEdits Is Nothing Then GoTo SafeExit ' Exit if creation failed
        End If

        ' Load UserEdits into dictionary for faster lookup
        ' Note: Consider if loading the dictionary on every change is efficient enough.
        ' If performance is critical with many rapid changes, caching the dictionary might be needed,
        ' but that adds complexity (e.g., invalidating cache on refresh).
        Set userEditsDict = Module_Dashboard_UserEdits.LoadUserEditsToDictionary(wsEdits) ' Use helper from Module_Dashboard_UserEdits

        ' Get the last row with data in column A
        Dim lastRow As Long
        lastRow = wsEdits.Cells(wsEdits.Rows.Count, UE_COL_DOCNUM).End(xlUp).Row ' Use Constant
        If lastRow < 1 Then lastRow = 1

        ' Find existing row using dictionary
        If userEditsDict.Exists(docNum) Then
            editRow = userEditsDict(docNum) ' Get existing row number
        Else
            editRow = 0 ' Flag as not found
        End If

        Dim destRow As Long
        If editRow = 0 Then
            ' Document number not found, create new entry
            destRow = lastRow + 1
            wsEdits.Cells(destRow, UE_COL_DOCNUM).Value = docNum ' Use Constant
        Else
            ' Document number found, update existing entry
            destRow = editRow
        End If

        ' Map dashboard columns to UserEdits with standardized mapping using constants:
        wsEdits.Cells(destRow, UE_COL_PHASE).Value = Me.Cells(Target.Row, DB_COL_PHASE).Value       ' Engagement Phase (B)
        wsEdits.Cells(destRow, UE_COL_LASTCONTACT).Value = Me.Cells(Target.Row, DB_COL_LASTCONTACT).Value ' Last Contact Date (C)
        wsEdits.Cells(destRow, UE_COL_COMMENTS).Value = Me.Cells(Target.Row, DB_COL_COMMENTS).Value     ' User Comments (D)

        ' Set ChangeSource to workbook identity and add timestamp using constants
        wsEdits.Cells(destRow, UE_COL_SOURCE).Value = Module_Identity.GetWorkbookIdentity()  ' Source (E)
        wsEdits.Cells(destRow, UE_COL_TIMESTAMP).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss") ' Timestamp (F)

        Module_Dashboard_UserEdits.LogUserEditsOperation "Updated UserEdits for DocNumber " & docNum & " with attribution " & Module_Identity.GetWorkbookIdentity() ' Use Module_Dashboard_UserEdits logger
    End If

SafeExit:
    Set userEditsDict = Nothing ' Clean up dictionary
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
        Module_Dashboard_UserEdits.LogUserEditsOperation "ERROR in Worksheet_Change: " & Err.Description ' Use Module_Dashboard_UserEdits logger
    End If
End Sub

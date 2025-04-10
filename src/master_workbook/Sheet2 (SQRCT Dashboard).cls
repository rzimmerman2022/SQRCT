VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' --- Constants ---
' Note: These could be moved to a central module if preferred
Private Const USEREDITS_SHEET_NAME As String = "UserEdits"
Private Const MASTER_ATTRIBUTION As String = "MASTER" ' Specific attribution for Master

Private Const UE_COL_DOCNUM As String = "A"
Private Const UE_COL_PHASE As String = "B"
Private Const UE_COL_LASTCONTACT As String = "C"
Private Const UE_COL_EMAIL As String = "D"
Private Const UE_COL_COMMENTS As String = "E"
Private Const UE_COL_SOURCE As String = "F"
Private Const UE_COL_TIMESTAMP As String = "G"

Private Const DB_COL_PHASE As String = "K"
Private Const DB_COL_LASTCONTACT As String = "L"
Private Const DB_COL_EMAIL As String = "M"
Private Const DB_COL_COMMENTS As String = "N"
' --- End Constants ---


Private Sub Worksheet_Change(ByVal Target As Range)
    ' Only run if changes are in columns K:N (user-editable columns) - Use Constants
    If Intersect(Target, Me.Range(DB_COL_PHASE & ":" & DB_COL_COMMENTS)) Is Nothing Then Exit Sub
    ' Skip rows above row 4 (header rows)
    If Target.Row < 4 Then Exit Sub
    Application.EnableEvents = False
    On Error GoTo SafeExit

    ' Note: Logging omitted in Master sheet code

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
            ' Call SetupUserEditsSheet from Module1 to ensure consistency
            Module1.SetupUserEditsSheet
            Set wsEdits = ThisWorkbook.Sheets(USEREDITS_SHEET_NAME) ' Re-set reference after creation
            If wsEdits Is Nothing Then GoTo SafeExit ' Exit if creation failed
        End If

        ' Load UserEdits into dictionary for faster lookup
        Set userEditsDict = Module1.LoadUserEditsToDictionary(wsEdits) ' Use helper from Module1

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
        wsEdits.Cells(destRow, UE_COL_PHASE).Value = Me.Cells(Target.Row, DB_COL_PHASE).Value       ' Engagement Phase
        wsEdits.Cells(destRow, UE_COL_LASTCONTACT).Value = Me.Cells(Target.Row, DB_COL_LASTCONTACT).Value ' Last Contact Date
        wsEdits.Cells(destRow, UE_COL_EMAIL).Value = Me.Cells(Target.Row, DB_COL_EMAIL).Value         ' Email Contact
        wsEdits.Cells(destRow, UE_COL_COMMENTS).Value = Me.Cells(Target.Row, DB_COL_COMMENTS).Value     ' User Comments

        ' Set ChangeSource to "MASTER" and add timestamp using constants
        wsEdits.Cells(destRow, UE_COL_SOURCE).Value = MASTER_ATTRIBUTION ' Use Master constant
        wsEdits.Cells(destRow, UE_COL_TIMESTAMP).Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        ' Note: Logging omitted
    End If

SafeExit:
    Set userEditsDict = Nothing ' Clean up dictionary
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
        ' Note: Logging omitted
        ' Removed invalid/orphaned .Font reference here
    End If
End Sub

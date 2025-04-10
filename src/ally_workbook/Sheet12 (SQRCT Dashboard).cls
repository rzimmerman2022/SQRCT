VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Only run if changes are in columns K:N (user-editable columns)
    If Intersect(Target, Me.Range("K:N")) Is Nothing Then Exit Sub
    ' Skip rows above row 4 (header rows)
    If Target.Row < 4 Then Exit Sub
    
    Application.EnableEvents = False
    On Error GoTo SafeExit
    
    ' Log the change operation
    LogUserEditsOperation "Worksheet change in row " & Target.Row & ", column " & Target.Column
    
    ' Grab the Document Number from column A (unique key)
    Dim docNum As String
    docNum = Trim(CStr(Me.Cells(Target.Row, "A").Value))
    
    If docNum <> "" And docNum <> "Document Number" Then
        Dim wsEdits As Worksheet
        On Error Resume Next
        Set wsEdits = ThisWorkbook.Sheets("UserEdits")
        On Error GoTo SafeExit
        
        If wsEdits Is Nothing Then
            ' Create UserEdits sheet directly with standardized structure
            Set wsEdits = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsEdits.Name = "UserEdits"
            
            ' Set up headers with standardized structure
            With wsEdits.Range("A1:G1")
                .Value = Array("DocNumber", "Engagement Phase", "Last Contact Date", _
                               "Email Contact", "User Comments", "ChangeSource", "Timestamp")
                .Font.Bold = True
                .Interior.Color = RGB(16, 107, 193)
                .Font.Color = RGB(255, 255, 255)
            End With
            
            wsEdits.Visible = xlSheetHidden
        End If
        
        ' Get the last row with data in column A
        Dim lastRow As Long
        lastRow = wsEdits.Cells(wsEdits.Rows.Count, "A").End(xlUp).Row
        If lastRow < 1 Then lastRow = 1
        
        ' More robust method to find existing document number
        Dim foundRow As Long
        foundRow = 0
        
        If lastRow > 1 Then
            Dim i As Long
            For i = 2 To lastRow
                ' Compare cleaned document numbers for exact match
                If Trim(LCase(wsEdits.Cells(i, "A").Value)) = Trim(LCase(docNum)) Then
                    foundRow = i
                    Exit For
                End If
            Next i
        End If
        
        Dim destRow As Long
        If foundRow = 0 Then
            ' Document number not found, create new entry
            destRow = lastRow + 1
            wsEdits.Cells(destRow, "A").Value = docNum
        Else
            ' Document number found, update existing entry
            destRow = foundRow
        End If
        
        ' Dashboard K ? UserEdits B (Engagement Phase)
        ' Dashboard L ? UserEdits C (Last Contact Date)
        ' Dashboard M ? UserEdits D (Email Contact)
        ' Dashboard N ? UserEdits E (User Comments)
        wsEdits.Cells(destRow, "B").Value = Me.Cells(Target.Row, "K").Value  ' Engagement Phase
        wsEdits.Cells(destRow, "C").Value = Me.Cells(Target.Row, "L").Value  ' Last Contact Date
        wsEdits.Cells(destRow, "D").Value = Me.Cells(Target.Row, "M").Value  ' Email Contact
        wsEdits.Cells(destRow, "E").Value = Me.Cells(Target.Row, "N").Value  ' User Comments
        
        ' Set ChangeSource to workbook identity and add timestamp
        wsEdits.Cells(destRow, "F").Value = GetWorkbookIdentity()  ' Use workbook identity from Module_Identity
        wsEdits.Cells(destRow, "G").Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss")
        
        LogUserEditsOperation "Updated UserEdits for DocNumber " & docNum & " with attribution " & GetWorkbookIdentity()
    End If
    
SafeExit:
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
        LogUserEditsOperation "ERROR in Worksheet_Change: " & Err.Description
    End If
End Sub




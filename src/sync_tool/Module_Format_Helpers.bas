Attribute VB_Name = "Module_Format_Helpers"
Option Explicit

'===============================================================================
' MODULE_FORMAT_HELPERS
'-------------------------------------------------------------------------------
' Purpose:
' Provides helper routines for applying formatting to cells based on their
' content. In particular, it implements ApplyAttributionFormatting, which
' sets a cell's background color based on the ChangeSource attribution value.
'
' Design Considerations / Best Practices:
' - Uses Option Explicit to enforce variable declaration.
' - Implements robust error handling.
' - Relies on well-named constants for attribution codes and colors.
'
' Usage:
' Call ApplyAttributionFormatting(targetCell) to format a cell (typically
' in WriteUserEditsToWorkbook) based on its attribution value.
'===============================================================================

'-------------------------------------------------------------------------------
' ApplyAttributionFormatting
'-------------------------------------------------------------------------------
' Applies background color formatting to a cell based on its ChangeSource
' attribution value.
'
' Expected behavior:
' - "AF" ? sets background to COLOR_ALLY
' - "RZ" ? sets background to COLOR_RYAN
' - "MASTER" ? sets background to COLOR_MASTER
' - "AF+RZ", etc ? sets background to COLOR_COMBINED
'
' If the cell is empty or its value doesn't match any known code, any
' existing formatting is removed.
'-------------------------------------------------------------------------------
Public Sub ApplyAttributionFormatting(targetCell As Range)
    On Error GoTo ErrorHandler
    
    ' Validate the input cell.
    If targetCell Is Nothing Then Exit Sub
    
    If Module_Utilities.IsNullOrEmpty(targetCell.value) Then
        targetCell.Interior.ColorIndex = xlNone
        Exit Sub
    End If
    
    Dim attribValue As String
    attribValue = UCase(Trim(targetCell.value))
    
    ' Apply color based on attribution.
    Select Case True
        Case attribValue = ATTRIBUTION_ALLY
            targetCell.Interior.Color = COLOR_ALLY
        Case attribValue = ATTRIBUTION_RYAN
            targetCell.Interior.Color = COLOR_RYAN
        Case attribValue = ATTRIBUTION_MASTER
            targetCell.Interior.Color = COLOR_MASTER
        Case InStr(attribValue, "+") > 0
            targetCell.Interior.Color = COLOR_COMBINED
        Case Else
            targetCell.Interior.ColorIndex = xlNone
    End Select
    
    Exit Sub
    
ErrorHandler:
    ' Log the error using your logger module if available; otherwise, show a message.
    On Error Resume Next
    Module_SyncTool_Logger.LogMessage "Error in ApplyAttributionFormatting: " & err.Description & _
        " (Error " & err.Number & ")", "ERROR"
End Sub


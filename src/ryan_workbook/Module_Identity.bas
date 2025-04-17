Attribute VB_Name = "Module_Identity"
Option Explicit

' ===============================================================================
' MODULE_IDENTITY - Identifies which user's workbook this is
' THIS CONSTANT MUST BE DIFFERENT IN EACH WORKBOOK:
' - Use "RZ" for Ryan's workbook
' - Use "AF" for Ally's workbook
' - Use "MASTER" for the Master workbook
' ===============================================================================
Public Const WORKBOOK_IDENTITY As String = "RZ"  ' CHANGE THIS VALUE PER WORKBOOK

' ===============================================================================
' GET_WORKBOOK_IDENTITY - Returns the identity of this workbook
' ===============================================================================
Public Function GetWorkbookIdentity() As String
    GetWorkbookIdentity = WORKBOOK_IDENTITY
End Function

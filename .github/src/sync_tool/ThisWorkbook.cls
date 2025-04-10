VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' THISWORKBOOK MODULE
' Contains Workbook event handlers
' Note: To be added directly to the ThisWorkbook module in the VBA editor
'===============================================================================

'===============================================================================
' WORKBOOK_OPEN - Handles the workbook open event safely
'===============================================================================
Private Sub Workbook_Open()
    ' Defer initialization to avoid timing issues with module loading
    Application.OnTime Now + TimeValue("00:00:01"), "SafeStartup"
End Sub

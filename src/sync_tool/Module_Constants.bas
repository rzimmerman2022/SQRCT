Attribute VB_Name = "MODULE_CONSTANTS"
Option Explicit

'===============================================================================
' MODULE_CONSTANTS
'-------------------------------------------------------------------------------
' Purpose:
' Centralizes all constant definitions used throughout the SyncTool
' application. This includes sheet names, cell references, color values,
' formatting standards, and error message templates.
'
' Design Considerations:
' - No executable code is included; this module acts as a single source
'   of truth for global configuration values.
' - All other modules reference these constants directly, ensuring
'   consistency and maintainability.
'===============================================================================

'-------------------------------------------------------------------------------
' Sheet Names in the SyncTool workbook (internal sheets)
'-------------------------------------------------------------------------------
Public Const SYNCTOOL_LOG_SHEET As String = "SyncLog"
Public Const SYNCTOOL_HISTORY_SHEET As String = "DocChangeHistory"
Public Const SYNCTOOL_MERGEDATA_SHEET As String = "MergeData"
Public Const SYNCTOOL_DASHBOARD_SHEET As String = "SQRCT SyncTool Dashboard"
Public Const ERROR_LOG_SHEET As String = "ErrorLog"  ' Added this constant

'-------------------------------------------------------------------------------
' External Workbook Sheet Name (for UserEdits data)
'-------------------------------------------------------------------------------
Public Const EXTERNAL_USEREDITS_SHEET As String = "UserEdits"

'-------------------------------------------------------------------------------
' User Attribution Codes - used to tag the source of data changes
'-------------------------------------------------------------------------------
Public Const ATTRIBUTION_ALLY As String = "AF"
Public Const ATTRIBUTION_RYAN As String = "RZ"
Public Const ATTRIBUTION_MASTER As String = "MASTER"

'-------------------------------------------------------------------------------
' SyncTool Dashboard Cell References - layout details for the dashboard
'-------------------------------------------------------------------------------
Public Const CELL_ALLY_PATH As String = "C3"
Public Const CELL_RYAN_PATH As String = "C4"
Public Const CELL_MASTER_PATH As String = "C5"
Public Const CELL_STATUS_DISPLAY As String = "B10"
Public Const CELL_LAST_SYNC_TIME As String = "B11"

'-------------------------------------------------------------------------------
' UserEdits Sheet Column References - mapping for data extraction and writing
'-------------------------------------------------------------------------------
Public Const COL_DOCNUMBER As String = "A"
Public Const COL_ENGAGEMENTPHASE As String = "B"
Public Const COL_LASTCONTACTDATE As String = "C"
Public Const COL_EMAILCONTACT As String = "D"
Public Const COL_USERCOMMENTS As String = "E"
Public Const COL_CHANGESOURCE As String = "F"
Public Const COL_TIMESTAMP As String = "G"

'-------------------------------------------------------------------------------
' Color Values (as Long) - Pre-calculated values for consistent UI styling
'-------------------------------------------------------------------------------
Public Const COLOR_HEADER_BLUE As Long = 12656656 ' RGB(16, 107, 193) - Professional blue
Public Const COLOR_HEADER_TEXT As Long = 16777215 ' RGB(255, 255, 255) - White text
Public Const COLOR_ALLY As Long = 16764876 ' RGB(204, 229, 255) - Light blue for Ally
Public Const COLOR_RYAN As Long = 13434828 ' RGB(204, 255, 204) - Light green for Ryan
Public Const COLOR_MASTER As Long = 15921906 ' RGB(242, 242, 242) - Light gray for Master
Public Const COLOR_COMBINED As Long = 10281215 ' RGB(255, 235, 156) - Light yellow for combined
Public Const COLOR_ERROR As Long = 13537535 ' RGB(255, 199, 206) - Light red for errors
Public Const COLOR_WARNING As Long = 10281215 ' RGB(255, 235, 156) - Light yellow for warnings

'-------------------------------------------------------------------------------
' Date Format Standards - Ensures consistent date/time formatting throughout the app
'-------------------------------------------------------------------------------
Public Const FORMAT_TIMESTAMP As String = "yyyy-mm-dd hh:mm:ss"
Public Const FORMAT_DATE As String = "mm/dd/yyyy"

'-------------------------------------------------------------------------------
' Error Message Templates - Templates with placeholders for dynamic formatting
'-------------------------------------------------------------------------------
Public Const ERR_FILE_NOT_FOUND As String = "File not found: {0}"
Public Const ERR_INVALID_PATH As String = "Invalid file path: {0}"
Public Const ERR_SHEET_NOT_FOUND As String = "Sheet not found: {0} in {1}"
Public Const ERR_DUPLICATE_PATH As String = "{0} file and {1} file cannot be the same file."

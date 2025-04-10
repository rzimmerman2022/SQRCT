Attribute VB_Name = "Module_Utilities"
Option Explicit

'===============================================================================
' MODULE_UTILITIES
' Contains general helper functions used across multiple modules.
' These functions do not modify any workbooks directly.
'===============================================================================

'===============================================================================
' IS_NULL_OR_EMPTY - Checks if a value is null, empty, or missing.
'
' Parameters:
' value - A Variant representing the value to check.
'
' Returns:
' Boolean - True if the value is Null, an empty string, or missing; False otherwise.
'===============================================================================
Public Function IsNullOrEmpty(value As Variant) As Boolean
    If IsNull(value) Then
        IsNullOrEmpty = True
    ElseIf IsObject(value) Then
        IsNullOrEmpty = (value Is Nothing)
    ElseIf VarType(value) = vbString Then
        IsNullOrEmpty = (Len(Trim(CStr(value))) = 0)
    Else
        IsNullOrEmpty = False
    End If
End Function

'===============================================================================
' FORMAT_ERROR_MESSAGE - Replaces placeholders in an error message template.
'
' Parameters:
' template - A string containing placeholders (e.g., "{0}", "{1}").
' args() - A ParamArray of values to replace the placeholders.
'
' Returns:
' String - The formatted error message with all placeholders replaced.
'===============================================================================
Public Function FormatErrorMessage(template As String, ParamArray args() As Variant) As String
    Dim result As String
    Dim i As Long
    
    result = template
    For i = LBound(args) To UBound(args)
        result = Replace(result, "{" & i & "}", CStr(args(i)))
    Next i
    
    FormatErrorMessage = result
End Function

'===============================================================================
' IS_VALID_ATTRIBUTION - Validates source attribution codes.
'
' Parameters:
' attribution - A string representing the source attribution (e.g., "AF", "RZ", "MASTER").
'
' Returns:
' Boolean - True if the attribution is valid (either a valid single code or a valid
' combination separated by "+"); False otherwise.
'===============================================================================
Public Function IsValidAttribution(attribution As String) As Boolean
    ' An empty or blank attribution is invalid.
    If Trim(attribution) = "" Then
        IsValidAttribution = False
        Exit Function
    End If
    
    ' Check for a valid single source.
    If attribution = ATTRIBUTION_ALLY Or _
       attribution = ATTRIBUTION_RYAN Or _
       attribution = ATTRIBUTION_MASTER Then
        IsValidAttribution = True
        Exit Function
    End If
    
    ' If attribution contains a "+" sign, validate each part.
    If InStr(attribution, "+") > 0 Then
        Dim parts As Variant
        parts = Split(attribution, "+")
        
        Dim allValid As Boolean
        allValid = True
        
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            Dim part As String
            part = Trim(parts(i))
            
            If part <> ATTRIBUTION_ALLY And _
               part <> ATTRIBUTION_RYAN And _
               part <> ATTRIBUTION_MASTER Then
                allValid = False
                Exit For
            End If
        Next i
        
        IsValidAttribution = allValid
        Exit Function
    End If
    
    ' Otherwise, the attribution is not valid.
    IsValidAttribution = False
End Function

'===============================================================================
' COMBINE_ATTRIBUTIONS - Combines multiple attribution codes into one string,
' removing duplicates.
'
' Parameters:
' attributions() - A ParamArray of attribution strings.
'
' Returns:
' String - A combined string of unique attribution codes separated by "+".
'===============================================================================
Public Function CombineAttributions(ParamArray attributions() As Variant) As String
    Dim result As String
    Dim uniqueDict As Object
    Dim i As Long
    Dim attribution As Variant
    
    ' Create a dictionary to store unique attribution codes.
    Set uniqueDict = CreateObject("Scripting.Dictionary")
    
    ' Add each non-empty attribution to the dictionary.
    For i = LBound(attributions) To UBound(attributions)
        attribution = CStr(attributions(i))
        If Trim(attribution) <> "" And Not uniqueDict.Exists(attribution) Then
            uniqueDict.Add attribution, True
        End If
    Next i
    
    ' Process any combined attributions (e.g., "AF+RZ") to split and add individually.
    Dim tempSources As New Collection
    For Each attribution In uniqueDict.keys
        If InStr(attribution, "+") > 0 Then
            Dim parts As Variant
            parts = Split(attribution, "+")
            
            Dim j As Long
            For j = LBound(parts) To UBound(parts)
                If Trim(parts(j)) <> "" Then
                    If Not uniqueDict.Exists(Trim(parts(j))) Then
                        uniqueDict.Add Trim(parts(j)), True
                    End If
                End If
            Next j
            
            uniqueDict.Remove attribution
        End If
    Next attribution
    
    ' Join all unique attribution codes using "+" as a delimiter.
    result = ""
    For Each attribution In uniqueDict.keys
        If result <> "" Then result = result & "+"
        result = result & attribution
    Next attribution
    
    CombineAttributions = result
End Function

'===============================================================================
' COLUMN_LETTER_FROM_NUMBER - Converts a column number to its corresponding Excel column letter.
'
' Parameters:
' columnNumber - A Long representing the column number (e.g., 1 for A, 2 for B).
'
' Returns:
' String - The corresponding Excel column letter.
'===============================================================================
Public Function ColumnLetterFromNumber(columnNumber As Long) As String
    ColumnLetterFromNumber = Split(Cells(1, columnNumber).Address, "$")(1)
End Function

'===============================================================================
' GET_FILE_NAME - Extracts the file name from a full file path.
'
' Parameters:
' fullPath - A string containing the full path to a file.
'
' Returns:
' String - The file name extracted from the path. If the input is empty, returns "[Empty Path]".
'===============================================================================
Public Function GetFileName(fullPath As String) As String
    If IsNullOrEmpty(fullPath) Then
        GetFileName = "[Empty Path]"
        Exit Function
    End If
    
    ' Use InStrRev to locate the last "\" and extract the file name.
    If InStr(fullPath, "\") > 0 Then
        GetFileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
    Else
        GetFileName = fullPath
    End If
End Function

'===============================================================================
' FORMAT_HEADERS - Applies standard formatting to sheet headers.
'
' Parameters:
' ws - A Worksheet where headers will be formatted.
' headerRange - A Range object representing the header cells.
' headerText - (Optional) An array of header titles to be applied to headerRange.
'
' Returns:
' Nothing.
'
' Purpose:
' Applies bold formatting, a standard background color, and text color to header cells.
' If headerText is provided, it populates the headerRange with these values.
'===============================================================================
Public Sub FormatHeaders(ws As Worksheet, headerRange As Range, Optional headerText As Variant)
    ' Validate that the worksheet and range are provided.
    If ws Is Nothing Then Exit Sub
    If headerRange Is Nothing Then Exit Sub
    
    ' If header text is provided as an array, apply it to the header range.
    If Not IsMissing(headerText) Then
        If IsArray(headerText) Then
            headerRange.value = headerText
        End If
    End If
    
    ' Apply standard header formatting: bold text, background color, and text color.
    With headerRange
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER_BLUE
        .Font.Color = COLOR_HEADER_TEXT
    End With
    
    ' Auto-fit the column(s) for a neat appearance.
    headerRange.EntireColumn.AutoFit
End Sub


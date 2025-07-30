Option Explicit

' --- Enums and Constants ---
Public Enum SqrctColors
    DarkBlue = &H874B0A
    MediumBlue = &HD47800
    LightBlue = &HF5BB5E
    SuccessGreen = &H50AF4C
    WarningOrange = &H809FF
    ErrorRed = &H3935E5
    TextDark = &H363534
    TextLight = &HFFFFFF
    BorderGrey = &HC8C8C8
    BackgroundGrey = &HFAF9F8
End Enum
Private Const CARD_WIDTH As Long = 5
Private Const CARD_HEIGHT As Long = 4
Private Const SECTION_MARGIN As Long = 2
Private Const FONT_HEADER As String = "Segoe UI Light"
Private Const FONT_TITLE As String = "Segoe UI Semibold"
Private Const FONT_BODY As String = "Segoe UI"
Private Const FONT_NUMBER As String = "Segoe UI Light"
#Const DEBUG_MODE = True  ' Set to True when debugging, False for release
' --- End Constants ---


Public Sub BuildModernPerfDashboard()
    Dim wsPerf As Worksheet
    Dim dataModel As Object ' Scripting.Dictionary (late bound)
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    Dim t0 As Double

    t0 = Timer ' Initialize timer

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation

    On Error GoTo CleanFail_Perf
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' --- Ensure Config Sheet Exists ---
    ' Assumes EnsureConfigSheet is Public in Module_Dashboard or copied locally
    Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Ensuring Config Sheet..." ' Use DebugLog for major step
    On Error Resume Next ' Try calling it
    Module_Dashboard.EnsureConfigSheet
    If Err.Number <> 0 Then ' If it failed (e.g., not Public), try calling a local version
        Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Module_Dashboard.EnsureConfigSheet failed. Trying local..." ' Log attempt
        Err.Clear
        EnsureConfigSheet_Local ' Call local copy if Module_Dashboard one fails/doesn't exist
        If Err.Number <> 0 Then
             Module_Dashboard.DebugLog "BuildModernPerfDashboard", "WARNING: EnsureConfigSheet_Local also failed or missing. Err: " & Err.Description ' Log final failure
             Err.Clear
        End If
    End If
    On Error GoTo CleanFail_Perf ' Restore error handler

    Const PERF_DASH_SHEET_NAME As String = "SQRCT PERF DASH"
    Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Getting/Creating Perf Dash Sheet: " & PERF_DASH_SHEET_NAME ' Log step
    On Error Resume Next
    Set wsPerf = ThisWorkbook.Sheets(PERF_DASH_SHEET_NAME)
    On Error GoTo 0 ' Clear Error from sheet check

    If wsPerf Is Nothing Then
        Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Sheet not found. Adding new sheet..." ' Log add
        On Error GoTo CleanFail_Perf ' Errors during Add are critical
        Set wsPerf = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If wsPerf Is Nothing Then
             Module_Dashboard.DebugLog "BuildModernPerfDashboard", "FATAL: Failed to add new worksheet." ' Log fatal error
             GoTo ErrorExit_Perf ' Exit if Add failed
        End If
        On Error Resume Next ' Naming can fail but is less critical
        wsPerf.Name = PERF_DASH_SHEET_NAME
         If Err.Number <> 0 Then
              Module_Dashboard.DebugLog "BuildModernPerfDashboard", "WARNING: Failed to name sheet '" & PERF_DASH_SHEET_NAME & "'. Using '" & wsPerf.Name & "'. Err: " & Err.Description ' Log name fail
              Err.Clear
         End If
         On Error GoTo CleanFail_Perf ' Restore critical error handler
    Else
        Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Sheet found. Clearing contents..." ' Log clear
        ClearPerfDashboardContents wsPerf ' Use local clear sub (Code Provided Below)
    End If

    ' --- Double check wsPerf object after Get/Create ---
    If wsPerf Is Nothing Then
        Module_Dashboard.DebugLog "BuildModernPerfDashboard", "FATAL: wsPerf object Is Nothing after Get/Create attempt." ' Log fatal error
        GoTo ErrorExit_Perf
    End If
    Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Perf Dash sheet obtained: '" & wsPerf.Name & "'" ' Log success

    wsPerf.Activate
    ApplyModernDashboardTheme wsPerf ' Use local theme sub (Code Provided Below)

    ' --- Build the specific data model reading from "SQRCT Dashboard" ---
    Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Building Performance Data Model..." ' Log step
    Set dataModel = BuildPerformanceDataModel_FromDashboard() ' Call revised builder below
    If dataModel Is Nothing Then
         Module_Dashboard.DebugLog "BuildModernPerfDashboard", "FATAL: BuildPerformanceDataModel_FromDashboard returned Nothing." ' Log fatal error
         GoTo ErrorExit_Perf ' Exit if data model failed
    End If
     Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Data Model built successfully." ' Log success

    ' --- Create Visualization Sections with Charts (only if needed) ---
    Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Creating Visualizations..." ' Log step
    Dim forceRebuild As Boolean
    forceRebuild = (wsPerf.ChartObjects.Count = 0) ' Rebuild only if no charts exist
    Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Force chart rebuild = " & forceRebuild ' Log rebuild status

    ' --- Calls to Visualization Functions (Code Provided Below) ---
    CreateExecutiveSnapshot wsPerf, dataModel ', forceRebuild ' ForceRebuild currently unused in helpers
    CreatePipelineAnalytics wsPerf, dataModel ', forceRebuild
    CreatePerformanceMetrics wsPerf, dataModel ', forceRebuild
    CreateForecastTrends wsPerf, dataModel ', forceRebuild
    Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Visualizations created." ' Log completion

    AddDashboardControls wsPerf ' Use local controls sub (Code Provided Below)
    AddPerfRefreshButton wsPerf ' Use local button sub (Code Provided Below)
     Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Controls and Button added." ' Log completion

    ' --- Update Historical Metrics ---
    Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Updating Historical Metrics..." ' Log step
    ' NOTE: UpdatePerfHistoricalMetrics was added to Module_Dashboard previously.
    ' Calling the version in Module_Dashboard.
    Module_Dashboard.UpdateHistoricalMetrics dataModel
    Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Called Module_Dashboard.UpdateHistoricalMetrics." ' Log call


    wsPerf.Range("A1").Select
    Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Process completed successfully (before cleanup)." ' Log success

CleanExit_Perf:
    On Error Resume Next ' Prevent cleanup errors from masking original error
    #If DEBUG_MODE Then
        Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Cleanup: Restoring App Settings..." ' Log cleanup step
    #End If
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation
    Set wsPerf = Nothing
    Set dataModel = Nothing
    #If DEBUG_MODE Then
        Module_Dashboard.DebugLog "BuildModernPerfDashboard", "EXIT (Clean). Total Time: " & format(Timer - t0, "0.00") & "s" ' Log exit time
    #End If
    Exit Sub

ErrorExit_Perf:
    MsgBox "Performance Dashboard generation failed. Please check the 'SQRCT Dashboard' sheet exists and has data.", vbCritical, "Perf Dash Failed"
    GoTo CleanExit_Perf ' Go to cleanup even after MsgBox

CleanFail_Perf:
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    Dim errSource As String: errSource = Err.Source
    Dim errLine As Long: errLine = Erl
    #If DEBUG_MODE Then
        Module_Dashboard.DebugLog "BuildModernPerfDashboard", "ERROR Handler! Err=" & errNum & ": " & errDesc & " (" & errSource & ", Line: " & errLine & ")" ' Log error details
    #End If
    MsgBox "An error occurred in BuildModernPerfDashboard:" & vbCrLf & _
           "Error #" & errNum & ": " & errDesc, vbExclamation, "Perf Dash Runtime Error"
    Resume CleanExit_Perf ' Use Resume to ensure cleanup code runs
End Sub

' Build the data model reading directly from "SQRCT Dashboard" sheet
Public Function BuildPerformanceDataModel_FromDashboard() As Object
    Dim model As Object: Set model = CreateObject("Scripting.Dictionary")
    Dim wsDash As Worksheet
    Dim arrDash As Variant
    Dim dataRange As Range
    Dim lastDashRow As Long
    Dim proceed As Boolean ' Flag to control execution flow
    Dim procName As String: procName = "BuildPerfModel" ' For Logging

    ' --- Use Module_Dashboard.DebugLog for major steps (writes to sheet) ---
    Module_Dashboard.DebugLog procName, "Starting function..."

    ' --- Get the main dashboard sheet ---
    On Error Resume Next ' Enable error trapping *just* for getting the sheet
    Set wsDash = ThisWorkbook.Sheets(Module_Dashboard.DASHBOARD_SHEET_NAME) ' Use constant
    If Err.Number <> 0 Or wsDash Is Nothing Then
        Module_Dashboard.DebugLog procName, "ERROR: Main dashboard sheet '" & Module_Dashboard.DASHBOARD_SHEET_NAME & "' not found or inaccessible. Err: " & Err.Number
        proceed = False ' Don't proceed with reading data
        Err.Clear ' Clear the error
    Else
        Module_Dashboard.DebugLog procName, "Found main dashboard sheet '" & wsDash.Name & "'" ' Log sheet found
        proceed = True
    End If
    On Error GoTo 0 ' Restore default error handling NOW

    If Not proceed Then GoTo AddPerfDefaults ' Jump to default setup if sheet wasn't found

    ' --- Read data from the dashboard sheet (A4:N<lastRow>) ---
    lastDashRow = wsDash.Cells(wsDash.rows.Count, "A").End(xlUp).Row
    Module_Dashboard.DebugLog procName, "Last row on dashboard (Col A) = " & lastDashRow ' Log last row
    If lastDashRow < 4 Then ' Check if there's any data below headers
        Module_Dashboard.DebugLog procName, "No data found on '" & wsDash.Name & "' (A4:N). Building empty model."
        proceed = False ' Don't proceed with processing
    End If

    If Not proceed Then GoTo AddPerfDefaults ' Jump to default setup if no data rows


    ' Try reading the data range
    Module_Dashboard.DebugLog procName, "Attempting to read range A4:" & Module_Dashboard.DB_COL_COMMENTS & lastDashRow ' Log read attempt
    On Error Resume Next
    Set dataRange = wsDash.Range("A4:" & Module_Dashboard.DB_COL_COMMENTS & lastDashRow) ' A4:N<lastRow>
    arrDash = dataRange.Value2 ' Read values
    If Err.Number <> 0 Then
        Module_Dashboard.DebugLog procName, "ERROR reading data from '" & wsDash.Name & "' Range(" & dataRange.Address & "). Err: " & Err.Description ' Log error
        proceed = False ' Don't proceed if read fails
        Err.Clear
    End If
    On Error GoTo 0 ' Restore default handling

    If Not proceed Then GoTo AddPerfDefaults ' Jump to default setup if read fails

    Module_Dashboard.DebugLog procName, "Successfully read data into arrDash." ' Log success

    ' --- Variable Declarations ---
    Dim totalRows As Long, openRows As Long, convRows As Long, decRows As Long, wwomRows As Long
    Dim firstFURows As Long, secondFURows As Long, thirdFURows As Long, pendingRows As Long
    Dim sumOpenAmt As Double, maxOpenAmt As Double, totalConvAmt As Double
    Dim totalCycleDays As Double, cycleCount As Long
    Dim dictStageCount As Object: Set dictStageCount = CreateObject("Scripting.Dictionary")
    Dim dictStageAmount As Object: Set dictStageAmount = CreateObject("Scripting.Dictionary")
    Dim ageBuckets(1 To 4) As Long, ageBucketValues(1 To 4) As Double
    Dim stageCycleTimes As Object: Set stageCycleTimes = CreateObject("Scripting.Dictionary")
    stageCycleTimes.Add "First F/U", Array(0, 9999, 0, 10) ' Avg, Min, Max, Target
    stageCycleTimes.Add "Second F/U", Array(0, 9999, 0, 7)
    stageCycleTimes.Add "Third F/U", Array(0, 9999, 0, 10)
    stageCycleTimes.Add "Pending", Array(0, 9999, 0, 5)
    ' --- End Declarations ---

    ' --- Determine Array Bounds and Total Rows ---
    Dim r As Long, lowerBound As Long, upperBound As Long
    Dim isDataArray2D As Boolean
    totalRows = 0 ' Initialize
    If IsArray(arrDash) Then
        On Error Resume Next ' Handle potential errors getting bounds
        lowerBound = LBound(arrDash, 1)
        upperBound = UBound(arrDash, 1)
        isDataArray2D = (Err.Number = 0)
        If Not isDataArray2D Then
            Err.Clear
            lowerBound = LBound(arrDash) ' Try 1D bounds
            upperBound = UBound(arrDash)
            If Err.Number = 0 Then ' Successfully got 1D bounds
                totalRows = upperBound - lowerBound + 1 ' Calculate total rows for 1D
                If totalRows < 0 Then totalRows = 0 ' Handle empty 1D array edge case
            Else ' Failed to get bounds even for 1D
                 Module_Dashboard.DebugLog procName, "ERROR: Cannot determine bounds of arrDash. Err: " & Err.Description
                 totalRows = 0
                 Err.Clear
            End If
        Else ' Successfully got 2D bounds
            totalRows = upperBound - lowerBound + 1 ' Calculate total rows for 2D
            If totalRows < 0 Then totalRows = 0 ' Handle empty 2D array edge case
        End If
        On Error GoTo 0 ' Restore default handling
    Else
         Module_Dashboard.DebugLog procName, "arrDash is not an array after reading."
    End If
    Module_Dashboard.DebugLog procName, "Determined: Is2D=" & isDataArray2D & ", LowerBound=" & lowerBound & ", UpperBound=" & upperBound & ", TotalRows=" & totalRows ' Log bounds info

    ' --- LOOP START ---
    If totalRows > 0 Then
        Module_Dashboard.DebugLog procName, "Starting data processing loop for " & totalRows & " rows..." ' Log loop start
        ' Define column indices (A=1 to N=14)
        Const IDX_DOCNUM As Long = 1
        Const IDX_AMOUNT As Long = 4
        Const IDX_DOCDATE As Long = 5
        Const IDX_PHASE As Long = 12
        Const IDX_LASTCONTACT As Long = 13
        Dim phase As String
        Dim amt As Double
        Dim docDate As Date
        Dim lastContactDate As Date
        Dim stageDays As Double
        Dim createDate As Date
        Dim convDate As Date
        Dim quoteDate As Date
        Dim quoteAge As Long

        For r = 1 To totalRows ' Loop based on calculated totalRows
            ' --- Use Debug.Print for high-frequency output inside the loop ---
            #If DEBUG_MODE Then
                Debug.Print procName & " --- Processing Row " & r & "/" & totalRows & " ---"
            #End If
            phase = ""
            amt = 0
            docDate = 0
            lastContactDate = 0
            stageDays = 0
            createDate = 0
            convDate = 0
            quoteDate = 0
            quoteAge = -1
            On Error Resume Next ' Handle potential errors within loop for THIS row

            ' --- Read Raw Values ---
            Dim rawPhase As String, rawAmt As String, rawDocDate As String, rawLastContact As String
             If isDataArray2D Then
                 rawPhase = CStr(arrDash(r + lowerBound - 1, IDX_PHASE))
                 rawAmt = CStr(arrDash(r + lowerBound - 1, IDX_AMOUNT))
                 rawDocDate = CStr(arrDash(r + lowerBound - 1, IDX_DOCDATE))
                 rawLastContact = CStr(arrDash(r + lowerBound - 1, IDX_LASTCONTACT))
             Else ' Single row read as 1D array
                 rawPhase = CStr(arrDash(IDX_PHASE + lowerBound - 1))
                 rawAmt = CStr(arrDash(IDX_AMOUNT + lowerBound - 1))
                 rawDocDate = CStr(arrDash(IDX_DOCDATE + lowerBound - 1))
                 rawLastContact = CStr(arrDash(IDX_LASTCONTACT + lowerBound - 1))
             End If
             #If DEBUG_MODE Then
                 Debug.Print procName & " Row " & r & ": Raw Phase='" & rawPhase & "', Amt='" & rawAmt & "', DocDate='" & rawDocDate & "', LastContact='" & rawLastContact & "'"
             #End If

            ' --- Convert Values ---
             phase = Trim$(rawPhase)
             If isNumeric(rawAmt) Then amt = CDbl(rawAmt) Else amt = 0
             If IsDate(rawDocDate) Then docDate = CDate(rawDocDate) Else docDate = 0
             If IsDate(rawLastContact) Then lastContactDate = CDate(rawLastContact) Else lastContactDate = 0
             #If DEBUG_MODE Then
                 Debug.Print procName & " Row " & r & ": Converted phase='" & phase & "', amt=" & amt & ", docDate=" & CStr(docDate) & ", lastContactDate=" & CStr(lastContactDate)
             #End If

             If Err.Number <> 0 Then ' Check if conversion caused error
                  #If DEBUG_MODE Then
                      Debug.Print procName & " Row " & r & ": ERROR during value conversion. Err: " & Err.Description
                  #End If
                  Err.Clear
                  GoTo NextIteration_PerfModel ' Skip rest of processing for this row
             End If

            createDate = docDate
            convDate = lastContactDate
            quoteDate = docDate
            If phase = "" Then phase = "Undefined"

            ' --- Aggregation Logic ---
            #If DEBUG_MODE Then
                Debug.Print procName & " Row " & r & ": Start Aggregation for phase '" & phase & "'"
            #End If

            ' Stage Count & Amount Aggregation
            If Not dictStageCount.Exists(phase) Then
                dictStageCount.Add phase, CLng(0)
                dictStageAmount.Add phase, CDbl(0)
            End If
            dictStageCount(phase) = dictStageCount(phase) + 1
            dictStageAmount(phase) = dictStageAmount(phase) + amt

            ' Cycle Time By Stage Aggregation
            If docDate > 0 And lastContactDate > 0 And lastContactDate >= docDate Then
                stageDays = CDbl(lastContactDate - docDate)
                #If DEBUG_MODE Then
                    Debug.Print procName & " Row " & r & ": Calculated stageDays=" & stageDays
                #End If
                If stageDays >= 0 And stageDays <= 365 Then ' Validate duration
                    #If DEBUG_MODE Then
                        Debug.Print procName & " Row " & r & ": Valid stageDays. Checking phase..."
                    #End If
                    Select Case phase
                        Case "First F/U"
                            #If DEBUG_MODE Then
                                Debug.Print procName & " Row " & r & ": Calling UpdateStageCycleTimes('First F/U', " & stageDays & ")"
                            #End If
                            UpdateStageCycleTimes stageCycleTimes, "First F/U", stageDays
                        Case "Second F/U"
                            #If DEBUG_MODE Then
                                Debug.Print procName & " Row " & r & ": Calling UpdateStageCycleTimes('First F/U', " & stageDays * 0.6 & ") & ('Second F/U', " & stageDays * 0.4 & ")"
                            #End If
                            UpdateStageCycleTimes stageCycleTimes, "First F/U", stageDays * 0.6
                            UpdateStageCycleTimes stageCycleTimes, "Second F/U", stageDays * 0.4
                        Case "Third F/U"
                            #If DEBUG_MODE Then
                                Debug.Print procName & " Row " & r & ": Calling UpdateStageCycleTimes('First F/U', ..), ('Second F/U', ..), ('Third F/U', ..)"
                            #End If
                            UpdateStageCycleTimes stageCycleTimes, "First F/U", stageDays * 0.4
                            UpdateStageCycleTimes stageCycleTimes, "Second F/U", stageDays * 0.3
                            UpdateStageCycleTimes stageCycleTimes, "Third F/U", stageDays * 0.3
                        Case "Pending"
                            #If DEBUG_MODE Then
                                Debug.Print procName & " Row " & r & ": Calling UpdateStageCycleTimes('First F/U', ..), ('Second F/U', ..), ('Third F/U', ..), ('Pending', ..)"
                            #End If
                            UpdateStageCycleTimes stageCycleTimes, "First F/U", stageDays * 0.3
                            UpdateStageCycleTimes stageCycleTimes, "Second F/U", stageDays * 0.3
                            UpdateStageCycleTimes stageCycleTimes, "Third F/U", stageDays * 0.2
                            UpdateStageCycleTimes stageCycleTimes, "Pending", stageDays * 0.2
                        Case "Converted" ' If converted, assume it passed through all stages prior
                            #If DEBUG_MODE Then
                                Debug.Print procName & " Row " & r & ": Calling UpdateStageCycleTimes for all stages (Converted)"
                            #End If
                            UpdateStageCycleTimes stageCycleTimes, "First F/U", stageDays * 0.3
                            UpdateStageCycleTimes stageCycleTimes, "Second F/U", stageDays * 0.25
                            UpdateStageCycleTimes stageCycleTimes, "Third F/U", stageDays * 0.25
                            UpdateStageCycleTimes stageCycleTimes, "Pending", stageDays * 0.2
                    End Select
                End If
            End If

            ' Categorize by Phase & Calculate Overall Cycle Time
            Select Case phase
                Case "Converted"
                    convRows = convRows + 1
                    totalConvAmt = totalConvAmt + amt
                    If createDate > 0 And convDate > 0 And convDate > createDate Then
                        totalCycleDays = totalCycleDays + CDbl(convDate - createDate)
                        cycleCount = cycleCount + 1
                    End If
                Case "Declined"
                    decRows = decRows + 1
                Case "WWOM" ' Assuming closed state
                    wwomRows = wwomRows + 1
                Case "First F/U"
                    firstFURows = firstFURows + 1
                Case "Second F/U"
                    secondFURows = secondFURows + 1
                Case "Third F/U"
                    thirdFURows = thirdFURows + 1
                Case "Pending"
                    pendingRows = pendingRows + 1
            End Select

            ' Accumulate Open Quotes & Value & Aging
            ' Check if the phase is NOT one of the defined closed/terminal states or Undefined
            If phase <> "Converted" And phase <> "Declined" And phase <> "WWOM" And phase <> "Undefined" Then
                openRows = openRows + 1
                sumOpenAmt = sumOpenAmt + amt
                If amt > maxOpenAmt Then maxOpenAmt = amt

                ' --- Aging Calculation for Open Quotes ---
                If quoteDate > 0 Then
                    quoteAge = Date - quoteDate ' Age in days
                    If quoteAge >= 0 Then ' Ensure age is valid
                        Select Case quoteAge
                            Case 0 To 30
                                ageBuckets(1) = ageBuckets(1) + 1
                                ageBucketValues(1) = ageBucketValues(1) + amt
                            Case 31 To 60
                                ageBuckets(2) = ageBuckets(2) + 1
                                ageBucketValues(2) = ageBucketValues(2) + amt
                            Case 61 To 90
                                ageBuckets(3) = ageBuckets(3) + 1
                                ageBucketValues(3) = ageBucketValues(3) + amt
                            Case Else ' Over 90
                                ageBuckets(4) = ageBuckets(4) + 1
                                ageBucketValues(4) = ageBucketValues(4) + amt
                        End Select
                    End If ' quoteAge >= 0
                End If ' quoteDate > 0
            End If ' End Open Quote Check
            ' --- End Aggregation Logic Snippet ---

NextIteration_PerfModel: ' Label to jump to if error occurred in loop
             On Error GoTo 0 ' Ensure default error handling is back on for the next loop iteration
             #If DEBUG_MODE Then
                 Debug.Print procName & " Row " & r & ": Finished Processing." ' Mark end of row
             #End If
         Next r ' Next row in arrDash
         Module_Dashboard.DebugLog procName, "Finished data processing loop." ' Log loop end to sheet
         If IsArray(arrDash) Then Erase arrDash ' Release array memory
     Else
         Module_Dashboard.DebugLog procName, "Skipping processing loop as totalRows = 0." ' Log skip to sheet
         totalRows = 0 ' Ensure totalRows is 0 if loop skipped
     End If
    ' --- End Data Processing ---

AddPerfDefaults:
    Module_Dashboard.DebugLog procName, "Reached AddPerfDefaults section." ' Log milestone to sheet

    ' --- Finalize Stage Times Array ---
    Dim avgStageTimesArray() As Variant
    Dim stageIdx As Long
    Dim stageCount As Long
    Dim stageNameKey As Variant

    stageCount = 0 ' Initialize counter
    ' --- Loop 1: Count actual stages ---
    For Each stageNameKey In stageCycleTimes.Keys
        If Right$(CStr(stageNameKey), 6) <> "_Count" Then stageCount = stageCount + 1
    Next stageNameKey
    Module_Dashboard.DebugLog procName, "Counted " & stageCount & " stages for cycle time array."

    ' --- Check if stages were found ---
    If stageCount > 0 Then
        ReDim avgStageTimesArray(1 To stageCount, 1 To 5) ' Size the output array
        stageIdx = 0
        ' --- Loop 2: Populate the output array ---
        For Each stageNameKey In stageCycleTimes.Keys ' Start Loop 2
            If Right$(CStr(stageNameKey), 6) <> "_Count" Then ' Skip counter keys
                stageIdx = stageIdx + 1
                Dim stageDataArray As Variant
                stageDataArray = stageCycleTimes(stageNameKey) ' Get array(Avg, Min, Max, Target)
                avgStageTimesArray(stageIdx, 1) = stageNameKey ' Col 1: Name
                avgStageTimesArray(stageIdx, 2) = Round(stageDataArray(0), 1) ' Col 2: Avg
                If stageDataArray(1) = 9999 Then avgStageTimesArray(stageIdx, 3) = 0 Else avgStageTimesArray(stageIdx, 3) = stageDataArray(1) ' Col 3: Min
                avgStageTimesArray(stageIdx, 4) = stageDataArray(2) ' Col 4: Max
                avgStageTimesArray(stageIdx, 5) = stageDataArray(3) ' Col 5: Target
            End If
        Next stageNameKey
        Module_Dashboard.DebugLog procName, "Populated avgStageTimesArray."
    Else
        avgStageTimesArray = Array() ' Return empty array if no stage data
        Module_Dashboard.DebugLog procName, "No stage data found, avgStageTimesArray is empty."
    End If

    ' --- Read Pipeline Target ---
    Module_Dashboard.DebugLog procName, "Reading Pipeline Target from Config Sheet..."
    Dim pipelineTargetValue As Double
    Dim wsConfig As Worksheet
    Dim cfgRow As Long
    Dim lastCfgRow As Long
    Dim targetFound As Boolean

    pipelineTargetValue = 1000000 ' Default value if not found or sheet missing
    targetFound = False         ' Initialize flag

    On Error Resume Next ' Temporarily ignore errors finding the sheet
    Set wsConfig = ThisWorkbook.Worksheets(Module_Dashboard.CONFIG_SHEET_NAME) ' Use Public Constant

    ' Check if sheet was found and accessible
    If Err.Number = 0 And Not wsConfig Is Nothing Then
        ' Sheet found, now find the setting
        lastCfgRow = wsConfig.Cells(wsConfig.rows.Count, "A").End(xlUp).Row ' Find last row in Col A

        ' Check if there are any settings below the header row (Row 1)
        If lastCfgRow >= 2 Then
            ' Loop through rows starting from Row 2
            For cfgRow = 2 To lastCfgRow
                ' Check if the setting name in Column A matches
                If Trim$(CStr(wsConfig.Cells(cfgRow, "A").value)) = "Pipeline Target" Then
                    ' Found the setting row, now check if Column B has a numeric value
                    If isNumeric(wsConfig.Cells(cfgRow, "B").value) Then
                        pipelineTargetValue = CDbl(wsConfig.Cells(cfgRow, "B").value) ' Read the value
                        targetFound = True ' Set the flag indicating we found it
                        Module_Dashboard.DebugLog procName, "Found Pipeline Target in Config: " & pipelineTargetValue
                        Exit For ' Stop searching once found
                    Else
                         Module_Dashboard.DebugLog procName, "Pipeline Target setting found, but value in Col B is not numeric."
                    End If ' IsNumeric check
                End If ' Setting name check
            Next cfgRow ' Next row in Config sheet
        End If ' lastCfgRow >= 2 check
        If Not targetFound Then Module_Dashboard.DebugLog procName, "'Pipeline Target' setting not found in Config sheet rows 2:" & lastCfgRow & ". Using default."
    Else
        ' Log if sheet wasn't found (Error occurred or wsConfig is Nothing)
        Module_Dashboard.DebugLog procName, "WARNING: Config sheet '" & Module_Dashboard.CONFIG_SHEET_NAME & "' not found or inaccessible (Err=" & Err.Number & "). Using default Pipeline Target."
    End If ' wsConfig exists check

    Err.Clear ' Clear any potential error from the On Error Resume Next above
    On Error GoTo 0 ' Restore default error handling for the rest of the function
     Module_Dashboard.DebugLog procName, "Using Pipeline Target Value: " & pipelineTargetValue


    ' --- Calculate Derived Metrics ---
    Module_Dashboard.DebugLog procName, "Calculating Derived Metrics..."
    Dim pipelineVsTarget As Double
    Dim avgCycleTime As Double
    Dim conversionRate As Double
    Dim avgOpenAmt As Double ' Added

    ' --- Pipeline vs Target ---
    If pipelineTargetValue = 0 Then
        pipelineVsTarget = 0 ' Avoid division by zero
         Module_Dashboard.DebugLog procName, "Pipeline Target is 0, setting vs Target to 0."
    Else
        pipelineVsTarget = sumOpenAmt / pipelineTargetValue ' Calculate ratio
         Module_Dashboard.DebugLog procName, "Calculated PipelineVsTarget Ratio: " & pipelineVsTarget & " (Value=" & sumOpenAmt & ", Target=" & pipelineTargetValue & ")"
    End If

    ' --- Average cycle time ---
    If cycleCount = 0 Then
        avgCycleTime = 0
    Else
        avgCycleTime = totalCycleDays / cycleCount
    End If
    Module_Dashboard.DebugLog procName, "Calculated AvgCycleTime: " & avgCycleTime & " (Days=" & totalCycleDays & ", Count=" & cycleCount & ")"

    ' --- Conversion rate ---
    If totalRows = 0 Then
        conversionRate = 0
    Else
        conversionRate = convRows / totalRows
    End If
     Module_Dashboard.DebugLog procName, "Calculated ConversionRate: " & conversionRate & " (Converted=" & convRows & ", Total=" & totalRows & ")"

     ' --- Average Open Amount ---
     If openRows = 0 Then
         avgOpenAmt = 0
     Else
         avgOpenAmt = sumOpenAmt / openRows
     End If
      Module_Dashboard.DebugLog procName, "Calculated AvgOpenAmt: " & avgOpenAmt & " (Value=" & sumOpenAmt & ", Open Count=" & openRows & ")"


    ' --- Calculate Historical Comparisons ---
    Module_Dashboard.DebugLog procName, "Calculating Historical Comparisons..."
    Dim histData As Variant ' Declare variable to hold history data
    Dim prevMonthConvRate As Double ' Variable for last month's rate
    Dim prevYearCycleTime As Double ' Variable for last year's time
    Dim convRateMoM As Double       ' Result: Month-over-Month % change
    Dim cycleTimeYoY As Double      ' Result: Year-over-Year % change

    ' Declare loop/processing variables for history check
    Dim histRow As Long
    Dim histDate As Date
    Dim currentMonthDate As Date
    Dim prevMonthDate As Date
    Dim prevYearDate As Date

    ' Initialize comparison values to default/zero
    prevMonthConvRate = 0
    prevYearCycleTime = 0
    convRateMoM = 0
    cycleTimeYoY = 0

    ' Attempt to read the historical data
    histData = ReadPerfHistoricalData() ' Assumes this function exists in this module or is Public

    ' Check if history data is a valid array before processing
    If IsArray(histData) Then
        ' Check array bounds and dimensions (needs at least 7 columns for dates and metrics)
        If UBound(histData, 1) >= LBound(histData, 1) And UBound(histData, 2) >= 7 Then
             Module_Dashboard.DebugLog procName, "Historical data array found (" & UBound(histData, 1) & " rows, " & UBound(histData, 2) & " cols). Processing..."

            ' Calculate reference dates needed for comparison
            currentMonthDate = DateSerial(Year(Date), Month(Date), 1) ' First day of current month
            prevMonthDate = DateAdd("m", -1, currentMonthDate)        ' First day of previous month
            prevYearDate = DateAdd("yyyy", -1, currentMonthDate)      ' First day of same month last year
             Module_Dashboard.DebugLog procName, "Ref Dates: Current=" & format(currentMonthDate, "yyyy-mm") & ", PrevM=" & format(prevMonthDate, "yyyy-mm") & ", PrevY=" & format(prevYearDate, "yyyy-mm")

            ' Loop through historical data backwards to find the most recent matching entries
            For histRow = UBound(histData, 1) To LBound(histData, 1) Step -1

                On Error Resume Next ' Handle potential errors reading/converting data for this specific row
                histDate = 0 ' Reset date variable
                histDate = CDate(histData(histRow, 1)) ' Column 1 = Date in history sheet

                ' Proceed only if date conversion was successful and date is valid
                If Err.Number = 0 And histDate > 0 Then

                    ' Check for previous month's data (only if we haven't found it yet)
                    If prevMonthConvRate = 0 Then ' Optimization: Stop checking once found
                        If Year(histDate) = Year(prevMonthDate) And Month(histDate) = Month(prevMonthDate) Then
                             Module_Dashboard.DebugLog procName, "Found potential Prev Month match at histRow " & histRow
                            ' Found matching month/year, now check metric value
                            If isNumeric(histData(histRow, 7)) Then ' Column 7 = Conversion Rate
                                prevMonthConvRate = CDbl(histData(histRow, 7))
                                 Module_Dashboard.DebugLog procName, "-> Set prevMonthConvRate = " & prevMonthConvRate
                            Else
                                Module_Dashboard.DebugLog procName, "-> Value in Col 7 not numeric."
                            End If
                        End If
                    End If ' End check for previous month rate

                    ' Check for previous year's data (only if we haven't found it yet)
                    If prevYearCycleTime = 0 Then ' Optimization: Stop checking once found
                         If Year(histDate) = Year(prevYearDate) And Month(histDate) = Month(prevYearDate) Then
                             Module_Dashboard.DebugLog procName, "Found potential Prev Year match at histRow " & histRow
                            ' Found matching month/year, now check metric value
                            If isNumeric(histData(histRow, 8)) Then ' Column 8 = Avg Cycle Time
                                prevYearCycleTime = CDbl(histData(histRow, 8))
                                Module_Dashboard.DebugLog procName, "-> Set prevYearCycleTime = " & prevYearCycleTime
                            Else
                                 Module_Dashboard.DebugLog procName, "-> Value in Col 8 not numeric."
                            End If
                        End If
                    End If ' End check for previous year time

                    ' Optimization: Exit loop early if we have found both needed values
                    If prevMonthConvRate <> 0 And prevYearCycleTime <> 0 Then
                        Module_Dashboard.DebugLog procName, "Found both historical metrics. Exiting history loop early."
                        Exit For
                    End If

                End If ' End if date was valid for this row

                Err.Clear ' Clear any error from processing this row before the next iteration
            Next histRow ' Next historical row

            On Error GoTo 0 ' Restore default error handling after the loop finishes

        Else
            Module_Dashboard.DebugLog procName, "Historical data array dimensions invalid (Rows:" & UBound(histData, 1) & ", Cols:" & UBound(histData, 2) & "). Skipping history calc."
        End If ' End check array bounds
    Else
        Module_Dashboard.DebugLog procName, "No valid historical data array returned by ReadPerfHistoricalData. Skipping history calc."
    End If ' End if IsArray

    ' --- Month-over-Month ---
    If prevMonthConvRate <> 0 Then
        convRateMoM = (conversionRate / prevMonthConvRate) - 1
    Else
        convRateMoM = 0 ' Or handle as "N/A" if preferred
    End If
    Module_Dashboard.DebugLog procName, "Calculated ConvRateMoM: " & convRateMoM & " (Current=" & conversionRate & ", PrevM=" & prevMonthConvRate & ")"

    ' --- Year-over-Year ---
    If prevYearCycleTime <> 0 Then
        cycleTimeYoY = (avgCycleTime / prevYearCycleTime) - 1
    Else
        cycleTimeYoY = 0 ' Or handle as "N/A"
    End If
    Module_Dashboard.DebugLog procName, "Calculated CycleTimeYoY: " & cycleTimeYoY & " (Current=" & avgCycleTime & ", PrevY=" & prevYearCycleTime & ")"

    ' --- Populate the Model Dictionary ---
    Module_Dashboard.DebugLog procName, "Populating final model dictionary..."
    With model
        .Add "totalRows", totalRows
        .Add "openRows", openRows
        .Add "convRows", convRows
        .Add "decRows", decRows
        .Add "wwomRows", wwomRows
        .Add "firstFURows", firstFURows
        .Add "secondFURows", secondFURows
        .Add "thirdFURows", thirdFURows
        .Add "pendingRows", pendingRows

        ' Financial metrics
        .Add "sumOpenAmt", sumOpenAmt
        .Add "maxOpenAmt", maxOpenAmt
        .Add "totalConvAmt", totalConvAmt
        .Add "avgOpenAmt", avgOpenAmt ' Added

        ' Derived metrics
        .Add "conversionRate", conversionRate
        .Add "avgCycleTime", Round(avgCycleTime, 1) ' Round for display consistency
        .Add "pipelineTargetValue", pipelineTargetValue ' Added
        .Add "pipelineVsTarget", pipelineVsTarget ' Ratio vs target
        .Add "convRateMoM", convRateMoM
        .Add "cycleTimeYoY", cycleTimeYoY

        ' Array/Dictionary metrics
        .Add "ageBuckets", ageBuckets
        .Add "ageBucketValues", ageBucketValues
        .Add "dictStageCount", dictStageCount
        .Add "dictStageAmount", dictStageAmount
        .Add "avgStageTimes", avgStageTimesArray ' Array(Stage, Avg, Min, Max, Target)

        ' Example additional derived metric
        ' .Add "stageConversionRate", IIf(firstFURows = 0, 0, secondFURows / firstFURows)
    End With
     Module_Dashboard.DebugLog procName, "Model dictionary populated."

    Set BuildPerformanceDataModel_FromDashboard = model

    ' --- Cleanup within Function ---
    Set dictStageCount = Nothing
    Set dictStageAmount = Nothing
    Set stageCycleTimes = Nothing
    Set wsDash = Nothing
    Set dataRange = Nothing
    If IsArray(arrDash) Then Erase arrDash
    If IsArray(histData) Then Erase histData
    If IsArray(avgStageTimesArray) Then Erase avgStageTimesArray

    Module_Dashboard.DebugLog procName, "Exiting function normally." ' Log normal exit
    Exit Function

' Removed GoTo Cleanup pattern here, use Exit Function for normal exit

End Function


' --- Visual Rendering & Chart Functions ---

Private Sub ApplyModernDashboardTheme(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next ' Handle potential errors during formatting
    ws.Cells.ClearFormats ' Clear existing formats first
    ws.Cells.Interior.ColorIndex = xlNone ' Remove all background colors
    With ws
        .Tab.Color = SqrctColors.DarkBlue
        With .Cells.Font
            .Name = FONT_BODY
            .Size = 10
            .Color = SqrctColors.TextDark
            .Bold = False
            .Italic = False
        End With
        .Cells.Interior.Color = SqrctColors.BackgroundGrey ' Apply background grey
        .StandardWidth = 8.43 ' Reset standard width
    End With
    Application.DisplayGridlines = False ' Turn off gridlines for this sheet

    ' Format Title Row (Row 1)
    With ws.Range("A1:Z1") ' Apply across a wide range
        .Merge
        .value = "SQRCT PERFORMANCE DASHBOARD" ' Set title here
        .Font.Name = FONT_HEADER
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = SqrctColors.DarkBlue
        .Font.Color = SqrctColors.TextLight
        .RowHeight = 36
        .WrapText = False
        .Borders.LineStyle = xlNone
    End With
    ' Optional: Can set ws.Range("A1").Value explicitly if needed later

    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ApplyModernDashboardTheme", "WARNING: Error applying theme. Err: " & Err.Description: Err.Clear
    On Error GoTo 0 ' Restore default error handling
End Sub

Private Sub ClearPerfDashboardContents(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim shp As Shape
    Dim chartObj As ChartObject ' Use specific type for charts
    Dim obj As Object         ' Generic object for iteration safety

    Module_Dashboard.DebugLog "ClearPerfDashboardContents", "Clearing shapes and charts from '" & ws.Name & "'..."
    On Error Resume Next ' Ignore errors deleting individual items

    ' Loop backwards when deleting to avoid skipping items
    For Each obj In ws.Shapes
        ' Check specific types if needed, otherwise delete all
        If obj.Name <> "RefreshPerfButton" Then ' Keep the refresh button
            obj.Delete
        End If
    Next obj

    ' Clear Charts separately if Shapes loop doesn't get them
    For Each chartObj In ws.ChartObjects
         chartObj.Delete
    Next chartObj

    ' Clear cell contents below Row 1
    Dim lastRow As Long
    lastRow = 1 ' Default if sheet is empty
    If ws.UsedRange.rows.Count > 1 Then
        lastRow = ws.UsedRange.rows.Count + ws.UsedRange.Row - 1
    End If
    If lastRow < 2 Then lastRow = 2 ' Ensure we clear at least down to row 2

    ws.Range("A2:Z" & lastRow + 50).Clear ' Clear a generous range below title row

    If Err.Number <> 0 Then Module_Dashboard.DebugLog "ClearPerfDashboardContents", "WARNING: Error during clear operation. Err: " & Err.Description: Err.Clear
    Module_Dashboard.DebugLog "ClearPerfDashboardContents", "Finished clearing."
    On Error GoTo 0 ' Restore default error handling
End Sub

Private Sub CreateExecutiveSnapshot(ws As Worksheet, model As Object)
    If ws Is Nothing Or model Is Nothing Then Exit Sub
    Dim startRow As Long: startRow = 3 ' Start below title row
    Dim startCol As Long: startCol = 2 ' Indent a bit
    Dim procName As String: procName = "CreateExecSnapshot"

    Module_Dashboard.DebugLog procName, "Creating section at row " & startRow & ", col " & startCol
    ' --- Section Title ---
    With ws.Cells(startRow, startCol).Resize(1, (CARD_WIDTH * 2) + 1) ' Span across two cards width + margin
        .Merge
        .value = "EXECUTIVE SUMMARY"
        .Font.Name = FONT_TITLE
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = SqrctColors.DarkBlue
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = SqrctColors.BackgroundGrey ' Match sheet background
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = SqrctColors.MediumBlue
        .Borders(xlEdgeBottom).Weight = xlThin
        .RowHeight = 25
    End With

    ' --- Card Positions ---
    Dim cardRow1 As Long: cardRow1 = startRow + 1
    Dim cardRow2 As Long: cardRow2 = cardRow1 + CARD_HEIGHT + 1 ' Add space below first row of cards
    Dim cardCol1 As Long: cardCol1 = startCol
    Dim cardCol2 As Long: cardCol2 = cardCol1 + CARD_WIDTH + 1 ' Add space between cards

    ' --- Get Data Safely ---
    Dim openRowsVal As Long
    Dim sumOpenAmtVal As Double
    Dim convRateVal As Double
    Dim avgCycleTimeVal As Double
    Dim convRateMoMVal As Double
    Dim cycleTimeYoYVal As Double

    On Error Resume Next ' Handle potential errors if keys don't exist in model
    openRowsVal = model("openRows")
    sumOpenAmtVal = model("sumOpenAmt")
    convRateVal = model("conversionRate")
    avgCycleTimeVal = model("avgCycleTime")
    convRateMoMVal = model("convRateMoM")
    cycleTimeYoYVal = model("cycleTimeYoY")
    If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Error reading metric from data model. Err: " & Err.Description: Err.Clear
    On Error GoTo 0

    Module_Dashboard.DebugLog procName, "Creating Metric Cards..."
    ' --- Create Cards ---
    CreateMetricCard ws, cardRow1, cardCol1, "OPEN QUOTES", openRowsVal, "#,##0", "", SqrctColors.MediumBlue, "Total active quotes"
    CreateMetricCard ws, cardRow1, cardCol2, "PIPELINE VALUE", sumOpenAmtVal, "$#,##0", "", SqrctColors.DarkBlue, "Value of open quotes"
    CreateMetricCard ws, cardRow2, cardCol1, "CONVERSION RATE", convRateVal, "0.0%", "", SqrctColors.SuccessGreen, IIf(convRateMoMVal >= 0, "+", "") & format(convRateMoMVal, "0.0%") & " MoM"
    CreateMetricCard ws, cardRow2, cardCol2, "AVG CYCLE TIME", avgCycleTimeVal, "0.0", " days", SqrctColors.LightBlue, IIf(cycleTimeYoYVal >= 0, "+", "") & format(cycleTimeYoYVal, "0.0%") & " YoY"
    Module_Dashboard.DebugLog procName, "Finished creating section."
End Sub

Private Sub CreatePipelineAnalytics(ws As Worksheet, model As Object)
    If ws Is Nothing Or model Is Nothing Then Exit Sub
    Dim startRow As Long
    Dim startCol As Long
    Dim procName As String: procName = "CreatePipelineAnalytics"

    On Error Resume Next ' Find last used row safely
    startRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If Err.Number <> 0 Then startRow = 1 Else startRow = startRow + SECTION_MARGIN ' Add margin below last content
    Err.Clear
    On Error GoTo 0 ' Restore error handling

    startCol = 2 ' Align with section above
    Module_Dashboard.DebugLog procName, "Creating section at row " & startRow & ", col " & startCol

    ' --- Section Title ---
    With ws.Cells(startRow, startCol).Resize(1, (CARD_WIDTH * 2) + 1)
        .Merge
        .value = "PIPELINE ANALYTICS"
        .Font.Name = FONT_TITLE
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = SqrctColors.DarkBlue
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = SqrctColors.BackgroundGrey
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = SqrctColors.MediumBlue
        .Borders(xlEdgeBottom).Weight = xlThin
        .RowHeight = 25
    End With

    ' --- Chart Positions ---
    Dim chartRow As Long: chartRow = startRow + 1
    Dim chartCol1 As Long: chartCol1 = startCol
    Dim chartCol2 As Long: chartCol2 = chartCol1 + CARD_WIDTH + 1
    Module_Dashboard.DebugLog procName, "Creating Pipeline Charts..."

    ' --- Create Charts ---
    CreatePipelineDistributionChart ws, chartRow, chartCol1, model
    CreatePipelineAgeChart ws, chartRow, chartCol2, model
    Module_Dashboard.DebugLog procName, "Finished creating section."
End Sub

Private Sub CreatePerformanceMetrics(ws As Worksheet, model As Object)
    If ws Is Nothing Or model Is Nothing Then Exit Sub
    Dim startRow As Long
    Dim startCol As Long
    Dim procName As String: procName = "CreatePerfMetrics"

    On Error Resume Next ' Find last used row safely
    startRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If Err.Number <> 0 Then startRow = 1 Else startRow = startRow + SECTION_MARGIN
    Err.Clear
    On Error GoTo 0

    startCol = 2
    Module_Dashboard.DebugLog procName, "Creating section at row " & startRow & ", col " & startCol

    ' --- Section Title ---
    With ws.Cells(startRow, startCol).Resize(1, (CARD_WIDTH * 2) + 1)
        .Merge
        .value = "PERFORMANCE METRICS"
        .Font.Name = FONT_TITLE
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = SqrctColors.DarkBlue
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = SqrctColors.BackgroundGrey
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = SqrctColors.MediumBlue
        .Borders(xlEdgeBottom).Weight = xlThin
        .RowHeight = 25
    End With

    ' --- Chart Positions ---
    Dim chartRow As Long: chartRow = startRow + 1
    Dim chartCol1 As Long: chartCol1 = startCol
    Dim chartCol2 As Long: chartCol2 = chartCol1 + CARD_WIDTH + 1
    Module_Dashboard.DebugLog procName, "Creating Performance Charts..."

    ' --- Create Charts ---
    CreateCycleTimeChart ws, chartRow, chartCol1, model
    CreateConversionFunnelChart ws, chartRow, chartCol2, model
    Module_Dashboard.DebugLog procName, "Finished creating section."
End Sub

Private Sub CreateForecastTrends(ws As Worksheet, model As Object)
    If ws Is Nothing Or model Is Nothing Then Exit Sub
    Dim startRow As Long
    Dim startCol As Long
    Dim procName As String: procName = "CreateForecastTrends"

    On Error Resume Next ' Find last used row safely
    startRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If Err.Number <> 0 Then startRow = 1 Else startRow = startRow + SECTION_MARGIN
    Err.Clear
    On Error GoTo 0

    startCol = 2
    Module_Dashboard.DebugLog procName, "Creating section at row " & startRow & ", col " & startCol

    ' --- Section Title ---
    With ws.Cells(startRow, startCol).Resize(1, (CARD_WIDTH * 2) + 1)
        .Merge
        .value = "FORECAST & HISTORICAL TRENDS"
        .Font.Name = FONT_TITLE
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = SqrctColors.DarkBlue
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = SqrctColors.BackgroundGrey
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = SqrctColors.MediumBlue
        .Borders(xlEdgeBottom).Weight = xlThin
        .RowHeight = 25
    End With

    ' --- Card/Chart Positions ---
    Dim itemRow As Long: itemRow = startRow + 1
    Dim itemCol1 As Long: itemCol1 = startCol
    Dim itemCol2 As Long: itemCol2 = itemCol1 + CARD_WIDTH + 1
    Module_Dashboard.DebugLog procName, "Creating Target Card and History Chart..."

    ' --- Create Items ---
    CreatePipelineTargetCard ws, itemRow, itemCol1, model ' Keep as card
    CreateHistoricalTrendsChart ws, itemRow, itemCol2, model
    Module_Dashboard.DebugLog procName, "Finished creating section."
End Sub


' --- CHARTING FUNCTIONS ---

Private Sub CreatePipelineDistributionChart(ws As Worksheet, topRow As Long, leftCol As Long, model As Object)
    Dim chtObj As ChartObject, cht As Chart
    Dim stages(1 To 4) As String
    Dim counts(1 To 4) As Long
    Dim chartLeft As Double, chartTop As Double, chartWidth As Double, chartHeight As Double
    Dim tempSheet As Worksheet, tempRange As Range
    Dim procName As String: procName = "CreatePipelineDistChart"

    On Error GoTo ChartErrorHandler_Perf
    Module_Dashboard.DebugLog procName, "Starting chart creation..."

    ' --- Get Data Safely ---
    On Error Resume Next
    stages(1) = "First F/U":  counts(1) = model("firstFURows")
    stages(2) = "Second F/U": counts(2) = model("secondFURows")
    stages(3) = "Third F/U":  counts(3) = model("thirdFURows")
    stages(4) = "Pending":    counts(4) = model("pendingRows")
    If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Error reading stage counts from model. Err: " & Err.Description: Err.Clear
    On Error GoTo ChartErrorHandler_Perf

    ' --- Calculate Chart Position ---
    chartLeft = ws.Cells(topRow, leftCol).left
    chartTop = ws.Cells(topRow, leftCol).top
    chartWidth = ws.Cells(topRow, leftCol).Resize(1, CARD_WIDTH).width
    chartHeight = ws.Cells(topRow, leftCol).Resize(CARD_HEIGHT, 1).height
    Module_Dashboard.DebugLog procName, "Calculated chart position: L=" & chartLeft & ", T=" & chartTop & ", W=" & chartWidth & ", H=" & chartHeight

    ' --- Prepare Data on Temp Sheet ---
    Set tempSheet = Module_Dashboard.GetOrCreateTempSheet("PerfChartDataTemp") ' Use shared helper
    If tempSheet Is Nothing Then Module_Dashboard.DebugLog procName, "FATAL: Failed to get temp sheet.": Exit Sub
    Set tempRange = tempSheet.Range("A1").Resize(4, 2)
    tempRange.Columns(1).value = Application.Transpose(stages)
    tempRange.Columns(2).value = Application.Transpose(counts)
    Module_Dashboard.DebugLog procName, "Data written to temp sheet range " & tempRange.Address(False, False)

    ' --- Create Chart Object ---
    Set chtObj = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
    Set cht = chtObj.Chart
    Module_Dashboard.DebugLog procName, "Chart object added."

    ' --- Format Chart ---
    With cht
        .SetSourceData Source:=tempRange
        .ChartType = xlPie
        .hasTitle = True
        .ChartTitle.Text = "Open Quotes by Stage"
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Name = FONT_TITLE
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .PlotArea.format.Fill.Visible = msoFalse
        .ChartArea.format.Fill.Visible = msoFalse
        .ChartArea.format.Line.Visible = msoFalse

        ' Apply formatting safely
        On Error Resume Next
        If .SeriesCollection.Count > 0 Then
            With .SeriesCollection(1)
                .Points(1).format.Fill.ForeColor.RGB = SqrctColors.LightBlue
                .Points(2).format.Fill.ForeColor.RGB = SqrctColors.MediumBlue
                .Points(3).format.Fill.ForeColor.RGB = SqrctColors.DarkBlue
                .Points(4).format.Fill.ForeColor.RGB = SqrctColors.WarningOrange
                .ApplyDataLabels
                .DataLabels.format.TextFrame2.TextRange.Font.Size = 8
                .DataLabels.NumberFormat = "#,##0" ' Format labels as numbers
            End With
        End If
        If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Error formatting chart series/labels. Err: " & Err.Description: Err.Clear
        On Error GoTo ChartErrorHandler_Perf ' Restore specific handler
    End With
    Module_Dashboard.DebugLog procName, "Chart formatting applied."

    ' --- Cleanup Temp Data ---
    tempRange.ClearContents
    Module_Dashboard.DebugLog procName, "Temp data cleared."

ChartExit_Perf:
    On Error Resume Next ' Final cleanup safety
    Set cht = Nothing: Set chtObj = Nothing: Set tempRange = Nothing: Set tempSheet = Nothing
    Module_Dashboard.DebugLog procName, "Finished chart creation."
    Exit Sub

ChartErrorHandler_Perf:
    Module_Dashboard.DebugLog procName, "ERROR [" & Err.Number & "]: " & Err.Description & " (Line: " & Erl & ")"
    Resume ChartExit_Perf ' Go to cleanup on error
End Sub

Private Sub CreatePipelineAgeChart(ws As Worksheet, topRow As Long, leftCol As Long, model As Object)
    Dim chtObj As ChartObject, cht As Chart
    Dim categories(1 To 4) As String
    Dim counts As Variant ' Array from model
    Dim chartLeft As Double, chartTop As Double, chartWidth As Double, chartHeight As Double
    Dim tempSheet As Worksheet, tempRange As Range
    Dim procName As String: procName = "CreatePipelineAgeChart"

    On Error GoTo ChartErrorHandler_Perf
    Module_Dashboard.DebugLog procName, "Starting chart creation..."

    ' --- Get Data Safely ---
    On Error Resume Next
    categories(1) = "0-30d": categories(2) = "31-60d": categories(3) = "61-90d": categories(4) = "90d+"
    counts = model("ageBuckets")
    If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Error reading ageBuckets from model. Err: " & Err.Description: Err.Clear: GoTo ChartExit_Perf ' Exit if data missing
    If Not IsArray(counts) Then Module_Dashboard.DebugLog procName, "WARNING: ageBuckets is not an array.": GoTo ChartExit_Perf
    On Error GoTo ChartErrorHandler_Perf

    ' --- Calculate Chart Position ---
    chartLeft = ws.Cells(topRow, leftCol).left
    chartTop = ws.Cells(topRow, leftCol).top
    chartWidth = ws.Cells(topRow, leftCol).Resize(1, CARD_WIDTH).width
    chartHeight = ws.Cells(topRow, leftCol).Resize(CARD_HEIGHT, 1).height
    Module_Dashboard.DebugLog procName, "Calculated chart position."

    ' --- Prepare Data on Temp Sheet ---
    Set tempSheet = Module_Dashboard.GetOrCreateTempSheet("PerfChartDataTemp")
    If tempSheet Is Nothing Then Module_Dashboard.DebugLog procName, "FATAL: Failed to get temp sheet.": Exit Sub
    Set tempRange = tempSheet.Range("C1").Resize(4, 2) ' Use different range on temp sheet
    tempRange.Columns(1).value = Application.Transpose(categories)
    tempRange.Columns(2).value = Application.Transpose(counts)
    Module_Dashboard.DebugLog procName, "Data written to temp sheet range " & tempRange.Address(False, False)

    ' --- Create Chart Object ---
    Set chtObj = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
    Set cht = chtObj.Chart
    Module_Dashboard.DebugLog procName, "Chart object added."

    ' --- Format Chart ---
    With cht
        .SetSourceData Source:=tempRange
        .ChartType = xlColumnClustered
        .hasTitle = True
        .ChartTitle.Text = "Open Quotes by Age"
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Name = FONT_TITLE
        .HasLegend = False
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorGridlines.Delete
        .Axes(xlValue).DisplayUnit = xlNone ' Ensure no units displayed (like Thousands)
        .Axes(xlValue).HasDisplayUnitLabel = False
        .PlotArea.format.Fill.Visible = msoFalse
        .ChartArea.format.Fill.Visible = msoFalse
        .ChartArea.format.Line.Visible = msoFalse

        ' Apply formatting safely
        On Error Resume Next
        If .SeriesCollection.Count > 0 Then
            With .SeriesCollection(1)
                .format.Fill.ForeColor.RGB = SqrctColors.MediumBlue
                .Points(4).format.Fill.ForeColor.RGB = SqrctColors.WarningOrange ' Highlight >90d
                .ApplyDataLabels
                .DataLabels.format.TextFrame2.TextRange.Font.Size = 8
                .DataLabels.NumberFormat = "#,##0"
                .DataLabels.Position = xlLabelPositionOutsideEnd ' Position outside bars
            End With
        End If
        If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Error formatting chart series/labels. Err: " & Err.Description: Err.Clear
        On Error GoTo ChartErrorHandler_Perf ' Restore specific handler
    End With
    Module_Dashboard.DebugLog procName, "Chart formatting applied."

    ' --- Cleanup Temp Data ---
    tempRange.ClearContents
    Module_Dashboard.DebugLog procName, "Temp data cleared."

ChartExit_Perf:
    On Error Resume Next
    Set cht = Nothing: Set chtObj = Nothing: Set tempRange = Nothing: Set tempSheet = Nothing
    Module_Dashboard.DebugLog procName, "Finished chart creation."
    Exit Sub

ChartErrorHandler_Perf:
    Module_Dashboard.DebugLog procName, "ERROR [" & Err.Number & "]: " & Err.Description & " (Line: " & Erl & ")"
    Resume ChartExit_Perf
End Sub

Private Sub CreateCycleTimeChart(ws As Worksheet, topRow As Long, leftCol As Long, model As Object)
    Dim avgStageTimes As Variant
    Dim chtObj As ChartObject, cht As Chart
    Dim i As Long, rowCount As Long
    Dim chartLeft As Double, chartTop As Double, chartWidth As Double, chartHeight As Double
    Dim tempSheet As Worksheet, tempRange As Range
    Dim procName As String: procName = "CreateCycleTimeChart"

    On Error GoTo ChartErrorHandler_Perf
    Module_Dashboard.DebugLog procName, "Starting chart creation..."

    ' --- Get Data Safely ---
    On Error Resume Next
    avgStageTimes = model("avgStageTimes") ' Expects Array(Stage, Avg, Min, Max, Target)
    If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Error reading avgStageTimes from model. Err: " & Err.Description: Err.Clear: GoTo ChartExit_Perf
    If Not IsArray(avgStageTimes) Then Module_Dashboard.DebugLog procName, "WARNING: avgStageTimes is not an array.": GoTo ChartExit_Perf
    rowCount = 0
    On Error Resume Next
    rowCount = UBound(avgStageTimes, 1)
    If Err.Number <> 0 Or rowCount < LBound(avgStageTimes, 1) Then Module_Dashboard.DebugLog procName, "WARNING: avgStageTimes array is empty or invalid bounds.": Err.Clear: GoTo ChartExit_Perf
    On Error GoTo ChartErrorHandler_Perf
    Module_Dashboard.DebugLog procName, "Retrieved avgStageTimes array with " & rowCount & " rows."

    ' --- Calculate Chart Position ---
    chartLeft = ws.Cells(topRow, leftCol).left
    chartTop = ws.Cells(topRow, leftCol).top
    chartWidth = ws.Cells(topRow, leftCol).Resize(1, CARD_WIDTH).width
    chartHeight = ws.Cells(topRow, leftCol).Resize(CARD_HEIGHT, 1).height
    Module_Dashboard.DebugLog procName, "Calculated chart position."

    ' --- Prepare Data on Temp Sheet ---
    Set tempSheet = Module_Dashboard.GetOrCreateTempSheet("PerfChartDataTemp")
    If tempSheet Is Nothing Then Module_Dashboard.DebugLog procName, "FATAL: Failed to get temp sheet.": Exit Sub
    Set tempRange = tempSheet.Range("E1").Resize(rowCount + 1, 3) ' Use different range
    tempRange.rows(1).value = Array("Stage", "Avg Days", "Target Days") ' Header row
    For i = 1 To rowCount
        tempRange.Cells(i + 1, 1).value = avgStageTimes(i, 1) ' Stage Name (Index 1)
        tempRange.Cells(i + 1, 2).value = avgStageTimes(i, 2) ' Avg Days (Index 2)
        tempRange.Cells(i + 1, 3).value = avgStageTimes(i, 5) ' Target Days (Index 5)
    Next i
    Module_Dashboard.DebugLog procName, "Data written to temp sheet range " & tempRange.Address(False, False)

    ' --- Create Chart Object ---
    Set chtObj = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
    Set cht = chtObj.Chart
    Module_Dashboard.DebugLog procName, "Chart object added."

    ' --- Format Chart ---
    With cht
        .SetSourceData Source:=tempRange
        .ChartType = xlColumnClustered
        .hasTitle = True
        .ChartTitle.Text = "Avg vs Target Cycle Time (Days)"
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Name = FONT_TITLE
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorGridlines.Delete
        .PlotArea.format.Fill.Visible = msoFalse
        .ChartArea.format.Fill.Visible = msoFalse
        .ChartArea.format.Line.Visible = msoFalse

        ' Apply formatting safely
        On Error Resume Next
        If .SeriesCollection.Count >= 2 Then ' Need at least 2 series
            .SeriesCollection(1).Name = "Avg Days" ' Assumes Series 1 is Avg
            .SeriesCollection(1).format.Fill.ForeColor.RGB = SqrctColors.MediumBlue
            .SeriesCollection(2).Name = "Target Days" ' Assumes Series 2 is Target
            .SeriesCollection(2).format.Fill.ForeColor.RGB = SqrctColors.ErrorRed
            ' Optional: Add Data Labels
            ' .SeriesCollection(1).ApplyDataLabels
            ' .SeriesCollection(1).DataLabels.NumberFormat = "0.0"
            ' .SeriesCollection(1).DataLabels.Font.Size = 8
        End If
        If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Error formatting chart series. Err: " & Err.Description: Err.Clear
        On Error GoTo ChartErrorHandler_Perf ' Restore specific handler
    End With
    Module_Dashboard.DebugLog procName, "Chart formatting applied."

    ' --- Cleanup Temp Data ---
    tempRange.ClearContents
    Module_Dashboard.DebugLog procName, "Temp data cleared."

ChartExit_Perf:
    On Error Resume Next
    Set cht = Nothing: Set chtObj = Nothing: Set tempRange = Nothing: Set tempSheet = Nothing
    Module_Dashboard.DebugLog procName, "Finished chart creation."
    Exit Sub

ChartErrorHandler_Perf:
    Module_Dashboard.DebugLog procName, "ERROR [" & Err.Number & "]: " & Err.Description & " (Line: " & Erl & ")"
    Resume ChartExit_Perf
End Sub

Private Sub CreateConversionFunnelChart(ws As Worksheet, topRow As Long, leftCol As Long, model As Object)
    Dim chtObj As ChartObject, cht As Chart
    Dim stages(1 To 5) As String
    Dim counts(1 To 5) As Long
    Dim chartLeft As Double, chartTop As Double, chartWidth As Double, chartHeight As Double
    Dim tempSheet As Worksheet, tempRange As Range
    Dim procName As String: procName = "CreateConvFunnelChart"

    On Error GoTo ChartErrorHandler_Perf
    Module_Dashboard.DebugLog procName, "Starting chart creation..."

    ' --- Get Data Safely ---
    On Error Resume Next
    stages(1) = "1. First F/U": counts(1) = model("firstFURows")
    stages(2) = "2. Second F/U": counts(2) = model("secondFURows")
    stages(3) = "3. Third F/U": counts(3) = model("thirdFURows")
    stages(4) = "4. Pending": counts(4) = model("pendingRows")
    stages(5) = "5. Converted": counts(5) = model("convRows")
    If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Error reading funnel counts from model. Err: " & Err.Description: Err.Clear: GoTo ChartExit_Perf
    On Error GoTo ChartErrorHandler_Perf

    ' --- Calculate Chart Position ---
    chartLeft = ws.Cells(topRow, leftCol).left
    chartTop = ws.Cells(topRow, leftCol).top
    chartWidth = ws.Cells(topRow, leftCol).Resize(1, CARD_WIDTH).width
    chartHeight = ws.Cells(topRow, leftCol).Resize(CARD_HEIGHT, 1).height
    Module_Dashboard.DebugLog procName, "Calculated chart position."

    ' --- Prepare Data on Temp Sheet ---
    Set tempSheet = Module_Dashboard.GetOrCreateTempSheet("PerfChartDataTemp")
    If tempSheet Is Nothing Then Module_Dashboard.DebugLog procName, "FATAL: Failed to get temp sheet.": Exit Sub
    Set tempRange = tempSheet.Range("G1").Resize(5, 2) ' Use different range
    tempRange.Columns(1).value = Application.Transpose(stages)
    tempRange.Columns(2).value = Application.Transpose(counts)
    Module_Dashboard.DebugLog procName, "Data written to temp sheet range " & tempRange.Address(False, False)

    ' --- Create Chart Object ---
    Set chtObj = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
    Set cht = chtObj.Chart
    Module_Dashboard.DebugLog procName, "Chart object added."

    ' --- Format Chart ---
    With cht
        .SetSourceData Source:=tempRange
        .ChartType = xlBarClustered ' Use Bar chart for funnel stages
        .hasTitle = True
        .ChartTitle.Text = "Conversion Funnel (by Stage Count)"
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Name = FONT_TITLE
        .HasLegend = False
        .Axes(xlCategory).ReversePlotOrder = True ' Start funnel from top
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorGridlines.Delete
        .PlotArea.format.Fill.Visible = msoFalse
        .ChartArea.format.Fill.Visible = msoFalse
        .ChartArea.format.Line.Visible = msoFalse

        ' Apply formatting safely
        On Error Resume Next
        If .SeriesCollection.Count > 0 Then
            With .SeriesCollection(1)
                .format.Fill.ForeColor.RGB = SqrctColors.SuccessGreen
                .ApplyDataLabels
                .DataLabels.Position = xlLabelPositionInsideEnd
                .DataLabels.format.TextFrame2.TextRange.Font.Size = 8
                .DataLabels.NumberFormat = "#,##0"
            End With
        End If
        If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Error formatting chart series/labels. Err: " & Err.Description: Err.Clear
        On Error GoTo ChartErrorHandler_Perf ' Restore specific handler
    End With
    Module_Dashboard.DebugLog procName, "Chart formatting applied."

    ' --- Cleanup Temp Data ---
    tempRange.ClearContents
    Module_Dashboard.DebugLog procName, "Temp data cleared."

ChartExit_Perf:
    On Error Resume Next
    Set cht = Nothing: Set chtObj = Nothing: Set tempRange = Nothing: Set tempSheet = Nothing
    Module_Dashboard.DebugLog procName, "Finished chart creation."
    Exit Sub

ChartErrorHandler_Perf:
    Module_Dashboard.DebugLog procName, "ERROR [" & Err.Number & "]: " & Err.Description & " (Line: " & Erl & ")"
    Resume ChartExit_Perf
End Sub

Private Sub CreateHistoricalTrendsChart(ws As Worksheet, topRow As Long, leftCol As Long, model As Object)
    Dim histWS As Worksheet
    Dim dateCol As Range, valueCol As Range, rateCol As Range ' Add more series if needed
    Dim chtObj As ChartObject, cht As Chart
    Dim lastHistRow As Long
    Dim chartLeft As Double, chartTop As Double, chartWidth As Double, chartHeight As Double
    Dim procName As String: procName = "CreateHistTrendsChart"

    On Error GoTo ChartErrorHandler_Perf
    Module_Dashboard.DebugLog procName, "Starting chart creation..."

    ' --- Get History Sheet ---
    On Error Resume Next
    Set histWS = ThisWorkbook.Sheets(Module_Dashboard.PERF_HISTORY_SHEET_NAME) ' Use constant from Module_Dashboard
    If histWS Is Nothing Then
        Module_Dashboard.DebugLog procName, "INFO: Cannot create historical chart: '" & Module_Dashboard.PERF_HISTORY_SHEET_NAME & "' sheet not found."
        MsgBox "Cannot create historical chart: '" & Module_Dashboard.PERF_HISTORY_SHEET_NAME & "' sheet not found.", vbInformation
        GoTo ChartExit_Perf
    End If
    On Error GoTo ChartErrorHandler_Perf

    ' --- Check for Data on History Sheet ---
    lastHistRow = histWS.Cells(histWS.rows.Count, "A").End(xlUp).Row
    If lastHistRow <= 1 Then
        Module_Dashboard.DebugLog procName, "INFO: Cannot create historical chart: No data found on '" & Module_Dashboard.PERF_HISTORY_SHEET_NAME & "' sheet."
        MsgBox "Cannot create historical chart: No data found on '" & Module_Dashboard.PERF_HISTORY_SHEET_NAME & "' sheet.", vbInformation
        GoTo ChartExit_Perf
    End If
    Module_Dashboard.DebugLog procName, "Found " & lastHistRow - 1 & " data rows on history sheet."

    ' --- Define Data Ranges (Adjust column letters/indices as per your history sheet) ---
    Set dateCol = histWS.Range("A2:A" & lastHistRow)    ' Column 1 = Timestamp
    Set valueCol = histWS.Range("F2:F" & lastHistRow)   ' Column 6 = Pipeline Value
    Set rateCol = histWS.Range("G2:G" & lastHistRow)    ' Column 7 = Conversion Rate
    Module_Dashboard.DebugLog procName, "History Ranges: Date=" & dateCol.Address(False, False) & ", Value=" & valueCol.Address(False, False) & ", Rate=" & rateCol.Address(False, False)

    ' --- Calculate Chart Position ---
    chartLeft = ws.Cells(topRow, leftCol).left
    chartTop = ws.Cells(topRow, leftCol).top
    chartWidth = ws.Cells(topRow, leftCol).Resize(1, CARD_WIDTH).width
    chartHeight = ws.Cells(topRow, leftCol).Resize(CARD_HEIGHT, 1).height
    Module_Dashboard.DebugLog procName, "Calculated chart position."

    ' --- Create Chart Object ---
    Set chtObj = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
    Set cht = chtObj.Chart
    Module_Dashboard.DebugLog procName, "Chart object added."

    ' --- Format Chart ---
    With cht
        .ChartType = xlLineMarkers ' Line chart with markers
        .hasTitle = True
        .ChartTitle.Text = "Historical Trends"
        .ChartTitle.Font.Size = 10
        .ChartTitle.Font.Name = FONT_TITLE
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8

        ' Clear existing series before adding new ones
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        ' Add Pipeline Value Series (Primary Axis)
        With .SeriesCollection.NewSeries
            .Name = "Pipeline Value"
            .XValues = dateCol
            .Values = valueCol
            .AxisGroup = xlPrimary
            .format.Line.ForeColor.RGB = SqrctColors.DarkBlue
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
        End With

        ' Add Conversion Rate Series (Secondary Axis)
        With .SeriesCollection.NewSeries
            .Name = "Conv. Rate"
            .XValues = dateCol
            .Values = rateCol
            .AxisGroup = xlSecondary ' Put on secondary axis
            .format.Line.ForeColor.RGB = SqrctColors.SuccessGreen
            .MarkerStyle = xlMarkerStyleSquare
            .MarkerSize = 5
        End With

        ' Format Axes
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        .Axes(xlValue, xlPrimary).TickLabels.Font.Size = 8
        .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "$#,##0" ' Currency format
        .Axes(xlValue, xlPrimary).MajorGridlines.Delete
        .Axes(xlValue, xlSecondary).TickLabels.Font.Size = 8
        .Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0.0%" ' Percentage format
        .Axes(xlValue, xlSecondary).MajorGridlines.Delete ' Remove gridlines for secondary axis too

        ' Format Plot/Chart Area
        .PlotArea.format.Fill.Visible = msoFalse
        .ChartArea.format.Fill.Visible = msoFalse
        .ChartArea.format.Line.Visible = msoFalse
    End With
    Module_Dashboard.DebugLog procName, "Chart formatting applied."


ChartExit_Perf:
    On Error Resume Next
    Set cht = Nothing: Set chtObj = Nothing: Set histWS = Nothing
    Set dateCol = Nothing: Set valueCol = Nothing: Set rateCol = Nothing
     Module_Dashboard.DebugLog procName, "Finished chart creation."
    Exit Sub
ChartErrorHandler_Perf:
     Module_Dashboard.DebugLog procName, "ERROR [" & Err.Number & "]: " & Err.Description & " (Line: " & Erl & ")"
     Resume ChartExit_Perf
End Sub


' --- CARD & CONTROL FUNCTIONS ---

Private Sub CreateMetricCard(ws As Worksheet, topRow As Long, leftCol As Long, title As String, value As Variant, valueFormat As String, suffix As String, titleColor As Long, subtitle As String)
    If ws Is Nothing Then Exit Sub
    Dim cardWidthPixels As Double, cardHeightPixels As Double
    Dim cardLeft As Double, cardTop As Double
    Dim cardName As String: cardName = "Card_" & Replace(Replace(title, " ", "_"), "/", "_") ' Make name valid
    Dim txtBoxName As String: txtBoxName = "Txt_" & Replace(Replace(title, " ", "_"), "/", "_")
    Dim procName As String: procName = "CreateMetricCard_" & cardName

    Module_Dashboard.DebugLog procName, "Creating card: '" & title & "'"
    ' --- Calculate Position & Size ---
    On Error Resume Next ' Handle errors if cells are hidden/invalid
    cardLeft = ws.Cells(topRow, leftCol).left
    cardTop = ws.Cells(topRow, leftCol).top
    cardWidthPixels = ws.Cells(topRow, leftCol).Resize(1, CARD_WIDTH).width
    cardHeightPixels = ws.Cells(topRow, leftCol).Resize(CARD_HEIGHT, 1).height
    If Err.Number <> 0 Or cardWidthPixels <= 0 Or cardHeightPixels <= 0 Then
        Module_Dashboard.DebugLog procName, "WARNING: Error calculating card dimensions. Using defaults. Err: " & Err.Description
        cardWidthPixels = CARD_WIDTH * 60 ' Fallback width
        cardHeightPixels = CARD_HEIGHT * 15 ' Fallback height
        If cardLeft <= 0 Then cardLeft = 10 ' Fallback position
        If cardTop <= 0 Then cardTop = 10
        Err.Clear
    End If
    On Error GoTo 0 ' Restore default handling

     Module_Dashboard.DebugLog procName, "Card Pos: L=" & cardLeft & ", T=" & cardTop & ", W=" & cardWidthPixels & ", H=" & cardHeightPixels

    ' --- Delete Existing Shapes with Same Name ---
    On Error Resume Next
    ws.Shapes(cardName).Delete
    ws.Shapes(txtBoxName).Delete
    If Err.Number <> 0 Then Err.Clear ' Ignore error if shapes didn't exist
    On Error GoTo CardErrorHandler_Perf ' Use specific handler

    ' --- Create Card Shape ---
    Dim card As Shape
    Set card = ws.Shapes.AddShape(msoShapeRoundedRectangle, cardLeft, cardTop, cardWidthPixels, cardHeightPixels)
    card.Name = cardName
    With card
        .Fill.Visible = msoTrue ' Ensure fill is on
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255) ' White fill
        .Line.ForeColor.RGB = SqrctColors.BorderGrey
        .Line.Weight = 0.75
        .Line.Visible = msoTrue ' Ensure line is on
        .Shadow.Type = msoShadow21 ' Bottom right offset shadow
        .Shadow.Visible = msoTrue
        .Shadow.Style = msoShadowStyleOuterShadow   '? correct
        .Shadow.Blur = 5
        .Shadow.OffsetX = 2
        .Shadow.OffsetY = 2
        .Shadow.Transparency = 0.7
        .Adjustments(1) = 0.1 ' Adjust corner rounding (0=sharp, 0.5=fully rounded)
    End With
     Module_Dashboard.DebugLog procName, "Card shape created."

    ' --- Create Text Box ---
    Dim txtBox As Shape
    Dim padding As Single: padding = 5 ' Padding inside card
    Set txtBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, cardLeft + padding, cardTop + padding, cardWidthPixels - (padding * 2), cardHeightPixels - (padding * 2))
    txtBox.Name = txtBoxName
    With txtBox
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoTrue
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter

            ' Format Value Safely
            Dim displayValue As String
            On Error Resume Next
            If isNumeric(value) Then
                displayValue = format(value, valueFormat) & suffix
            ElseIf IsDate(value) Then
                displayValue = format(value, valueFormat) & suffix
            ElseIf IsEmpty(value) Or IsNull(value) Then
                displayValue = "N/A" & suffix
            Else
                displayValue = CStr(value) & suffix
            End If
            If Err.Number <> 0 Then
                 Module_Dashboard.DebugLog procName, "WARNING: Error formatting value for card '" & title & "'. Value='" & value & "', Format='" & valueFormat & "'. Err: " & Err.Description
                 displayValue = "Error!" & suffix
                 Err.Clear
            End If
            On Error GoTo CardErrorHandler_Perf ' Restore handler

            ' Set Text and Format Paragraphs
            .TextRange.Text = UCase(title) & vbCrLf & displayValue & vbCrLf & subtitle
            If .TextRange.Paragraphs.Count >= 3 Then ' Check paragraph count before formatting
                With .TextRange.Paragraphs(1).Font ' Title
                    .Name = FONT_TITLE
                    .Size = 10
                    .Bold = msoTrue
                    .Fill.Visible = msoTrue: .Fill.Solid: .Fill.ForeColor.RGB = titleColor
                End With
                With .TextRange.Paragraphs(2).Font ' Value
                    .Name = FONT_NUMBER
                    .Size = 22
                    .Bold = msoTrue
                     .Fill.Visible = msoTrue: .Fill.Solid: .Fill.ForeColor.RGB = SqrctColors.TextDark
                End With
                With .TextRange.Paragraphs(3).Font ' Subtitle
                    .Name = FONT_BODY
                    .Size = 9
                    .Italic = msoTrue
                    .Fill.Visible = msoTrue: .Fill.Solid: .Fill.ForeColor.RGB = SqrctColors.TextDark
                End With
            Else
                 Module_Dashboard.DebugLog procName, "WARNING: Textbox for card '" & title & "' has less than 3 paragraphs. Formatting skipped."
            End If
        End With
    End With
     Module_Dashboard.DebugLog procName, "Textbox created and formatted."

CardExit_Perf:
    On Error Resume Next
    Set card = Nothing
    Set txtBox = Nothing
    Exit Sub
CardErrorHandler_Perf:
    Module_Dashboard.DebugLog procName, "ERROR [" & Err.Number & "]: " & Err.Description & " (Line: " & Erl & ")"
    Resume CardExit_Perf ' Go to cleanup on error
End Sub


Private Sub CreatePipelineTargetCard(ws As Worksheet, topRow As Long, leftCol As Long, model As Object)
    Dim pipelineRatio As Double
    Dim pipelineVal As Double
    Dim titleColor As Long
    Dim procName As String: procName = "CreatePipelineTargetCard"

    Module_Dashboard.DebugLog procName, "Creating card..."
    ' --- Get Data Safely ---
    On Error Resume Next
    pipelineVal = model("sumOpenAmt")
    pipelineRatio = model("pipelineVsTarget") ' This is the Ratio
    If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Error reading target metrics from model. Err: " & Err.Description: Err.Clear
    On Error GoTo 0

    ' --- Determine Color ---
    If pipelineRatio >= 1 Then
        titleColor = SqrctColors.SuccessGreen ' Green if >= 100%
    ElseIf pipelineRatio >= 0.8 Then
        titleColor = SqrctColors.WarningOrange ' Orange if 80-99%
    Else
        titleColor = SqrctColors.ErrorRed ' Red if < 80%
    End If

    ' --- Create Card ---
    CreateMetricCard ws, topRow, leftCol, "PIPELINE vs TARGET", pipelineRatio, "0%", "", titleColor, "Current Value: " & format(pipelineVal, "$#,##0")
    Module_Dashboard.DebugLog procName, "Finished creating card."
End Sub

Private Sub AddDashboardControls(ws As Worksheet)
    ' Adds timestamp below content
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long
    Dim timeStampCell As Range
    Dim procName As String: procName = "AddDashboardControls"

    Module_Dashboard.DebugLog procName, "Adding controls/timestamp..."
    On Error Resume Next ' Find last used row safely
    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If Err.Number <> 0 Then lastRow = 1 ' Default if sheet is empty or find fails
    Err.Clear
    On Error GoTo 0

    Set timeStampCell = ws.Cells(lastRow + 2, 2) ' Add 2 rows below last content, in column B
    With timeStampCell
        .value = "Last Updated: " & format(Now, "mm/dd/yyyy hh:nn AM/PM")
        .ClearFormats ' Clear previous formats first
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = SqrctColors.TextDark
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    Module_Dashboard.DebugLog procName, "Timestamp added to " & timeStampCell.Address(False, False)
    Set timeStampCell = Nothing
End Sub

Private Sub AddPerfRefreshButton(ws As Worksheet)
    ' Adds the "Refresh Perf Dash" button, aligned with the timestamp if possible
    If ws Is Nothing Then Exit Sub
    Dim btn As Shape
    Dim ctrlRow As Long
    Dim btnLeft As Double, btnTop As Double
    Dim btnWidth As Double: btnWidth = 120
    Dim btnHeight As Double: btnHeight = 24
    Const BUTTON_NAME As String = "RefreshPerfButton"
    Dim timeStampCell As Range
    Dim procName As String: procName = "AddPerfRefreshButton"

    Module_Dashboard.DebugLog procName, "Adding refresh button..."

    ' --- Find Timestamp Row ---
    On Error Resume Next
    Set timeStampCell = ws.Columns(2).Find("Last Updated:", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Err.Number = 0 And Not timeStampCell Is Nothing Then
        ctrlRow = timeStampCell.Row ' Use same row as timestamp
        Module_Dashboard.DebugLog procName, "Found timestamp in row " & ctrlRow
    Else
        ' Fallback: Find last row + 2 (same logic as AddDashboardControls)
        Err.Clear
        On Error Resume Next
        ctrlRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        If Err.Number <> 0 Then ctrlRow = 1
        Err.Clear
        On Error GoTo 0
        ctrlRow = ctrlRow + 2
        Module_Dashboard.DebugLog procName, "Timestamp not found, using fallback row " & ctrlRow
    End If
    On Error GoTo 0 ' Restore default

    ' --- Calculate Position (e.g., align right) ---
    On Error Resume Next ' Handle errors accessing column/cell properties
    ' Position button towards the right, e.g., starting at Column J's left edge
    btnLeft = ws.Columns("J").left
    btnTop = ws.Cells(ctrlRow, "J").top ' Align vertically with timestamp row
    If Err.Number <> 0 Then
        Module_Dashboard.DebugLog procName, "WARNING: Error calculating button position. Using defaults. Err: " & Err.Description
        btnLeft = 400 ' Fallback position
        btnTop = ctrlRow * 15 ' Approximate Y position
        Err.Clear
    End If
    On Error GoTo ButtonErrorHandler_Perf ' Specific handler

    ' --- Delete Existing Button ---
    On Error Resume Next
    ws.Shapes(BUTTON_NAME).Delete
    If Err.Number <> 0 Then Err.Clear ' Ignore error if not found
    On Error GoTo ButtonErrorHandler_Perf

    ' --- Create Button ---
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    With btn
        .Name = BUTTON_NAME
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = SqrctColors.MediumBlue
        .Line.Visible = msoFalse
        .Adjustments(1) = 0.1 ' Corner rounding

        ' Add Text
        With .TextFrame2
            .TextRange.Text = "Refresh Perf Dash"
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .WordWrap = msoFalse
            With .TextRange.Font
                .Fill.Visible = msoTrue
                .Fill.Solid
                .Fill.ForeColor.RGB = SqrctColors.TextLight
                .Size = 11
                .Name = FONT_BODY
                .Bold = msoTrue
            End With
        End With

        ' Assign Macro - Ensure correct quoting for workbook name
        .OnAction = "'" & ThisWorkbook.Name & "'!modPerformanceDashboard.BuildModernPerfDashboard"
    End With
    Module_Dashboard.DebugLog procName, "Button created and macro assigned."

ButtonExit_Perf:
    On Error Resume Next
    Set btn = Nothing
    Set timeStampCell = Nothing
    Exit Sub

ButtonErrorHandler_Perf:
    Module_Dashboard.DebugLog procName, "ERROR [" & Err.Number & "]: " & Err.Description & " (Line: " & Erl & ")"
    Resume ButtonExit_Perf
End Sub


' --- HELPER FUNCTIONS ---

Private Sub UpdateStageCycleTimes(ByRef cycleTimes As Object, ByVal stageName As String, ByVal duration As Double)
    ' Updates Avg/Min/Max cycle time for a given stage in the dictionary
    ' Expects cycleTimes dictionary item for stageName to be Array(Avg, Min, Max, Target)
    If cycleTimes Is Nothing Or duration <= 0 Or duration > 365 Or Trim$(stageName) = "" Then Exit Sub
    If Not cycleTimes.Exists(stageName) Then Exit Sub ' Only update existing stages

    Dim stageData As Variant
    stageData = cycleTimes(stageName)

    ' Validate array structure
    If Not IsArray(stageData) Then Exit Sub
    If UBound(stageData) < 2 Then Exit Sub ' Need at least Avg, Min, Max (indices 0, 1, 2)

    Dim countKey As String: countKey = stageName & "_Count" ' Use related key for count

    ' Initialize count if it doesn't exist
    If Not cycleTimes.Exists(countKey) Then
        cycleTimes.Add countKey, CLng(0)
    End If

    ' Increment count
    cycleTimes(countKey) = cycleTimes(countKey) + 1
    Dim currentCount As Long: currentCount = cycleTimes(countKey)

    ' Update Min
    If duration < stageData(1) Or stageData(1) = 9999 Then ' 9999 is initial Min placeholder
        stageData(1) = duration
    End If

    ' Update Max
    If duration > stageData(2) Then
        stageData(2) = duration
    End If

    ' Update Average (Running Average)
    If currentCount = 1 Then
        stageData(0) = duration ' First entry is the average
    Else
        ' Formula: NewAvg = ((OldAvg * (N-1)) + NewValue) / N
        stageData(0) = ((stageData(0) * (currentCount - 1)) + duration) / currentCount
    End If

    ' Write updated array back to dictionary
    cycleTimes(stageName) = stageData

End Sub

Private Function ReadPerfHistoricalData() As Variant
    ' Reads history - Requires ReadHistoricalData to be Public in Module_Dashboard
    ' Return value needs to be Variant to handle potential errors or empty arrays
    Dim histData As Variant
    histData = False ' Default to False or an empty array to indicate failure/no data

    On Error Resume Next ' Handle error if function doesn't exist or fails
    histData = Module_Dashboard.ReadHistoricalData()
    If Err.Number <> 0 Then
        Module_Dashboard.DebugLog "ReadPerfHistoricalData", "ERROR calling Module_Dashboard.ReadHistoricalData. Err: " & Err.Description
        histData = False ' Ensure failure indicated
        Err.Clear
    End If
    On Error GoTo 0

    ReadPerfHistoricalData = histData
End Function

' Provide a local copy of EnsureConfigSheet if the one in Module_Dashboard is not Public
' Ensure it matches the logic in Module_Dashboard if you copy it
Private Sub EnsureConfigSheet_Local()
    Dim ws As Worksheet
    Dim wsExists As Boolean
    Const TARGET_SETTING_NAME As String = "Pipeline Target"
    Const TARGET_DEFAULT_VALUE As Double = 1000000
    Dim procName As String: procName = "EnsureConfigSheet_Local"

    Module_Dashboard.DebugLog procName, "Running local version..."
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Module_Dashboard.CONFIG_SHEET_NAME)
    wsExists = (Err.Number = 0 And Not ws Is Nothing)
    Err.Clear
    On Error GoTo ConfigSheetErrorHandler_Local

    If Not wsExists Then
         Module_Dashboard.DebugLog procName, "Config sheet '" & Module_Dashboard.CONFIG_SHEET_NAME & "' not found. Creating..."
        Dim currentProtection As Boolean
        currentProtection = ThisWorkbook.ProtectStructure Or ThisWorkbook.ProtectWindows
        If currentProtection Then
            Module_Dashboard.DebugLog procName, "Workbook protected. Attempting unprotect..."
            On Error Resume Next
            ThisWorkbook.Unprotect Password:=Module_Dashboard.PW_WORKBOOK ' Use constant
            If Err.Number <> 0 Then
                 Module_Dashboard.DebugLog procName, "ERROR: Failed to unprotect workbook structure. Cannot create config sheet. Err: " & Err.Description
                 Exit Sub ' Cannot proceed
            End If
            Err.Clear
            On Error GoTo ConfigSheetErrorHandler_Local
             Module_Dashboard.DebugLog procName, "Workbook unprotected."
        End If

        ' Add Sheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        On Error Resume Next ' Naming can fail
        ws.Name = Module_Dashboard.CONFIG_SHEET_NAME
        If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Failed to name config sheet. Using '" & ws.Name & "'. Err: " & Err.Description: Err.Clear
        On Error GoTo ConfigSheetErrorHandler_Local ' Restore main handler

        ' Format and add default
        ws.Visible = xlSheetVeryHidden
        ws.Range("A1:C1").value = Array("Setting", "Value", "Notes")
        With ws.Range("A1:C1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = SqrctColors.DarkBlue ' Use Enum
            .Font.Color = SqrctColors.TextLight   ' Use Enum
            .Borders.LineStyle = xlContinuous
        End With
        ws.Range("A2").value = TARGET_SETTING_NAME
        ws.Range("B2").value = TARGET_DEFAULT_VALUE
        ws.Range("B2").NumberFormat = "$#,##0"
        ws.Range("C2").value = "Default pipeline target. Adjust Cell B2."
        ws.Columns("A:C").AutoFit

        ' Re-apply protection if it was on
        If currentProtection Then
            Module_Dashboard.DebugLog procName, "Re-protecting workbook..."
            On Error Resume Next
            ThisWorkbook.Protect Password:=Module_Dashboard.PW_WORKBOOK, Structure:=True, Windows:=True ' Reapply protection
            If Err.Number <> 0 Then Module_Dashboard.DebugLog procName, "WARNING: Failed to re-protect workbook. Err: " & Err.Description: Err.Clear
            On Error GoTo ConfigSheetErrorHandler_Local
            Module_Dashboard.DebugLog procName, "Workbook re-protected."
        End If
        Module_Dashboard.DebugLog procName, "Config sheet created and populated."
    Else
        Module_Dashboard.DebugLog procName, "Config sheet '" & Module_Dashboard.CONFIG_SHEET_NAME & "' found. Verifying setting..."
        ' --- Verify Target Setting Exists ---
        Dim targetFound As Boolean: targetFound = False
        Dim r As Long, lastCfgRowCheck As Long
        On Error Resume Next ' Handle errors reading sheet
        lastCfgRowCheck = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
        If Err.Number = 0 And lastCfgRowCheck >= 2 Then
            For r = 2 To lastCfgRowCheck
                If Trim$(CStr(ws.Cells(r, "A").value)) = TARGET_SETTING_NAME Then
                    targetFound = True
                    Exit For
                End If
            Next r
        End If
        Err.Clear
        On Error GoTo ConfigSheetErrorHandler_Local ' Restore handler

        If Not targetFound Then
             Module_Dashboard.DebugLog procName, "'" & TARGET_SETTING_NAME & "' not found. Adding default entry..."
            ' Add the default setting if missing
            On Error Resume Next ' Handle errors writing to sheet
            Dim lastRowAdd As Long
            lastRowAdd = ws.Cells(ws.rows.Count, "A").End(xlUp).Row + 1
            If Err.Number = 0 Then
                ws.Cells(lastRowAdd, "A").value = TARGET_SETTING_NAME
                ws.Cells(lastRowAdd, "B").value = TARGET_DEFAULT_VALUE
                ws.Cells(lastRowAdd, "B").NumberFormat = "$#,##0"
                ws.Cells(lastRowAdd, "C").value = "Default pipeline target. Adjust Cell B" & lastRowAdd & "."
                ws.Columns("A:C").AutoFit
                 Module_Dashboard.DebugLog procName, "Added default '" & TARGET_SETTING_NAME & "' to row " & lastRowAdd
            Else
                 Module_Dashboard.DebugLog procName, "ERROR: Could not add default setting. Sheet might be protected. Err: " & Err.Description
            End If
            Err.Clear
            On Error GoTo ConfigSheetErrorHandler_Local ' Restore handler
        Else
             Module_Dashboard.DebugLog procName, "'" & TARGET_SETTING_NAME & "' setting found."
        End If
    End If

ConfigSheetExit_Local:
    On Error Resume Next
    Set ws = Nothing
    Exit Sub

ConfigSheetErrorHandler_Local:
    Module_Dashboard.DebugLog procName, "ERROR [" & Err.Number & "]: " & Err.Description & " (Line: " & Erl & ")"
    Resume ConfigSheetExit_Local ' Go to cleanup on error
End Sub



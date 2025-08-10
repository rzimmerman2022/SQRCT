Option Explicit

' Color constants for consistent theming
Public Enum SqrctColors
    DarkBlue = &H874B0A     ' RGB(10, 75, 135)
    MediumBlue = &HD47800   ' RGB(0, 120, 212)
    LightBlue = &HF5BB5E    ' RGB(94, 187, 245)
    SuccessGreen = &H50AF4C ' RGB(76, 175, 80)
    WarningOrange = &H809FF ' RGB(255, 152, 0)
    ErrorRed = &H3935E5     ' RGB(229, 57, 53)
    TextDark = &H363534     ' RGB(52, 53, 54)
    TextLight = &HFFFFFF    ' RGB(255, 255, 255)
    BorderGrey = &HC8C8C8   ' RGB(200, 200, 200)
    BackgroundGrey = &HFAF9F8 ' RGB(248, 249, 250)
End Enum

' Card size constants
Private Const CARD_WIDTH As Long = 5
Private Const CARD_HEIGHT As Long = 4
Private Const SECTION_MARGIN As Long = 2

' Font settings
Private Const FONT_HEADER As String = "Segoe UI Light"
Private Const FONT_TITLE As String = "Segoe UI Semibold"
Private Const FONT_BODY As String = "Segoe UI"
Private Const FONT_NUMBER As String = "Segoe UI Light"

Public Sub BuildModernPerfDashboard()
    Dim wsPerf As Worksheet
    Dim dataModel As Scripting.Dictionary
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    Dim t0 As Double: t0 = Timer
    
    ' Speed wrapper to improve performance
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation
    
    On Error GoTo CleanFail
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' --- Ensure Config Sheet Exists ---
    EnsureConfigSheet ' Add this call early
    
    ' Get or create the sheet
    Set wsPerf = GetOrCreateSheet("SQRCT PERF DASH", _
                "SQRCT PERFORMANCE DASHBOARD", _
                SqrctColors.DarkBlue)
    
    ' Apply modern theme
    ApplyModernDashboardTheme wsPerf
    
    ' Clear existing content and shapes
    ClearDashboardContents wsPerf
    
    ' Build the data model from dashboard array
    Set dataModel = BuildDashboardDataModel()
    
    ' Create visualization sections
    CreateExecutiveSnapshot wsPerf, dataModel
    CreatePipelineAnalytics wsPerf, dataModel
    CreatePerformanceMetrics wsPerf, dataModel
    CreateForecastTrends wsPerf, dataModel
    
    ' Add timestamp and refresh controls
    AddDashboardControls wsPerf
    
    ' Add refresh button
    AddRefreshButton wsPerf

    ' --- Update Historical Metrics ---
    UpdateHistoricalMetrics dataModel ' Call after model is built
    
    ' Log completion
    #If DEBUG Then
        Module_Dashboard.DebugLog "BuildModernPerfDashboard", "Completed in " & Format(Timer - t0, "0.00") & " s"
#EndIf
    
CleanExit:
    ' Restore original settings
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation
    Exit Sub
    
CleanFail:
    ' Log error and clean up
    #If DEBUG Then
        Module_Dashboard.DebugLog "BuildModernPerfDashboard", "ERROR: " & Err.Source & " - " & Err.Description
    #EndIf
    
    MsgBox "An error occurred while building the performance dashboard:" & vbCrLf & _
           Err.Description, vbExclamation, "Dashboard Error"
    Resume CleanExit
End Sub

' Helper function to update the stageCycleTimes dictionary with new duration data
' Parameters:
'   cycleTimes - Dictionary of arrays where key = stage name, value = array(AvgDays, MinDays, MaxDays, TargetDays)
'   stageName - The stage name to update
'   duration - The duration in days to add to this stage's metrics
Private Sub UpdateStageCycleTimes(ByRef cycleTimes As Scripting.Dictionary, ByVal stageName As String, ByVal duration As Double)
    ' Skip if duration is invalid
    If duration <= 0 Or duration > 365 Then Exit Sub
    
    ' Verify stage exists in dictionary
    If Not cycleTimes.Exists(stageName) Then Exit Sub
    
    ' Get current values
    Dim stageData As Variant
    stageData = cycleTimes(stageName)
    
    ' Current values in array:
    ' 0 = Average days (weighted sum, needs to be divided by count later)
    ' 1 = Minimum days
    ' 2 = Maximum days
    ' 3 = Target days (constant, no change)
    
    ' Create counter in dictionary if not exists
    Dim countKey As String: countKey = stageName & "_Count"
    If Not cycleTimes.Exists(countKey) Then
        cycleTimes.Add countKey, 0
    End If
    
    ' Update count
    cycleTimes(countKey) = cycleTimes(countKey) + 1
    
    ' Update min (if new duration is lower)
    If duration < stageData(1) Then stageData(1) = duration
    
    ' Update max (if new duration is higher)
    If duration > stageData(2) Then stageData(2) = duration
    
    ' Update average - we're storing a running sum; will divide by count at the end
    stageData(0) = ((stageData(0) * (cycleTimes(countKey) - 1)) + duration) / cycleTimes(countKey)
    
    ' Update the array in the dictionary
    cycleTimes(stageName) = stageData
End Sub

' Build the dashboard data model from source data
Private Function BuildDashboardDataModel() As Scripting.Dictionary
    ' Create a dictionary object to hold all metrics
    Dim model As New Scripting.Dictionary
    
    ' Get dashboard data array
    Dim arrDash As Variant
    arrDash = Module_Dashboard.BuildDashboardDataArray()
    
    ' Initialize counters
    Dim totalRows As Long, openRows As Long, convRows As Long
    Dim decRows As Long, wwomRows As Long, firstFURows As Long
    Dim secondFURows As Long, thirdFURows As Long, pendingRows As Long
    Dim sumOpenAmt As Double, maxOpenAmt As Double, totalConvAmt As Double
    
    ' Stage dictionaries
    Dim dictStageCount As New Scripting.Dictionary
    Dim dictStageAmount As New Scripting.Dictionary
    
    ' Age buckets (0-30, 31-60, 61-90, 90+)
    Dim ageBuckets(1 To 4) As Long
    Dim ageBucketValues(1 To 4) As Double
    
    ' Cycle time tracking
    Dim totalCycleDays As Long, cycleCount As Long
    Dim stageTimeDays As New Scripting.Dictionary
    Dim stageTimeCount As New Scripting.Dictionary
    
    ' *** NEW: Stage Cycle Time Tracking ***
    ' Track transitions between phases to calculate avg time in each stage
    ' Structure: Dictionary of arrays where key = stage name, value = array(AvgDays, MinDays, MaxDays, TargetDays)
    Dim stageCycleTimes As New Scripting.Dictionary
    ' Initialize with stages we want to track
    stageCycleTimes.Add "First F/U", Array(0, 9999, 0, 10) ' Default target of 10 days per stage
    stageCycleTimes.Add "Second F/U", Array(0, 9999, 0, 7)
    stageCycleTimes.Add "Third F/U", Array(0, 9999, 0, 10)
    stageCycleTimes.Add "Pending", Array(0, 9999, 0, 5)
    
    ' Transition tracking
    ' For each quote, we'll track what stage(s) it went through and how long in each
    Dim quoteStageLog As New Scripting.Dictionary ' Key = DocNum, Value = Dictionary of stage transitions
    
    ' Process rows from source data
    On Error Resume Next ' Handle potential array errors
    totalRows = UBound(arrDash, 1)
    If Err.Number <> 0 Then
        ' Array may be empty or invalid
        #If DEBUG Then
            Module_Dashboard.DebugLog "BuildDashboardDataModel", "Error accessing dashboard array: " & Err.Description
        #End If
        Err.Clear
        totalRows = 0
    End If
    On Error GoTo 0
    
    If totalRows <= 0 Then
        ' No data, return empty model with zeroed metrics
        model.Add "totalRows", 0
        model.Add "openRows", 0
        model.Add "convRows", 0
        model.Add "decRows", 0
        model.Add "wwomRows", 0
        model.Add "firstFURows", 0
        model.Add "secondFURows", 0
        model.Add "thirdFURows", 0
        model.Add "pendingRows", 0
        model.Add "sumOpenAmt", 0
        model.Add "maxOpenAmt", 0
        model.Add "totalConvAmt", 0
        model.Add "avgOpenAmt", 0
        model.Add "conversionRate", 0
        model.Add "avgCycleTime", 0
        model.Add "pipelineVsTarget", 0
        model.Add "convRateMoM", 0
        model.Add "cycleTimeYoY", 0
        model.Add "ageBuckets", ageBuckets
        model.Add "ageBucketValues", ageBucketValues
        model.Add "dictStageCount", dictStageCount
        model.Add "dictStageAmount", dictStageAmount
        model.Add "stageConversionRate", 0
        
        ' Add empty arrays for charts
        ReDim Preserve ageBuckets(1 To 4) ' Age buckets
        model.Add "avgStageTimes", Array() ' Empty stage times array
        
        ' Return the empty model
        Set BuildDashboardDataModel = model
        Exit Function
    End If
    
    ' Process rows
    Dim r As Long, phase As String, amt As Double
    Dim docNum As String, prevPhase As String, stageKey As String
    
    For r = 1 To totalRows
        On Error Resume Next
        phase = CStr(arrDash(r, 12))  ' Engagement Phase (L)
        docNum = CStr(arrDash(r, 1))  ' Document Number (A)
        
        ' Handle empty phase as active
        If Trim(phase) = "" Then phase = "Undefined"
        
        ' Ensure phase exists in dictionaries
        If Not dictStageCount.Exists(phase) Then
            dictStageCount.Add phase, 0
            dictStageAmount.Add phase, 0
        End If
        
        ' Increment stage counters
        dictStageCount(phase) = dictStageCount(phase) + 1
        
        ' Add stage amount (with error handling)
        amt = CDbl(arrDash(r, 4))  ' Document Amount (D)
        If Err.Number <> 0 Then
            amt = 0
            Err.Clear
        End If
        
        dictStageAmount(phase) = dictStageAmount(phase) + amt
        
        ' *** Cycle Time By Stage Calculation ***
        ' Calculate stage duration from document date to last contact date
        Dim docDate As Date, lastContactDate As Date, stageDays As Long
        docDate = CDate(arrDash(r, 5))  ' Document Date (E)
        lastContactDate = CDate(arrDash(r, 13))  ' Last Contact Date (M)
        
        If Err.Number = 0 And lastContactDate >= docDate Then
            ' Get total days from creation to last contact
            stageDays = lastContactDate - docDate
            
            ' Skip if days calculation is invalid (negative or excessive)
            If stageDays >= 0 And stageDays <= 365 Then
                ' Infer stage duration based on the current phase
                If phase = "First F/U" Then
                    ' Still in first follow-up - all time spent here
                    UpdateStageCycleTimes stageCycleTimes, "First F/U", stageDays
                ElseIf phase = "Second F/U" Then
                    ' Allocate 60% to First F/U and 40% to current stage
                    UpdateStageCycleTimes stageCycleTimes, "First F/U", stageDays * 0.6
                    UpdateStageCycleTimes stageCycleTimes, "Second F/U", stageDays * 0.4
                ElseIf phase = "Third F/U" Then
                    ' Split across all three stages
                    UpdateStageCycleTimes stageCycleTimes, "First F/U", stageDays * 0.4
                    UpdateStageCycleTimes stageCycleTimes, "Second F/U", stageDays * 0.3
                    UpdateStageCycleTimes stageCycleTimes, "Third F/U", stageDays * 0.3
                ElseIf phase = "Pending" Then
                    ' Split across four stages
                    UpdateStageCycleTimes stageCycleTimes, "First F/U", stageDays * 0.3
                    UpdateStageCycleTimes stageCycleTimes, "Second F/U", stageDays * 0.3
                    UpdateStageCycleTimes stageCycleTimes, "Third F/U", stageDays * 0.2
                    UpdateStageCycleTimes stageCycleTimes, "Pending", stageDays * 0.2
                ElseIf phase = "Converted" Then
                    ' Split across all stages to end
                    UpdateStageCycleTimes stageCycleTimes, "First F/U", stageDays * 0.3
                    UpdateStageCycleTimes stageCycleTimes, "Second F/U", stageDays * 0.25
                    UpdateStageCycleTimes stageCycleTimes, "Third F/U", stageDays * 0.25
                    UpdateStageCycleTimes stageCycleTimes, "Pending", stageDays * 0.2
                End If
            End If
        End If
        
        ' Categorize by phase for counting
        Select Case phase
            Case "Converted"
                convRows = convRows + 1
                totalConvAmt = totalConvAmt + amt
                
                ' Calculate cycle time for converted quotes
                ' Get document date (creation date) and last contact date (for converted quotes)
                Dim createDate As Date, convDate As Date
                On Error Resume Next
                createDate = CDate(arrDash(r, 5))  ' Document Date (E)
                convDate = CDate(arrDash(r, 13))   ' Last Contact Date (M)
                
                ' If dates are valid and in correct order, calculate cycle time
                If Err.Number = 0 And convDate > createDate Then
                    totalCycleDays = totalCycleDays + (convDate - createDate)
                    cycleCount = cycleCount + 1
                End If
                Err.Clear
                
            Case "Declined"
                decRows = decRows + 1
                
            Case "WWOM"
                wwomRows = wwomRows + 1
                
            Case "First F/U"
                firstFURows = firstFURows + 1
                openRows = openRows + 1
                sumOpenAmt = sumOpenAmt + amt
                If amt > maxOpenAmt Then maxOpenAmt = amt
                
            Case "Second F/U"
                secondFURows = secondFURows + 1
                openRows = openRows + 1
                sumOpenAmt = sumOpenAmt + amt
                If amt > maxOpenAmt Then maxOpenAmt = amt
                
            Case "Third F/U"
                thirdFURows = thirdFURows + 1
                openRows = openRows + 1
                sumOpenAmt = sumOpenAmt + amt
                If amt > maxOpenAmt Then maxOpenAmt = amt
                
            Case "Pending"
                pendingRows = pendingRows + 1
                openRows = openRows + 1
                sumOpenAmt = sumOpenAmt + amt
                If amt > maxOpenAmt Then maxOpenAmt = amt
                
            Case Else  ' Other active phases
                openRows = openRows + 1
                sumOpenAmt = sumOpenAmt + amt
                If amt > maxOpenAmt Then maxOpenAmt = amt
        End Select
        
        ' Calculate quote age and add to bucket
        Dim quoteDate As Date
        quoteDate = CDate(arrDash(r, 5))  ' Document Date (E)
        If Err.Number = 0 Then
            Dim quoteAge As Long: quoteAge = Date - quoteDate
            
            ' Add to appropriate aging bucket
            If phase <> "Converted" And phase <> "Declined" And phase <> "WWOM" Then
                If quoteAge <= 30 Then
                    ageBuckets(1) = ageBuckets(1) + 1
                    ageBucketValues(1) = ageBucketValues(1) + amt
                ElseIf quoteAge <= 60 Then
                    ageBuckets(2) = ageBuckets(2) + 1
                    ageBucketValues(2) = ageBucketValues(2) + amt
                ElseIf quoteAge <= 90 Then
                    ageBuckets(3) = ageBuckets(3) + 1
                    ageBucketValues(3) = ageBucketValues(3) + amt
                Else
                    ageBuckets(4) = ageBuckets(4) + 1
                    ageBucketValues(4) = ageBucketValues(4) + amt
                End If
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next r
    
    ' *** Finalize Stage Cycle Times ***
    ' Create the stage times data for the chart
    Dim avgStageTimesArray() As Variant
    Dim stageIdx As Long, stageCount As Long
    
    ' Count actual stages (exclude _Count entries)
    stageCount = 0
    Dim stageNameKey As Variant
    For Each stageNameKey In stageCycleTimes.Keys
        If Right(stageNameKey, 6) <> "_Count" Then
            stageCount = stageCount + 1
        End If
    Next stageNameKey
    
    ' Create properly sized array for stage times
    If stageCount > 0 Then
        ReDim avgStageTimesArray(1 To stageCount, 1 To 5) ' 5 columns: Stage, AvgDays, MinDays, MaxDays, Target
        
        stageIdx = 0
        For Each stageNameKey In stageCycleTimes.Keys
            ' Skip counter entries
            If Right(stageNameKey, 6) <> "_Count" Then
                stageIdx = stageIdx + 1
                Dim stageDataArray As Variant
                stageDataArray = stageCycleTimes(stageNameKey)
                
                avgStageTimesArray(stageIdx, 1) = stageNameKey ' Stage Name
                avgStageTimesArray(stageIdx, 2) = Round(stageDataArray(0), 1) ' Avg Days
                avgStageTimesArray(stageIdx, 3) = stageDataArray(1) ' Min Days
                avgStageTimesArray(stageIdx, 4) = stageDataArray(2) ' Max Days
                avgStageTimesArray(stageIdx, 5) = stageDataArray(3) ' Target Days
            End If
        Next stageNameKey
    Else
        ' No stage data, return empty array
        ReDim avgStageTimesArray(0, 0)
    End If
    
    ' --- Read Pipeline Target from Config Sheet ---
    Dim pipelineTargetValue As Double
    Dim wsConfig As Worksheet
    Dim targetFound As Boolean
    
    ' Default value if not found
    pipelineTargetValue = 1000000 ' $1M default
    targetFound = False
    
    ' Try to read from Config sheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("Config")
    If Err.Number = 0 And Not wsConfig Is Nothing Then
        ' Find Pipeline Target setting
        Dim cfgRow As Long
        For cfgRow = 2 To wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row
            If CStr(wsConfig.Cells(cfgRow, "A").Value) = "Pipeline Target" Then
                ' Found it - try to read the value
                If IsNumeric(wsConfig.Cells(cfgRow, "B").Value) Then
                    pipelineTargetValue = CDbl(wsConfig.Cells(cfgRow, "B").Value)
                    targetFound = True
                End If
                Exit For
            End If
        Next cfgRow
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Calculate Pipeline vs Target percentage
    Dim pipelineVsTarget As Double
    If pipelineTargetValue <> 0 Then
        pipelineVsTarget = (sumOpenAmt / pipelineTargetValue) - 1
    Else
        pipelineVsTarget = 0 ' Avoid division by zero
    End If
    
    ' Calculate derived metrics
    Dim avgCycleTime As Double
    avgCycleTime = IIf(cycleCount = 0, 0, totalCycleDays / cycleCount)
    
    Dim conversionRate As Double
    conversionRate = IIf(totalRows = 0, 0, convRows / totalRows)
    
    ' Calculate historical comparison data
    Dim histData As Variant
    Dim prevMonthConvRate As Double, prevYearCycleTime As Double
    Dim convRateMoM As Double, cycleTimeYoY As Double
    
    histData = ReadHistoricalData()
    
    ' Default to 0 if no historical data
    prevMonthConvRate = 0
    prevYearCycleTime = 0
    
    ' If we have historical data, find previous month and year data
    If IsArray(histData) Then
        Dim histRow As Long, histDate As Date
        Dim currentMonthDate As Date, prevMonthDate As Date, prevYearDate As Date
        
        ' Calculate reference dates
        currentMonthDate = DateSerial(Year(Date), Month(Date), 1)
        prevMonthDate = DateAdd("m", -1, currentMonthDate)
        prevYearDate = DateAdd("yyyy", -1, currentMonthDate)
        
        ' Look for matches in historical data
        For histRow = UBound(histData, 1) To LBound(histData, 1) Step -1
            On Error Resume Next
            histDate = CDate(histData(histRow, 1)) ' Column 1 = Date
            If Err.Number = 0 Then
                ' Check for previous month entry
                If prevMonthConvRate = 0 And Year(histDate) = Year(prevMonthDate) And Month(histDate) = Month(prevMonthDate) Then
                    prevMonthConvRate = CDbl(histData(histRow, 6)) ' Column 6 = Conversion Rate
                End If
                
                ' Check for previous year entry
                If prevYearCycleTime = 0 And Year(histDate) = Year(prevYearDate) And Month(histDate) = Month(prevYearDate) Then
                    prevYearCycleTime = CDbl(histData(histRow, 7)) ' Column 7 = Avg Cycle Time
                End If
                
                ' Exit if we found both values
                If prevMonthConvRate <> 0 And prevYearCycleTime <> 0 Then Exit For
            End If
            Err.Clear
            On Error GoTo 0
        Next histRow
    End If
    
    ' Calculate MoM and YoY changes
    If prevMonthConvRate <> 0 Then
        convRateMoM = (conversionRate / prevMonthConvRate) - 1
    Else
        convRateMoM = 0
    End If
    
    If prevYearCycleTime <> 0 Then
        cycleTimeYoY = (avgCycleTime / prevYearCycleTime) - 1
    Else
        cycleTimeYoY = 0
    End If
    
    ' Store all metrics in the model
    With model
        ' Core metrics
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
        .Add "avgOpenAmt", IIf(openRows = 0, 0, sumOpenAmt / openRows)
        
        ' Rate and time metrics
        .Add "conversionRate", conversionRate
        .Add "avgCycleTime", Round(avgCycleTime, 1)
        
        ' Comparisons
        .Add "pipelineVsTarget", pipelineVsTarget
        .Add "convRateMoM", convRateMoM
        .Add "cycleTimeYoY", cycleTimeYoY
        
        ' Age buckets
        .Add "ageBuckets", ageBuckets
        .Add "ageBucketValues", ageBucketValues
        
        ' Stage distributions
        .Add "dictStageCount", dictStageCount
        .Add "dictStageAmount", dictStageAmount
        
        ' Stage cycle times
        .Add "avgStageTimes", avgStageTimesArray
        
        ' Funnel metrics
        .Add "stageConversionRate", IIf(firstFURows = 0, 0, secondFURows / firstFURows)
    End With
    
    ' Return the populated data model
    Set BuildDashboardDataModel = model
End Function

' ----------------------------------------------------------------------------------
' Visual Rendering and UI Functions
' ----------------------------------------------------------------------------------

' Apply modern dashboard theme with consistent styling
Private Sub ApplyModernDashboardTheme(ws As Worksheet)
    ' Clear any existing theme elements
    On Error Resume Next
    ws.Cells.Interior.ColorIndex = xlNone
    On Error GoTo 0
    
    ' Apply modern color scheme to entire sheet
    With ws
        .Tab.Color = SqrctColors.DarkBlue
        .Cells.Font.Name = FONT_BODY
        .Cells.Font.Size = 10
    End With
    
    ' Apply background color to entire sheet
    ws.Cells.Interior.Color = SqrctColors.BackgroundGrey
    
    ' Apply header row styling
    With ws.Range("A1:Z1")
        .Font.Name = FONT_HEADER
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = SqrctColors.DarkBlue
        .Font.Color = SqrctColors.TextLight
        .RowHeight = 36
    End With
    
    ' Log theme application
    #If DEBUG Then
        Module_Dashboard.DebugLog "ApplyModernDashboardTheme", "Applied modern theme to worksheet " & ws.Name
    #EndIf
End Sub

' Clear existing dashboard content before rebuilding
Private Sub ClearDashboardContents(ws As Worksheet)
    Dim shp As Shape
    
    ' Clear all existing shapes/charts
    On Error Resume Next
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0
    
    ' Clear cell contents below header row
    ws.Range("A2:Z1000").ClearContents
    
    ' Log clearing operation
    #If DEBUG Then
        Module_Dashboard.DebugLog "ClearDashboardContents", "Cleared dashboard contents"
    #EndIf
End Sub

' Create the executive summary section with key metrics
Private Sub CreateExecutiveSnapshot(ws As Worksheet, model As Scripting.Dictionary)
    Dim startRow As Long: startRow = 2
    Dim startCol As Long: startCol = 2
    Dim r As Long, c As Long
    
    ' Create section title
    With ws.Cells(startRow, startCol)
        .Value = "EXECUTIVE SUMMARY"
        .Font.Name = FONT_TITLE
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Font.Color = SqrctColors.DarkBlue
    End With
    
    ' Create 4 key metric cards in a 2x2 grid
    CreateMetricCard ws, startRow + 1, startCol, "OPEN QUOTES", _
        model("openRows"), "#,##0", "quotes", SqrctColors.MediumBlue, _
        "Total open quotes in pipeline"
        
    CreateMetricCard ws, startRow + 1, startCol + CARD_WIDTH, "PIPELINE VALUE", _
        model("sumOpenAmt"), "$#,##0", "", SqrctColors.DarkBlue, _
        "Total value of open quotes"
        
    CreateMetricCard ws, startRow + 1 + CARD_HEIGHT, startCol, "CONVERSION RATE", _
        model("conversionRate"), "0.0%", "", SqrctColors.SuccessGreen, _
        IIf(model("convRateMoM") >= 0, "↑ ", "↓ ") & Format(Abs(model("convRateMoM")), "0.0%") & " vs last month"
        
    CreateMetricCard ws, startRow + 1 + CARD_HEIGHT, startCol + CARD_WIDTH, "AVG CYCLE TIME", _
        model("avgCycleTime"), "0.0", " days", SqrctColors.LightBlue, _
        IIf(model("cycleTimeYoY") <= 0, "↑ ", "↓ ") & Format(Abs(model("cycleTimeYoY")), "0.0%") & " vs last year"
    
    #If DEBUG Then
        Module_Dashboard.DebugLog "CreateExecutiveSnapshot", "Created executive summary"
    #EndIf
End Sub

' Create the pipeline analytics section with charts
Private Sub CreatePipelineAnalytics(ws As Worksheet, model As Scripting.Dictionary)
    Dim startRow As Long: startRow = 2 + (CARD_HEIGHT * 2) + SECTION_MARGIN
    Dim startCol As Long: startCol = 2
    
    ' Create section title
    With ws.Cells(startRow, startCol)
        .Value = "PIPELINE ANALYTICS"
        .Font.Name = FONT_TITLE
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Font.Color = SqrctColors.DarkBlue
    End With
    
    ' Create pipeline stage distribution card
    CreatePipelineDistributionCard ws, startRow + 1, startCol, model
    
    ' Create pipeline age distribution card
    CreatePipelineAgeCard ws, startRow + 1, startCol + CARD_WIDTH + 1, model
    
    #If DEBUG Then
        Module_Dashboard.DebugLog "CreatePipelineAnalytics", "Created pipeline analytics section"
    #EndIf
End Sub

' Create the performance metrics section
Private Sub CreatePerformanceMetrics(ws As Worksheet, model As Scripting.Dictionary)
    Dim startRow As Long: startRow = 2 + (CARD_HEIGHT * 4) + (SECTION_MARGIN * 2)
    Dim startCol As Long: startCol = 2
    
    ' Create section title
    With ws.Cells(startRow, startCol)
        .Value = "PERFORMANCE METRICS"
        .Font.Name = FONT_TITLE
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Font.Color = SqrctColors.DarkBlue
    End With
    
    ' Create cycle time by stage card
    CreateCycleTimeCard ws, startRow + 1, startCol, model
    
    ' Create conversion funnel card
    CreateConversionFunnelCard ws, startRow + 1, startCol + CARD_WIDTH + 1, model
    
    #If DEBUG Then
        Module_Dashboard.DebugLog "CreatePerformanceMetrics", "Created performance metrics section"
    #EndIf
End Sub

' Create forecast trends section
Private Sub CreateForecastTrends(ws As Worksheet, model As Scripting.Dictionary)
    Dim startRow As Long: startRow = 2 + (CARD_HEIGHT * 6) + (SECTION_MARGIN * 3)
    Dim startCol As Long: startCol = 2
    
    ' Create section title
    With ws.Cells(startRow, startCol)
        .Value = "FORECAST TRENDS"
        .Font.Name = FONT_TITLE
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Font.Color = SqrctColors.DarkBlue
    End With
    
    ' Create pipeline vs target card
    CreatePipelineTargetCard ws, startRow + 1, startCol, model
    
    ' Create historical trends card
    CreateHistoricalTrendsCard ws, startRow + 1, startCol + CARD_WIDTH + 1, model
    
    #If DEBUG Then
        Module_Dashboard.DebugLog "CreateForecastTrends", "Created forecast trends section"
    #EndIf
End Sub

' Add timestamp and refresh controls to dashboard
Private Sub AddDashboardControls(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 2
    
    ' Add timestamp
    With ws.Cells(lastRow, 2)
        .Value = "Last Updated: " & Format(Now, "mm/dd/yyyy hh:nn AM/PM")
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = SqrctColors.TextDark
    End With
    
    #If DEBUG Then
        Module_Dashboard.DebugLog "AddDashboardControls", "Added dashboard controls"
    #EndIf
End Sub

' Add refresh button to dashboard
Private Sub AddRefreshButton(ws As Worksheet)
    Dim btn As Shape
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 2
    
    ' Create refresh button
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Cells(lastRow, 8).Left, ws.Cells(lastRow, 8).Top, 100, 24)
    
    ' Format button
    With btn
        .Fill.ForeColor.RGB = SqrctColors.MediumBlue
        .Line.Visible = msoFalse
        
        With .TextFrame2
            .TextRange.Text = "Refresh Dashboard"
            .TextRange.Font.Fill.ForeColor.RGB = SqrctColors.TextLight
            .TextRange.Font.Size = 11
            .TextRange.Font.Name = FONT_BODY
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
        End With
        
        .OnAction = "BuildModernPerfDashboard"
    End With
    
    #If DEBUG Then
        Module_Dashboard.DebugLog "AddRefreshButton", "Added refresh button"
    #EndIf
End Sub

' Helper function to create a metric card
Private Sub CreateMetricCard(ws As Worksheet, topRow As Long, leftCol As Long, _
                           title As String, value As Variant, format As String, _
                           suffix As String, colorCode As Long, subtitle As String)
    Dim card As Shape
    Dim titleCell As Range
    Dim valueCell As Range
    Dim subtitleCell As Range
    
    ' Set cell references
    Set titleCell = ws.Cells(topRow, leftCol)
    Set valueCell = ws.Cells(topRow + 1, leftCol)
    Set subtitleCell = ws.Cells(topRow + 2, leftCol)
    
    ' Create card outline
    Set card = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        titleCell.Left, titleCell.Top, _
        ws.Columns(leftCol).Width * CARD_WIDTH - 10, _
        ws.Rows(topRow).Height * CARD_HEIGHT - 10)
    
    ' Format card
    With card
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = SqrctColors.BorderGrey
        .Line.Weight = 1
        .Line.Visible = msoTrue
        .Shadow.Type = msoShadow21
        .Name = "Card_" & Replace(title, " ", "_")
    End With
    
    ' Set title
    With titleCell
        .Value = UCase(title)
        .Font.Name = FONT_TITLE
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Font.Color = colorCode
    End With
    
    ' Set value
    With valueCell
        If IsNumeric(value) Then
            .Value = Format(value, format) & suffix
        Else
            .Value = value & suffix
        End If
        .Font.Name = FONT_NUMBER
        .Font.Size = 24
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Color = SqrctColors.TextDark
    End With
    
    ' Set subtitle
    With subtitleCell
        .Value = subtitle
        .Font.Name = FONT_BODY
        .Font.Size = 9
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .Font.Color = SqrctColors.TextDark
    End With
End Sub

' Helper functions for chart cards
Private Sub CreatePipelineDistributionCard(ws As Worksheet, topRow As Long, leftCol As Long, model As Scripting.Dictionary)
    ' Placeholder for pipeline distribution chart
    CreateMetricCard ws, topRow, leftCol, "STAGE DISTRIBUTION", _
        "Data Available", "", "", SqrctColors.MediumBlue, _
        Format(model("firstFURows") + model("secondFURows") + model("thirdFURows") + model("pendingRows"), "#,##0") & " active quotes"
End Sub

Private Sub CreatePipelineAgeCard(ws As Worksheet, topRow As Long, leftCol As Long, model As Scripting.Dictionary)
    ' Placeholder for pipeline age chart
    CreateMetricCard ws, topRow, leftCol, "QUOTE AGING", _
        "Data Available", "", "", SqrctColors.MediumBlue, _
        Format(model("ageBuckets")(4), "#,##0") & " quotes over 90 days"
End Sub

Private Sub CreateCycleTimeCard(ws As Worksheet, topRow As Long, leftCol As Long, model As Scripting.Dictionary)
    ' Placeholder for cycle time by stage chart
    CreateMetricCard ws, topRow, leftCol, "CYCLE TIME BY STAGE", _
        "Data Available", "", "", SqrctColors.MediumBlue, _
        "Average: " & Format(model("avgCycleTime"), "0.0") & " days"
End Sub

Private Sub CreateConversionFunnelCard(ws As Worksheet, topRow As Long, leftCol As Long, model As Scripting.Dictionary)
    ' Placeholder for conversion funnel chart
    CreateMetricCard ws, topRow, leftCol, "CONVERSION FUNNEL", _
        "Data Available", "", "", SqrctColors.MediumBlue, _
        Format(model("conversionRate") * 100, "0.0") & "% overall conversion"
End Sub

Private Sub CreatePipelineTargetCard(ws As Worksheet, topRow As Long, leftCol As Long, model As Scripting.Dictionary)
    ' Placeholder for pipeline vs target chart
    Dim pctLabel As String
    
    If model("pipelineVsTarget") >= 0 Then
        pctLabel = "+" & Format(model("pipelineVsTarget") * 100, "0.0") & "%"
    Else
        pctLabel = Format(model("pipelineVsTarget") * 100, "0.0") & "%"
    End If
    
    CreateMetricCard ws, topRow, leftCol, "PIPELINE VS TARGET", _
        pctLabel, "", " vs goal", _
        IIf(model("pipelineVsTarget") >= 0, SqrctColors.SuccessGreen, SqrctColors.ErrorRed), _
        "Current: $" & Format(model("sumOpenAmt"), "#,##0")
End Sub

Private Sub CreateHistoricalTrendsCard(ws As Worksheet, topRow As Long, leftCol As Long, model As Scripting.Dictionary)
    ' Placeholder for historical trends chart
    CreateMetricCard ws, topRow, leftCol, "HISTORICAL TRENDS", _
        "Data Available", "", "", SqrctColors.MediumBlue, _
        Format(model("convRows"), "#,##0") & " quotes converted YTD"
End Sub

' Create a configurable settings sheet for dashboard parameters
Private Sub EnsureConfigSheet()
    Dim ws As Worksheet
    Dim wsExists As Boolean
    Const CONFIG_SHEET_NAME As String = "Config"
    Const TARGET_SETTING_NAME As String = "Pipeline Target"
    Const TARGET_DEFAULT_VALUE As Double = 1000000 ' $1M default target
    
    ' Check if Config sheet already exists
    wsExists = False
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    If Err.Number = 0 And Not ws Is Nothing Then wsExists = True
    Err.Clear
    On Error GoTo 0
    
    If Not wsExists Then
        ' Create new Config sheet (very hidden)
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONFIG_SHEET_NAME
        ws.Visible = xlSheetVeryHidden
        
        ' Add headers
        ws.Range("A1:C1").Value = Array("Setting", "Value", "Notes")
        With ws.Range("A1:C1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = SqrctColors.DarkBlue
            .Font.Color = SqrctColors.TextLight
        End With
        
        ' Add Pipeline Target setting
        ws.Range("A2").Value = TARGET_SETTING_NAME
        ws.Range("B2").Value = TARGET_DEFAULT_VALUE
        ws.Range("C2").Value = "Placeholder target value. Adjust to your specific quarterly/annual revenue goals."
        
        ' Format value as currency
        ws.Range("B2").NumberFormat = "$#,##0"
        
        ' Auto-size columns
        ws.Columns("A:C").AutoFit
        
        ' Log sheet creation
        #If DEBUG Then
            Module_Dashboard.DebugLog "EnsureConfigSheet", "Created Config sheet with default settings."
        #EndIf
    Else
        ' Config sheet exists - verify required settings are present
        Dim targetFound As Boolean
        targetFound = False
        
        ' Check for Pipeline Target setting
        Dim r As Long
        For r = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            If ws.Cells(r, "A").Value = TARGET_SETTING_NAME Then
                targetFound = True
                Exit For
            End If
        Next r
        
        ' Add setting if not found
        If Not targetFound Then
            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
            ws.Range("A" & lastRow).Value = TARGET_SETTING_NAME
            ws.Range("B" & lastRow).Value = TARGET_DEFAULT_VALUE
            ws.Range("B" & lastRow).NumberFormat = "$#,##0"
            ws.Range("C" & lastRow).Value = "Placeholder target value. Adjust to your specific quarterly/annual revenue goals."
            
            ' Log setting addition
            #If DEBUG Then
                Module_Dashboard.DebugLog "EnsureConfigSheet", "Added missing Pipeline Target setting to Config sheet."
            #EndIf
        End If
    End If
End Sub

' Function to read historical performance data
Private Function ReadHistoricalData() As Variant
    ' Read data from PerfHistory sheet if it exists
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim histData As Variant
    Const HISTORY_SHEET_NAME As String = "PerfHistory"
    
    ' Check if history sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HISTORY_SHEET_NAME)
    If Err.Number <> 0 Or ws Is Nothing Then
        ' Sheet doesn't exist - return empty array
        ReadHistoricalData = Empty
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Get data range
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Check if we have any data (beyond headers)
    If lastRow <= 1 Then
        ReadHistoricalData = Empty
        Exit Function
    End If
    
    ' Read the data
    Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
    histData = dataRange.Value
    
    ' Return the data array
    ReadHistoricalData = histData
End Function

' Function to update historical performance data
Private Sub UpdateHistoricalMetrics(ByRef model As Scripting.Dictionary)
    ' Update or create PerfHistory sheet with current metrics
    Const HISTORY_SHEET_NAME As String = "PerfHistory"
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentDate As Date: currentDate = Date
    
    ' Check if history sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HISTORY_SHEET_NAME)
    If Err.Number <> 0 Or ws Is Nothing Then
        ' Create new history sheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = HISTORY_SHEET_NAME
        
        ' Add headers
        ws.Range("A1:J1").Value = Array("Date", "Total Quotes", "Open Quotes", "Converted", _
                                         "Pipeline Value", "Conversion Rate", "Avg Cycle Time", _
                                         "First F/U Count", "Second F/U Count", "Pending Count")
        With ws.Range("A1:J1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = SqrctColors.DarkBlue
            .Font.Color = SqrctColors.TextLight
        End With
        
        ' Format columns
        ws.Columns("A").NumberFormat = "mm/dd/yyyy"
        ws.Columns("E").NumberFormat = "$#,##0.00"
        ws.Columns("F").NumberFormat = "0.0%"
        
        ' Auto-size columns
        ws.Columns("A:J").AutoFit
        
        #If DEBUG Then
            Module_Dashboard.DebugLog "UpdateHistoricalMetrics", "Created PerfHistory sheet with headers."
        #EndIf
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Find last row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 1 Then lastRow = 1 ' Start at row 2 if no data yet
    
    ' Check if we already have an entry for today
    Dim today As Boolean: today = False
    Dim r As Long
    
    If lastRow > 1 Then
        For r = 2 To lastRow
            If DateValue(ws.Cells(r, "A").Value) = DateValue(currentDate) Then
                today = True
                lastRow = r ' Overwrite today's row
                Exit For
            End If
        Next r
    End If
    
    ' Add new row or update today's entry
    If Not today Then lastRow = lastRow + 1
    
    ' Write metrics to history sheet
    ws.Cells(lastRow, "A").Value = currentDate
    ws.Cells(lastRow, "B").Value = model("totalRows")
    ws.Cells(lastRow, "C").Value = model("openRows")
    ws.Cells(lastRow, "D").Value = model("convRows")
    ws.Cells(lastRow, "E").Value = model("sumOpenAmt")
    ws.Cells(lastRow, "F").Value = model("conversionRate")
    ws.Cells(lastRow, "G").Value = model("avgCycleTime")
    ws.Cells(lastRow, "H").Value = model("firstFURows")
    ws.Cells(lastRow, "I").Value = model("secondFURows")
    ws.Cells(lastRow, "J").Value = model("pendingRows")
    
    ' Format the row
    ws.Range(ws.Cells(lastRow, "A"), ws.Cells(lastRow, "J")).Interior.Color = RGB(240, 240, 240)
    
    #If DEBUG Then
        Module_Dashboard.DebugLog "UpdateHistoricalMetrics", IIf(today, "Updated", "Added") & " history entry for " & Format(currentDate, "mm/dd/yyyy")
    #EndIf

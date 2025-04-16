// Query Name: OrderConf_Combined_Base
let
    // ========================================================================
    // Step 1: Define Source Queries and Expected Schema
    // ========================================================================
    SourceQueryNames = {
        "OrderConf_New",
        "OrderConf_Open",
        "OrderConf_AsAvailable",
        "OrderConf_Hold",
      //"OrderConf_Closed", // Keep commented out
        "OrderConf_Declined"
    },
    ExpectedSchema = type table [
        FileName = text, FolderLocation = text, FileExtension = text,
        FileDateModified = nullable datetime, FileDateCreated = nullable datetime
        // Add other columns if your source queries consistently provide them
    ],

    // ========================================================================
    // Step 2: Safely Get Data from Each Source Query
    // ========================================================================
    SafeGetTable = (queryName as text) as table =>
        let
            SourceAttempt = try Expression.Evaluate("#""" & queryName & """", #shared),
            Result = if SourceAttempt[HasError] or not (Value.Is(SourceAttempt[Value], type table)) then
                        Table.FromRows({}, ExpectedSchema)
                     else
                        // Ensure conformance to schema even on success
                        Table.Combine({Table.FromRows({}, ExpectedSchema), SourceAttempt[Value]})
        in
            Result,
    ListOfSourceTables = List.Transform(SourceQueryNames, each SafeGetTable(_)),

    // ========================================================================
    // Step 3: Combine the Safely Retrieved Tables
    // ========================================================================
    CombinedData = Table.Combine(ListOfSourceTables),

    // ========================================================================
    // Step 4: Select and Type Core Columns Needed for Processing
    // ========================================================================
    CoreColumnsToKeep = {"FileName", "FolderLocation", "FileExtension", "FileDateModified", "FileDateCreated"},
    EnsureCoreColumns = List.Accumulate(
        CoreColumnsToKeep,
        CombinedData,
        (tableState, columnName) =>
            if Table.HasColumns(tableState, {columnName}) then tableState
            else Table.AddColumn(tableState, columnName, each null)
    ),
    SelectedCore = Table.SelectColumns(EnsureCoreColumns, CoreColumnsToKeep),
    TypedCore = Table.TransformColumnTypes(SelectedCore, {
        {"FileName", type text},
        {"FolderLocation", type text},
        {"FileExtension", type text},
        {"FileDateModified", type nullable datetime},
        {"FileDateCreated", type nullable datetime}
    }),

    // ========================================================================
    // Step 5: Extract Identifiers (with Error Handling per Row)
    // ========================================================================

    // --- Extract Primary Document Number (SCxxxx / BSxxxx) ---
    AddPrimaryDocNumber = Table.AddColumn(TypedCore, "PrimaryDocNumber", each
        let
            fileName = [FileName],
            attempt = try
                if fileName = null then null else
                let
                  upperFileName = Text.Upper(fileName),
                  separators = {" ", "_", ".", "(", ")", "#", "-"},
                  parts = Text.SplitAny(upperFileName, Text.Combine(separators)),
                  CheckPart = (part as text) as logical =>
                      let
                          startsWithPrefix = Text.StartsWith(part, "SC") or Text.StartsWith(part, "BS"),
                          hasMinLength = Text.Length(part) >= 5
                      in startsWithPrefix and hasMinLength,
                  matchingParts = List.Select(parts, each CheckPart(_)),
                  result = if List.Count(matchingParts) > 0 then List.First(matchingParts) else null
                in result
            otherwise null
        in attempt
    , type text),

    // --- Extract TRF/ERF Number ---
    // MODIFIED: Removed fallback logic. Only extracts number if "TRF" or "ERF" marker is found.
    AddTempFormNumber = Table.AddColumn(AddPrimaryDocNumber, "_TempFormNumber", each
        let
            fileName = [FileName],
            attempt = try
                if fileName = null then null else
                let
                    upperFileName = Text.Upper(fileName),
                    markers = {"TRF", "ERF"},
                    markerPositions = List.Transform(markers, each Text.PositionOf(upperFileName, _)),
                    validPositions = List.Select(markerPositions, each _ >= 0),
                    firstMarkerPos = if List.Count(validPositions) > 0 then List.Min(validPositions) else -1,
                    result = if firstMarkerPos >= 0 then // Only proceed if a marker was found
                        let
                           potentialNumberString = Text.TrimStart(Text.Middle(upperFileName, firstMarkerPos + 3), {"_", " ", "#"}),
                           extractedDigits = if potentialNumberString <> null then Text.Select(potentialNumberString, {"0".."9"}) else null,
                           limitedDigits = if extractedDigits <> null then Text.Start(extractedDigits, 10) else null
                        in if limitedDigits <> null and Text.Length(limitedDigits) >= 6 then limitedDigits else null // Check length 6-10
                    else
                       null // If no marker found, result is null
                in result
            otherwise null
        in attempt
    , type text),

    // --- Extract Form Type (TRF or ERF) ---
    AddFormType = Table.AddColumn(AddTempFormNumber, "FormType", each
        let
            fileName = [FileName],
            attempt = try
                if fileName = null then null else
                let
                    upperFileName = Text.Upper(fileName),
                    markers = {"TRF", "ERF"},
                    markerPositions = List.Transform(markers, (m) => {m, Text.PositionOf(upperFileName, m)}),
                    validMarkers = List.Select(markerPositions, each _{1} >= 0),
                    firstMarkerPos = if List.Count(validMarkers) > 0 then List.Min(List.Transform(validMarkers, each _{1})) else -1,
                    foundMarkerRecord = List.FirstN(List.Select(validMarkers, each _{1} = firstMarkerPos), 1),
                    formTypeResult = if List.Count(foundMarkerRecord) > 0 then foundMarkerRecord{0}{0} else null
                in formTypeResult
            otherwise null
        in attempt
    , type text),

    // ========================================================================
    // Step 6: Rename Temp Column and Select Final Output
    // ========================================================================
    RenameTempCol = Table.RenameColumns(AddFormType, {{"_TempFormNumber", "FormNumber"}}),

    FinalOutput = Table.SelectColumns(RenameTempCol, {
        "PrimaryDocNumber", "FormNumber", "FormType", // Identifiers
        "FolderLocation", "FileName", "FileExtension", "FileDateModified", "FileDateCreated" // Context
    })

in
    FinalOutput


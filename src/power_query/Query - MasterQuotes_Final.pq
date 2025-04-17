let
    //////////////////////////////////////////////////////////////////////////
    // STEP A: Reference your raw data (MasterQuotes_Raw)
    //////////////////////////////////////////////////////////////////////////
    Source = MasterQuotes_Raw,

    //////////////////////////////////////////////////////////////////////////
    // STEP B: Identify the overall LatestPullDate from the entire table
    //////////////////////////////////////////////////////////////////////////
    LatestPullDate = List.Max(Source[Date Pulled]),

    //////////////////////////////////////////////////////////////////////////
    // STEP C: Sort by Date Pulled (oldest to newest) then Document Amount (highest to lowest)
    //////////////////////////////////////////////////////////////////////////
    Sorted = Table.Sort(
        Source,
        {
            {"Date Pulled", Order.Ascending},
            {"Document Amount", Order.Descending}
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP D: Filter out rows where [User To Enter] is one of the undesired names
    //////////////////////////////////////////////////////////////////////////
    Filtered = Table.SelectRows(
        Sorted,
        each not List.Contains(
            {"rosannai", "kathrynh", "ryanz", "alexandriaf"},
            Text.Lower([User To Enter])
        )
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP E: Group by Document Number; capture:
    //         - FirstPullForDoc = min date
    //         - MaxPullForDoc   = max date
    //         - SubRows = the rest
    //////////////////////////////////////////////////////////////////////////
    Grouped = Table.Group(
        Filtered,
        {"Document Number"},
        {
            {"SubRows", each Table.RemoveColumns(_, {"Document Number"})},
            {"FirstPullForDoc", each List.Min([Date Pulled])},
            {"MaxPullForDoc", each List.Max([Date Pulled])}
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP F: Mark "MissingInLatest" if max date < overall LatestPullDate
    //////////////////////////////////////////////////////////////////////////
    WithMissing = Table.AddColumn(
        Grouped,
        "MissingInLatest",
        each [MaxPullForDoc] < LatestPullDate,
        type logical
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP G: For each sub-table, add an "Occurrence" and a column carrying the
    //         earliest date (FirstPullForDoc) so we can expand it later.
    //////////////////////////////////////////////////////////////////////////
    WithIndex = Table.AddColumn(
        WithMissing,
        "IndexedSub",
        each
            let
                tbl = [SubRows],
                withOcc = Table.AddIndexColumn(tbl, "Occurrence", 1, 1, Int64.Type),
                withFirst = Table.AddColumn(withOcc, "GroupFirstPulled", (x) => [FirstPullForDoc])
            in
                withFirst
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP H: Expand sub-tables (EXCLUDING the original Date Pulled),
    //         bringing along "GroupFirstPulled" and "DataSource" as well
    //////////////////////////////////////////////////////////////////////////
    Expanded = Table.ExpandTableColumn(
        WithIndex,
        "IndexedSub",
        {
            "Document Date",
            "Customer Number",
            "Customer Name",
            "User To Enter",
            "Document Amount",
            "Salesperson ID",
            "Occurrence",
            "Historic Stage",
            "GroupFirstPulled",
            "DataSource"
        },
        {
            "Document Date",
            "Customer Number",
            "Customer Name",
            "User To Enter",
            "Document Amount",
            "Salesperson ID",
            "Occurrence",
            "Historic Stage",
            "GroupFirstPulled",
            "DataSource"
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP I: Add "Auto Stage" column based on Historic Stage and Occurrence
    //////////////////////////////////////////////////////////////////////////
    AddStage = Table.AddColumn(
        Expanded,
        "Auto Stage",
        each
            let
                hist = [Historic Stage],
                occ = [Occurrence]
            in
                if hist = null or Text.Trim(hist) = "" then
                    if occ = 1 then "First F/U"
                    else if occ = 2 then "Second F/U"
                    else if occ = 3 then "Third F/U"
                    else null // Changed from "Long-Term F/U" for occ >= 4
                else
                    hist,
        type text
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP J: Add "Auto Note" if Historic Stage is blank and MissingInLatest is true
    //////////////////////////////////////////////////////////////////////////
    AddNote = Table.AddColumn(
        AddStage,
        "Auto Note",
        each
            if ([Historic Stage] = null or Text.Trim([Historic Stage]) = "") then
                if [MissingInLatest] then "Confirm converted/voided."
                else null
            else null,
        type text
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP K: Rename columns (rename GroupFirstPulled to First Date Pulled, etc.)
    //////////////////////////////////////////////////////////////////////////
    RenameCols = Table.RenameColumns(
        AddNote,
        {
            {"Auto Stage", "AutoStage"},
            {"Auto Note", "AutoNote"},
            {"MissingInLatest", "IsMissing"},
            {"GroupFirstPulled", "First Date Pulled"}
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP K2: Clean AutoNote column to ensure blanks are empty strings and type is text
    //////////////////////////////////////////////////////////////////////////
    CleanAutoNote = Table.ReplaceValue(
         RenameCols, 
         null, 
         "", 
         Replacer.ReplaceValue, 
         {"AutoNote"}
    ),
    EnsureAutoNoteText = Table.TransformColumnTypes(
         CleanAutoNote,
         {{"AutoNote", type text}}
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP L: Group by Document Number with custom aggregator to choose final row 
    //         based on DataSource and stage:
    //         1) If any CSV row has a non-blank stage, pick the one w/ highest Occurrence
    //         2) Else if any Legacy row is "Converted", pick that w/ highest Occurrence
    //         3) Else fallback to highest Occurrence overall
    //////////////////////////////////////////////////////////////////////////
    GroupedFinal = Table.Group(
        EnsureAutoNoteText,
        {"Document Number"},
        {
            {
                "FinalRow",
                (tbl) =>
                    let
                        csvRows         = Table.SelectRows(tbl, each [DataSource] = "CSV"),
                        legacyRows      = Table.SelectRows(tbl, each [DataSource] = "Legacy"),
                        csvNonBlank     = Table.SelectRows(csvRows, each [Historic Stage] <> null and Text.Trim([Historic Stage]) <> ""),
                        chosenCSV       = if Table.RowCount(csvNonBlank) > 0 then Table.Max(csvNonBlank, "Occurrence") else null,
                        legacyConverted = Table.SelectRows(legacyRows, each [Historic Stage] = "Converted"),
                        chosenLegacy    = if Table.RowCount(legacyConverted) > 0 then Table.Max(legacyConverted, "Occurrence") else null,
                        fallback        = Table.Max(tbl, "Occurrence"),
                        finalPick       = if chosenCSV <> null then chosenCSV else if chosenLegacy <> null then chosenLegacy else fallback
                    in
                        finalPick
            }
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP M: Expand the selected final row back into columns
    //////////////////////////////////////////////////////////////////////////
    ExpandedFinal = Table.ExpandRecordColumn(
        GroupedFinal,
        "FinalRow",
        {
            "First Date Pulled",
            "Document Date",
            "Customer Number",
            "Customer Name",
            "User To Enter",
            "Document Amount",
            "Salesperson ID",
            "Occurrence",
            "Historic Stage",
            "AutoStage",
            "AutoNote",
            "IsMissing",
            "DataSource"
        },
        {
            "First Date Pulled",
            "Document Date",
            "Customer Number",
            "Customer Name",
            "User To Enter",
            "Document Amount",
            "Salesperson ID",
            "Occurrence",
            "Historic Stage",
            "AutoStage",
            "AutoNote",
            "IsMissing",
            "DataSource"
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP N: Rename Occurrence to Pull Count
    //////////////////////////////////////////////////////////////////////////
    RenameOccurrence = Table.RenameColumns(ExpandedFinal, {{"Occurrence", "Pull Count"}}),

    //////////////////////////////////////////////////////////////////////////
    // STEP O: Reorder columns to include Pull Count
    //////////////////////////////////////////////////////////////////////////
    ReorderedColumns = Table.ReorderColumns(
        RenameOccurrence, // Changed source from RemoveTemp
        {
            "First Date Pulled",
            "Document Number",
            "Document Date",
            "Customer Number",
            "Customer Name",
            "User To Enter",
            "Document Amount",
            "Salesperson ID",
            "Pull Count", // Added Pull Count
            "Historic Stage",
            "AutoStage",
            "AutoNote",
            "IsMissing",
            "DataSource"
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP P: Convert date columns to date type
    //////////////////////////////////////////////////////////////////////////
    DatesFixed = Table.TransformColumnTypes(
        ReorderedColumns, // Source is correct (was ReorderedColumns)
        {
            {"First Date Pulled", type date},
            {"Document Date", type date}
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP Q: Clean "Document Number" safely (avoid offset errors)
    //////////////////////////////////////////////////////////////////////////
    CleanDocNumber = Table.TransformColumns(
        DatesFixed, // Source is correct (was DatesFixed)
        {
            {
                "Document Number",
                each
                    let
                        original     = _,
                        len          = Text.Length(original),
                        prefix       = if len >= 5 then Text.Start(original, 5) else original,
                        numericPart  = if len > 5 then Text.Range(original, 5) else null,
                        parsed       = try Number.From(numericPart) otherwise null,
                        padded       = if parsed <> null then Text.PadStart(Text.From(parsed), 5, "0") else ""
                    in
                        if prefix = "BSMOQ" then
                            original
                        else if parsed = null then
                            original
                        else
                            prefix & padded,
                type text
            }
        }
    ),
    ChangedType = Table.TransformColumnTypes(CleanDocNumber, {{"Document Amount", Currency.Type}}), // Source is correct (was CleanDocNumber)

    //////////////////////////////////////////////////////////////////////////
    // STEP R: Final sort: by First Date Pulled (oldest to newest) then Document Number (alphabetically)
    //////////////////////////////////////////////////////////////////////////
    FinalSorted = Table.Sort(
        ChangedType, // Source is correct (was ChangedType)
        {
            {"First Date Pulled", Order.Ascending},
            {"Document Number", Order.Ascending}
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP S: Add Explicit Blank Row Filtering
    //////////////////////////////////////////////////////////////////////////
    RemoveBlankRows = Table.SelectRows(
        FinalSorted, // Source is correct (was FinalSorted)
        each ([Document Number] <> null and [Document Number] <> "") and // Standard null/empty check
             ([Document Date] <> null or [Customer Name] <> null or [Document Amount] <> null)
    )

in
    RemoveBlankRows


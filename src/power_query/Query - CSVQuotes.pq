let
    //////////////////////////////////////////////////////////////////////////
    // STEP 1: Get all files in the target folder
    //////////////////////////////////////////////////////////////////////////
     Source = Folder.Files("R:\Projects\Strategic Quote Recovery and Conversion Tracker (SQRCT)\30DayExports_DailyPull"),

    //////////////////////////////////////////////////////////////////////////
    // STEP 2: Filter to only .csv files
    //////////////////////////////////////////////////////////////////////////
    OnlyCSVs = Table.SelectRows(Source, each [Extension] = ".csv"),

    //////////////////////////////////////////////////////////////////////////
    // STEP 3: For each CSV, parse with Csv.Document,
    //         then skip the first row (CSV's own headers)
    //////////////////////////////////////////////////////////////////////////
    AddParsedCSV = Table.AddColumn(
        OnlyCSVs,
        "Data",
        each
            let
                rawCSV = Csv.Document(
                    [Content],
                    [
                        Delimiter = ",",
                        Encoding = 65001,
                        QuoteStyle = QuoteStyle.None
                    ]
                ),
                // skip the very first row (header line)
                removeHeaderRow = Table.Skip(rawCSV, 1)
            in
                removeHeaderRow
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 4: Expand that "Data" column
    // Adjust if CSV has more/fewer columns
    //////////////////////////////////////////////////////////////////////////
    ExpandRows = Table.ExpandTableColumn(
        AddParsedCSV,
        "Data",
        {
            "Column1","Column2","Column3","Column4",
            "Column5","Column6","Column7","Column8","Column9"
        },
        {
            "Batch Number","Document Date","Document Number",
            "Customer Number","Customer Name","Document Type",
            "User To Enter","Document Amount","Salesperson ID from Customer Master"
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 5: Remove columns not needed (Batch Number, Document Type)
    //////////////////////////////////////////////////////////////////////////
    RemoveUnneeded = Table.RemoveColumns(
        ExpandRows,
        {"Batch Number","Document Type"}
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 6: Rename columns to match your final schema
    //////////////////////////////////////////////////////////////////////////
    RenameCols = Table.RenameColumns(
        RemoveUnneeded,
        {
            {"Document Date", "Document Date"},
            {"Document Number", "Document Number"},
            {"Customer Number", "Customer Number"},
            {"Customer Name", "Customer Name"},
            {"User To Enter", "User To Enter"},
            {"Document Amount", "Document Amount"},
            {"Salesperson ID from Customer Master", "Salesperson ID"}
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 7: Parse "Date Pulled" from the file name (e.g., SQRCT_20250210.csv)
    //////////////////////////////////////////////////////////////////////////
    AddDatePulled = Table.AddColumn(
        RenameCols,
        "Date Pulled",
        each
            let
                nameNoExt = Text.BeforeDelimiter([Name], ".csv"),  // e.g. "SQRCT_20250210"
                datePart = Text.AfterDelimiter(nameNoExt, "_")     // e.g. "20250210"
            in
                // Convert "20250210" => 2025-02-10
                Date.FromText(datePart, [Format="yyyyMMdd"])
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 8: Convert date columns to Date type
    //////////////////////////////////////////////////////////////////////////
    SetDateType = Table.TransformColumnTypes(
        AddDatePulled,
        {
            {"Date Pulled", type date},
            {"Document Date", type date}
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 9: Add F/U columns not in CSV
    //////////////////////////////////////////////////////////////////////////
    AddFUFld1 = Table.AddColumn(SetDateType, "First F/U", each null),
    AddFUFld2 = Table.AddColumn(AddFUFld1, "Second F/U", each null),
    AddFUFld3 = Table.AddColumn(AddFUFld2, "Third F/U", each null),
    AddFUFld4 = Table.AddColumn(AddFUFld3, "Long-Term F/U", each null),
    AddEmail = Table.AddColumn(AddFUFld4, "Email Contact", each null),

    //////////////////////////////////////////////////////////////////////////
    // STEP 10: Sort by Date Pulled, then Document Date
    //////////////////////////////////////////////////////////////////////////
    SortedByDate = Table.Sort(
        AddEmail,
        {
            {"Date Pulled", Order.Ascending},
            {"Document Date", Order.Ascending}
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 11: Clean "Document Amount" (remove $ and commas)
    //////////////////////////////////////////////////////////////////////////
    RemoveDollars = Table.ReplaceValue(
        SortedByDate,
        "$",
        "",
        Replacer.ReplaceText,
        {"Document Amount"}
    ),
    RemoveCommas = Table.ReplaceValue(
        RemoveDollars,
        ",",
        "",
        Replacer.ReplaceText,
        {"Document Amount"}
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 12: Convert "Document Amount" to Currency
    //////////////////////////////////////////////////////////////////////////
    FinalCurrency = Table.TransformColumnTypes(
        RemoveCommas,
        {{"Document Amount", Currency.Type}}
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 13: Remove leftover folder columns
    // (Content, Name, Extension, Dates, Attributes, etc.)
    //////////////////////////////////////////////////////////////////////////
    RemoveFileCols = Table.RemoveColumns(
        FinalCurrency,
        {
            "Content","Name","Extension",
            "Date accessed","Date modified","Date created",
            "Attributes","Folder Path"
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 14: Reorder columns in final order
    //////////////////////////////////////////////////////////////////////////
    ReorderCols = Table.ReorderColumns(
        RemoveFileCols,
        {
            "Date Pulled",
            "Document Date",
            "Document Number",
            "Customer Number",
            "Customer Name",
            "User To Enter",
            "Document Amount",
            "Salesperson ID",
            "First F/U",
            "Second F/U",
            "Third F/U",
            "Long-Term F/U",
            "Email Contact"
        }
    )

    // If you want to remove those new F/U or Email columns, 
    // add another line:
    //    Table.RemoveColumns(ReorderCols, {"First F/U", "Second F/U", 
    //                                     "Third F/U", "Long-Term F/U", "Email Contact"})

in
    //////////////////////////////////////////////////////////////////////////
    // Final Output
    //////////////////////////////////////////////////////////////////////////
    ReorderCols
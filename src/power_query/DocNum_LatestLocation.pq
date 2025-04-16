// Query Name: DocNum_LatestLocation
let
    // 1. Reference the base query containing all file details and extracted identifiers
    //    Ensure 'OrderConf_Combined_Base' is the correct name of your base query.
    Source = #"OrderConf_Combined_Base",

    // 2. REVERTED: Filter for rows where PrimaryDocNumber is populated
    //    This ensures we find the latest location for ANY document with an SC/BS number,
    //    regardless of whether a FormNumber was also found.
    FilterValidDocNums = Table.SelectRows(Source, each ([PrimaryDocNumber] <> null)),

    // 3. Group by the Primary Document Number
    //    This gathers all files related to the same document number.
    GroupedByDocNum = Table.Group(
        FilterValidDocNums, // Use the original filtered table
        {"PrimaryDocNumber"}, // Column(s) to group by
        {
            // For each group (each unique PrimaryDocNumber), find the row with the latest 'FileDateModified'
            {"LatestRecord", each
                let
                    // Sort the current group's table descending by modification date
                    SortedTable = Table.Sort(_, {{"FileDateModified", Order.Descending}}),
                    // Get the first row (which is the latest)
                    LatestRow = Table.FirstN(SortedTable, 1)
                in
                    LatestRow,
             // Define expected table structure within the group (ensure this matches columns in FilterValidDocNums)
             type table [PrimaryDocNumber=nullable text, FormNumber=nullable text, FormType=nullable text, FolderLocation=text, FileName=text, FileExtension=text, FileDateModified=nullable datetime, FileDateCreated=nullable datetime]
            }
        }
    ),

    // 4. Expand the 'LatestRecord' column
    //    Expand details of the most recent file for each PrimaryDocNumber found.
    ExpandLatestRecord = Table.ExpandTableColumn(
        GroupedByDocNum,
        "LatestRecord",
        {"FolderLocation", "FileName", "FileDateModified"}, // Columns to bring out from the latest record
        {"MostRecent_FolderLocation", "MostRecent_FileName", "MostRecent_FileDateModified"} // New names for the expanded columns
    ),

    // 5. Select Final Columns (Optional but good practice)
    //    Output will now contain all PrimaryDocNumbers found in the source data and their latest location info.
    SelectFinal = Table.SelectColumns(ExpandLatestRecord, {"PrimaryDocNumber", "MostRecent_FolderLocation", "MostRecent_FileName", "MostRecent_FileDateModified"})

in
    SelectFinal



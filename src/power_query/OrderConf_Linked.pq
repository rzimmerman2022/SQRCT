// Query Name: OrderConf_Linked
let
    // 1. Start with the base data containing all files and extracted identifiers
    // Ensure 'OrderConf_Combined_Base' is the correct name of your base query
    Source = #"OrderConf_Combined_Base",

    // 2. Reference the mapping table created in the 'Map_Form_DocNum' query
    // Ensure 'Map_Form_DocNum' is the correct name of your mapping query
    MapTable = Map_Form_DocNum,

    // 3. Merge (Left Outer Join) the base data with the map table
    // This looks up the PrimaryDocNumber associated with each row's FormNumber
    MergeMap = Table.NestedJoin(
        Source, {"FormNumber"},         // Key column(s) from the Source table (left)
        MapTable, {"FormNumber"},       // Key column(s) from the MapTable (right)
        "MapResult",                    // Name for the new column holding the nested table result
        JoinKind.LeftOuter              // Keep all rows from Source, add matches from MapTable
    ),

    // 4. Expand the merged column to get the mapped PrimaryDocNumber
    // We only need the PrimaryDocNumber column from the map result.
    ExpandMappedDocNum = Table.ExpandTableColumn(
        MergeMap,
        "MapResult",                  // The column containing the nested tables
        {"PrimaryDocNumber"},         // The column(s) to expand from the nested table
        {"MappedPrimaryDocNum"}       // The name(s) for the newly expanded column(s)
    ),

    // 5. Determine the Final Transaction Group ID
    // Apply priority logic to establish a consistent ID for related documents.
    AddTransactionGroupID = Table.AddColumn(ExpandMappedDocNum, "TransactionGroupID", each
        // Priority 1: Use the PrimaryDocNumber if found directly on this file
        if [PrimaryDocNumber] <> null then
            [PrimaryDocNumber]
        // Priority 2: If no direct PrimaryDocNumber, use the one found via the FormNumber map
        else if [MappedPrimaryDocNum] <> null then
            [MappedPrimaryDocNum]
        // Priority 3: If no PrimaryDocNumber link found, use the FormNumber itself
        // Prefixing with FormType helps distinguish it (e.g., "TRF:1234567")
        else if [FormNumber] <> null then
             ([FormType] ?? "FORM") & ":" & [FormNumber]
        // Priority 4: If no identifiers found at all
        else
            null,
    type text), // Set the type for the new Group ID column

    // 6. Select and Order the Final Columns for Output
    // Choose the columns you want in the final result table.
    SelectOutputColumns = Table.SelectColumns(AddTransactionGroupID, {
        // --- Core Identifiers & Grouping ---
        "TransactionGroupID",   // The unified ID for grouping related files
        "PrimaryDocNumber",     // The extracted SC/BS number (Quote/Order)
        "FormNumber",           // The extracted TRF/ERF number
        "FormType",             // "TRF" or "ERF"
        // --- File Context ---
        "FolderLocation",       // e.g., "1. New Orders"
        "FileName",             // Original filename
        "FileExtension",
        "FileDateModified",
        "FileDateCreated"
        // Add any other columns from 'Source' you wish to keep
        // Note: "MappedPrimaryDocNum" is automatically excluded here unless listed
    }),

    // 7. Optional: Sort the Final Table
    // Sorting by the TransactionGroupID makes it easy to see related files together.
    FinalSort = Table.Sort(SelectOutputColumns, {
        {"TransactionGroupID", Order.Ascending},
        {"FileDateModified", Order.Ascending}, // Secondary sort by date
        {"FileName", Order.Ascending}          // Tertiary sort by name
    })

in
    FinalSort



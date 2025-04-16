// Query Name: Map_Form_DocNum
let
    // 1. Reference the base query which contains extracted identifiers
    // Ensure 'OrderConf_Combined_Base' is the correct name of your base query
    Source = #"OrderConf_Combined_Base",

    // 2. Filter for rows that contain BOTH types of identifiers
    // These rows act as the bridge/link between FormNumber and PrimaryDocNumber.
    // This step finds files whose names contain both an SC/BS number AND a TRF/ERF number.
    FilterLinkingRows = Table.SelectRows(Source, each
        ([FormNumber] <> null and [PrimaryDocNumber] <> null)
    ),

    // 3. Select only the identifier columns needed for the map
    // We only need the pairs of linked numbers.
    SelectMappingColumns = Table.SelectColumns(FilterLinkingRows, {"FormNumber", "PrimaryDocNumber"}),

    // 4. Remove duplicate mapping pairs to create a clean lookup table
    // This ensures each unique FormNumber <-> PrimaryDocNumber link appears only once in our map.
    RemoveDuplicates = Table.Distinct(SelectMappingColumns)

in
    // Output the final mapping table
    RemoveDuplicates

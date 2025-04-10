let
    //////////////////////////////////////////////////////////////////////////
    // STEP 1: Load the "ExistingQuotes" Excel table
    //////////////////////////////////////////////////////////////////////////
    Source = Excel.CurrentWorkbook(){[Name="ExistingQuotes"]}[Content],

    //////////////////////////////////////////////////////////////////////////
    // STEP 2: Convert column types
    //////////////////////////////////////////////////////////////////////////
    #"Changed Type" = Table.TransformColumnTypes(
        Source,
        {
            {"Date Pulled",     type date},
            {"Document Date",   type date},
            {"Document Number", type text},
            {"Customer Number", type text},
            {"Customer Name",   type text},
            {"User To Enter",   type text},
            {"Document Amount", type number},
            {"Salesperson ID",  type text},

            // Old follow-up columns (to be removed):
            {"First F/U",       type text},
            {"Second F/U",      type text},
            {"Third F/U",       type text},
            {"Long-Term F/U",   type text},
            {"Email Contact",   type text},

            // The single new override column:
            {"Historic Stage",  type text}
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 3: Remove old follow-up columns
    //////////////////////////////////////////////////////////////////////////
    #"Removed Columns" = Table.RemoveColumns(
        #"Changed Type",
        {
            "First F/U",
            "Second F/U",
            "Third F/U",
            "Long-Term F/U",
            "Email Contact"
        }
    ),

    //////////////////////////////////////////////////////////////////////////
    // STEP 4: Convert Document Amount to currency (optional)
    //////////////////////////////////////////////////////////////////////////
    #"Changed Type1" = Table.TransformColumnTypes(
        #"Removed Columns",
        {
            {"Document Amount", Currency.Type}
        }
    )
in
    #"Changed Type1"
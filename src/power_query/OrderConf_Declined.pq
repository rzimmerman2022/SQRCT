// Suggested Query Name: OrderConf_Declined
let
    // 1. CONFIGURATION - Pointing to original folder name
    TargetFolderPath = "S:\Order Management\Order Confirmation\6. Declined Orders",
    FolderLocationName = "6. Declined Orders", // Reflects actual folder name

    // --- Start Query ---

    // 2. GET FILES FROM THE SPECIFIC FOLDER
    Source = try Folder.Files(TargetFolderPath) otherwise null,
    CheckSource = if Source = null then Table.FromRows({}, type table [QuoteNumber=text, FormNumber=text, FormType=text, FolderLocation=text, FileName=text, FileExtension=text, FileDateModified=nullable datetime, FileDateCreated=nullable datetime]) else Source,

    // 3. EXTRACT KEY IDENTIFIERS
    AddQuoteNumber = if Table.RowCount(CheckSource) > 0 then Table.AddColumn(CheckSource, "QuoteNumber", each let fileName = [Name] ?? "", upperFileName = Text.Upper(fileName), separators = {" ", "_", ".", "(", ")", "#", "-"}, parts = Text.SplitAny(upperFileName, Text.Combine(separators)), CheckPart = (part as text) as logical => let startsWithPrefix = Text.StartsWith(part, "SC") or Text.StartsWith(part, "BS"), hasMinLength = Text.Length(part) >= 5 in startsWithPrefix and hasMinLength, matchingParts = List.Select(parts, each CheckPart(_)), result = if List.Count(matchingParts) > 0 then matchingParts{0} else null in result, type text) else Table.AddColumn(CheckSource, "QuoteNumber", each null, type text),
    AddFormNumber = if Table.RowCount(AddQuoteNumber) > 0 then Table.AddColumn(AddQuoteNumber, "FormNumber", each let fileName = Text.Upper([Name] ?? ""), markers = {"TRF", "ERF"}, markerInfo = List.Transform(markers, each {_, Text.PositionOf(fileName, _)}), foundMarkers = List.Select(markerInfo, each _{1} >= 0), firstMarkerPos = if List.Count(foundMarkers) > 0 then List.Min(List.Transform(foundMarkers, each _{1})) else -1, potentialNumberString = if firstMarkerPos >= 0 then Text.TrimStart(Text.Middle(fileName, firstMarkerPos + 3), {"_", " ", "#"}) else null, extractedDigits = if potentialNumberString <> null then Text.Select(potentialNumberString, {"0".."9"}) else null, limitedDigits = if extractedDigits <> null then Text.Start(extractedDigits, 10) else null, result = if limitedDigits <> null and Text.Length(limitedDigits) >= 6 then limitedDigits else null, fallbackResult = if result = null then let parts = Text.SplitAny(fileName, "_ .()#"), numericParts = List.Select(parts, each Value.Is(Value.FromText(_), type number) and Text.Length(_) >= 6 and Text.Length(_) <= 10) in if List.Count(numericParts) > 0 then List.First(numericParts) else null else result in fallbackResult, type text) else Table.AddColumn(AddQuoteNumber, "FormNumber", each null, type text),
    AddFormType = if Table.RowCount(AddFormNumber) > 0 then Table.AddColumn(AddFormNumber, "FormType", each let fileName = Text.Upper([Name] ?? ""), markers = {"TRF", "ERF"}, markerInfo = List.Transform(markers, each {_, Text.PositionOf(fileName, _)}), foundMarkers = List.Select(markerInfo, each _{1} >= 0), firstMarkerPos = if List.Count(foundMarkers) > 0 then List.Min(List.Transform(foundMarkers, each _{1})) else -1, foundMarkerRecord = List.FirstN(List.Select(foundMarkers, each _{1} = firstMarkerPos), 1), formTypeResult = if List.Count(foundMarkerRecord) > 0 then foundMarkerRecord{0}{0} else null in formTypeResult, type text) else Table.AddColumn(AddFormNumber, "FormType", each null, type text),
    AddFolderLocation = Table.AddColumn(AddFormType, "FolderLocation", each FolderLocationName, type text),

    // 4. FINAL CLEANUP & SELECT
    RemoveErrorRows = Table.SelectRows(AddFolderLocation, each try [Extension] <> ".err" otherwise true),
    SelectFinalColumns = Table.SelectColumns( RemoveErrorRows, {"QuoteNumber", "FormNumber", "FormType", "FolderLocation", "Name", "Extension", "Date modified", "Date created"}),
    RenamedColumns = Table.RenameColumns(SelectFinalColumns, {{"Name", "FileName"}, {"Extension", "FileExtension"}, {"Date modified", "FileDateModified"}, {"Date created", "FileDateCreated"}}),
    SortedTable = Table.Sort(RenamedColumns, {{"FormType", Order.Ascending}, {"FormNumber", Order.Ascending},{"FileName", Order.Ascending}})
in
    SortedTable



## PowerQuery Snippets

***

### ZIP Batch Extractor

```powerquery
/* -------------------------------------------------------------------------
   Source: ERPM NF-XML reader, updated at Ampara Data Project
   Function: fnExtractFromZIP
   Purpose: Extracts and decompresses files from ALL .zip archives in a folder 
            using native Power Query binary logic (No external tools required).
   Context: Bypasses the need for Python/Shell scripts when processing 
            large amount of files compressed into zip to allowing easy upload.

   Attribution:
   - Logic adapted from: Mark White (sql10.blogspot.com -> Reading Zip files in PowerQuery / M)
   - Refactor: Wrapped raw binary logic into a Folder Iterator to handle 
     batch processing of multiple ZIPs simultaneously, and improved usage interface.
   ------------------------------------------------------------------------- */

(bvFolderPath as text) =>
let
    vFolderPath = Folder.Files(bvFolderPath),
    vZipFiles = Table.SelectRows(vFolderPath, each Text.Lower([Extension]) = ".zip"),

    // --- Internal Binary Decoder ---
    fnExtractFromZip = (bvZIPFile as binary) =>
    let
        Header = BinaryFormat.Record([
            MiscHeader = BinaryFormat.Binary(14),
            BinarySize = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
            FileSize   = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
            FileNameLen= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16, ByteOrder.LittleEndian),
            ExtrasLen  = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16, ByteOrder.LittleEndian)    
        ]),

        HeaderChoice = BinaryFormat.Choice(
            BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
            each if _ <> 67324752 // ZIP Signature Check
                then BinaryFormat.Record([IsValid = false, Filename=null, Content=null])
                else BinaryFormat.Choice(
                        BinaryFormat.Binary(26),
                        each BinaryFormat.Record([
                            IsValid  = true,
                            Filename = BinaryFormat.Text(Header(_)[FileNameLen]),
                            Extras   = BinaryFormat.Text(Header(_)[ExtrasLen]),
                            Content  = BinaryFormat.Transform(
                                BinaryFormat.Binary(Header(_)[BinarySize]),
                                (x) => try Binary.Buffer(Binary.Decompress(x, Compression.Deflate)) otherwise null
                            )
                        ]),
                        type binary
                    )
        ),

        ZipFormat = BinaryFormat.List(HeaderChoice, each _[IsValid] = true),

        Entries = List.Transform(
            List.RemoveLastN(ZipFormat(bvZIPFile), 1),
            (e) => [FileName = e[Filename], Content = e[Content]]
        )
    in
        Table.FromRecords(Entries),
    // -----------------------------

    // Execute Extraction across all found ZIPs
    vExtractedData = Table.AddColumn(vZipFiles, "ExtractedData", each try fnExtractFromZip([Content]) otherwise null),
    vCleanColumns = Table.SelectColumns(vExtractedData,{"ExtractedData"}),
    vTable = Table.ExpandTableColumn(vCleanColumns, "ExtractedData", {"FileName", "Content"}, {"Name", "Content"})
in
    vTable
```

### Text Normalizer (Whitespace Cleaner)

```powerquery
/* -------------------------------------------------------------------------
   Function: fnNormalizeSpaces
   Purpose: Sanitizes text by removing leading/trailing whitespace and 
            replacing multiple internal spaces with a single space.
   Context: Essential for "Fuzzy Matching" preparation when joining datasets 
            from different systems (e.g., User Input vs. Database Records).]

   //Future Improvements:
     -[ ] Enables the function to work on multiple columns 
   ------------------------------------------------------------------------- */

(bvInputTable as table, bvColumnName as text) as table => 
let
    normalizedTable = Table.TransformColumns(
        bvInputTable,
        {
            {bvColumnName,
                each Text.Combine(
                    List.Select(
                        Text.Split(Text.Trim(_), " "), each _ <> ""
                    ), " "
                ), type text
            }
        }
    )
in
    normalizedTable
```

### Accent Remover (Data Standardization)

```powerquery
/* -------------------------------------------------------------------------
   Function: fnRemoveAccents
   Purpose: Replaces accented characters (Latin-1) with their ASCII equivalents.
   Context: Critical for standardizing names/addresses before performing Database
            Merges or ID generation.
   ------------------------------------------------------------------------- */

(fnText as nullable text) as text =>
let
    vTtxt = if fnText = null then "" else fnText,
    vPairs = {
        {"á","a"},{"à","a"},{"ã","a"},{"â","a"},{"ä","a"},
        {"Á","A"},{"À","A"},{"Ã","A"},{"Â","A"},{"Ä","A"},
        {"é","e"},{"è","e"},{"ê","e"},{"ë","e"},
        {"É","E"},{"È","E"},{"Ê","E"},{"Ë","E"},
        {"í","i"},{"ì","i"},{"î","i"},{"ï","i"},
        {"Í","I"},{"Ì","I"},{"Î","I"},{"Ï","I"},
        {"ó","o"},{"ò","o"},{"õ","o"},{"ô","o"},{"ö","o"},
        {"Ó","O"},{"Ò","O"},{"Õ","O"},{"Ô","O"},{"Ö","O"},
        {"ú","u"},{"ù","u"},{"û","u"},{"ü","u"},
        {"Ú","U"},{"Ù","U"},{"Û","U"},{"Ü","U"},
        {"ç","c"},{"Ç","C"},{"ñ","n"},{"Ñ","N"},
        {"ý","y"},{"ỳ","y"},{"ÿ","y"},{"Ý","Y"}
    },
    result = List.Accumulate(vPairs, vTtxt , (state, pair) => Text.Replace(state, pair{0}, pair{1}))
in
    result
```

### Text Scraper

```powerquery
/* -------------------------------------------------------------------------
   Source: Loggi Email Parser Tool
   Function: fnExtractBetween
   Purpose: Robustly extracts text occurring between two delimiters from a 
            list of candidate columns (also works with 1).
   Context: Used to parse unstructured .EML (email) bodies where the target 
            data (e.g., "Fechamento R$ 500.00") might shift between columns 
            due to inconsistent email formatting.
   ------------------------------------------------------------------------- */

(row as record, cols as list, startDelim as text, endDelim as text) as nullable text =>
let
    results = 
        List.Transform(
            cols,
            (colName) =>
                let
                    val = Record.FieldOrDefault(row, colName, null),
                    txt = if val <> null and Text.Length(Text.From(val)) > 0 then Text.From(val) else null,
                    rawExtract = if txt <> null then try Text.BetweenDelimiters(txt, startDelim, endDelim) otherwise null else null,
                    res = if rawExtract <> null and rawExtract <> "" then Text.Trim(rawExtract) else null
                in if res = "" then null else res
        ),
    nonNull = List.RemoveNulls(results)
in
    if List.Count(nonNull) > 0 then nonNull{0} else null

/* -------------
   Usage Example
   ------------- 
#"Total dos serviços" = Table.AddColumn(
    SourceTable,
    "Total dos serviços",
    each try fnExtractBetween(_, {"Column305", "Column306", "ColumnN..."}, "Fechamento R$ ", " +<br>") otherwise null,
    type text
)
*/
```



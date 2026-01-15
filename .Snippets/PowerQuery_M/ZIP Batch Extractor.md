### ZIP Batch Extractor

> Purpose: Extracts and decompresses files from ALL .zip archives in a folder using native Power Query binary logic (No external tools required).
> 
> Source: ERPM NF-XML reader, updated at Ampara Data Project
> 
> Context: Bypasses the need for Python/Shell scripts when processing large amount of files compressed into zip to allowing easy upload.

```powerquery
/* Call: fnExtractFromZIP 
   ---------------------- */

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

#### Attribution

- Logic adapted from: Mark White (sql10.blogspot.com -> Reading Zip files in PowerQuery / M)
- Refactor: Wrapped raw binary logic into a Folder Iterator to handle 
  batch processing of multiple ZIPs simultaneously, and improved usage interface.

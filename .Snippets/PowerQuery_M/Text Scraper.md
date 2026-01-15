### Text Scraper

> Purpose: Robustly extracts text occurring between two delimiters from a list of candidate columns (also works with 1).
> 
> Source: Loggi Email Parser Tool
> 
> Context: Used to parse unstructured .EML (email) bodies where the target data (e.g., "Fechamento R$ xxx") might shift between columns due to inconsistent email formatting.

```powerquery
/* Call: fnExtractWithin
   ---------------------- */

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
    each try fnExtractWithin(_, {"Column305", "Column306", "ColumnN..."}, "Fechamento R$ ", " +<br>") otherwise null,
    type text
)
*/
```

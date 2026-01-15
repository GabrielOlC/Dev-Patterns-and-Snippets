### Text Normalizer (Whitespace Cleaner)

> Purpose: Sanitizes text by removing leading/trailing whitespace and replacing multiple internal spaces with a single space.
> 
> Context: Essential for "Fuzzy Matching" preparation when joining datasets from different systems (e.g., User Input vs. Database Records).

```powerquery
/* Call: fnNormalizeSpaces
   ----------------------- */

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

#### Future Improvements:

- [ ]  Enables the function to work on multiple columns.

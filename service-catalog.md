# Service Catalog

Quick index of the reusable snippets in [`.Snippets`](.Snippets)

---

## VBA

### [vfCopyToClipboard](.Snippets/VBA/Clipboard%20Utility.bas)

Copies a string to the system clipboard through `MSForms.DataObject`

### [vfHexToRGB](.Snippets/VBA/Hex%20to%20RGB%20Converter.bas)

Converts a hex color code (with or without `#`) into the `Long` value used by VBA's `RGB`.

### [vfResizeListObject](.Snippets/VBA/Table%20Resizer.bas)

Resizes an Excel table (ListObject) down to its header plus one data row, keeping formulas and formats, before a new dataset is imported via SQL/Power Query.

### [vsForceTextFormat](.Snippets/VBA/Type%20Enforcer.bas)

Rewrites every value in the current selection as a string so numeric IDs are not auto-converted to numbers by Excel.

### [vsExportUniqueIDs](.Snippets/VBA/Unique%20Filter%20%26%20Clipboard%20Export.bas)

Collects unique key/value pairs from the visible (filtered) rows of an Excel table and copies them to the clipboard as tab-separated lines.

---

## Power Query (M)

### [fnExtractBetweenStr](.Snippets/PowerQuery_M/fnExtractBetweenStr.pq)

Returns the first non-empty, trimmed text found between two delimiters across a list of candidate columns of a row, or `null` if none match.

### [fnExtractFromZIPBty](.Snippets/PowerQuery_M/fnExtractFromZIPBty.pq)

Extracts and decompresses the files inside every `.zip` in a folder using native M binary parsing, returning one row per file with the source ZIP name, file name and content.

### [fnNormalizeTblSpaces](.Snippets/PowerQuery_M/fnNormalizeTblSpaces.pq)

Trims the text of a table column and collapses repeated internal spaces into a single space.

### [fnRemoveColAccents](.Snippets/PowerQuery_M/fnRemoveColAccents.pq)

Replaces accented (Latin-1) characters with their ASCII equivalents across a list of table columns.

### [fnRemoveStrAccents](.Snippets/PowerQuery_M/fnRemoveStrAccents.pq)

Replaces accented (Latin-1) characters in a text value with their ASCII equivalents; `null` becomes an empty string.

### [fnReplaceInCol](.Snippets/PowerQuery_M/fnReplaceInCol.pq)

Applies a list of find → replace rules per column (plain, whole-word, whole-cell, length or contains modes), lowercasing and trimming the text, removing double spaces and turning empty cells into `null`.

---

## Excel Formulas

### [Advanced Pricing Algorithm (Fallback Logic)](.Snippets/Excel/Advanced%20Pricing%20Algorithm%20%28Fallback%20Logic%29.md)

Looks up a product's price for the selected tier and, if that tier is empty, falls back to the next lower tier, returning the fallback price as a negative number to flag it as an estimate.

### [Dynamic 2D Lookup](.Snippets/Excel/Dynamic%202D%20Lookup.md)

Returns the value at the intersection of a row and a column of an Excel table, using the included helper formulas to generate the row and column indices.

### [Not Empty or Zero](.Snippets/Excel/Not%20Empty%20or%20Zero.md)

Named LAMBDA (`lmdNotEmptyOrZero`, alias `lmdNEZ`) that returns `TRUE` when a cell is neither blank nor zero.

### [Alphanumeric Validator](.Snippets/Excel/Text%20Entries%20Validator.md)

Returns `TRUE` when a text contains only letters (A–Z, a–z) and digits (0–9), with toggles to allow spaces and empty strings.

### [ASCII Validator](.Snippets/Excel/Text%20Entries%20Validator.md)

Returns `TRUE` when every character of a text is standard ASCII (codes 0–127), with a toggle to allow empty strings.

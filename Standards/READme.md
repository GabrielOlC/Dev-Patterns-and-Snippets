# 📏 Naming Conventions & Patterns

This repository defines a strict **Notation** tailored for multi-language data projects. This ensures that any variable's scope and type are instantly recognizable without navigating to its definition. It also allows fast CTRL-F across multi-language projects.

**Note: This is a Living Document. Standards may evolve as new technologies are integrated.**

* * *

## Universal Prefixes

>  **These prefixes apply across all languages (Excel, VBA, SQL, M) to denote data relationships and system configurations.**

* `sys_` → **System Variables:** Configuration settings defined globally (e.g., `sys_cnFilePath`, `sys_DateLimit`). **Use situation:** Using a dynamic in‑book cell to set a VBA header name/range. A normal cell name would be `cnCellName`, but since this cell is a static reference in code used to dynamically set the header ranges, it is called `sys_cnCellName`.

* `FK_` → **Foreign Keys:** Explicitly marking relational data IDs in tables or SQL queries (e.g., `FK_CustomerID`, `FK_OrderID`).

* `aux_` or `_` or `.` → **Auxiliary Elements:** Helper tables, temporary calculations, or intermediate staging data (e.g., `_tbCalculation`, `.MidCalc`, `aux_FilterCol`). **Use situation:** Python’s personalized functions may use `_`, Notion’s user‑visible properties may use `.`, and Excel formulas may use `aux_`.
  
  * **Note:** This convention was, for a long time, referred to as `sup`, `sub`.

* * *

## Universal suffix

> **This is a summary of every code language variable**

* `v` → **Variable(Standard):** Basic data types like Integer, String, Boolean. **e.g.:**`vRowCount`, `vUserName`
* `s` → **Set (Objects):** Object variables that require the `Set` keyword. **e.g.:**`sFileDialog`, `sRange`
* `c`  → **Constant:** Canstant variable. **e.g.:** `ctbID` 
* `rs` → **Recordset:** Variables of type `ADODB.Recordset`.
* `wb` → **Connection:** Variables of type `ADODB.Connection` (Database connectivity).
* `cn` → **Cell Name:** Named ranges referencing specific cells. **e.g.:**`cnTaxRate`, `cnFilePath`
* `vf` → **Functions:** Procedures that return a value by given an argument **e.g.:**`vfGetLastRow`, `vfCalculateTax`
* `m`→ **Measures**: Measures returning a value. **e.g.:** `mAnimalCount`
* `ar` → **Array:** Variable with multiple values. **e.g.:** `arItems`, `arPrices`
* `dt` → **Dictionary:** Key-value pair collection. **e.g.:** `dtConfig`, `dtUserData`

## VBA Specifics

**My VBA architecture follows Object Calisthenics principles: Sheet modules handle triggers, while logic is encapsulated in distinct "WorksheetFunctions" or "CrossModules". Refer to `Architecture_Principles.md`**

### 🔹Components & Constructs

* `vb` → **Sheet Modules:** Representing the code behind a specific worksheet. **e.g.:**`vbDashboard`, `vbDataInput`

* `vf` → **VBA Functions:** Procedures that return a value. **e.g.:**`vfGetLastRow`, `vfCalculateTax`

* `vs` → **VBA Sub:** Procedures that perform an action without returning a value. **e.g.:**`vsExportPDF`, `vsClearInputs`

* `cl` → **Class:** Variables representing a custom Class Object. **e.g.:**`clCustomer`

* `mcl` → **Module Class Variable:** Private variables defined within a Class module. **e.g.:**`Private mclID as String`

* `bt_`→**Buttons:** Subs or Functions that are triggered by a worksheet button. **e.g.:**`bt_vsUpdateAllTokens()`

* `fm`→ **User Forms:** Custom dialog windows. **e.g.:** `fmTransferTokens`

### 🔹Arguments & Parameters

* `bv` → **ByValue:** Arguments passed by value; changes inside the function do **not** affect the original variable. **e.g.:**`Sub vsProcessData(bvInputValue As String)`
* `bf` → **ByRef:** Arguments passed by reference; changes inside the function **do** affect the original variable. **e.g.:**`Function vfUpdateCount(bfCounter As Integer)`

### 🔹Variables

* `v` → **Dim Variable(Standard):** Basic data types like Integer, String, Boolean. **e.g.:**`vRowCount`, `vUserName`
* `s` → **Set (Objects):** Object variables that require the `Set` keyword. **e.g.:**`sFileDialog`, `sRange`
* `c` → **Constant:** Canstant variable. **e.g.:** `ctbID`
* `rs` → **Recordset:** Variables of type `ADODB.Recordset`
* `wb` → **Connection:** Variables of type `ADODB.Connection` (Database connectivity).
* `ar` → **Array:** Variable with multiple values. **e.g.:**  `arItems`, `arPrices`
* `dt` → **Dictionary:** Key-value pair collection. **e.g.:** `dtConfig`, `dtUserData`

* * *

## Excel Formulas & Structure

**Patterns for named ranges, tables, and LAMBDA functions to keep the spreadsheet layer clean.**

### 🔹Tables & Objects

    **Structure:** [`SheetInitials` all in CAPS] + [First letters of `Name` in CAPS].

* `tb[SheetInitials][Name]` → **Tables:** Structured Excel Tables. **e.g.:** `tbSDDash` (Sheet: **S**ummed **D**ata, Table: **DASH**).
* `ptb[SheetInitials][Name]` → **PivotTables:** Structured Excel Pivot tables. **e.g.:** `ptbSDDash` (Sheet: **S**ummed **D**ata, Table: **DASH**)
* `obj[SheetInitials][Name]` → **Objects:** Shapes, buttons, or form controls. **e.g.:** `objMMButton` (Sheet: **M**ain **M**enu, Object: **BUTTON**).

### 🔹Named Ranges & Variables

* `cn` → **Cell Name:** Named ranges referencing specific cells (**e.g.:**`cnTaxRate`, `cnFilePath`).
* `v` → **LET Variable:** Variables defined inside a `LET` formula function. (**e.g.:** `=LET(vTotal, SUM(A:A), vTotal * 0.1)`)
  * Note: All equations should be inside a variable; the last argument which can be a function without a variable should not exist, instead, replace it with a simpler `v` to facilitate testing. The maximum complexity the last equation should be formatting. **e.g.:** `=LET(v, complex equation, FORMAT(v,"yyyy-mm-dd"))` 

### 🔹LAMBDA Functions

* `lmd` → **Lambda Definition:** The name of the custom LAMBDA function in the Name Manager. **e.g.:** `lmdNotEmptyOrZero` (Shortened: `lmdNEZ`)
* `lm` → **Lambda Parameter:** Internal variables used within the LAMBDA logic. **e.g.:** `=LAMBDA(lmCell, IF(lmCell<>0...))`

* * *

## Power Query (M) & DAX

**Standards for ETL processes and Data Analysis Expressions.**

### 🔹Queries

* `qr_` → **Query Tables:** Final output queries loaded to the grid or data model. **e.g.:** `qr_SalesData`, `qr_DimProducts`.

### 🔹Functions

* `fn` → **M Functions:** Custom Power Query functions. **e.g.:** `fnUnzipXML`, `fnParseDate`.

### 🔹Measures

* `m`→ **Measures**: Measures returning a value. **e.g.:** `mAnimalCount`



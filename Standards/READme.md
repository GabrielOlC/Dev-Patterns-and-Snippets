# 📏 Naming Conventions & Patterns

This repository defines a strict **Notation** tailored for multi-language projects. This ensures that any variable's scope and type are instantly recognizable without navigating to its definition. It also allows fast CTRL-F across multi-language projects.

**Complementary:** `Architecture_Principles.md` defines where code lives, the layers that `wf_`, `cm_` and `sys_` point to.

**Note: This is a Living Document. Standards may evolve as new technologies are integrated. This might result in mixed scopes like: spreadsheets (from excel), pages (from web development), tables (from databases) and so on.**

---

## Universal Prefixes

> Note: A prefix that can be followed by another token ends in `_`, so both stay readable. **e.g.:** `qr_tbSDDash` is a query (`qr_`) loading the Excel table `tbSDDash`.

### Anatomy

> **A name is a token (scope and type) followed by a descriptive part (what the thing is). Casing follows the language.**

- **Descriptive part:** English, acronyms in caps. **e.g.:** `cTermsURL`, `fmCNPJ`, `FK_OrderID`
- **Booleans:** The descriptive part opens with `Is`, `Has`, `Needs`, `Show` or `Skip`. **e.g.:** `vIsValid`, `vHasExpired`, `cShowCountdown`
- **Units:** A quantity ends with its unit: `Ms`, `Bytes`, `Days`, `Years`. **e.g.:** `cRequestTimeoutMs`, `cMaxPDFBytes`
- **Second instance:** The same thing, repeated, takes a number instead of a synonym; a different thing gets its own name. **e.g.:** two personal phones are `fmRespWhatsApp` and `fmRespWhatsApp2`; the institutional one is `fmWhatsAppInstitucional`.
- **Same fact, same name:** The original caller is preferable. When a value is stored again, to shorten a long caller or avoid repeating heavy work, it keeps its descriptive part, so one search finds both. A new name is only for a new fact. `sys_` is never stored again. **e.g.:** `vFormsPath = sURL.pathname` keeps *Path*, the same fact; `dtRoute = dtRoutes[vFormsPath]` keeps *Route*, one item of the same collection; `vIsDebug = cDebugMode && !vIsProduction` derives a new fact, so it gets a new name.
- **Native names:** Names set by the language or by an API stay as they are. **e.g.:** `self`, `__init__`, `sURL.pathname`
- **Loop counters:** Short loops, comprehensions and lambdas may use `z`, `x`, `c` and `v`; nested ones combine them (`zz`, `zx`, `zc`…). Anything longer gets a descriptive name. **e.g.:** `for z in arWaves`

### Values

> **Values that travel between systems are written the same way everywhere.**

- **Dates:** ISO 8601 with an explicit offset, wherever the format allows. **e.g.:** `2026-10-08T16:00:00-03:00`, never `2026-10-08 16:00`
- **Stored options:** Lower snake\_case in the domain's language, without accents. **e.g.:** `sem_preferencia`
- **Metadata keys:** Keys written by the system, rather than taken from a field, are lower snake\_case. **e.g.:** `form_id`, `debug_mode`, `received_at`

### Primary

> **These prefixes apply across all languages to denote data relationships and system configurations.**

- `sys_` → **System Variables:** Configuration settings defined globally, always called in full, never stored in another variable or renamed. **e.g.:** `sys_cnFilePath`, `sys_DateLimit`, `sys_config.js`. **Use situation:** Using a dynamic in‑book cell to set a VBA header name/range. A normal cell name would be `cnCellName`, but since this cell is a static reference in code used to dynamically set the header ranges, it is called `sys_cnCellName`.
- `FK_` → **Foreign Keys:** Explicitly marking relational data IDs in tables or SQL queries. **e.g.:** `FK_CustomerID`, `FK_OrderID`
- `aux_` or `_` or `.` → **Auxiliary Elements:** Helper tables, temporary calculations, or intermediate staging data. **e.g.:** `_tbCalculation`, `.MidCalc`, `aux_FilterCol`. **Use situation:** Python’s personalized functions may use `_`, Notion’s user‑visible properties may use `.`, and Excel formulas may use `aux_`. *(legacy: `sup`, `sub`)*

### Files

> **The layer a file belongs to is in its name, in every language.** The layers themselves are defined in `Architecture_Principles.md`; `sys_` and `aux_` (above) name files too.

- `wf_` → **Controller Files:** **e.g.:** a VBA module `wf_Dashboard`, a page script `wf_afiliado.js`
- `cm_` → **Service Files:** **e.g.:** a VBA module `cm_Calculations`, a script `cm_forms.js`, a SQL file `cm_<structure>.sql` that builds a reusable database structure
- **Plugin folders:** A folder of interchangeable services with one scope keeps plain file names; the folder states the role. **e.g.:** `/layouts/biomes_confluence.py`

---

### Default

> **Tokens used in more than one language.** A token starts in the language that first needed it and moves here once a second language adopts it. For reference, each language chapter lists, by name, the tokens it takes from here.
> 
> Where a language gives a token a local meaning, the local one wins inside that language, **e.g.:** `c` for columns in SQL.

- `v` → **Variable (Standard):** Basic values like integers, strings and booleans; also a `LET` variable, an M step or a SQL CTE. **e.g.:** `vRowCount`, `vUserName`
- `s` → **Object References:** Any object, including one made from a class; in VBA, an object assigned with `Set`; in JavaScript, DOM elements too. **e.g.:** `sRange`, `sBtnSubmit`, `sCustomer = clCustomer()`
- `c` → **Constant:** A single setting or limit, fixed on purpose. The one token that never stacks: a fixed dictionary or list keeps its own token. **e.g.:** `cMaxPDFBytes`, `cDebugMode`, but `dtRoutes`
- `ar` → **Array:** Variable with multiple values. **e.g.:** `arItems`, `arPrices`
- `dt` → **Dictionary:** Key-value pair collection; in JavaScript, also any plain object used as a record or a lookup table. A property read like a variable takes its token; a key looked up as data takes none and follows *Values*. **e.g.:** `dtRoutes`, `sys_config.dtPages['cadastro-afiliado'].cDevMode`
- `fn` → **Functions:** Procedures that return a value. A variable that holds a function takes `fn`, by what the function does. **e.g.:** `fnGetLastRow`, `fnIsValidCPF`, `for (const fnCheck of [fnIsValidCPF, fnIsValidCNPJ])` *(legacy: `vf`)*
- `vs` → **Subs:** Procedures that perform an action without returning a value. **e.g.:** `vsExportPDF`, `vsRenderMessage`
- `bt` → **Buttons:** A button, or the procedure it fires, **e.g.:** `btSubmit` (the element), `bt_vsSubmit()` (the procedure) — followed by another token, it takes `_`&#32;
- `bv_` → **Parameters:** Every parameter starts with `bv_`, followed by its own token. **e.g.:** `s``u``b v``sRenderCallout(bv_sHost, bv_dtCopy, bv_vIsAssertive)`, `fu``nction vf``HexToRGB(``b``yV`` al  ``b``v``_``v``HexColor` ` As String) As  ``Lo``n``g`
- `lm` → **Lambda Parameter:** Parameters inside an Excel `LAMBDA` or an M function. **e.g.:** `=LAMBDA(lmCell, IF(lmCell<>0...))`
- `fm` → **Forms:** A dialog (VBA UserForm) or one of its fields (web). A field's `name`, `id` and payload key are one string, and the same fact keeps it in every form. **e.g.:** `fmTransferTokens`, `fmCNPJ`
- `cl` → **Class:** The class itself, or a variable holding a class; an object made from it is an `s`. **e.g.:** `clCustomer`, `clLayout = dtLayouts[bv_vKey]`

## VBA Specifics

**My VBA architecture follows** ***O****bject Calisthenics **P**rinciples***: sheet modules handle triggers, while logic lives in distinct WorksheetFunctions (`wf_`) and CrossModules (`cm_`) modules. Refer to `Architecture_Principles.md`.**

**From Default:** `v`, `s`, `c`, `ar`, `dt`, `fn`, `vs`, `bt`, `bv_`, `fm`, `cl`

### 🔹Components & Constructs

- `vb` → **Sheet Modules:** Representing the code behind a specific worksheet. **e.g.:** `vbDashboard`, `vbDataInput`
- `mcl` → **Class Member:** A private variable inside a class module; VBA has no `self.` or `this.` to mark it. **e.g.:** `Private mclDebt As Double`

### 🔹Arguments & Parameters

- `bf_` → **ByRef:** Arguments passed by reference; changes inside the function **do** affect the original variable. Like `bv_`, followed by its own token. **e.g.:** `Function fnUpdateCount(bf_vCounter As Integer)`

### 🔹Variables

- `rs` → **Recordset:** Variables of type `ADODB.Recordset`.
- `wb` → **Connection:** Variables of type `ADODB.Connection` (Database connectivity).

---

## Excel Formulas & Structure

**Patterns for named ranges, tables, and LAMBDA functions to keep the spreadsheet layer clean.**

**From Default:** `v`, `lm`

### 🔹Tables & Objects

**Structure:** `<SheetInitials>` in caps, then the first letters of `<Name>` in caps.

- `tb<SheetInitials><Name>` → **Tables:** Structured Excel Tables. **e.g.:** `tbSDDash` (Sheet: **S**ummed **D**ata, Table: **DASH**).
- `ptb<SheetInitials><Name>` → **PivotTables:** Structured Excel Pivot tables. **e.g.:** `ptbSDDash` (Sheet: **S**ummed **D**ata, Table: **DASH**)
- `obj<SheetInitials><Name>` → **Objects:** Shapes, buttons, or form controls. **e.g.:** `objMMButton` (Sheet: **M**ain **M**enu, Object: **BUTTON**).

### 🔹Named Ranges & Variables

- `cn` → **Cell Name:** Named ranges referencing specific cells (**e.g.:** `cnTaxRate`, `cnFilePath`).
- **LET Rule:** Every equation inside a `LET` sits in a `v` variable; the last argument only returns a variable, at most wrapped in formatting, so each step can be tested on its own. **e.g.:** `=LET(vDueDate, complex equation, FORMAT(vDueDate,"yyyy-mm-dd"))`

### 🔹LAMBDA Functions

- `lmd` → **Lambda Definition:** A custom LAMBDA, named in the Name Manager or held in a `LET` variable. **e.g.:** `lmdNotEmptyOrZero` (Alias: `lmdNEZ`), `=LET(lmdTax, LAMBDA(lmValue, lmValue*0.1), lmdTax(A1))`

---

## Power Query (M) & DAX

**Standards for ETL processes and Data Analysis Expressions.**

**From Default:** `v`, `bv_`, `fn`, `lm`

### 🔹Queries

- `qr_` → **Query Tables:** Final output queries loaded to the grid or data model. **e.g.:** `qr_tbSDDash`, `qr_Products`.

### 🔹Measures

- `m` → **Measures:** DAX measures returning a value. **e.g.:** `mAnimalCount`

---

## SQL & Database

**Postgres folds unquoted names to lower case, so every SQL name is lower snake\_case, and every prefix ends in `_`. A name has at most 63 characters; Postgres cuts longer ones without warning.**

**From Default:** `v`, `c` (inside functions and CTEs), joined to the name like columns

### 🔹Objects

- `sch_` → **Schemas:** One per business domain, plus a shared one for what every domain uses, like a `cm_`. **e.g.:** `sch_forms`, `sch_shared`
- `tb_` → **Tables:** **e.g.:** `sch_forms.tb_affiliated`, `sch_forms.tb_affiliated_proxy`
- `vw_` → **Views:** A stored `SELECT` behaving as a read-only table. **e.g.:** `vw_ongs_active`
- `vwm_` → **Materialized Views:** Views stored on disk, refreshed on command. **e.g.:** `vwm_campaign_performance`
- `dm_` → **Domains:** Reusable formats. **e.g.:** `sch_shared.dm_cnpj`, `sch_shared.dm_email`
- `tp_` → **Types:** Enumerations. **e.g.:** `sch_forms.tp_event_format`, `sch_forms.tp_api_cnpj`
- `fn_` → **Functions:** Routines returning a value, including trigger bodies. **e.g.:** `fn_set_updated_at`
- `trg_` → **Triggers:** Fired by an insert, update or delete; they hold no logic and execute an `fn_`. **e.g.:** `trg_affiliated_set_updated`
- `pol_` → **Policies:** Row-level security rules. **e.g.:** `pol_affiliated_select_auth`
- `ix_<TableInitials>_<columns or scope>` → **Indexes:** The table's initials, then its columns as named (functions dropped), or a short scope when the columns are too many. **e.g.:** `ix_zwca_fkdoctoken_clogat`, `ix_n_cmail`, `ix_zwca_latest`
- `gr_` → **Group Roles:** No login; they hold the permissions. **e.g.:** `gr_newsletters`, `gr_backups`
- `usr_` → **Login Roles:** One per application or person; they join groups. **e.g.:** `usr_n8n_main`, `usr_febraca_owner`

### 🔹Columns

- `c` → **Columns:** `c` joined to the name, with no `_`, so a column never clashes with a Postgres function or keyword. Columns with their own prefix (`fk_`, `aux_`) don't take it. **e.g.:** `ccnpj`, `clegal_name`, `ccreated_at`
- `cpk` → **Surrogate Primary Key:** Time-ordered UUID. **e.g.:** `cpk uuid PRIMARY KEY DEFAULT uuidv7()`
- `fk_` → **Foreign Keys:** The referenced table without `tb_`, or the external system's ID. **e.g.:** `fk_affiliated`, `fk_zapsigndoc`
- `aux_` → **Temporary Columns:** **e.g.:** `aux_good_cnpj`

### 🔹Queries

- `<TableInitials>` → **Alias:** The table's initials, declared per statement. **e.g.:** `from sch_forms.tb_affiliated_proxy as ap`
- `v` → **CTE:** Named intermediate result; the SQL `LET`. **e.g.:** `with vopened as (...)`

---

## JavaScript & Browser

**From Default:** `v`, `s`, `c`, `ar`, `dt`, `fn`, `vs`, `bt`, `bv_`, `fm`, `cl`

### 🔹Casing

- **Files and objects:** Lower case after the prefix; a `cm_` or `sys_` file exposes one global object with the same name. **e.g.:** `cm_forms.js` → `cm_forms.vsRenderMessage`; `sys_config.js` → `sys_config.dtPages`

### 🔹Markup

- `<fieldId>-<role>` → **Satellites:** Elements serving one field take its id, a hyphen and their role. **e.g.:** `fmCNPJ-hint`, `fmCNPJ-message`
- `data-<key>` → **Declarations & State:** Attributes in kebab-case; a state value is the same word the code and the payload use. **e.g.:** `data-page="cadastro-afiliado"`, `data-state="already_registered"`

---

## Python Specifics

**From Default:** `v`, `s`, `c`, `ar`, `dt`, `fn`, `vs`, `bv_`, `cl`

### 🔹Components & Constructs

- `fnMain` → **Entry Point:** The function a script runs first; it returns the exit code. **e.g.:** `sys.exit(fnMain())`

### 🔹Variables

- `ar` → **Sequences:** Any list, tuple, set, deque or array. **e.g.:** `arWaves`, `arQueue = deque([bv_vStart])`

### 🔹Kept from Python

- **Files:** snake\_case, since they are import names. **e.g.:** `wf_generate.py`, `cm_pipeline.py`
- **`_`**** **→ **Throwaway ****:**`_` alone is a value thrown away, not the auxiliary prefix. **e.g.:** `for vA, _, vName in arPasses`
- **Command-line options:** kebab-case, so the attributes argparse makes from them are snake\_case. **e.g.:** `--ore-rolls` → `sArgs.ore_rolls`

---

## CSS

**Classes say what a thing is; custom properties say where a value comes from.**

### 🔹Classes

- `block__element--modifier` → **BEM:** Lower case, English. **e.g.:** `.field__message--error`, `.choice--required`

### 🔹Custom Properties

- `--brand-*` → **Brand:** The declared palette, verbatim, never derived. **e.g.:** `--brand-navy`
- `--sem-*` → **Meaning:** Semantic colour and its background. **e.g.:** `--sem-error`, `--sem-error-bg`
- `--ui-*` → **Interface:** Neutrals and surfaces that are not brand colour. **e.g.:** `--ui-border`, `--ui-text-muted`
- `--fs-*`, `--sp-*` → **Scales:** Type sizes in hundreds, spacing in steps. **e.g.:** `--fs-300`, `--sp-4`

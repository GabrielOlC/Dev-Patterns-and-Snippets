# 🏗️ Architecture Principles

**Strategies for building robust, scalable, and maintainable systems within low-code environments.**

> "Complexity is the enemy of execution. This architecture decouples Logic, Data, and Interface to turn spreadsheets into software."

---

## Visualizing the Module Hierarchy

```markup-templating
/VBA_Project_Structure Example
│
├── /Microsoft Excel Objects (vb_ & Buttons) # Event listeners and triggers
│   ├── vbDashboard       # Event Listeners Only
│   └── vbDataInput       # UI Triggers
│
├── /Controllers-WorksheetFunctions (wf_)  # Event controllers (logic)
│   ├── wf_Dashboard      # Traffic Control for Dashboard
│   └── wf_DataValidation # Input Rules & Routing
│
├── /Services-CrossModules (cm_ & cl) # Reusable code
│   ├── cm_Calculations   # Pure Math Logic
│   ├── cm_Database       # ADODB Connection & SQL
│   ├── cm_Files          # FileSystemObject Logic
|   ├── clRange           # Class with range references
|   └── clString          # Class with name references
│
└── /Core (sys_) # Global setups
    ├── sys_ErrorHandler  # Centralized Logging
    └── sys_Config        # Global Constants
```

## The Separation of Concerns (The "MVC" Adaptation)

In many Excel projects, logic is often buried inside Button Clicks or Worksheet Events, leading to unmaintainable "Spaghetti Code." This framework enforces a strict separation between the **User Interface**, the **Business Logic**, and the **Data Layer**.

### 🔹 The View (Interface)

* **Worksheets & UserForms:** These are designated strictly for display and user-related input.
* **Rule:** Heavy calculation logic is **forbidden** in this layer.
* **Triggers:** Sheet Modules (`vbDashboard`, `vbInput`) contain *only* Event Listeners (e.g., `SelectionChange`, `Button_Click`). These events immediately delegate execution to a Controller.

### 🔹 The Controller (Orchestration)

* **Module:** `WorksheetFunctions (wf_)`
* **Purpose:** Handles operations specific to a sheet's workflow. Acts as the traffic controller—validating inputs from the interface and invoking the necessary services.
* **Naming:** Functions mirror the Sheet's purpose. **e.g.:** `wf_ValidateDashboardInputs`

### 🔹 The Model/Service (Business Logic)

* **Module:** `CrossModules (cm_)` (and Class Modules)
* **Purpose:** Pure logic and reusable algorithms. This code is agnostic of the specific sheet calling it.
* **Rule:** Functions must accept arguments (`ByVal`) rather than reading directly from grid ranges, ensuring testability and reusability.
* **Naming:** Grouped by functional domain. **e.g.:** `cm_calculations`

---

## Object Calisthenics (Adaptation)

Principles from *Object Calisthenics* are applied to enforce code hygiene and readability:

1. **One Level of Indentation per Method:** Deeply nested `If/Loop` structures are refactored into smaller, distinct helper functions.
2. **No "Magic Numbers":** Hardcoded row numbers or column indices are prohibited. All references must be named (e.g., `sys_cnTaxRate` or `tbSales[Amount]`).
3. **Self-Documenting Code:** Variable names must be descriptive (`vRowCount` vs `r`). Comments explain *Why* a decision was made, not *What* the code is doing.

---

## Data Integrity & Interaction

Data is treated as structured records, never as arbitrary "cells on a grid."

### 🔹 Database First (ADODB)

* **SQL** via **ADODB connections** is utilized to manipulate data instead of looping through Excel rows.
* This approach decouples data storage (Access, SQLite, or external Excel files) from the interface, enabling ACID-like transactions (Commit/Rollback) and significantly higher processing speeds.

### 🔹 ListObjects (Tables) over Ranges

* When data must reside in Excel, it is stored in **Structured Tables** (`ListObjects`).
* **Why:** Tables ensure automatic expansion, consistent headers, and structured referencing (e.g., `Table1[Column1]`), eliminating reference errors common in standard ranges. With in-cell header references, the table is complety dynamic, where the user may change its headers name, order and table position without compromising the code logic. 

---

## Error Handling Strategy (Poka-Yoke)

Errors are managed through a dual strategy of **Prevention** and **Graceful Failure**.

### 🔹 Prevention (Poka-Yoke)

* **Input Validation:** The system physically prevents entry of invalid data via Data Validation, UI restrictions, and immediate feedback loops (visual/audio alerts) before execution logic is triggered.

### 🔹 Centralized Handling

* **Module:** `ErrorHandler`
* **Pattern:** All primary procedures implement a standardized error trapping block.
* **Logging:** Exceptions are not merely displayed; they are logged (timestamp, user, error description) to a secure log sheet or database for auditing purposes.

```vba
' Standard Pattern Example
Sub vsProcessAction()
    On Error GoTo ErrorHandler
    ' [Main Logic Here]

ExitSub:
    Exit Sub

ErrorHandler:
    ' Calls the centralized handler to log and display user-friendly message
    vsLogAndAlert(Err.Number, Err.Description, "vsProcessAction")
    Resume ExitSub
End Sub
```

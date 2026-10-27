# 🏗️ Architecture Principles

**Strategies for building robust, scalable, and maintainable systems within low-code environments.**

> "Complexity is the enemy of execution. This architecture decouples Logic, Data, and Interface, providing a universal pattern for building maintainable, scalable systems across spreadsheets, applications, databases, and automated workflows."

---

## Visualizing the Module Hierarchy

Example: a web project. In a VBA project, the same layers are `vb` sheet modules, `wf_` and `cm_` modules, and the `sys_` core.

```
/General_Project_Structure
│
├── /docs                            # Knowledge that outlives the code
│   ├── /clients/<client>.md         # Manual — one client's current state
│   ├── design-decisions.md          # Global decisions
│   └── service-catalog.md           # One line per service
│
├── /public                          # Served as-is (what the user sees)
│   ├── index.html                   # Entry — the front door to the pages
│   └── .assetsignore                # What must never be served
│
├── /public/<page> (index & wf_)     # One folder per page — view and controller side by side
│   ├── index.html                   # View — markup only, decides nothing
│   ├── wf_<page>.js                 # Controller — traffic control for the page
│   ├── style.css                    # Optional — only what this page alone needs
│   └── README.md                    # Placard — names the page, links its manual
│
├── /public/shared/js (cm_ & sys_)   # Reusable code
│   ├── cm_<function>.js                # Service — one concern each (rules, masks, lookups…)
│   └── sys_config.js                # Core — global setup, one section per page
│
├── /public/shared (shared assets)   # Reusable assets — no logic
│   ├── /css                         # Base look every page loads
│   ├── /img                         # Brand assets — logos, icons
│   └── /data                        # Reference data, split so a page loads only what it needs
│
├── /src                             # Gateway
│   └── index.js                     # Secrets, proof of origin, refusals
│
└── README, wrangler, gitignore, etc # General files of the project
```

## The Separation of Concerns (The "MVC" Adaptation)

In many projects, logic is often buried inside Button Clicks or Worksheet Events leading to unmaintainable "Spaghetti code" and duplicated logic. This framework enforces a strict separation between the **User Interface**, the **Business Logic**, and the **Data Layer**. A **client** is anything the project serves on its own: a sheet, a page, a form, a report, an integration. Every client goes through a Controller (`wf_`), and the Controller picks the Services (`cm_`) that build the result. Whatever is needed twice is built once, so a new client *configures* existing code instead of rewriting it. The principle has a limit: a client with needs of its own gets its own `wf_ / cm_`, called only by that client.

### 🔹 The View (Interface - the client)

- **Worksheets, UserForms, Pages:** These are designated strictly for display and user-related input.
- **Rule:** Heavy calculation logic is **forbidden** in this layer.
- **Triggers:** Sheet Modules (`vbDashboard`, `vbInput`) contain *only* Event Listeners (e.g., `SelectionChange`, `Button_Click`). These events immediately delegate execution to a Controller.

### 🔹 The Controller (Orchestration)

- **Module:** `WorksheetFunctions (wf_)`
- **Purpose:** Handles operations specific to a client's workflow. Acts as the traffic controller—validating inputs from the interface and invoking the necessary services.
- **Naming:** Functions mirror the client's purpose. **e.g.:** `wf_ValidateDashboardInputs`
- **Rule:** The controller chooses; the service builds. It picks which outcome to show and in which words, and never assembles a component a service already renders. **e.g.:** `wf_`afiliado`` picks the outcome; `cm_``f``orms.vsRenderCallout` draws the box.

### 🔹 The Model/Service (Business Logic)

- **Module:** `CrossModules (cm_)` (and Class Modules)
- **Purpose:** Reusable logic and algorithms. This code is agnostic of the specific client calling it. **Why:**  Services act as the project's toolkit of trusted behaviours and patterns. They are built to be reused, evolved, and carried into future projects. The most valuable ones are preserved in [Dev-Patterns-and-Snippets](https://github.com/GabrielOlC/Dev-Patterns-and-Snippets/tree/main/.Snippets), enabling new projects to begin with proven code rather than empty modules.
- **Rule:** Functions must accept arguments (`ByVal`) rather than reading directly from grid ranges, ensuring testability and reusability. That is: the service has no idea who called it, never fetches its own inputs from the client and therefore is usable by any controller. It may write to the client, but only where the controller points — or, by fixed name, to an element's satellite (e.g., `fmCNPJ` → `fmCNPJ-message`) or to a fixture every page carries (e.g., `#live-region-polite`). **e.g.:** `vsAnnounce(bvMessage, bvIsAssertive)`
- **Naming:** Grouped by functional domain. **e.g.:** `cm_calculations`
- **Ready for Many, Built for One:** A service is shaped for a second client from the start, but its machinery must wait until that client exists. Arguments replace hard-coded values and the signature can grow optional arguments; mode switches, tables and registries arrive with the second real client. **e.g.:** `fnReplaceInCol` took a list of rules from day one, so new modes were added later without touching existing calls.
- **Extend Without Breaking:** New behaviour arrives as an optional argument whose default keeps every existing call unchanged. **e.g.:** `formatField(value, formatter, options = {})`. Also, a service may call another service. **e.g.:** A client needs to submit and normalize a CNPJ. Since a normalization service already exists, this service first calls the normalization service and then submits the normalized CNPJ.
- **Documentation:** Every service file opens with the header in [In code Snippets](https://github.com/GabrielOlC/Dev-Patterns-and-Snippets/blob/main/Templates/In%20code%20Snippets.md). One header per file, because each file solves one concern.

### 🔹 Limits of Reuse of a service

- **Configure, Don't Copy:** Preferably, a new client requirement is added as an entry in a table, not as new code. **e.g.:** a new CEP provider is a new entry to the argument  `bvProviders` not a new `cm_FetchAddress`.
- **Same Look, Different Job:** Components that look alike but serve different purposes are kept separate rather than merged into one service with branches. **e.g.,** two clients that need a CEP API, but one needs a specific fallback order, while the other needs a single API with no fallback.

---

## Object Calisthenics (Adaptation)

Principles from *Object Calisthenics* are applied to enforce code hygiene and readability:

1. **One Level of Indentation per Method:** Deeply nested `If/Loop` structures are refactored into smaller, distinct helper functions.
2. **No "Magic Values":** Hardcoded values are prohibited. Numbers, strings, limits, timeouts, sizes, e-mail addresses, URLs, file paths, and similar configuration data must be defined in sys\_ and referenced through named constants. (**e.g.:** `sys_config.cRequestTimeoutMs`, `sys_config.cMaxPDFBytes`, `sys_config.cContactEmail`, `sys_config.cSupportURL`).
3. **Explicit Declaration:** Each client names its own configuration instead of having it inferred from a file path or position. **e.g.:**  In the page:`<html data-page="cadastro-afiliado">` in the sys\_  `sys_config.dtPages = { 'cadastro-afiliado': { cDevMode: false } }`
4. **Self-Documenting Code:** Variable names must be descriptive (`vRowCount` vs `r`). A comment is added only where the code cannot say it, and then it gives the reason in a line or two.
5. **Lists over Booleans:** Membership is expressed as a list, so changing it means adding or removing a name rather than flipping flags. **e.g.:** `arRequiredFields`, `vAcceptsFiles`
6. **Label vs. Value:** The text a person reads can be rewritten at any time; the value that is stored is the contract. Stored values are canonical (digits only, no masks); formatting belongs to the interface. **e.g.:** the page shows *Não tenho preferência* and stores `sem_preferencia`.
7. **Clients Name, Never Choose:** Anything sent from a browser can be forged, so a page only names a route (through `sys_config`) and sends its content; the gateway maps that route to its destination and credentials, adds the facts a client could fake, and refuses any route it doesn't know. **e.g.:** pages post to `/api/<form>/<action>`; only the gateway knows which API that route reaches and which key unlocks it.

---

## Data Integrity

Data is treated as structured records, never as arbitrary "cells on a grid." The data layer outlives every interface built on it, so it is designed to stand on its own.

- **One Source of Truth:** A database holds the canonical record; sheets, forms, apps and workflows only read, write or display it. **Why:** interfaces change often and data has to outlive them, so replacing a tool never means migrating the data. **e.g.:** Excel works through SQL (ADODB) instead of looping over rows, and keeps any local data in structured tables (`ListObjects`); a web form writes to Postgres through the gateway.
- **One Scope per Table:** Each table holds one scope, and an outside system is a scope of its own, so what it knows about a record lives in its own table, related to the record through its key, never as extra columns on it. **Why:** adding or replacing a system means adding or dropping its table, never changing the tables it relates to. **e.g.:** `affiliates` holds the CNPJ, name and e-mail; the payment platform’s customer ID and status live in `affiliates_payment`, related through the CNPJ.
- **The Database Has the Last Word:** Every rule about what a valid value looks like (a format, or a list of options) is also declared in the database, so it holds no matter which client writes: a form, a workflow, an import or a manual fix. Each rule follows the official specification stated in the database.
- **Idempotent, Parameterized Writes:** A write applied twice leaves the same result, and each field's merge rule (overwrite, keep, accumulate, or append as history) is stated explicitly, as a business decision. Values always travel as parameters, never concatenated into a query. **Why:** retries are normal, and concatenation invites injection.
- **Closed by Default:** Access is granted on purpose: each application or person has its own login, permissions belong to group roles, and each role gets only the operations it needs. **Why:** activity stays traceable, and a fault or a leaked credential stays confined to one login.
- **Safe in Production:** A change to a live database never holds clients hostage. It waits only briefly for what it needs and fails fast if it can't get it, to be retried later, instead of queuing every client behind it until they time out. **e.g.:** schema changes run with a short lock timeout; indexes are built concurrently.

---

## Error Handling Strategy (Poka-Yoke)

Errors are managed through a dual strategy of **Prevention** and **Graceful Failure**.

### 🔹 Prevention (Poka-Yoke)

- **Input Validation:** The system physically prevents entry of invalid data via Data Validation, UI restrictions, and immediate feedback loops (visual/audio alerts) before execution logic is triggered.

### 🔹 Centralized Handling

- **Module:** `ErrorHandler`
- **Pattern:** All primary procedures implement a standardized error trapping block.
- **Logging:** Exceptions are not merely displayed; they are logged (timestamp, user, error description) to a secure log sheet or database for auditing purposes.

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

---

## Accessibility

Every form must work for someone who can't see the screen. A screen reader only says what the markup tells it to.

### 🔹 The Rules

- **Say it, don't only show it:** Every outcome that appears on screen is also announced. **e.g.:** after submit, `cm_forms.vsAnnounce('Inscrição enviada.', true)`
- **Live regions are a page fixture:** Every page carries `#live-region-polite` and `#live-region-assertive`, hidden from sight. Polite waits its turn; assertive interrupts, so it is kept for answers to something the person just did.
- **Messages belong to their field:** Each input has a `<id>-message` satellite, linked through `aria-describedby`, so the reader speaks the error together with the field.
- **Never colour or icon alone:** State is carried in words and in `aria-invalid`. Icons repeat the text and are marked `aria-hidden`.
- **Take them to the problem:** On a failed submit, focus moves to the first invalid field. **e.g.:** `cm_forms.vsFocusFirstInvalid(arInputs)`

---

## Documentation

Each document has one job, and each fact has one home. A fact written in two places is eventually wrong in one of them, because a copy goes stale without anyone noticing. The documents below run from the broadest scope to the narrowest: a fact is written at the narrowest level where it is true, and moves up only when it proves true more widely.

### 🔹 Where Things Live

- **Standards** (this repository) → **How I Build:** Principles and notation that hold and evolve with every project. A project never restates them.
- `README.md` → **Structure:** What the project is, its layout, how to run it, how to publish it and how to add a client. Nothing about one client.
- `docs/service-catalog.md` → **Catalog:** One entry per service the project uses: a short summary of what it does. The full documentation lives in the service's own header, never here. **e.g.:** `cm_Validators` — validates CNPJ numbers, including the new alphanumeric format.
- `docs/design-decisions.md` → **Global Decisions:** What holds for every client of this project, and why (business logic).
- `docs/clients/<client>.md` → **Manual:** One client's current state, including the decisions that are its alone.
- `<client folder>/README.md` → **Placard:** Optional, used where a client has a folder of its own. It names the client and links to its manual. **Why:** client folders usually sit in `./public`, which is published, so the manual is kept in `docs/`.
- **The code comments** → **Pending Work:** What is missing or wrong, as markers where the fix happens. (see *Comments (Pending Work)*)

---

### 🔹 Comments (Pending Work)

A task to be done or something to be revised lives in the code, as a marker in the file where it is resolved.

- **Shape:** `<KIND>(<priority> | <create-YYMMDD> | <updated-YYMMDD (optional)> | <KEY (optional)>): <what> — <why / what changes with it and its location (optional)> — <more details (optional)>`. **e.g.:** `// TODO(high | 260902 | 261014 | H2609-01): sign a short-lived session for the CNPJ lookup — today only a rate limit guards it.` KEY = `<Priority first letter><YYMM>-<Sequence>`

**Kinds**

- `TODO` → **Work that needs doing.**
- `FIXME` → **Something that behaves wrong today.**
- `HACK` → **A deliberate workaround that should go away.** It names what removes it.
- `ASK` → **A decision someone else must make before the code changes.** **e.g.,** a business question
- `[ ] and [x]` → `CHECK`: **A permanent pre-release gate**, such as a debug flag or a test endpoint. It carries no priority or dates, because every `CHECK` blocks a release, and it is never deleted: it sits next to the switch and states the safe value. Its box shows the state: `[ ]` still has to be set before a release, `[x]` is set, and when the switch changes back, the box is unticked again. **e.g.:** `// [ ] Turn off cDebugMode — skips origin and captcha checks.` Every `CHECK` is read against its line before a release.

**Priorities**

- `urgent` → **Causing harm now, or at the next deploy:** data exposure or loss, a security or legal gap (e.g., LGPD), production broken for real users, or wrong information shown to them. Fixed ASAP, urgent markers live for days, not weeks. Requires a why and more details in the label.
- `high` → **Blocks shipping this part, or breaks at a known event** (a date, a volume, a dependency changing), or hits real users who have a workaround. Fixed in the current cycle, before that part ships again. Requires a why in the label.
- `medium` → **A worse result with nothing lost:** confusing UX, wrong but recoverable data, slow or fragile code. Planned, or fixed the next time that area is touched.
- `low` → **Everything else:** naming, clean-up, nice-to-have ideas.

**Writing Rules**

- **One Home:** When the work spans files, the full text sits where the fix happens; the other spots say ` see TODO in <file>: <Key>  `.
- **Completed Tasks:** They are deleted with the exception of the checks. Learnings are saved on the docs following *where things live*.
- **Checkboxes Stay in Markdown:** `- [ ]` and `- [x]` belong to Markdown checklists (release notes, pull requests), never to code. `CHECK` is the one exception, since its box records whether the gate is set.

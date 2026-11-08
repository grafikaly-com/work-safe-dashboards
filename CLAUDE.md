# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

- `pbix/*.pbix`: 9 Power BI reports.
- `xlsx/`: two hand-maintained registers:
  - `Dashboards Master File.xlsx`: one sheet per dashboard, listing each table's SharePoint location, owner and associated flow.
  - `Power Automate Flow Master Register.xlsx`: the flows.
- `scripts/`: Python tooling that generates schema and lineage docs into `docs/`.
- There is no build, lint or test setup.

## Commands

The Python environment is a conda env with the exact Python version in its name (conda isn't on Git Bash's PATH, so call it from PowerShell):

```powershell
& "$env:USERPROFILE\anaconda3\condabin\conda.bat" create -y -n work-safe-dashboards-py3.12.16 python=3.12.16 pip
& "$env:USERPROFILE\anaconda3\envs\work-safe-dashboards-py3.12.16\python.exe" -m pip install -r scripts/requirements.txt
```

Regenerate the docs with `$PY` = that env's `python.exe`:

```sh
$PY scripts/extract_pbix.py                          # all pbix -> docs/models/_extracted/<Report>.json
$PY scripts/extract_pbix.py "pbix/Survey Dash.pbix"  # one model
$PY scripts/build_docs.py                            # JSON + xlsx -> docs/*.md, docs/models/*.md, docs/Model Schema Register.xlsx
```

`extract_pbix.py` prints per-model counts plus `unresolved=N`, the number of report fields that point at tables or fields missing from the model. There are no tests. To check a change, rerun both scripts and diff `docs/`. To check Mermaid syntax, render the diagrams with `@mermaid-js/mermaid-cli`.

## Doc generation pipeline

- **`extract_pbix.py`** reads each model with `pbixray`:
  - `x.tables` minus auto date/time tables (`LocalDateTable_*`, `DateTableTemplate_*`) gives the table list. Don't use `tmschema_tables`, which leaves out DAX calculated tables.
  - `x.schema` provides columns, enriched from `tmschema_columns`. That source also holds internal `H$`/`R$` tables, so filter it to the known table names.
  - Relationships, M and DAX are read too, and report JSON gives the fields each page uses.
  - The Power Query parser classifies each table's source (SharePoint list by GUID `Source{[Id=...]}`, Excel on SharePoint, Web API, Enter Data, append or derived).
  - It also finds Power Query-only "staging" queries, which aren't loaded to the model (Permit Forecast has several).
- **Privacy rules:**
  - Never call `get_table()`/`iter_table()`, which would read row data.
  - All M passes through `redact_m()`, which strips URL query strings (the ubiquity survey API URLs), key-like values, and the base64 rows of "Enter Data" tables (e.g. `Salaries`).
  - Never follow hyperlinks in the xlsx files. Load them with `openpyxl(..., data_only=True, keep_links=False)`.
- **`build_docs.py`** links model tables to Master File rows by table name and to flows in two ways:
  - **Documented:** the Master File row names the flow. Fuzzy matching covers typos such as `Commentary Reminder` vs the register's `Commnetary Reminder`.
  - **Inferred:** a register description mentions "<list name> list", or the flow is named after the list.
- SharePoint lists are identified across models by (site, list GUID); display names come from the Master File's `Site - List - View` location.
- Each model page's "Business purpose" section is hand-written in `docs/models/_purpose/<Report>.md`, and `build_docs.py` inserts it if that file exists. Edit that file, not the generated page.
- Master File sheet names don't match pbix names exactly. When a dashboard is added or renamed, update `SHEET_TO_REPORT` in `build_docs.py`.

## Working with `.pbix` files

A `.pbix` is a zip archive. Power BI Desktop is the only supported editor. Re-zipping modified contents by hand can corrupt the file, so treat the archives as read-only and inspect them with Python's `zipfile`:

```sh
python -I -c "import zipfile,sys; print(*zipfile.ZipFile(sys.argv[1]).namelist(), sep='\n')" "pbix/Survey Dash.pbix"
```

What is inside each archive, and how it is encoded:

- `Report/definition/` holds the report in PBIR format: `report.json`, `pages/pages.json`, `pages/<pageId>/page.json` and `pages/<pageId>/visuals/<visualId>/visual.json`. These are UTF-8 JSON. Eight of the nine reports use this layout.
- **Exception:** `ELT Frontpage.pbix` still uses the legacy single-file `Report/Layout` (UTF-16LE JSON, with visual configs stored as JSON strings inside it).
- Visuals point at the model through `SourceRef.Entity` (table) and `Property` (column or measure). Grep these to find which tables a report depends on.
- `DataModel` is an XPress9-compressed Analysis Services backup (binary). Read it only through `pbixray` (see above), and never edit it.
- `Version`, `Metadata`, `Settings` and `DiagramLayout` are UTF-16LE. `Connections` is UTF-8.

## How the reports relate

- All nine were downloaded from the Power BI Service: `Metadata` has `"CreatedFrom":"Cloud"`, and `Connections` points at the same `OriginalWorkspaceObjectId` (`93a819f8-…`). Each report has its own `DatasetId`/`ReportId` and an embedded import-mode model; none is a thin report on a shared dataset. A change made here only takes effect once it is republished to that workspace.
- Common look: every report uses the built-in base theme `CY24SU10` (no custom theme) and the same logo resource, `WS_Primary_White*.PNG`. Every report except Programme Risk and Survey Dash uses a single-value `Query1` table, shown in a card.
- Main tables per report. The same source tables are re-imported into several separate models:

  | Report | Main tables |
  |---|---|
  | PEPA Management Dash | `TimesheetData`, `Applications`, `Complaints&Audits`, `RoleTable`, stage-order tables |
  | PEPA Exec Dash | `Applications`, `TimesheetData`, `Abbreviations`, `RoleTable` |
  | Compliance Monitoring Dash | `TimesheetData`, `Complaints&Audits`, `Applications`, stage-order tables |
  | ELT Frontpage | `AllApplications`, `TimesheetData`, `FinanceTable`, `PermitForecast` (its visuals also reference `A&A *`, `Technicall` and `AllResponses` tables that are no longer in the model) |
  | ELT Permit Timeliness Report | `AllApplications`, `FakeWarehouse`, `RiskAssessments`, `Commentary` |
  | Permit SPE Dash | `AllApplications`, `Survey`, `DateTable` |
  | Permit Forecast | `* Forecast` / `* Renewal Forecast` tables (ASB, AAO, Diving, HS), `DateTable` |
  | Programme Risk | `RiskRegister`, `TreatmentRegister`, `StatusSort`, `ScoreSort` |
  | Survey Dash | `All Responses`, `Regimes` |

- `Compliance Monitoring Dash` was built from pages of `PEPA Management Dash`. They share page IDs (`676ff357…` Timesheet, `717c05d4…` Complaint & Audit Time, `93362c88…` Audit & Complaint Status), so a fix to one of those pages probably belongs in both.
- Several reports still contain leftover pages such as `Page 1`, `Duplicate of …`, and tooltip or drillthrough pages (`Tooltip1/2`, `RiskDrillthrough`, `TreatDrillthrough`). Check `page.json` visibility before treating a page as user-facing.
- Six reports carry a stale conditional-format binding to `Adventure Activities[CurrentMonth]`, copied from a template. None of those six models has that table. `docs/gaps.md` lists these and the other mismatches.

## Git notes

- `pbix/` is untracked so far. The files are 0.5–8 MB each (about 34 MB in total), and Git LFS is not configured (there is no `.gitattributes`). Every save in Power BI Desktop rewrites the whole binary, so commits produce no meaningful diffs.

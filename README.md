# tool

Actuarial Excel VBA toolkit: navigation helpers, cross-workbook import/export, and formula-tracing utilities. Each tool ships as a macro-enabled workbook (`.xlsm`) with VBA source exported alongside for version control.

## Repository layout

| Path | Workbook | Purpose |
|------|----------|---------|
| [`vba Library/`](vba%20Library/) | `vba library.xlsm` | Shared utilities, navigation user forms, bookmarks, external-link helper |
| [`ImportExport/`](ImportExport/) | `ImportExport.xlsm` | Bulk push values or formulas from a control workbook into linked external files |
| [`TraceFormulaTool/`](TraceFormulaTool/) | `TraceFormulaTool.xlsm` | Formula simplification and goto-history for auditing model logic |
| [`extractMacro.py`](extractMacro.py) | — | Extract VBA from `.xlsm` files into the folders above |

## vba Library

General-purpose macros intended as a personal add-in or reference library.

### Navigation (user forms)

| Macro | Shortcut | Form | Description |
|-------|----------|------|-------------|
| `ShowrecentWB` | Ctrl+Alt+T | `RecentWB` | Open a recent workbook from Excel’s recent-files list (filterable) |
| `ShowSheetNavigator` | Ctrl+Alt+W | `SheetNavigator` | Jump to a sheet in the active workbook (hidden sheets excluded) |
| `ShowWorkbookSelector` | Ctrl+Alt+B | `WorkbookSelector` | Switch to an open workbook |
| `ShowBookmarkManagerSetMode` | Ctrl+Alt+J | `Bookmarks` | Save up to 9 cell bookmarks (workbook, sheet, address) |
| `ShowBookmarkManagerGetMode` | Ctrl+Alt+K | `Bookmarks` | Go to a saved bookmark |

`createUserForm*.bas` modules are builders that recreate the user forms programmatically if you need to regenerate UI.

### Utilities (`utils.bas`)

Scriptable Excel operations used in models and UAT:

- **File:** `xopen`, `xclose`, `xsaveclose`, `xsaveas`
- **Ranges:** `xrange`, `xrangeSelect`, `xrangeSet` — address mini-language with `^` `v` `<` `>` and doubled arrows (e.g. `A1>>`, `A1:A9&vv`) to extend selections like keyboard navigation
- **Clipboard:** `xcopy`, `xpaste`, `xpasteA` / `F` / `V` / `T` / `W` / `L` (paste special variants)
- **Other:** `xgoalseek`, `xreplace`, `xchangeLink`, `xchangeLinkOpen`, `xrun`, `xlog`

### Other modules

- **`externalref.bas`** — `ShowAndOpenExternalReferences` (Ctrl+Shift+Q): list external paths in the active cell’s formula and open or navigate to the chosen link
- **`str.bas`** — `FStr` for lightweight string formatting
- **`UAT.bas`** — manual test macros for `utils` (paths are machine-specific; edit before running)
- **`getUserFormUI.bas`** — `ExportFormGeometryToCSV` for form layout export

## ImportExport

Pushes data from a control workbook into many linked workbooks in one run.

### Setup

1. **Import** sheet — cells with external-link formulas, e.g. `='C:\Folder\[Book.xlsx]Sheet1'!A1`
2. **Export** sheet — same layout; plain values (or formulas for formula export) to write into those linked cells
3. Target workbooks must exist on disk; they are opened automatically if not already open

### Macros

| Macro | Mode | Action |
|-------|------|--------|
| `RunImportExportValue` | `exportValue` | Write Export **values** into linked cells |
| `RunImportExportFormula` | `exportFormula` | Write Export **formulas** (`Formula2`) into linked cells |
| `DryRun` | `dryRun` | Validate links and underline matched pairs; no writes |

Before updating, the tool backs up Import and Export as values-only sheets, confirms the target workbook list, and optionally saves and closes external workbooks after export.

### Link maintenance (`Module2.bas`)

- `ExtractLinksShort` — list external link sources into the sheet
- `UpdateLinks` — batch `ChangeLink` from a Yes/flag column and new path column

## TraceFormulaTool

Helpers for reviewing and simplifying formulas when tracing model logic.

| Module | Macro | Description |
|--------|-------|-------------|
| `Module3.bas` | `GotoHistory` (Ctrl+Alt+G) | Open `frmGotoHistory` — stack of recent cell locations; back navigation, manual goto, bookmark current selection |
| `ifsimplify.bas` | `SimplifyIFFormulas` | On selection: peel nested `IF`s when branches can be evaluated to literals |
| `lookupsimplify.bas` | `ConvertNestedLookupsToDirectReferences` | On selection: replace resolvable `VLOOKUP` / `XLOOKUP` / `INDEX` chains with direct cell references |

`DEBUG_IF_MODE` and `DEBUG_MODE` in those modules log parsing steps to the Immediate window (Ctrl+G).

## Development workflow

VBA is edited in Excel; source in this repo is kept in sync via extraction.

### Extract macros from workbooks

Requires [oletools](https://github.com/decalage2/oletools):

```bash
pip install oletools
python extractMacro.py
```

The script uses `olevba` to write each module (`.bas`, `.frm`, etc.) into the folder that contains the corresponding `.xlsm`. Paths in `extractMacro.py` use Windows-style separators; on macOS, adjust paths or run from a Windows environment if extraction fails.

### Import macros back into Excel

1. Open the `.xlsm` in Excel.
2. Alt+F11 → File → Import File, or paste module contents into existing modules.
3. For `.frm` files, import the form and ensure code-behind matches.
4. Save the workbook.

User-form builder modules (`createUserForm*.bas`, `CreateUserFormfrmGotoHistory.bas`) regenerate forms when run inside the VBA editor; use them only when rebuilding UI from scratch.

## Requirements

- Microsoft Excel with VBA macros enabled (`.xlsm`)
- Windows paths in formulas and UAT macros assume a typical actuarial model layout; adapt paths as needed
- Python 3 + `oletools` only for `extractMacro.py`

## License

Not specified in this repository; add a license file if you intend to share or publish the code.

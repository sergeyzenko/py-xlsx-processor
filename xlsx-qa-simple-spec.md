# XLSX Questionnaire Processor - Simple Interactive CLI

## Overview

A Python terminal app that:

- Loads an XLSX questionnaire.
- Extracts every textual cell and maintains a CSV “catalog” for human review.
- Prompts the user only for rows marked as questions (with optional default answers).
- Writes collected answers back to the workbook at user-specified locations.

The workflow is deliberately two-phase:

1. **Catalog preparation.** First run mirrors workbook text to `<source>_catalog.csv`. Analysts edit this file (classify `isQuestion`, supply default answers, etc.).
2. **Interactive answering.** Subsequent runs consume the curated catalog, prompt for outstanding answers, update the CSV, and persist responses into a copy of the workbook.

## Flow

```
┌─────────────┐     ┌─────────────┐     ┌──────────────┐     ┌──────────────┐     ┌─────────────┐     ┌─────────────┐
│  Bootstrap  │────▶│  Load XLSX  │────▶│  Extract Text │────▶│  Sync Catalog │────▶│  Interactive│────▶│  Write Back │
│  venv       │     │             │     │  Cells        │     │  (CSV)        │     │  Q&A Loop   │     │  & Save     │
└─────────────┘     └─────────────┘     └──────────────┘     └──────────────┘     └─────────────┘     └─────────────┘
```
### Startup Sequence

```
┌─────────────────────────────────────────────────────────────────────┐
│                         App Startup                                  │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  1. Check: Am I running from ./venv/bin/python?                     │
│     │                                                               │
│     ├─ NO ──▶ Does ./venv exist?                                    │
│     │         │                                                     │
│     │         ├─ NO ──▶ Create venv: python3 -m venv ./venv         │
│     │         │                                                     │
│     │         └─ YES ─┐                                             │
│     │                 │                                             │
│     │         ◀───────┘                                             │
│     │         Re-exec: os.execv('./venv/bin/python', [... args])    │
│     │         (Script restarts inside venv, flow continues)         │
│     │                                                               │
│     └─ YES ─▶ 2. Check: Is openpyxl importable?                     │
│               │                                                     │
│               ├─ NO ──▶ Run: pip install openpyxl                   │
│               │                                                     │
│               └─ YES ─▶ 3. Run main application                     │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

### venv Auto-Management

Bootstrap script handles virtual environment setup automatically prior to loading any third-party code.

1. Detect current interpreter. If not `./venv/bin/python`, create the venv (first run only) and `os.execv` back into it.
2. Once inside the managed venv, ensure required packages (`openpyxl`) are installed.
3. Proceed with application imports and execution.

Key properties:

- **No manual activation:** Users always invoke `python3 xlsx_qa.py ...`.
- **Process replacement:** `execv` guarantees subsequent logic runs inside the right interpreter.
- **Implicit teardown:** The venv is process-scoped; no explicit deactivation step is required.
- **Idempotent:** Repeated invocations skip creation/installation when already satisfied.

## Data Structures

### Catalog Record

```python
@dataclass
class CatalogRecord:
    tab_name: str
    text_location: str
    text_value: str
    is_question: bool | None = None
    text_response: str | None = None
    text_response_location: str | None = None
    is_default_answer: bool | None = None
    default_answer_question_location: str | None = None
    is_instruction: bool | None = None
```

- Records are keyed by `(tab_name, text_location)`.
- `text_value` captures the literal cell text (whitespace and newlines preserved).
- CSV columns match the fields above; boolean values are serialized as `true`/`false` for easy hand editing.
- Additional metadata (question classification, defaults, answers) is stored in-place and survives across runs.

---

## Module Breakdown

### 1. Loader

```python
def load_workbook(filepath: str) -> Workbook:
    """
    Load XLSX file, validate it's readable.
    Returns openpyxl Workbook object.
    Raises: FileNotFoundError, InvalidFileException
    """
```

### 2. Extractor

```python
def extract_catalog(workbook: Workbook) -> list[CatalogRecord]:
    """
    Iterate all sheets, all rows.
    Capture every textual cell (strings only) as a CatalogRecord.
    Preserve workbook traversal order: sheet order, row order, column order.
    Skip empty cells and non-text values.
    """
```

The extractor no longer classifies questions. Its sole purpose is to mirror worksheet text into an in-memory catalog that will later be merged with the persisted CSV.

### 3. Catalog Builder

```python
def sync_catalog(records: list[CatalogRecord], catalog_path: str) -> list[CatalogRecord]:
    """
    Load existing catalog CSV if present.
    Merge by (tab_name, text_location), preserving user-maintained metadata.
    Apply latest workbook text_value and append brand-new entries.
    Persist the refreshed catalog CSV (all fields quoted).
    Return the merged record list for downstream processing.
    """
```

CSV columns (all quoted to preserve whitespace and symbols):

| Column | Description |
|--------|-------------|
| `tabName` | Worksheet name |
| `textLocation` | Cell coordinate (e.g., `A13`) |
| `textValue` | Text content (newlines preserved) |
| `isQuestion` | `true` / `false` / blank |
| `textResponse` | Captured answer text |
| `textResponseLocation` | Target cell for the captured answer |
| `isDefaultAnswer` | Marks rows that supply default answers |
| `defaultAnswerQuestionLocation` | Coordinate the default answer applies to |
| `isInstruction` | Flags instructional rows |

### 4. Interactive Q&A

```python
def run_qa_session(records: list[CatalogRecord], default_answers: dict[str, str]) -> SessionResult:
    """
    Filter catalog records to those with is_question == True AND empty text_response.
    For each candidate:
      1. Display question context (sheet, location, text).
      2. Pre-populate the prompt with any default answer (lookup by defaultAnswerQuestionLocation)
         or previously captured text_response.
      3. Collect multiline answer input with navigation commands:
         - Enter immediately: accept default answer (if present) or skip.
         - "q" / "quit": abort session and persist progress.
         - "b" / "back": revisit prior question.
         - "s" / "skip": skip without saving.
      4. Prompt for textResponseLocation (suggest cell to the right if empty).
      5. Persist answer + location on the CatalogRecord.
    Return SessionResult describing whether the run finished or aborted.
    """
```

**Display format:**
```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Question 12/45  │  Sheet: Security  │  Cell: B15
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Do you have a documented incident response plan?

Answer (enter to accept default, 'q' to quit, 'b' to go back, 's' to skip):
> _
Response location [C15]:
> _
```

**Response location default:** When prompting for `textResponseLocation`, the CLI suggests the cell immediately to the right of the question cell. Users can override as needed.

### 5. Writer

```python
def write_answers(
    filepath: str,
    records: list[CatalogRecord],
    output_path: str | None = None
) -> str:
    """
    Load original workbook (fresh load, not from memory).
    For each record with both text_response and text_response_location,
    write the response to the target cell on the respective worksheet.
    Save to output_path (or filepath_answered.xlsx if not specified).
    
    Returns path to saved file.
    """
```

---

## CLI Interface

```bash
# Basic usage
python xlsx_qa.py questionnaire.xlsx

# Specify output file
python xlsx_qa.py questionnaire.xlsx -o completed.xlsx

# Override catalog location
python xlsx_qa.py questionnaire.xlsx --catalog custom/catalog.csv
```

## Edge Cases

| Situation | Handling |
|-----------|----------|
| Existing catalog CSV present | Merge rows by `(tabName, textLocation)` and preserve user edits |
| No rows flagged as questions | Skip Q&A loop, emit reminder, still persist catalog |
| Default answer rows defined | Auto-surface matching defaults based on `defaultAnswerQuestionLocation` |
| Response location omitted | Prompt until provided or user navigates back/quit |
| Non-XLSX file | Fail with clear error message |
| Empty workbook | Exit with "No textual entries found" |

---

## Dependencies

```
openpyxl>=3.1.0
```

That's it. No other dependencies needed for MVP.

**Note:** Dependencies are auto-installed by the bootstrap logic. User doesn't need to run `pip install` manually.

---

## File Structure

```
xlsx_qa/
├── xlsx_qa.py       # Entry point with bootstrap
├── __init__.py
├── app.py           # High-level orchestration
├── bootstrap.py     # Virtual environment management
├── cli.py           # Argument parsing
├── domain.py        # Catalog data models + helpers
├── extractor.py     # Workbook -> catalog record extraction
├── loader.py        # Workbook loading utilities
├── persistence.py   # Catalog CSV read/write
├── session.py       # Interactive Q&A loop
└── writer.py        # Workbook answer persistence
```

---

## Implementation Notes

1. **Extraction is exhaustive** - capture every textual cell and let users curate via the CSV rather than relying on heuristics.

2. **Provide sane defaults** - suggest the cell to the right for answers but always allow manual overrides for unusual layouts.

3. **Always save to new file** - never overwrite the source workbook when persisting responses.

4. **Catalog = progress** - write the CSV after every run (even on abort) so manual edits and captured answers are never lost.

---

## Example Session

```
$ python xlsx_qa.py vendor_questionnaire.xlsx

Creating virtual environment in .../venv...
Virtual environment created.
Switching to virtual environment...
Installing missing packages: openpyxl...
Packages installed.

Catalog saved to: vendor_questionnaire_catalog.csv
No unanswered questions flagged in catalog. Update the CSV and rerun when ready.

# analyst marks isQuestion=true for key prompts inside vendor_questionnaire_catalog.csv

$ python xlsx_qa.py vendor_questionnaire.xlsx

Catalog saved to: vendor_questionnaire_catalog.csv
Found 3 unanswered questions marked in catalog...

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Question 1/3  │  Sheet: Security  │  Cell: B3
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Does your organization have a formal information security policy?

Answer (enter to accept default, 'q' to quit, 'b' to go back, 's' to skip):
> Yes, we maintain an Information Security Policy reviewed annually.
Response location [C3]:
> 

...

Catalog saved to: vendor_questionnaire_catalog.csv
Answers written to: vendor_questionnaire_answered.xlsx
```

---

## Future Enhancements (Out of Scope for MVP)

- Answer suggestions from previous sessions
- Bulk import answers from JSON/YAML
- Custom answer column configuration
- XLS (legacy format) support
- Multi-line answer input mode

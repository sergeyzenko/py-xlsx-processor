# XLSX Questionnaire Processor - Simple Interactive CLI

## Overview

A Python terminal app that extracts questions from an XLSX file, prompts the user to answer each one interactively, and writes answers back to the spreadsheet.

---

## Flow

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│  Bootstrap  │────▶│  Load XLSX  │────▶│  Extract    │────▶│  Interactive│────▶│  Write Back │
│  venv       │     │             │     │  Questions  │     │  Q&A Loop   │     │  & Save     │
└─────────────┘     └─────────────┘     └─────────────┘     └─────────────┘     └─────────────┘
```

---

## venv Auto-Management

The app self-manages its virtual environment. User never manually activates/deactivates.

### Bootstrap Logic

```python
def bootstrap_venv():
    """
    Runs BEFORE any imports except stdlib.
    Must be at the very top of the entry point script.
    
    1. Detect if running inside the project's venv
    2. If not → create venv if missing, then re-exec into it
    3. Verify dependencies installed, install if missing
    """
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

### Implementation

```python
#!/usr/bin/env python3
"""
Entry point with venv auto-management.
This block runs before any third-party imports.
"""
import sys
import os
import subprocess

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VENV_DIR = os.path.join(SCRIPT_DIR, 'venv')
VENV_PYTHON = os.path.join(VENV_DIR, 'bin', 'python')
REQUIRED_PACKAGES = ['openpyxl']


def in_project_venv() -> bool:
    """Check if current Python executable is the project's venv."""
    return os.path.abspath(sys.executable) == os.path.abspath(VENV_PYTHON)


def venv_exists() -> bool:
    """Check if venv directory structure exists."""
    return (
        os.path.isdir(VENV_DIR) and
        os.path.isfile(VENV_PYTHON) and
        os.path.isfile(os.path.join(VENV_DIR, 'pyvenv.cfg'))
    )


def create_venv():
    """Create virtual environment in project root."""
    print(f"Creating virtual environment in {VENV_DIR}...")
    subprocess.check_call([sys.executable, '-m', 'venv', VENV_DIR])
    print("Virtual environment created.")


def reexec_in_venv():
    """Re-execute current script using venv Python."""
    print("Switching to virtual environment...")
    os.execv(VENV_PYTHON, [VENV_PYTHON] + sys.argv)


def ensure_packages():
    """Install missing packages in current venv."""
    missing = []
    for pkg in REQUIRED_PACKAGES:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    
    if missing:
        print(f"Installing missing packages: {', '.join(missing)}...")
        subprocess.check_call([
            sys.executable, '-m', 'pip', 'install', '--quiet', *missing
        ])
        print("Packages installed.")


def bootstrap():
    """Main bootstrap sequence."""
    if not in_project_venv():
        if not venv_exists():
            create_venv()
        reexec_in_venv()
        # Never reaches here - execv replaces process
    
    ensure_packages()


# === BOOTSTRAP RUNS HERE ===
bootstrap()

# === SAFE TO IMPORT THIRD-PARTY PACKAGES NOW ===
from openpyxl import load_workbook
# ... rest of application code


def main():
    # Application logic here
    pass


if __name__ == '__main__':
    try:
        main()
    finally:
        # No explicit deactivation needed - process exits, venv context dies
        pass
```

### Key Points

1. **No manual activation** - User just runs `python3 xlsx_qa.py`, bootstrap handles the rest

2. **Re-exec pattern** - Script re-launches itself using venv Python via `os.execv()`. This replaces the current process entirely.

3. **Deactivation is implicit** - When the script exits, the process ends. There's no shell session to "deactivate". The venv only exists for the script's lifetime.

4. **First run is slower** - Creates venv + installs packages. Subsequent runs skip straight to main.

5. **venv location** - Always `./venv` relative to script location, not cwd.

### User Experience

```bash
# First run (any Python 3)
$ python3 xlsx_qa.py questionnaire.xlsx
Creating virtual environment in /path/to/project/venv...
Virtual environment created.
Switching to virtual environment...
Installing missing packages: openpyxl...
Packages installed.

Loading: questionnaire.xlsx
Found 47 questions...

# Subsequent runs (instant start)
$ python3 xlsx_qa.py questionnaire.xlsx
Loading: questionnaire.xlsx
Found 47 questions...
```

### venv Directory Structure

After first run, project folder contains:

```
xlsx_qa/
├── xlsx_qa.py          # Main script
├── venv/               # Auto-created
│   ├── bin/
│   │   ├── python
│   │   ├── python3
│   │   ├── pip
│   │   └── ...
│   ├── lib/
│   │   └── python3.x/
│   │       └── site-packages/
│   │           └── openpyxl/
│   └── pyvenv.cfg
└── ...
```

### Cleanup

To reset environment:

```bash
rm -rf venv
# Next run will recreate it
```

---

## Data Structures

### Question Object

```python
@dataclass
class Question:
    id: int                    # Sequential ID for user reference
    sheet: str                 # Sheet name
    coord: str                 # Cell coordinate (e.g., "B5")
    row: int                   # Row number
    col: int                   # Column number
    text: str                  # Question text
    answer_coord: str          # Where to write the answer
    answer: str | None = None  # User's answer (filled during Q&A loop)
```

### Answer Target Strategy

Default: Answer goes in the cell immediately to the right of the question.

```
Question in B5  →  Answer in C5
Question in D10 →  Answer in E10
```

If the cell to the right already has content, flag it during extraction and let user confirm/override.

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
def extract_questions(workbook: Workbook) -> list[Question]:
    """
    Iterate all sheets, all rows.
    Identify cells that look like questions.
    Return ordered list of Question objects.
    
    Question identification heuristics:
    - Ends with "?"
    - Starts with question words (What, How, Do, Does, Is, Are, Can, Will, etc.)
    - Contains question patterns ("please describe", "please explain", "provide details")
    
    Skip:
    - Empty cells
    - Cells with only numbers
    - Cells with very short content (<10 chars) unless ends with "?"
    """
```

**Processing order:** Sheet by sheet (in workbook order), then row by row (top to bottom), then column by column (left to right).

### 3. Interactive Q&A

```python
def run_qa_session(questions: list[Question]) -> list[Question]:
    """
    For each question:
    1. Display: question number, sheet name, cell reference, question text
    2. Prompt user for answer
    3. Store answer in Question object
    
    Special inputs:
    - Empty input: skip question (leave answer as None)
    - "q" or "quit": exit early, save progress
    - "b" or "back": go back to previous question
    - "s" or "skip": skip current question
    
    Returns list with answers populated.
    """
```

**Display format:**
```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Question 12/45  │  Sheet: Security  │  Cell: B15
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Do you have a documented incident response plan?

Answer (enter to skip, 'q' to quit, 'b' to go back):
> _
```

### 4. Writer

```python
def write_answers(
    filepath: str,
    questions: list[Question],
    output_path: str | None = None
) -> str:
    """
    Load original workbook (fresh load, not from memory).
    Write each answer to its target cell.
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

# Resume from saved progress
python xlsx_qa.py questionnaire.xlsx --resume progress.json
```

---

## Session Persistence

If user quits mid-session, save progress to JSON:

```json
{
  "source_file": "questionnaire.xlsx",
  "timestamp": "2025-01-15T10:30:00Z",
  "current_index": 12,
  "questions": [
    {
      "id": 1,
      "sheet": "Security",
      "coord": "B5",
      "answer_coord": "C5",
      "text": "Do you have a security policy?",
      "answer": "Yes, documented and reviewed annually."
    },
    {
      "id": 2,
      "sheet": "Security", 
      "coord": "B6",
      "answer_coord": "C6",
      "text": "Is there an incident response plan?",
      "answer": null
    }
  ]
}
```

---

## Edge Cases

| Situation | Handling |
|-----------|----------|
| Answer cell already has content | Warn user, ask to confirm overwrite |
| Question cell is merged | Use top-left cell as reference, answer goes to right of merge region |
| Very long question text | Wrap display at terminal width |
| Multi-line answer needed | Support `\n` escape or multi-line input mode (triple-enter to finish) |
| Non-XLSX file | Fail with clear error message |
| Empty workbook | Exit with "No questions found" |

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
├── xlsx_qa.py        # Single file implementation, or split into:
├── loader.py         # Workbook loading
├── extractor.py      # Question extraction
├── session.py        # Interactive Q&A loop
├── writer.py         # Answer writing
└── models.py         # Question dataclass
```

For simplicity, can be a single ~200 line file.

---

## Implementation Notes

1. **Question detection is fuzzy** - start conservative (only "?" endings), let user feedback drive expansion of heuristics

2. **Don't get clever with answer placement** - "cell to the right" works for 90% of cases. If a document has a different pattern, user can process it manually or we add config later

3. **Always save to new file** - never overwrite input

4. **Progress saving is essential** - 100+ question documents are common, users will quit mid-session

---

## Example Session

```
$ python xlsx_qa.py vendor_questionnaire.xlsx

Loading: vendor_questionnaire.xlsx
Found 3 sheets: ['Security', 'Compliance', 'Technical']
Extracted 47 questions

Starting Q&A session...
(Commands: enter=skip, q=quit & save, b=back)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Question 1/47  │  Sheet: Security  │  Cell: B3
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Does your organization have a formal information security policy?

> Yes, we maintain an Information Security Policy reviewed annually.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Question 2/47  │  Sheet: Security  │  Cell: B4
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Is the policy approved by executive management?

> Yes, approved by CISO and CEO.

...

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Session complete!
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Answered: 42/47
Skipped: 5

Saved to: vendor_questionnaire_answered.xlsx
```

---

## Future Enhancements (Out of Scope for MVP)

- Answer suggestions from previous sessions
- Bulk import answers from JSON/YAML
- Custom answer column configuration
- XLS (legacy format) support
- Multi-line answer input mode

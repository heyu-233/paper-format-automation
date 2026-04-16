# paper-format-automation

Template-driven formatting automation for Chinese journal manuscripts.

This repository packages a local-first Codex skill plus supporting scripts for:
- extracting formatting rules from an official journal Word template
- checking a manuscript against those rules
- applying conservative auto-formatting to deterministic layout items
- producing a review report for anything that still needs manual confirmation

It is designed for the practical workflow: `template.docx/.doc + manuscript.docx -> rules + diff report + formatted docx`.

## What it does

- Extracts structured formatting rules from a journal template
- Detects major manuscript blocks such as title, abstract, headings, body, captions, references, and header/footer
- Generates a diff report between the manuscript and the extracted template rules
- Applies conservative formatting fixes for high-confidence items
- Supports `.doc` template preprocessing on Windows through Microsoft Word COM

## Current scope

v1 is focused on Chinese journal Word-template workflows:
- template: `.docx` or legacy `.doc`
- manuscript: `.docx`
- output:
  - `template_rules.json`
  - `diff_report.json`
  - `review-report.md`
  - `formatted.docx`

This is not a general-purpose "format any paper from any source" tool yet.

## Repository layout

```text
.
├─ SKILL.md                     # Skill entry and usage guidance
├─ agents/
│  └─ openai.yaml               # Skill metadata
├─ references/
│  ├─ workflow.md               # Workflow notes
│  ├─ tooling.md                # Local dependency notes
│  ├─ rule-schema.md            # Extracted rule structure
│  └─ supported-format-items.md # v1 support boundary
└─ scripts/
   ├─ extract_template_rules.py
   ├─ check_manuscript.py
   ├─ format_manuscript.py
   ├─ run_pipeline.py
   ├─ prepare_template.ps1
   ├─ run_docx4j_formatter.ps1
   └─ java/
      ├─ src/
      ├─ out/
      └─ formatter.jar
```

## Requirements

- Windows
- Python 3.11+ recommended
- Microsoft Word installed if you want to convert legacy `.doc` templates
- Java available if you want to use the Java-side formatter launcher path

Python dependencies are listed in `scripts/requirements.txt`.

## Quick start

### 1. Install dependencies

```powershell
pip install -r scripts/requirements.txt
```

### 2. Run the full pipeline

```powershell
python scripts/run_pipeline.py `
  --template "D:\path\to\journal-template.doc" `
  --manuscript "D:\path\to\manuscript.docx" `
  --outdir "D:\path\to\output" `
  --mode format
```

### 3. Check the outputs

Typical outputs inside `outdir`:
- `template_rules.json`
- `diff_report.json`
- `review-report.md`
- `formatted.docx`

## Core scripts

- `scripts/extract_template_rules.py`  
  Build a structured rule set from a template `.docx`

- `scripts/check_manuscript.py`  
  Compare a manuscript against extracted rules and generate machine-readable + markdown reports

- `scripts/format_manuscript.py`  
  Apply conservative formatting changes to a manuscript using extracted rules

- `scripts/run_pipeline.py`  
  Run preprocessing, extraction, checking, and formatting in one flow

- `scripts/prepare_template.ps1`  
  Convert `.doc` templates to `.docx` on Windows using Word COM

## Design principles

- Template-first instead of hardcoded journal constants
- Conservative auto-fix only for deterministic items
- Report unknown or risky cases instead of silently guessing
- Prefer block-level structure matching over single-paragraph guessing

## Current limitations

- v1 is centered on Word-based Chinese journal templates
- Complex front-page layouts may still need manual review
- Semantic correctness is out of scope; this tool focuses on formatting structure
- PDF-only template interpretation is not the main workflow

## For Codex users

This repository is also packaged as a local Codex skill. See:
- `SKILL.md`
- `agents/openai.yaml`

## Project status

This is the first open-source cut of the local skill and automation chain. It is already useful for real template-driven formatting work, but it is still evolving toward broader template coverage and stronger structure recognition.

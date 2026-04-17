# Examples

This directory documents example workflows for `paper-format-automation`.

The repository intentionally does not bundle your private manuscript or journal template files.
Instead, this folder explains how to run the tool on your own local inputs and what outputs to expect.

## Example workflow

Input:
- journal template: `template.docx` or legacy `template.doc`
- manuscript: `manuscript.docx`

Command:

```powershell
python ..\scripts\run_pipeline.py `
  --template "D:\path\to\template.doc" `
  --manuscript "D:\path\to\manuscript.docx" `
  --outdir "D:\path\to\example-output" `
  --mode format
```

Expected outputs:
- `template_rules.json`
- `diff_report.json`
- `review-report.md`
- `formatted.docx`

## Suggested example-output layout

```text
example-output/
├─ template_rules.json
├─ diff_report.json
├─ review-report.md
└─ formatted.docx
```

## What to inspect manually

- whether page layout and section layout match the journal template
- whether body, headings, abstract, captions, and references follow the template rules
- whether any remaining front-page items still require manual adjustment
- whether the review report flags anything that should not be silently auto-fixed

## Notes

- If your template is `.doc`, the pipeline converts it to `.docx` first on Windows
- For public sharing, prefer sanitized templates and manuscripts rather than real submission materials

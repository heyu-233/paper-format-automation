# Workflow

## Primary v1 workflow

1. Collect inputs
   - official journal template `.docx` or legacy `.doc`
   - manuscript `.docx`
   - optional notes from the journal submission guide
2. Pre-process template when needed
   - if the template is `.doc`, run `prepare_template.ps1` to convert it to `.docx`
3. Extract template rules
   - run `extract_template_rules.py`
   - review the generated `template_rules.json`
4. Check manuscript against template
   - run `check_manuscript.py`
   - inspect matched items, auto-fix candidates, and manual-review items
5. Auto-format conservatively
   - run `run_docx4j_formatter.ps1`
   - generate a formatted output `.docx`
6. Deliver results
   - formatted manuscript if available
   - markdown review report
   - note all skipped or risky items

## Decision rules

- If the template is `.doc`, convert it before any rule extraction.
- If template or manuscript is not a Word file the current pipeline supports, downgrade to report-only mode or stop with a clear boundary message.
- If template rules are weak or ambiguous, prefer reporting over auto-fixing.
- If a formatting change could alter meaning or structure, do not auto-apply it.

## Expected outputs

- `template_rules.json`
- `diff_report.json`
- `review-report.md`
- `formatted.docx` when the formatter runs

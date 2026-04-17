# Tooling

## Current v1 stack

- Java 17 runtime for the formatter launcher
- Microsoft Word COM automation for optional `.doc` to `.docx` conversion on Windows
- `python-docx` for `.docx` inspection and conservative formatting
- `docx2python` for content extraction support
- PowerShell wrappers for template conversion and formatter invocation

## Formatter behavior

The formatter entrypoint is exposed through:

`script_dir/java/formatter.jar`

The current launcher delegates to the local formatting script so that the full pipeline works without Maven. `scripts/java/lib/` remains reserved for a future bundled docx4j integration.

## Template preprocessing

Legacy `.doc` templates are not used directly in the main pipeline. Convert them first with:

`prepare_template.ps1 <template.doc> <template.docx>`

If Word COM is unavailable, the skill should stop and ask for a `.docx` template.

## Recommended future additions

- `docxtpl` for fixed metadata filling
- `docx-compare` for before/after validation reports
- vendored docx4j jars if we want a pure Java formatter core

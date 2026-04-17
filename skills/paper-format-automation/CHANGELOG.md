# Changelog

All notable changes to this project will be documented in this file.

The format is inspired by Keep a Changelog, with practical notes for this repository.

## [Unreleased]

### Planned
- Expand template coverage beyond the current core Chinese journal workflows
- Improve structure recognition for more front-page layouts and mixed-style references
- Add more example assets and regression fixtures

## [0.1.0] - 2026-04-16

### Added
- Initial open-source release of the `paper-format-automation` project
- Template rule extraction from Chinese journal `.docx` templates
- Manuscript/template diff reporting with JSON and Markdown outputs
- Conservative manuscript formatter pipeline
- Windows `.doc` template preprocessing through Word COM
- Local skill packaging via `SKILL.md` and `agents/openai.yaml`
- Java-side formatter launcher stub and Python fallback path

### Changed
- Improved style-aware extraction by reading inherited style-chain font settings instead of only direct run formatting
- Tightened alignment for core blocks including body, heading 1, heading 2, Chinese abstract, and reference entries
- Synced template style definitions into formatted output to reduce inheritance drift from the source manuscript
- Improved repository README and fixed skill metadata text encoding issues

### Notes
- This version is intended as a usable v1 for template-driven Chinese journal Word formatting
- The formatter remains conservative and still prioritizes explicit reporting over risky silent edits

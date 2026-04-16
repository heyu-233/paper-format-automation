---
name: paper-format-automation
description: Use when the user wants a Chinese journal paper or similar .docx manuscript checked against an official journal template, reformatted toward a template, or turned into a submission-ready version with a review report. Best for template-driven formatting, structure checks, and conservative auto-fixes on Word documents.
---

# Paper Format Automation

## Overview

Use this skill when the user provides a Chinese journal template and a manuscript, and wants a template-aligned submission version or a format-difference report.

This skill is template-first and conservative:
- extract rules from the official template
- compare the manuscript against those rules
- auto-fix deterministic formatting items
- report anything risky for manual review instead of guessing

V1 is optimized for Chinese journal `.docx` template + `.docx` manuscript workflows. It also supports a template pre-processing step for legacy `.doc` templates by converting them to `.docx` before the main pipeline runs.

## When to use this skill

Use this skill for requests like:
- “根据这个期刊模板检查论文格式”
- “把我的论文按投稿模板排版”
- “对比模板和论文，给我差异清单”
- “生成接近投稿版的 Word 文档，并标出还要人工确认的地方”

Do not use this skill as the primary workflow for:
- English journal LaTeX submission formatting
- scanned PDF-only template alignment
- content rewriting, polishing, or academic editing without a template-driven formatting goal

## Workflow

1. If the template is `.doc`, run `scripts/prepare_template.ps1` first to convert it to `.docx`.
2. Run `scripts/extract_template_rules.py` on the template `.docx` to generate `template_rules.json`.
3. Run `scripts/check_manuscript.py` with the manuscript and extracted rules to generate a structured diff report.
4. If the user wants an auto-formatted draft, call `scripts/run_docx4j_formatter.ps1` with the manuscript, rules, and output path.
5. Return both outputs when possible:
   - formatted `.docx`
   - review report listing matched items, auto-fixed items, and manual-review items

## Tooling rules

- Prefer the Python scripts for rule extraction, diff reporting, and conservative formatting.
- Prefer the PowerShell wrappers for template conversion and formatter invocation; they centralize environment-specific behavior.
- Never silently rewrite high-risk items such as unclear section structures, ambiguous header/footer logic, or reference content corrections.
- If the environment cannot convert `.doc` templates automatically, stop with a clear instruction instead of guessing.

## Supported v1 items

Read `references/supported-format-items.md` before promising automatic fixes. In short:
- auto-fix candidates: page layout, heading styles, body paragraph spacing, abstract/keyword labels, caption paragraph styles, reference paragraph styles
- report-only items: semantic section reordering, reference-content correctness, image/table content changes, unclear template-specific front matter

## References

- For the end-to-end flow and operator guidance, read `references/workflow.md`.
- For the rule JSON structure, read `references/rule-schema.md`.
- For support boundaries, read `references/supported-format-items.md`.
- For local dependency expectations, read `references/tooling.md`.

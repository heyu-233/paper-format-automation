# Roadmap

## Vision

Make `paper-format-automation` a practical, template-driven formatting toolkit for Chinese journal manuscripts that is reliable enough for real submission work and transparent enough for manual review when automation should stop.

## v0.1.x - Stabilize the current workflow

Focus:
- make the current `.doc/.docx template + .docx manuscript` workflow more reliable
- reduce false matches in structure recognition
- improve output consistency for high-value formatting blocks

Planned items:
- more regression tests using real-world journal templates
- stronger caption and reference block recognition
- better first-page/front-matter mapping
- cleaner repository documentation and examples

## v0.2 - Broader template coverage

Focus:
- support more template variants without per-journal hacks
- improve the distinction between template-driven values and generic application strategies

Planned items:
- stronger section and column handling
- richer page-header/page-footer reconstruction
- more robust handling of bilingual front matter
- improved handling of mixed reference styles

## v0.3 - Reviewability and verification

Focus:
- make output easier to trust and easier to inspect

Planned items:
- before/after formatting comparison report
- block-level verification summary
- more explicit manual-review categories
- sample regression suite and repeatable local validation workflow

## Future directions

- parse web/PDF submission instructions as supplemental hints
- support more journal families and house styles
- provide cleaner plugin/skill integration for reusable local workflows
- explore safer semi-automatic front-page reconstruction

## Non-goals for now

- full semantic editing or academic rewriting
- PDF-first exact visual imitation
- guaranteed one-click formatting for every journal template

# Supported Format Items

## Auto-fix candidates in v1

- first section page size, margins, and header/footer distance
- title, heading, and body paragraph formatting when the manuscript block is recognized structurally
- Chinese and English abstract / keyword paragraph formatting
- figure caption and table caption paragraph formatting
- reference-entry paragraph formatting when the reference block is recognized
- author bio paragraph formatting when the block is recognized

## Report-first items in v1

- author list correctness and affiliation content correctness
- fund metadata correctness
- semantic section reordering
- complex multi-section header/footer recreation
- reference content correctness and citation normalization
- image/table content editing
- scanned PDF template interpretation

## Default risk policy

- Prefer block-level recognition over single-paragraph guessing.
- Do not auto-fix a block if the checker cannot confidently locate the corresponding manuscript block.
- If block counts differ heavily from the template, report the issue even when some style fields match.

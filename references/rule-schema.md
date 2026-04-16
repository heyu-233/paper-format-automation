# Rule Schema

`template_rules.json` now uses a block-oriented structure instead of only a flat style snapshot.

```json
{
  "source": {
    "template_path": "..."
  },
  "page_layout": {
    "page_width": 0,
    "page_height": 0,
    "top_margin": 0,
    "bottom_margin": 0,
    "left_margin": 0,
    "right_margin": 0,
    "header_distance": 0,
    "footer_distance": 0
  },
  "section_layouts": [
    {
      "section_index": 0,
      "start_paragraph_index": 0,
      "end_paragraph_index": 0,
      "columns_num": 1,
      "columns_space_pt": 0
    }
  ],
  "blocks": {
    "title": {...},
    "authors": {...},
    "affiliations": {...},
    "abstract_cn": {...},
    "keywords_cn": {...},
    "title_en": {...},
    "abstract_en": {...},
    "heading_1": {...},
    "heading_2": {...},
    "body": {...},
    "caption_figure": {...},
    "caption_table": {...},
    "references_title": {...},
    "reference_entry": {...},
    "fund": {...},
    "author_bio_title": {...},
    "author_bio_entry": {...},
    "header": {...},
    "footer": {...}
  },
  "styles": {
    "title": {...},
    "heading_1": {...},
    "heading_2": {...},
    "body": {...},
    "abstract": {...},
    "keywords": {...},
    "caption": {...},
    "references": {...}
  },
  "notes": []
}
```

## Block object fields

Each block may include:
- `role`
- `role_label`
- `style_name`
- `alignment`
- `line_spacing_mode`
- `line_spacing_value`
- `space_before_pt`
- `space_after_pt`
- `left_indent_pt`
- `right_indent_pt`
- `first_line_indent_pt`
- `hanging_indent_pt`
- `font_east_asia`
- `font_ascii`
- `font_hansi`
- `font_size_pt`
- `bold`
- `italic`
- `sample_count`
- `sample_texts`
- `instructional_filtered_count`

## Design notes

- `blocks` is the primary source for extraction, checking, and later formatting.
- `section_layouts` captures section boundaries, column count, and column spacing so single-column vs double-column templates are visible in rules and reports.
- `styles` is kept as a compatibility view for older formatter/checker steps.
- Line spacing is now split into `line_spacing_mode` and `line_spacing_value` so exact spacing and multiple spacing are not mixed together.
- Chinese and Latin fonts are tracked separately so mixed-script paragraphs are less likely to be flattened incorrectly.

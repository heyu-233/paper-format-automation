from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict

from docx import Document

from docx_rule_utils import (
    ROLE_LABELS,
    build_block_rule,
    detect_blocks,
    filter_instructional_paragraphs,
    section_layouts,
)


def _fmt_length(length):
    return round(length.pt, 2) if length is not None else None


def _header_footer_rule(doc: Document, role: str) -> Dict[str, Any]:
    paragraphs = detect_blocks(doc).get(role, [])
    if not paragraphs:
        return {}
    rule = build_block_rule(paragraphs)
    rule["role_label"] = ROLE_LABELS[role]
    return rule


def _compat_styles(blocks: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    mapping = {
        "title": "title",
        "heading_1": "heading_1",
        "heading_2": "heading_2",
        "body": "body",
        "abstract": "abstract_cn",
        "keywords": "keywords_cn",
        "caption": "caption_figure",
        "references": "reference_entry",
    }
    styles: Dict[str, Dict[str, Any]] = {}
    for old_key, block_key in mapping.items():
        if block_key in blocks:
            styles[old_key] = {
                key: value
                for key, value in blocks[block_key].items()
                if key not in {"sample_count", "sample_texts", "role", "role_label"}
            }
    return styles


def build_rules(template_path: Path) -> Dict[str, Any]:
    doc = Document(str(template_path))
    section = doc.sections[0]
    detected = detect_blocks(doc)
    block_rules: Dict[str, Dict[str, Any]] = {}
    notes = []

    for role in ROLE_LABELS:
        paragraphs = filter_instructional_paragraphs(detected.get(role, []), role)
        if not paragraphs:
            notes.append(f"Missing template block: {role}")
            continue
        block_rules[role] = build_block_rule(paragraphs)
        block_rules[role]["role"] = role
        block_rules[role]["role_label"] = ROLE_LABELS[role]
        original_count = len(detected.get(role, []))
        if original_count != len(paragraphs):
            block_rules[role]["instructional_filtered_count"] = original_count - len(paragraphs)

    rules = {
        "source": {"template_path": str(template_path)},
        "page_layout": {
            "page_width": _fmt_length(section.page_width),
            "page_height": _fmt_length(section.page_height),
            "top_margin": _fmt_length(section.top_margin),
            "bottom_margin": _fmt_length(section.bottom_margin),
            "left_margin": _fmt_length(section.left_margin),
            "right_margin": _fmt_length(section.right_margin),
            "header_distance": _fmt_length(section.header_distance),
            "footer_distance": _fmt_length(section.footer_distance),
        },
        "section_layouts": section_layouts(doc),
        "blocks": block_rules,
        "styles": _compat_styles(block_rules),
        "notes": notes,
    }
    return rules


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract structured formatting rules from a journal template .docx")
    parser.add_argument("template", type=Path, help="Path to the template .docx")
    parser.add_argument("-o", "--output", type=Path, default=None, help="Output JSON path")
    args = parser.parse_args()

    if args.template.suffix.lower() != ".docx":
        raise SystemExit("Template must be a .docx file")

    output = args.output or args.template.with_name("template_rules.json")
    rules = build_rules(args.template)
    output.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote rules to {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

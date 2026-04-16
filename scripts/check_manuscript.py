from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List, Tuple

from docx import Document

from docx_rule_utils import ROLE_LABELS, build_block_rule, detect_blocks, section_layouts

SAFE_FORMAT_FIELDS = {
    "alignment",
    "line_spacing_mode",
    "line_spacing_value",
    "space_before_pt",
    "space_after_pt",
    "left_indent_pt",
    "right_indent_pt",
    "first_line_indent_pt",
    "hanging_indent_pt",
    "font_east_asia",
    "font_ascii",
    "font_hansi",
    "font_size_pt",
    "bold",
    "italic",
}

COUNT_TOLERANT_ROLES = {"reference_entry", "caption_figure", "caption_table", "heading_1", "heading_2", "body"}
STYLE_NAME_TOLERANT_ROLES = {"caption_figure", "caption_table", "references_title"}


def _fmt_length(length):
    return round(length.pt, 2) if length is not None else None


def _compare_value(field: str, expected: Any, actual: Any) -> Tuple[str, Dict[str, Any]]:
    if expected is None:
        return "skip", {}
    if actual == expected:
        return "matched", {"field": field, "expected": expected, "actual": actual}
    return "mismatch", {"field": field, "expected": expected, "actual": actual}


def _compare_block(role: str, rule: Dict[str, Any], actual_rule: Dict[str, Any], actual_count: int) -> Dict[str, Any]:
    result = {
        "role": role,
        "role_label": ROLE_LABELS.get(role, role),
        "status": "matched",
        "matched": [],
        "mismatches": [],
        "manual_review": [],
        "expected_sample_count": rule.get("sample_count"),
        "actual_sample_count": actual_count,
        "expected_samples": rule.get("sample_texts", []),
        "actual_samples": actual_rule.get("sample_texts", []),
    }

    if actual_count == 0:
        result["status"] = "manual_review"
        result["manual_review"].append("No matching block found in manuscript")
        return result

    for field, expected in rule.items():
        if field in {"role", "role_label", "sample_count", "sample_texts", "instructional_filtered_count"}:
            continue
        if field == "style_name" and role in STYLE_NAME_TOLERANT_ROLES:
            continue
        kind, payload = _compare_value(field, expected, actual_rule.get(field))
        if kind == "matched":
            result["matched"].append(payload)
        elif kind == "mismatch":
            result["mismatches"].append(payload)

    if (
        role not in COUNT_TOLERANT_ROLES
        and result["expected_sample_count"] is not None
        and actual_count != result["expected_sample_count"]
    ):
        result["manual_review"].append(
            f"Sample count differs: expected {result['expected_sample_count']}, actual {actual_count}"
        )

    if result["mismatches"]:
        risky = [item for item in result["mismatches"] if item["field"] not in SAFE_FORMAT_FIELDS]
        if risky or result["manual_review"]:
            result["status"] = "manual_review"
        else:
            result["status"] = "auto_fix_candidate"
    elif result["manual_review"]:
        result["status"] = "manual_review"

    return result


def build_report(manuscript: Path, rules_path: Path) -> Dict[str, Any]:
    rules = json.loads(rules_path.read_text(encoding="utf-8"))
    doc = Document(str(manuscript))
    section = doc.sections[0]

    manuscript_blocks = detect_blocks(doc)
    manuscript_sections = section_layouts(doc)
    page_actual = {
        "page_width": _fmt_length(section.page_width),
        "page_height": _fmt_length(section.page_height),
        "top_margin": _fmt_length(section.top_margin),
        "bottom_margin": _fmt_length(section.bottom_margin),
        "left_margin": _fmt_length(section.left_margin),
        "right_margin": _fmt_length(section.right_margin),
        "header_distance": _fmt_length(section.header_distance),
        "footer_distance": _fmt_length(section.footer_distance),
    }

    layout_mismatches = []
    for field, expected in rules.get("page_layout", {}).items():
        if expected is None:
            continue
        actual = page_actual.get(field)
        if actual != expected:
            layout_mismatches.append({"field": field, "expected": expected, "actual": actual})

    report = {
        "source": {
            "manuscript_path": str(manuscript),
            "rules_path": str(rules_path),
        },
        "page_layout": {
            "status": "matched" if not layout_mismatches else "auto_fix_candidate",
            "mismatches": layout_mismatches,
        },
        "section_layouts": {
            "status": "matched",
            "expected": rules.get("section_layouts", []),
            "actual": manuscript_sections,
            "mismatches": [],
        },
        "block_checks": [],
        "style_checks": [],
        "summary": {
            "matched": 0,
            "auto_fix_candidate": 0,
            "manual_review": 0,
        },
    }

    expected_sections = rules.get("section_layouts", [])
    if len(expected_sections) != len(manuscript_sections):
        report["section_layouts"]["status"] = "manual_review"
        report["section_layouts"]["mismatches"].append(
            {
                "field": "section_count",
                "expected": len(expected_sections),
                "actual": len(manuscript_sections),
            }
        )
    for index, expected_section in enumerate(expected_sections):
        if index >= len(manuscript_sections):
            break
        actual_section = manuscript_sections[index]
        for field in ("columns_num", "columns_space_pt", "header_distance", "footer_distance"):
            expected = expected_section.get(field)
            actual = actual_section.get(field)
            if expected != actual:
                report["section_layouts"]["mismatches"].append(
                    {
                        "section_index": index,
                        "field": field,
                        "expected": expected,
                        "actual": actual,
                    }
                )
    if report["section_layouts"]["mismatches"] and report["section_layouts"]["status"] != "manual_review":
        report["section_layouts"]["status"] = "manual_review"

    for role, rule in rules.get("blocks", {}).items():
        actual_paragraphs = manuscript_blocks.get(role, [])
        actual_rule = build_block_rule(actual_paragraphs) if actual_paragraphs else {}
        result = _compare_block(role, rule, actual_rule, len(actual_paragraphs))
        report["block_checks"].append(result)
        report["summary"][result["status"]] += 1

    for legacy_role, legacy_rule in rules.get("styles", {}).items():
        mapped_role = {
            "abstract": "abstract_cn",
            "keywords": "keywords_cn",
            "caption": "caption_figure",
            "references": "reference_entry",
        }.get(legacy_role, legacy_role)
        actual_paragraphs = manuscript_blocks.get(mapped_role, [])
        actual_rule = build_block_rule(actual_paragraphs) if actual_paragraphs else {}
        style_result = {
            "category": legacy_role,
            "mapped_role": mapped_role,
            "status": "manual_review" if not actual_paragraphs else "matched",
            "matched": [],
            "mismatches": [],
            "manual_review": [],
        }
        if not actual_paragraphs:
            style_result["manual_review"].append("No matching block found in manuscript")
        else:
            for field, expected in legacy_rule.items():
                if field == "style_name":
                    continue
                kind, payload = _compare_value(field, expected, actual_rule.get(field))
                if kind == "matched":
                    style_result["matched"].append(payload)
                elif kind == "mismatch":
                    style_result["mismatches"].append(payload)
            if style_result["mismatches"]:
                style_result["status"] = "auto_fix_candidate"
        report["style_checks"].append(style_result)

    report["summary"][report["page_layout"]["status"]] += 1
    report["summary"][report["section_layouts"]["status"]] += 1
    return report


def write_markdown(report: Dict[str, Any], output: Path) -> None:
    lines: List[str] = []
    lines.append("# Review Report")
    lines.append("")
    lines.append(f"- Manuscript: `{report['source']['manuscript_path']}`")
    lines.append(f"- Rules: `{report['source']['rules_path']}`")
    lines.append("")
    lines.append("## Summary")
    for key, value in report["summary"].items():
        lines.append(f"- {key}: {value}")
    lines.append("")
    lines.append("## Page Layout")
    if not report["page_layout"]["mismatches"]:
        lines.append("- matched")
    else:
        for item in report["page_layout"]["mismatches"]:
            lines.append(f"- {item['field']}: expected `{item['expected']}`, actual `{item['actual']}`")
    lines.append("")
    lines.append("## Section Layouts")
    if not report["section_layouts"]["mismatches"]:
        lines.append("- matched")
    else:
        for item in report["section_layouts"]["mismatches"]:
            prefix = f"section {item['section_index']} " if "section_index" in item else ""
            lines.append(f"- {prefix}{item['field']}: expected `{item['expected']}`, actual `{item['actual']}`")
    lines.append("")
    lines.append("## Block Checks")
    for item in report["block_checks"]:
        lines.append(f"### {item['role_label']} ({item['role']}) - {item['status']}")
        lines.append(f"- expected samples: {item['expected_sample_count']}, actual samples: {item['actual_sample_count']}")
        if item["expected_samples"]:
            lines.append(f"- template sample: `{item['expected_samples'][0]}`")
        if item["actual_samples"]:
            lines.append(f"- manuscript sample: `{item['actual_samples'][0]}`")
        for mismatch in item["mismatches"]:
            lines.append(f"- {mismatch['field']}: expected `{mismatch['expected']}`, actual `{mismatch['actual']}`")
        for note in item["manual_review"]:
            lines.append(f"- manual review: {note}")
        if not item["mismatches"] and not item["manual_review"]:
            lines.append("- matched")
        lines.append("")
    output.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Check a manuscript .docx against structured template rules")
    parser.add_argument("manuscript", type=Path)
    parser.add_argument("rules", type=Path)
    parser.add_argument("-o", "--output", type=Path, default=None, help="JSON output path")
    parser.add_argument("--markdown", type=Path, default=None, help="Optional markdown report path")
    args = parser.parse_args()

    report = build_report(args.manuscript, args.rules)
    json_out = args.output or args.manuscript.with_name("diff_report.json")
    json_out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.markdown:
        write_markdown(report, args.markdown)
    print(f"Wrote diff report to {json_out}")
    if args.markdown:
        print(f"Wrote markdown report to {args.markdown}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

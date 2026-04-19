from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
SUPPORTED_TEMPLATE_SUFFIXES = {".doc", ".docx"}


def run(cmd):
    print("RUN:", " ".join(str(c) for c in cmd))
    subprocess.run(cmd, check=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the paper format automation pipeline")
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--manuscript", required=True, type=Path)
    parser.add_argument("--outdir", required=True, type=Path)
    parser.add_argument("--mode", choices=["check", "format"], default="check")
    args = parser.parse_args()

    template_suffix = args.template.suffix.lower()
    manuscript_suffix = args.manuscript.suffix.lower()
    if template_suffix not in SUPPORTED_TEMPLATE_SUFFIXES:
        raise SystemExit("Template must be a .doc or .docx file")
    if manuscript_suffix != ".docx":
        raise SystemExit("Manuscript must be a .docx file")

    outdir = args.outdir
    outdir.mkdir(parents=True, exist_ok=True)

    if template_suffix == ".docx":
        prepared_template = args.template
    else:
        prepared_template = outdir / (args.template.stem + ".docx")
        run([
            "powershell", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPT_DIR / "prepare_template.ps1"),
            "-InputTemplate", str(args.template), "-OutputDocx", str(prepared_template)
        ])

    rules_path = outdir / "template_rules.json"
    diff_path = outdir / "diff_report.json"
    review_path = outdir / "review-report.md"
    formatted_path = outdir / "formatted.docx"

    run([sys.executable, str(SCRIPT_DIR / "extract_template_rules.py"), str(prepared_template), "-o", str(rules_path)])
    run([sys.executable, str(SCRIPT_DIR / "check_manuscript.py"), str(args.manuscript), str(rules_path), "-o", str(diff_path), "--markdown", str(review_path)])

    if args.mode == "format":
        run([
            "powershell", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPT_DIR / "run_docx4j_formatter.ps1"),
            "-InputDocx", str(args.manuscript), "-RulesJson", str(rules_path), "-OutputDocx", str(formatted_path),
            "-PythonExe", sys.executable,
        ])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = ROOT / "skills" / "paper-format-automation" / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import run_pipeline


class RunPipelineTests(unittest.TestCase):
    def test_rejects_non_word_template(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            template = tmp_path / "template.pdf"
            manuscript = tmp_path / "paper.docx"
            outdir = tmp_path / "out"
            template.write_text("x", encoding="utf-8")
            manuscript.write_text("x", encoding="utf-8")

            argv = [
                "run_pipeline.py",
                "--template",
                str(template),
                "--manuscript",
                str(manuscript),
                "--outdir",
                str(outdir),
            ]
            with patch.object(sys, "argv", argv):
                with self.assertRaises(SystemExit) as ctx:
                    run_pipeline.main()

            self.assertIn(".doc or .docx", str(ctx.exception))

    def test_format_mode_passes_active_python_to_powershell_wrapper(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            template = tmp_path / "template.docx"
            manuscript = tmp_path / "paper.docx"
            outdir = tmp_path / "out"
            template.write_text("x", encoding="utf-8")
            manuscript.write_text("x", encoding="utf-8")

            commands = []

            def fake_run(cmd):
                commands.append(cmd)

            argv = [
                "run_pipeline.py",
                "--template",
                str(template),
                "--manuscript",
                str(manuscript),
                "--outdir",
                str(outdir),
                "--mode",
                "format",
            ]
            with patch.object(sys, "argv", argv), patch.object(run_pipeline, "run", side_effect=fake_run):
                rc = run_pipeline.main()

            self.assertEqual(rc, 0)
            self.assertEqual(len(commands), 3)
            formatter_call = commands[-1]
            self.assertIn("-PythonExe", formatter_call)
            self.assertEqual(formatter_call[formatter_call.index("-PythonExe") + 1], sys.executable)


if __name__ == "__main__":
    unittest.main()

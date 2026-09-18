import tempfile
import unittest
from pathlib import Path

from audit_log import AuditLog


class AuditLogTests(unittest.TestCase):
    def test_saves_markdown_with_metadata_and_changes(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            dst = str(Path(tmp_dir) / "out.docx")
            audit = AuditLog(
                src="input.docx",
                dst=dst,
                formatting={"font_name": "Garamond", "font_size": 12},
                quote_language="serbian",
                output_format=".docx",
                options={"spacing": True},
            )
            audit.bump("blank_paragraphs_removed", 2)
            audit.record_change("text_normalization", 'a ,', 'a,', context="run 1")
            audit_path = audit.save()

            text = audit_path.read_text(encoding="utf-8")
            self.assertIn("# Audit Log", text)
            self.assertIn("input.docx", text)
            self.assertIn("blank_paragraphs_removed", text)
            self.assertIn("text_normalization", text)
            self.assertIn("a ,", text)
            self.assertIn("a,", text)
            self.assertIn("font_name", text)
            self.assertIn("Garamond", text)


if __name__ == "__main__":
    unittest.main()

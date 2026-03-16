import tempfile
import unittest
import zipfile
from pathlib import Path

from docx_notes import preserve_notes


CONTENT_TYPES_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

DOCUMENT_RELS_BASE = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"""

DOCUMENT_RELS_WITH_NOTES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>
"""


class PreserveNotesTests(unittest.TestCase):
    def test_preserves_notes_and_comments_parts_and_relationships(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            source = Path(tmp_dir) / "source.docx"
            target = Path(tmp_dir) / "target.docx"

            with zipfile.ZipFile(source, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
                zf.writestr("word/document.xml", "<w:document/>")
                zf.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS_WITH_NOTES)
                zf.writestr("word/footnotes.xml", "<w:footnotes/>")
                zf.writestr("word/endnotes.xml", "<w:endnotes/>")
                zf.writestr("word/comments.xml", "<w:comments/>")

            with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
                zf.writestr("word/document.xml", "<w:document/>")
                zf.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS_BASE)

            preserve_notes(source, target)

            with zipfile.ZipFile(target, "r") as zf:
                names = set(zf.namelist())
                self.assertIn("word/footnotes.xml", names)
                self.assertIn("word/endnotes.xml", names)
                self.assertIn("word/comments.xml", names)

                rels = zf.read("word/_rels/document.xml.rels").decode("utf-8")
                self.assertIn("relationships/footnotes", rels)
                self.assertIn("relationships/endnotes", rels)
                self.assertIn("relationships/comments", rels)

                content_types = zf.read("[Content_Types].xml").decode("utf-8")
                self.assertIn("/word/footnotes.xml", content_types)
                self.assertIn("/word/endnotes.xml", content_types)
                self.assertIn("/word/comments.xml", content_types)


if __name__ == "__main__":
    unittest.main()

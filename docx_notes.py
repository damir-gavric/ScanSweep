import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


CONTENT_TYPES_PATH = "[Content_Types].xml"
DOCUMENT_RELS_PATH = "word/_rels/document.xml.rels"
PRESERVED_PARTS = {
    "footnotes": {
        "part": "word/footnotes.xml",
        "rels": "word/_rels/footnotes.xml.rels",
        "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        "relationship_type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    },
    "endnotes": {
        "part": "word/endnotes.xml",
        "rels": "word/_rels/endnotes.xml.rels",
        "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
        "relationship_type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
    },
    "comments": {
        "part": "word/comments.xml",
        "rels": "word/_rels/comments.xml.rels",
        "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        "relationship_type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    },
}

PKG_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", PKG_CT_NS)
ET.register_namespace("", PKG_REL_NS)


def preserve_notes(source_docx, target_docx):
    source_docx = Path(source_docx)
    target_docx = Path(target_docx)
    if not source_docx.exists() or not target_docx.exists():
        return []

    with zipfile.ZipFile(source_docx, "r") as source_zip:
        source_names = set(source_zip.namelist())
        required = [cfg for cfg in PRESERVED_PARTS.values() if cfg["part"] in source_names]
        if not required:
            return []

        with zipfile.ZipFile(target_docx, "r") as target_zip:
            target_entries = {name: target_zip.read(name) for name in target_zip.namelist()}

        for cfg in required:
            target_entries[cfg["part"]] = source_zip.read(cfg["part"])
            if cfg["rels"] in source_names:
                target_entries[cfg["rels"]] = source_zip.read(cfg["rels"])

        target_entries[CONTENT_TYPES_PATH] = _merge_content_types(
            target_entries[CONTENT_TYPES_PATH],
            required,
        )

        source_rels_xml = source_zip.read(DOCUMENT_RELS_PATH)
        target_entries[DOCUMENT_RELS_PATH] = _merge_document_relationships(
            target_entries[DOCUMENT_RELS_PATH],
            source_rels_xml,
            required,
        )

    _rewrite_zip(target_docx, target_entries)
    return [name for name, cfg in PRESERVED_PARTS.items() if cfg in required]


def _merge_content_types(target_xml, required_parts):
    root = ET.fromstring(target_xml)
    existing_parts = {
        override.get("PartName")
        for override in root.findall(f"{{{PKG_CT_NS}}}Override")
    }

    for cfg in required_parts:
        part_name = f"/{cfg['part']}"
        if part_name in existing_parts:
            continue
        ET.SubElement(
            root,
            f"{{{PKG_CT_NS}}}Override",
            PartName=part_name,
            ContentType=cfg["content_type"],
        )

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _merge_document_relationships(target_xml, source_xml, required_parts):
    target_root = ET.fromstring(target_xml)
    source_root = ET.fromstring(source_xml)

    needed_types = {cfg["relationship_type"] for cfg in required_parts}
    existing_types = {
        rel.get("Type")
        for rel in target_root.findall(f"{{{PKG_REL_NS}}}Relationship")
    }

    for rel in source_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        if rel.get("Type") not in needed_types:
            continue
        if rel.get("Type") in existing_types:
            continue
        target_root.append(rel)

    return ET.tostring(target_root, encoding="utf-8", xml_declaration=True)


def _rewrite_zip(target_path, entries):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
        temp_path = Path(tmp_file.name)

    try:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as new_zip:
            for name, data in entries.items():
                new_zip.writestr(name, data)
        shutil.move(str(temp_path), str(target_path))
    finally:
        if temp_path.exists():
            temp_path.unlink(missing_ok=True)

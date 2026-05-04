import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from docx_heading_cleaner import DocxCleanerError, W, clean_docx_all


ET.register_namespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

WORD_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
"""


def make_docx(path: Path, styles_xml: str, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", CONTENT_TYPES)
        docx.writestr("_rels/.rels", RELS)
        docx.writestr("word/_rels/document.xml.rels", WORD_RELS)
        docx.writestr("word/styles.xml", styles_xml)
        docx.writestr("word/document.xml", document_xml)


def read_xml(docx_path: Path, part: str) -> ET.Element:
    with zipfile.ZipFile(docx_path) as docx:
        return ET.fromstring(docx.read(part))


def styles_xml(extra_styles: str = "") -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="240"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  {extra_styles}
</w:styles>
"""


def document_xml(paragraphs: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{paragraphs}</w:body>
</w:document>
"""


def paragraph(text: str, style_id: str | None = None, outline: str | None = None) -> str:
    p_style = f'<w:pStyle w:val="{style_id}"/>' if style_id else ""
    outline_el = f'<w:outlineLvl w:val="{outline}"/>' if outline is not None else ""
    p_pr = f"<w:pPr>{p_style}{outline_el}</w:pPr>" if p_style or outline_el else ""
    return f"<w:p>{p_pr}<w:r><w:t>{text}</w:t></w:r></w:p>"


class DocxHeadingCleanerTests(unittest.TestCase):
    def test_heading_style_is_replaced_with_no_nav_clone(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            src = Path(temp_dir) / "input.docx"
            out = Path(temp_dir) / "output.docx"
            make_docx(src, styles_xml(), document_xml(paragraph("Title", "Heading1")))

            report = clean_docx_all(src, out)

            self.assertEqual(report.changed_paragraphs, 1)
            self.assertEqual(report.cloned_styles, 1)
            doc = read_xml(out, "word/document.xml")
            p_style = doc.find(f".//{W}pStyle")
            self.assertEqual(p_style.get(f"{W}val"), "Heading1NoNav")

            styles = read_xml(out, "word/styles.xml")
            original_outline = styles.find(f".//{W}style[@{W}styleId='Heading1']/{W}pPr/{W}outlineLvl")
            clone_outline = styles.find(f".//{W}style[@{W}styleId='Heading1NoNav']/{W}pPr/{W}outlineLvl")
            clone_name = styles.find(f".//{W}style[@{W}styleId='Heading1NoNav']/{W}name")
            self.assertIsNotNone(original_outline)
            self.assertIsNone(clone_outline)
            self.assertEqual(clone_name.get(f"{W}val"), "Heading 1 (No Nav)")

    def test_direct_outline_level_is_removed_without_changing_style(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            src = Path(temp_dir) / "input.docx"
            out = Path(temp_dir) / "output.docx"
            make_docx(src, styles_xml(), document_xml(paragraph("Body", "Normal", "1")))

            report = clean_docx_all(src, out)

            self.assertEqual(report.changed_paragraphs, 1)
            self.assertEqual(report.direct_outline_removed, 1)
            doc = read_xml(out, "word/document.xml")
            self.assertIsNone(doc.find(f".//{W}outlineLvl"))
            self.assertEqual(doc.find(f".//{W}pStyle").get(f"{W}val"), "Normal")

    def test_style_inheriting_heading_points_to_no_nav_base_clone(self):
        extra = """
  <w:style w:type="paragraph" w:styleId="CustomLead">
    <w:name w:val="Custom Lead"/>
    <w:basedOn w:val="Heading1"/>
    <w:rPr><w:i/></w:rPr>
  </w:style>
"""
        with tempfile.TemporaryDirectory() as temp_dir:
            src = Path(temp_dir) / "input.docx"
            out = Path(temp_dir) / "output.docx"
            make_docx(src, styles_xml(extra), document_xml(paragraph("Long body", "CustomLead")))

            report = clean_docx_all(src, out)

            self.assertEqual(report.changed_paragraphs, 1)
            self.assertEqual(report.cloned_styles, 2)
            doc = read_xml(out, "word/document.xml")
            self.assertEqual(doc.find(f".//{W}pStyle").get(f"{W}val"), "CustomLeadNoNav")

            styles = read_xml(out, "word/styles.xml")
            custom_base = styles.find(f".//{W}style[@{W}styleId='CustomLeadNoNav']/{W}basedOn")
            heading_clone_outline = styles.find(f".//{W}style[@{W}styleId='Heading1NoNav']/{W}pPr/{W}outlineLvl")
            self.assertEqual(custom_base.get(f"{W}val"), "Heading1NoNav")
            self.assertIsNone(heading_clone_outline)

    def test_second_run_is_stable(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            src = Path(temp_dir) / "input.docx"
            first = Path(temp_dir) / "first.docx"
            second = Path(temp_dir) / "second.docx"
            make_docx(src, styles_xml(), document_xml(paragraph("Title", "Heading1")))

            clean_docx_all(src, first)
            report = clean_docx_all(first, second)

            self.assertEqual(report.changed_paragraphs, 0)
            self.assertEqual(report.cloned_styles, 0)
            doc = read_xml(second, "word/document.xml")
            self.assertEqual(doc.find(f".//{W}pStyle").get(f"{W}val"), "Heading1NoNav")

    def test_existing_output_requires_overwrite(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            src = Path(temp_dir) / "input.docx"
            out = Path(temp_dir) / "output.docx"
            make_docx(src, styles_xml(), document_xml(paragraph("Title", "Heading1")))
            out.write_text("existing", encoding="utf-8")

            with self.assertRaises(DocxCleanerError):
                clean_docx_all(src, out)

    def test_existing_output_can_be_overwritten(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            src = Path(temp_dir) / "input.docx"
            out = Path(temp_dir) / "output.docx"
            make_docx(src, styles_xml(), document_xml(paragraph("Title", "Heading1")))
            out.write_text("existing", encoding="utf-8")

            report = clean_docx_all(src, out, overwrite=True)

            self.assertEqual(report.changed_paragraphs, 1)
            doc = read_xml(out, "word/document.xml")
            self.assertEqual(doc.find(f".//{W}pStyle").get(f"{W}val"), "Heading1NoNav")


if __name__ == "__main__":
    unittest.main()

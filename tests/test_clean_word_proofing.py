import hashlib
import importlib.util
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from lxml import etree

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "clean_word_proofing.py"
spec = importlib.util.spec_from_file_location("clean_word_proofing", SCRIPT)
cleaner = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cleaner)
W = cleaner.W_NS
NS = cleaner.NS


def make_docx(path):
    content = f'''<w:document xmlns:w="{W}"><w:body><w:p>
      <w:proofErr w:type="spellStart"/><w:r><w:t>公司名 / ProjectXYZ</w:t><w:tab/><w:br/></w:r>
      <w:proofErr w:type="spellEnd"/>
      <w:r><w:rPr><w:b/><w:noProof w:val="false"/></w:rPr><w:t>保留加粗</w:t></w:r>
      <w:ins w:id="1" w:author="Test"><w:r><w:t>修订新增</w:t></w:r></w:ins>
      <w:del w:id="2" w:author="Test"><w:r><w:delText>修订删除</w:delText></w:r></w:del>
      </w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格内容</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
      <w:sectPr/></w:body></w:document>'''
    settings = f'''<w:settings xmlns:w="{W}"><w:hideSpellingErrors w:val="0"/>
      <w:hideGrammaticalErrors w:val="false"/><w:proofState w:spelling="dirty" w:grammar="dirty"/></w:settings>'''
    styles = f'''<w:styles xmlns:w="{W}"><w:docDefaults/>
      <w:style w:type="character" w:styleId="Custom"><w:rPr><w:i/><w:noProof w:val="off"/></w:rPr></w:style></w:styles>'''
    parts = {"word/document.xml": content.encode(), "word/settings.xml": settings.encode(), "word/styles.xml": styles.encode()}
    for name, tag in [("header1", "hdr"), ("footer1", "ftr"), ("footnotes", "footnotes"), ("endnotes", "endnotes"), ("comments", "comments")]:
        parts[f"word/{name}.xml"] = f'<w:{tag} xmlns:w="{W}"><w:p><w:proofErr w:type="gramStart"/><w:r><w:t>{name} 原文</w:t></w:r></w:p></w:{tag}>'.encode()
    parts["[Content_Types].xml"] = b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'
    parts["word/media/image1.png"] = b"\x89PNG\r\n\x1a\nTEST-UNCHANGED"
    parts["word/_rels/document.xml.rels"] = b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    with zipfile.ZipFile(path, "w") as zf:
        for name, value in parts.items(): zf.writestr(name, value)
    return parts


class CleanerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / "source.docx"
        self.output = self.root / "clean.docx"
        self.parts = make_docx(self.source)

    def run_cli(self, *args):
        return subprocess.run([sys.executable, str(SCRIPT), *map(str,args)], capture_output=True, text=True)

    def test_markers_and_plain_runs_in_all_parts(self):
        cleaner.clean_docx(self.source, self.output, verify_text=True)
        with zipfile.ZipFile(self.output) as zf:
            for name in self.parts:
                if name.startswith("word/") and name.endswith(".xml") and name != "word/settings.xml":
                    root = etree.fromstring(zf.read(name))
                    self.assertEqual(root.xpath("count(.//w:proofErr)", namespaces=NS), 0)
                    for run in root.iter(cleaner.qn("r")):
                        self.assertIsNotNone(run.find("w:rPr/w:noProof", namespaces=NS), name)

    def test_false_properties_are_enabled(self):
        cleaner.clean_docx(self.source, self.output)
        with zipfile.ZipFile(self.output) as zf:
            for name in ["word/document.xml", "word/styles.xml"]:
                root = etree.fromstring(zf.read(name))
                for node in root.iter(cleaner.qn("noProof")):
                    self.assertNotIn(node.get(cleaner.qn("val"),"true"), {"0","false","off"})
            root = etree.fromstring(zf.read("word/settings.xml"))
            for prop in ["hideSpellingErrors", "hideGrammaticalErrors"]:
                self.assertEqual(root.find(f"w:{prop}",NS).get(cleaner.qn("val")), "true")
            self.assertEqual(root.find("w:proofState",NS).get(cleaner.qn("grammar")), "clean")

    def test_original_text_media_reviews_and_formatting_preserved(self):
        original_hash = hashlib.sha256(self.source.read_bytes()).hexdigest()
        cleaner.clean_docx(self.source, self.output, verify_text=True)
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), original_hash)
        self.assertEqual(cleaner.visible_text(self.source), cleaner.visible_text(self.output))
        with zipfile.ZipFile(self.output) as zf:
            self.assertEqual(set(zf.namelist()), set(self.parts))
            self.assertEqual(zf.read("word/media/image1.png"), self.parts["word/media/image1.png"])
            self.assertEqual(zf.read("word/_rels/document.xml.rels"), self.parts["word/_rels/document.xml.rels"])
            root = etree.fromstring(zf.read("word/document.xml"))
            for element in ["b", "ins", "del", "tbl"]:
                self.assertEqual(root.xpath(f"count(.//w:{element})", namespaces=NS),1)
            self.assertEqual(root.find(".//w:delText",NS).text,"修订删除")
            self.assertEqual(zf.read("word/comments.xml").decode().count("comments 原文"),1)

    def test_default_style_and_second_run_idempotence(self):
        cleaner.clean_docx(self.source, self.output)
        again = self.root / "again.docx"
        stats = cleaner.clean_docx(self.output, again, verify_text=True)
        self.assertEqual(sum(stats.values()),0)
        with zipfile.ZipFile(self.output) as a, zipfile.ZipFile(again) as b:
            for name in a.namelist(): self.assertEqual(a.read(name),b.read(name))
            root = etree.fromstring(a.read("word/styles.xml"))
            self.assertIsNotNone(root.find("w:docDefaults/w:rPrDefault/w:rPr/w:noProof",NS))

    def test_cli_default_output_and_explicit_in_place(self):
        result = self.run_cli(self.source,"--verify-text")
        self.assertEqual(result.returncode,0,result.stderr)
        self.assertIn("visible_text_equal: True",result.stdout)
        self.assertTrue(self.source.with_name("source_no-proofing-lines.docx").exists())
        result = self.run_cli(self.source,"--in-place","--verify-text")
        self.assertEqual(result.returncode,0,result.stderr)
        with zipfile.ZipFile(self.source) as zf:
            self.assertNotIn(b"proofErr",zf.read("word/document.xml"))

    def test_unintended_overwrite_is_rejected(self):
        before=self.source.read_bytes()
        result=self.run_cli(self.source,self.source)
        self.assertNotEqual(result.returncode,0)
        self.assertEqual(self.source.read_bytes(),before)
        result=self.run_cli(self.source,self.output,"--in-place")
        self.assertNotEqual(result.returncode,0)
        self.assertEqual(self.source.read_bytes(),before)

    def test_verification_failure_preserves_both_files(self):
        self.output.write_bytes(b"previous destination")
        original=self.source.read_bytes()
        with patch.object(cleaner,"visible_text",side_effect=["before","after"]):
            with self.assertRaisesRegex(ValueError,"not replaced"):
                cleaner.clean_docx(self.source,self.output,verify_text=True)
        self.assertEqual(self.output.read_bytes(),b"previous destination")
        self.assertEqual(self.source.read_bytes(),original)
        with patch.object(cleaner,"visible_text",side_effect=["before","after"]):
            with self.assertRaises(ValueError):
                cleaner.clean_docx(self.source,self.source,verify_text=True)
        self.assertEqual(self.source.read_bytes(),original)

    def test_malformed_xml_does_not_replace_destination(self):
        broken=self.root/"broken.docx"
        with zipfile.ZipFile(broken,"w") as zf: zf.writestr("word/document.xml",b"<broken")
        self.output.write_bytes(b"preserve me")
        with self.assertRaises(etree.XMLSyntaxError): cleaner.clean_docx(broken,self.output)
        self.assertEqual(self.output.read_bytes(),b"preserve me")

    def test_non_docx_output_rejected(self):
        result=self.run_cli(self.source,self.root/"bad.pdf")
        self.assertNotEqual(result.returncode,0)
        self.assertFalse((self.root/"bad.pdf").exists())


if __name__ == "__main__": unittest.main()

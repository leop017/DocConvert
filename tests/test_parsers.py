"""Tests for ``docconvert.parsers``."""

import unittest

from docconvert.parsers import (
    BaseParser,
    Document,
    HtmlParser,
    MarkdownParser,
    PlainTextParser,
    get_parser,
)


class TestDocumentModel(unittest.TestCase):

    def test_defaults_are_empty(self):
        doc = Document()
        self.assertEqual(doc.text, "")
        self.assertEqual(doc.elements, [])
        self.assertEqual(doc.metadata, {})


class TestGetParserFactory(unittest.TestCase):

    def test_markdown_returns_markdown_parser(self):
        self.assertIsInstance(get_parser("markdown"), MarkdownParser)

    def test_md_alias(self):
        self.assertIsInstance(get_parser("md"), MarkdownParser)

    def test_html_returns_html_parser(self):
        self.assertIsInstance(get_parser("html"), HtmlParser)

    def test_text_returns_plain_parser(self):
        self.assertIsInstance(get_parser("text"), PlainTextParser)

    def test_unknown_format_raises(self):
        with self.assertRaises(ValueError):
            get_parser("xml")


class TestMarkdownParser(unittest.TestCase):

    def setUp(self):
        self.parser = MarkdownParser()

    def test_inherits_base(self):
        self.assertIsInstance(self.parser, BaseParser)

    def test_empty_input_returns_empty_doc(self):
        doc = self.parser.parse("")
        self.assertEqual(doc.text, "")
        self.assertEqual(doc.elements, [])

    def test_heading_is_extracted(self):
        doc = self.parser.parse("# Title\n\nBody")
        types = [e.element_type for e in doc.elements]
        self.assertIn("heading", types)

    def test_code_fence_captured(self):
        md = "```python\nprint('hi')\n```"
        doc = self.parser.parse(md)
        code_elems = [e for e in doc.elements if e.element_type == "code"]
        self.assertEqual(len(code_elems), 1)
        self.assertIn("print", code_elems[0].text)
        self.assertEqual(code_elems[0].metadata["language"], "python")

    def test_front_matter_becomes_metadata(self):
        md = "---\ntitle: Doc\nauthor: leop\n---\n# Body"
        doc = self.parser.parse(md)
        self.assertEqual(doc.metadata.get("title"), "Doc")
        self.assertEqual(doc.metadata.get("author"), "leop")

    def test_table_extracted(self):
        md = "| A | B |\n|---|---|\n| 1 | 2 |"
        doc = self.parser.parse(md)
        tables = [e for e in doc.elements if e.element_type == "table"]
        self.assertEqual(len(tables), 1)

    def test_list_extracted(self):
        md = "- one\n- two\n- three"
        doc = self.parser.parse(md)
        self.assertEqual(doc.elements[0].element_type, "list")


class TestHtmlParser(unittest.TestCase):

    def setUp(self):
        self.parser = HtmlParser()

    def test_strips_script_tags(self):
        html = "<p>hi</p><script>evil()</script>"
        doc = self.parser.parse(html)
        self.assertNotIn("evil", doc.text)

    def test_extracts_title_into_metadata(self):
        html = "<html><head><title>My Doc</title></head><body><p>body</p></body></html>"
        doc = self.parser.parse(html)
        self.assertEqual(doc.metadata.get("title"), "My Doc")

    def test_heading_level_metadata(self):
        html = "<h2>Section</h2>"
        doc = self.parser.parse(html)
        h = next(e for e in doc.elements if e.element_type == "heading")
        self.assertEqual(h.metadata.get("level"), 2)

    def test_table_rows_metadata(self):
        html = "<table><tr><td>A</td><td>B</td></tr></table>"
        doc = self.parser.parse(html)
        tables = [e for e in doc.elements if e.element_type == "table"]
        self.assertEqual(len(tables), 1)
        self.assertTrue(any("A" in r for r in tables[0].metadata["rows"]))


class TestPlainTextParser(unittest.TestCase):

    def setUp(self):
        self.parser = PlainTextParser()

    def test_collapses_whitespace(self):
        doc = self.parser.parse("a   b\t\tc")
        self.assertEqual(doc.text, "a b c")

    def test_collapses_blank_line_runs(self):
        doc = self.parser.parse("a\n\n\n\nb")
        self.assertEqual(doc.text, "a\n\nb")

    def test_empty_string(self):
        doc = self.parser.parse("")
        self.assertEqual(doc.text, "")
        self.assertEqual(doc.elements, [])


if __name__ == "__main__":
    unittest.main()

"""Tests for ``docconvert.chunkers``."""

import unittest

from docconvert.chunkers import (
    BaseChunker,
    FixedSizeChunker,
    MarkdownChunker,
    SentenceChunker,
    get_chunker,
)
from docconvert.parsers import MarkdownParser


class TestGetChunkerFactory(unittest.TestCase):

    def test_fixed_returns_fixed_chunker(self):
        self.assertIsInstance(get_chunker("fixed"), FixedSizeChunker)

    def test_sentence_returns_sentence_chunker(self):
        self.assertIsInstance(get_chunker("sentence"), SentenceChunker)

    def test_markdown_returns_markdown_chunker(self):
        self.assertIsInstance(get_chunker("markdown"), MarkdownChunker)

    def test_unknown_strategy_raises(self):
        with self.assertRaises(ValueError):
            get_chunker("magic")

    def test_kwargs_forwarded(self):
        chunker = get_chunker("fixed", chunk_size=128, chunk_overlap=16)
        self.assertEqual(chunker.chunk_size, 128)
        self.assertEqual(chunker.chunk_overlap, 16)


class TestFixedSizeChunker(unittest.TestCase):

    def test_invalid_chunk_size(self):
        with self.assertRaises(ValueError):
            FixedSizeChunker(chunk_size=0)
        with self.assertRaises(ValueError):
            FixedSizeChunker(chunk_size=-1)

    def test_invalid_overlap(self):
        with self.assertRaises(ValueError):
            FixedSizeChunker(chunk_size=10, chunk_overlap=-1)
        with self.assertRaises(ValueError):
            FixedSizeChunker(chunk_size=10, chunk_overlap=10)

    def test_short_text_no_split(self):
        chunker = FixedSizeChunker(chunk_size=100, chunk_overlap=10)
        chunks = chunker.chunk("hello world")
        self.assertEqual(len(chunks), 1)
        self.assertEqual(chunks[0].text, "hello world")
        self.assertEqual(chunks[0].metadata["chunk_index"], 0)
        self.assertEqual(chunks[0].metadata["chunk_count"], 1)

    def test_long_text_creates_overlapping_chunks(self):
        chunker = FixedSizeChunker(chunk_size=20, chunk_overlap=5)
        text = "a" * 100
        chunks = chunker.chunk(text)
        self.assertGreater(len(chunks), 1)
        for i, c in enumerate(chunks):
            self.assertEqual(c.metadata["chunk_index"], i)
            self.assertEqual(c.metadata["chunk_count"], len(chunks))
        for c in chunks[:-1]:
            self.assertEqual(len(c.text), 20)
        self.assertLessEqual(len(chunks[-1].text), 20)

    def test_empty_text_returns_no_chunks(self):
        chunker = FixedSizeChunker()
        self.assertEqual(chunker.chunk(""), [])
        # Pure-whitespace is still a single (small) chunk; the chunker
        # does not normalize input. PlainTextParser handles normalization.
        chunks = chunker.chunk("   ")
        self.assertEqual(len(chunks), 1)
        self.assertEqual(chunks[0].text, "   ")

    def test_metadata_propagates(self):
        chunker = FixedSizeChunker()
        from docconvert.parsers.models import Document
        doc = Document(text="abcdef", metadata={"origin": "unit-test"})
        chunks = chunker.chunk(doc)
        self.assertEqual(chunks[0].metadata["origin"], "unit-test")


class TestSentenceChunker(unittest.TestCase):

    def test_breaks_on_period(self):
        chunker = SentenceChunker(chunk_size=100, chunk_overlap=10)
        text = "First sentence. Second sentence. Third sentence."
        chunks = chunker.chunk(text)
        joined = " ".join(c.text for c in chunks)
        self.assertIn("First sentence", joined)
        self.assertIn("Third sentence", joined)

    def test_chinese_punctuation(self):
        chunker = SentenceChunker(chunk_size=200, chunk_overlap=20)
        text = "第一句。第二句！第三句？第四句。"
        chunks = chunker.chunk(text)
        joined = " ".join(c.text for c in chunks)
        self.assertIn("第一句", joined)
        self.assertIn("第四句", joined)

    def test_empty_text(self):
        chunker = SentenceChunker()
        self.assertEqual(chunker.chunk(""), [])

    def test_respects_chunk_size_upper_bound(self):
        chunker = SentenceChunker(chunk_size=40, chunk_overlap=5)
        text = ". ".join(["sentence"] * 30)
        chunks = chunker.chunk(text)
        for c in chunks:
            self.assertLessEqual(len(c.text), 80)


class TestMarkdownChunker(unittest.TestCase):

    def setUp(self):
        self.parser = MarkdownParser()

    def test_one_chunk_per_section(self):
        md = "# Title\n\nIntro.\n\n## Section A\n\nBody A.\n\n## Section B\n\nBody B."
        doc = self.parser.parse(md)
        chunks = MarkdownChunker().chunk(doc)
        self.assertGreaterEqual(len(chunks), 2)
        headers = [c.metadata["header"] for c in chunks if "header" in c.metadata]
        self.assertIn("Section A", headers)
        self.assertIn("Section B", headers)

    def test_empty_document(self):
        from docconvert.parsers.models import Document
        self.assertEqual(MarkdownChunker().chunk(Document()), [])

    def test_max_chunk_size_splits_long(self):
        from docconvert.parsers.models import Document
        md = "# H\n\n" + ("x" * 5000)
        doc = self.parser.parse(md)
        chunks = MarkdownChunker().chunk(doc, max_chunk_size=200)
        self.assertGreater(len(chunks), 1)


if __name__ == "__main__":
    unittest.main()

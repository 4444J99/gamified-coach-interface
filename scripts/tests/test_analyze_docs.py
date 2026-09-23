"""
Unit tests for DocumentAnalyzer in scripts/analyze_docs.py
"""

import os
import sys
import tempfile
import unittest
from io import StringIO
from pathlib import Path

from docx import Document

# Ensure scripts directory is in sys.path
sys.path.insert(0, str(Path(__file__).parent.parent))

from analyze_docs import DocumentAnalyzer  # noqa: E402


class TestDocumentAnalyzer(unittest.TestCase):
    """Test suite for DocumentAnalyzer class."""

    def setUp(self):
        """Set up temporary directory and sample documents for testing."""
        self.temp_dir = tempfile.TemporaryDirectory()
        self.base_path = Path(self.temp_dir.name)

        # Create a valid test docx file
        self.valid_docx_path = self.base_path / "valid_test.docx"
        doc = Document()
        doc.add_heading("Fitness Blueprint", level=1)
        doc.add_paragraph("This is a fitness training and workout exercise plan.")
        doc.add_paragraph("It provides coaching strategies and business models.")
        doc.save(self.valid_docx_path)

    def tearDown(self):
        """Clean up temporary directory."""
        self.temp_dir.cleanup()

    def test_ingest_and_digest_valid_docs(self):
        """Test ingesting and digesting valid docx files."""
        analyzer = DocumentAnalyzer(base_dir=str(self.base_path))
        ingested = analyzer.ingest_docs()

        self.assertIn("valid_test.docx", ingested)
        self.assertIn("valid_test.docx", analyzer.documents)
        self.assertEqual(analyzer.documents["valid_test.docx"]["num_paragraphs"], 3)

        analysis = analyzer.digest_docs()
        self.assertEqual(analysis["total_documents"], 1)
        self.assertIn("valid_test.docx", analysis["documents"])
        doc_analysis = analysis["documents"]["valid_test.docx"]
        self.assertGreater(doc_analysis["word_count"], 0)
        self.assertTrue(len(doc_analysis["key_topics"]) > 0)

    def test_oversized_file_rejection(self):
        """Test that files exceeding max_file_size_bytes are rejected."""
        # Set max_file_size_bytes to a small value (10 bytes)
        analyzer = DocumentAnalyzer(base_dir=str(self.base_path), max_file_size_bytes=10)

        captured_output = StringIO()
        sys.stdout = captured_output
        try:
            ingested = analyzer.ingest_docs()
        finally:
            sys.stdout = sys.__stdout__

        output = captured_output.getvalue()
        self.assertNotIn("valid_test.docx", ingested)
        self.assertIn("File size exceeds limit", output)

    def test_path_traversal_prevention(self):
        """Test that files located outside base_dir are rejected."""
        sub_dir = self.base_path / "sub"
        sub_dir.mkdir()
        analyzer = DocumentAnalyzer(base_dir=str(sub_dir))

        # Pass a document path outside base_dir manually to test logic
        captured_output = StringIO()
        sys.stdout = captured_output
        try:
            ingested = analyzer.ingest_docs()
        finally:
            sys.stdout = sys.__stdout__

        output = captured_output.getvalue()
        self.assertEqual(len(ingested), 0)

    def test_corrupted_file_handling(self):
        """Test that corrupted docx files are handled gracefully without crashing or leaking sensitive info."""
        corrupt_path = self.base_path / "corrupt.docx"
        with open(corrupt_path, "wb") as f:
            f.write(b"NOT_A_REAL_DOCX_FILE_DATA_1234567890")

        analyzer = DocumentAnalyzer(base_dir=str(self.base_path))

        captured_output = StringIO()
        sys.stdout = captured_output
        try:
            ingested = analyzer.ingest_docs()
        finally:
            sys.stdout = sys.__stdout__

        output = captured_output.getvalue()
        self.assertNotIn("corrupt.docx", ingested)
        self.assertIn("Failed to ingest corrupt.docx", output)
        # Ensure raw stack trace is not dumped to user output
        self.assertNotIn("Traceback", output)

    def test_key_topics_extraction(self):
        """Test topic extraction algorithm."""
        analyzer = DocumentAnalyzer(base_dir=str(self.base_path))
        paragraphs = [
            "Fitness fitness fitness training workout exercise.",
            "Gamified gamified battle quest gaming fitness.",
        ]
        topics = analyzer._extract_key_topics(paragraphs)
        self.assertIsInstance(topics, list)
        self.assertIn("fitnes", topics)  # Stemmed fitness

    def test_summary_creation_truncation(self):
        """Test summary creation and truncation behavior."""
        analyzer = DocumentAnalyzer(base_dir=str(self.base_path))
        long_paragraph = "word " * 100
        summary = analyzer._create_summary([long_paragraph])
        self.assertTrue(summary.endswith("..."))
        self.assertLessEqual(len(summary), 300)


if __name__ == "__main__":
    unittest.main()

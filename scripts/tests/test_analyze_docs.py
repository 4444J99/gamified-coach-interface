"""
Unit tests for DocumentAnalyzer in scripts/analyze_docs.py
"""

import os
import sys
import tempfile
import unittest
from io import StringIO
from pathlib import Path
from unittest.mock import patch

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

    def test_symlink_outside_base_dir_is_rejected(self):
        """A .docx symlink must not escape the configured source root."""
        with tempfile.TemporaryDirectory() as outside_temp_dir:
            outside_path = Path(outside_temp_dir) / "outside.docx"
            outside_doc = Document()
            outside_doc.add_paragraph("outside source root")
            outside_doc.save(outside_path)

            linked_path = self.base_path / "linked-outside.docx"
            try:
                os.symlink(outside_path, linked_path)
            except (OSError, NotImplementedError) as exc:
                self.skipTest(f"symlinks unavailable in this environment: {exc}")

            analyzer = DocumentAnalyzer(base_dir=str(self.base_path))
            captured_output = StringIO()
            sys.stdout = captured_output
            try:
                ingested = analyzer.ingest_docs()
            finally:
                sys.stdout = sys.__stdout__

            output = captured_output.getvalue()
            self.assertNotIn("linked-outside.docx", ingested)
            self.assertIn("linked-outside.docx", output)
            self.assertIn("Access denied (path outside base directory)", output)

    def test_non_regular_docx_path_is_rejected(self):
        """A directory with a .docx suffix must not be parsed as a document."""
        fake_docx_dir = self.base_path / "directory.docx"
        fake_docx_dir.mkdir()

        analyzer = DocumentAnalyzer(base_dir=str(self.base_path))
        captured_output = StringIO()
        sys.stdout = captured_output
        try:
            ingested = analyzer.ingest_docs()
        finally:
            sys.stdout = sys.__stdout__

        output = captured_output.getvalue()
        self.assertNotIn("directory.docx", ingested)
        self.assertIn("Path is not a regular file", output)

    def test_corrupted_file_handling(self):
        """Corrupted docx files fail safely without a traceback."""
        corrupt_path = self.base_path / "corrupt.docx"
        with open(corrupt_path, "wb") as corrupt_file:
            corrupt_file.write(b"NOT_A_REAL_DOCX_FILE_DATA_1234567890")

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
        self.assertNotIn("Traceback", output)

    def test_unexpected_parser_error_does_not_leak_exception_message(self):
        """Unexpected parser failures expose the class, not raw sensitive details."""
        sensitive_detail = "/private/customer/path/token-like-detail"
        analyzer = DocumentAnalyzer(base_dir=str(self.base_path))

        captured_output = StringIO()
        sys.stdout = captured_output
        try:
            with patch(
                "analyze_docs.Document",
                side_effect=RuntimeError(sensitive_detail),
            ):
                ingested = analyzer.ingest_docs()
        finally:
            sys.stdout = sys.__stdout__

        output = captured_output.getvalue()
        self.assertNotIn("valid_test.docx", ingested)
        self.assertIn("Ingestion error (RuntimeError)", output)
        self.assertNotIn(sensitive_detail, output)

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

#!/usr/bin/env python3
"""
Document Analysis System
Ingests, digests, and suggests paths based on Word documents in the repository.
"""

import sys
from pathlib import Path
from typing import Any, Dict, List

from docx import Document
from docx.opc.exceptions import OpcError, PackageNotFoundError

# Default maximum file size limit for ingested document files (10 MB)
DEFAULT_MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024


class DocumentAnalyzer:
    """Handles document ingestion, digestion, and path suggestion."""

    def __init__(self, base_dir: str = ".", max_file_size_bytes: int = DEFAULT_MAX_FILE_SIZE_BYTES):
        self.base_dir = Path(base_dir).resolve()
        if not self.base_dir.exists() or not self.base_dir.is_dir():
            print(
                f"Error: The base directory '{self.base_dir}' does not exist or is not a directory."
            )
            sys.exit(1)
        self.max_file_size_bytes = max_file_size_bytes
        self.documents: Dict[str, Any] = {}

    def ingest_docs(self) -> List[str]:
        """
        Ingest all .docx files from the repository base directory.
        Returns a list of ingested document names.
        """
        print("=" * 80)
        print("INGESTING DOCUMENTS")
        print("=" * 80)

        found_docx_paths = [
            file_path
            for file_path in self.base_dir.glob("**/*.docx")
            if not any(
                path_part.startswith(".") or path_part.startswith("~$")
                for path_part in file_path.parts
            )
        ]
        ingested_document_names = []

        for document_path in found_docx_paths:
            # Prevent directory traversal attacks by ensuring path containment
            resolved_document_path = document_path.resolve()
            if not (
                resolved_document_path == self.base_dir
                or self.base_dir in resolved_document_path.parents
            ):
                print(
                    f"✗ Failed to ingest {document_path.name}: Access denied (path outside base directory)"
                )
                continue

            if not resolved_document_path.is_file():
                print(
                    f"✗ Failed to ingest {document_path.name}: Path is not a regular file"
                )
                continue

            # Check file size limit to prevent resource exhaustion / DoS
            try:
                document_file_size = resolved_document_path.stat().st_size
            except (PermissionError, OSError):
                print(
                    f"✗ Failed to ingest {document_path.name}: Unable to query file stats"
                )
                continue

            if document_file_size > self.max_file_size_bytes:
                print(
                    f"✗ Failed to ingest {document_path.name}: File size exceeds limit ({document_file_size} bytes)"
                )
                continue

            try:
                parsed_docx = Document(resolved_document_path)
                paragraph_list = [
                    paragraph_element.text
                    for paragraph_element in parsed_docx.paragraphs
                    if paragraph_element.text.strip()
                ]
                self.documents[document_path.name] = {
                    "path": document_path,
                    "paragraphs": paragraph_list,
                    "num_paragraphs": len(paragraph_list),
                }
                ingested_document_names.append(document_path.name)
                print(f"✓ Ingested: {document_path.name}")
            except PackageNotFoundError:
                print(
                    f"✗ Failed to ingest {document_path.name}: Invalid or missing Word document package"
                )
            except OpcError:
                print(
                    f"✗ Failed to ingest {document_path.name}: Corrupted or malformed Word document structure"
                )
            except (PermissionError, OSError):
                print(
                    f"✗ Failed to ingest {document_path.name}: Unable to read file due to I/O or permission error"
                )
            except ValueError:
                print(
                    f"✗ Failed to ingest {document_path.name}: Invalid document parameter or format error"
                )
            except Exception as ingest_error:
                error_class_name = type(ingest_error).__name__
                print(
                    f"✗ Failed to ingest {document_path.name}: Ingestion error ({error_class_name})"
                )

        print(f"\nTotal documents ingested: {len(ingested_document_names)}")
        return ingested_document_names

    def digest_docs(self) -> Dict[str, Any]:
        """
        Digest the ingested documents by analyzing their content.
        Returns a summary of the analysis.

        Returns:
            Dict[str, Any]: A summary of the analysis with the following structure:
                {
                    'total_documents': int,  # Total number of ingested documents
                    'documents': {
                        <doc_name>: {
                            'num_paragraphs': int,      # Number of paragraphs in the document
                            'word_count': int,          # Total word count in the document
                            'key_topics': List[str],    # List of key topics extracted from the document
                            'summary': str              # Brief summary from the first few paragraphs
                        },
                        ...
                    }
                }
        """
        print("\n" + "=" * 80)
        print("DIGESTING DOCUMENTS")
        print("=" * 80)

        analysis = {"total_documents": len(self.documents), "documents": {}}

        for document_name, document_metadata in self.documents.items():
            paragraph_list = document_metadata["paragraphs"]

            # Extract key information
            document_analysis_summary = {
                "num_paragraphs": document_metadata["num_paragraphs"],
                "word_count": sum(
                    len(paragraph_text.split()) for paragraph_text in paragraph_list
                ),
                "key_topics": self._extract_key_topics(paragraph_list),
                "summary": self._create_summary(
                    paragraph_list[:3]
                ),  # First 3 paragraphs
            }

            analysis["documents"][document_name] = document_analysis_summary

            print(f"\n📄 {document_name}")
            print(
                f"   Paragraphs: {document_analysis_summary['num_paragraphs']}"
            )
            print(f"   Word Count: {document_analysis_summary['word_count']}")
            print(
                f"   Key Topics: {', '.join(document_analysis_summary['key_topics'][:5])}"
            )
            if document_analysis_summary["summary"]:
                print(f"   Summary: {document_analysis_summary['summary']}")

        return analysis

    def _extract_key_topics(self, paragraphs: List[str]) -> List[str]:
        """
        Extract key topics from text using simple frequency analysis.

        Args:
            paragraphs (List[str]): List of paragraph texts to analyze.

        Returns:
            List[str]: Top 10 most frequent words as key topics.
        """
        # Stop words list for filtering non-distinct words
        stop_words = {
            "the",
            "a",
            "an",
            "and",
            "or",
            "but",
            "in",
            "on",
            "at",
            "to",
            "for",
            "of",
            "with",
            "by",
            "from",
            "as",
            "is",
            "was",
            "are",
            "be",
            "been",
            "this",
            "that",
            "these",
            "those",
            "it",
            "its",
            "will",
            "can",
            "may",
            "would",
            "could",
            "should",
            "has",
            "have",
            "had",
            "do",
            "does",
            "did",
            "they",
            "their",
            "them",
            "we",
            "you",
            "your",
            "our",
            "who",
            "which",
            "what",
            "when",
            "where",
            "how",
            "all",
            "each",
            "some",
            "more",
            "most",
            "other",
            "into",
            "through",
            "during",
            "before",
            "after",
            "above",
            "below",
            "between",
            "under",
            "again",
            "further",
            "then",
            "once",
            "here",
            "there",
            "than",
            "such",
            "only",
            "very",
            "just",
            "also",
            "being",
            "both",
            "about",
            "over",
            "any",
            "same",
            "own",
            "while",
        }

        # Count word frequencies across all paragraphs
        word_frequency = {}
        for paragraph_text in paragraphs:
            raw_word_tokens = paragraph_text.lower().split()
            for raw_word_token in raw_word_tokens:
                # Remove punctuation
                clean_word_token = "".join(
                    char_symbol
                    for char_symbol in raw_word_token
                    if char_symbol.isalnum()
                )
                # Apply basic stemming (remove common suffixes)
                if clean_word_token.endswith("ing"):
                    clean_word_token = clean_word_token[:-3]
                elif clean_word_token.endswith("ed"):
                    clean_word_token = clean_word_token[:-2]
                elif clean_word_token.endswith("s") and len(clean_word_token) > 4:
                    clean_word_token = clean_word_token[:-1]

                if len(clean_word_token) > 3 and clean_word_token not in stop_words:
                    word_frequency[clean_word_token] = (
                        word_frequency.get(clean_word_token, 0) + 1
                    )

        # Sort by frequency and return top words
        sorted_word_frequencies = sorted(
            word_frequency.items(),
            key=lambda frequency_item: frequency_item[1],
            reverse=True,
        )
        return [
            word_token
            for word_token, frequency_count in sorted_word_frequencies[:10]
        ]

    def _create_summary(self, paragraphs: List[str]) -> str:
        """Create a brief summary from the first few paragraphs."""
        summary_text = " ".join(paragraphs)
        if len(summary_text) > 300:
            truncated_text = summary_text[:297]
            # Truncate at the last complete word before the cutoff
            if " " in truncated_text:
                truncated_text = truncated_text.rsplit(" ", 1)[0]
            summary_text = truncated_text + "..."
        return summary_text

    def suggest_path(self, analysis: Dict[str, Any]) -> None:
        """
        Suggest a development path based on the document analysis.

        Args:
            analysis (Dict[str, Any]): Dictionary containing document analysis results.
        """
        print("\n" + "=" * 80)
        print("SUGGESTED PATH")
        print("=" * 80)

        print(
            "\nBased on the analysis of the ingested documents, here are the recommendations:\n"
        )

        # Categorize documents by content
        fitness_docs = []
        business_docs = []
        gamification_docs = []

        for document_name, document_metadata in analysis["documents"].items():
            extracted_topic_list = [
                topic_token.lower()
                for topic_token in document_metadata["key_topics"]
            ]

            if any(
                target_topic in extracted_topic_list
                for target_topic in [
                    "fitness",
                    "training",
                    "workout",
                    "exercise",
                    "gym",
                ]
            ):
                fitness_docs.append(document_name)
            if any(
                target_topic in extracted_topic_list
                for target_topic in [
                    "business",
                    "enterprise",
                    "coaching",
                    "model",
                ]
            ):
                business_docs.append(document_name)
            if any(
                target_topic in extracted_topic_list
                for target_topic in ["gamified", "game", "battle", "quest"]
            ):
                gamification_docs.append(document_name)

        # Show document categorization first
        print("=" * 80)
        print("DOCUMENT CATEGORIZATION")
        print("=" * 80)

        if fitness_docs:
            print(f"\n💪 Fitness-focused documents ({len(fitness_docs)}):")
            for doc in fitness_docs:
                print(f"   - {doc}")

        if business_docs:
            print(f"\n💼 Business-focused documents ({len(business_docs)}):")
            for doc in business_docs:
                print(f"   - {doc}")

        if gamification_docs:
            print(f"\n🎮 Gamification-focused documents ({len(gamification_docs)}):")
            for doc in gamification_docs:
                print(f"   - {doc}")

        # Generate dynamic development path based on analysis
        print("\n" + "=" * 80)
        print("📋 RECOMMENDED DEVELOPMENT PATH (Generated from Analysis)")
        print("=" * 80)
        print()

        phase = 1

        # Always start with foundation if there are any documents
        if fitness_docs or business_docs or gamification_docs:
            print(f"{phase}. BUILD THE FOUNDATION")
            if fitness_docs:
                print("   └─ Create core data models for fitness coaching")
                print("   └─ Define user profiles and progress tracking structures")
            if business_docs:
                print("   └─ Establish client management and business workflows")
            if gamification_docs:
                print(
                    "   └─ Set up gamification mechanics (points, levels, achievements)"
                )
            phase += 1
            print()

        # Add interface development if we have fitness or business content
        if fitness_docs or business_docs:
            print(f"{phase}. DEVELOP THE INTERFACE")
            if fitness_docs:
                print("   └─ Design user dashboard for fitness tracking")
                print("   └─ Create workout logging interface")
                print("   └─ Build progress visualization tools")
            if business_docs:
                print("   └─ Develop coach management dashboard")
                print("   └─ Create client onboarding interface")
            phase += 1
            print()

        # Add coaching features if we have fitness content
        if fitness_docs:
            print(f"{phase}. IMPLEMENT COACHING FEATURES")
            print("   └─ Personalized workout recommendations")
            print("   └─ Goal setting and tracking")
            print("   └─ Motivational messaging system")
            print("   └─ Progress assessment and feedback")
            phase += 1
            print()

        # Add gamification layer if we detected gamification focus
        if gamification_docs:
            print(f"{phase}. ADD GAMIFICATION LAYER")
            print("   └─ Achievement system based on milestones")
            print("   └─ Challenge modes and battle scenarios")
            print("   └─ Leaderboards and social features")
            print("   └─ Reward mechanisms and badges")
            phase += 1
            print()
        else:
            # Suggest gamification as optional if not strongly detected
            if fitness_docs or business_docs:
                print(f"{phase}. CONSIDER GAMIFICATION (Optional)")
                print(
                    "   └─ NOTE: No strong gamification signals detected in documents"
                )
                print("   └─ Consider if gamification aligns with business goals")
                print("   └─ Could add achievement system for user engagement")
                phase += 1
                print()

        # Add business integration if we have business content
        if business_docs:
            print(f"{phase}. BUSINESS INTEGRATION")
            print("   └─ Payment and subscription management")
            print("   └─ Coach-client communication tools")
            print("   └─ Analytics and reporting dashboard")
            print("   └─ Marketing and customer acquisition features")
            phase += 1
            print()

        print("=" * 80)
        print("NEXT STEPS")
        print("=" * 80)

        # Generate dynamic next steps based on what was found
        next_steps = []

        if analysis["total_documents"] > 0:
            next_steps.append("Review the ingested document content in detail")

        if fitness_docs:
            next_steps.append(
                "Define fitness coaching data models and workout structures"
            )

        if gamification_docs:
            next_steps.append("Design gamification mechanics and reward systems")

        if business_docs:
            next_steps.append("Outline business requirements and revenue models")

        next_steps.extend(
            [
                "Create technical specifications for each component",
                "Set up project structure (frontend, backend, database)",
                "Begin iterative development with MVP features first",
            ]
        )

        for step_index, step_description in enumerate(next_steps, 1):
            print(f"\n{step_index}. {step_description}")

        print("\n✓ Analysis complete!")

    def run(self) -> None:
        """Execute the full pipeline: ingest, digest, suggest."""
        ingested_documents = self.ingest_docs()

        if not ingested_documents:
            print("\n⚠️  No documents found to analyze.")
            return

        analysis_result = self.digest_docs()
        self.suggest_path(analysis_result)


def main():
    """Main entry point for the document analyzer."""
    script_directory = Path(__file__).parent if "__file__" in globals() else Path.cwd()
    source_documents_directory = script_directory.parent / "docs" / "source-documents"

    if not source_documents_directory.exists():
        print(
            f"Error: Source documents directory not found: {source_documents_directory}"
        )
        print("Please ensure documents are in docs/source-documents/")
        sys.exit(1)

    document_analyzer = DocumentAnalyzer(
        base_dir=str(source_documents_directory)
    )
    document_analyzer.run()


if __name__ == "__main__":
    main()

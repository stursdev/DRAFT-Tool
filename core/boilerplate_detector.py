# =============================================================================
# core/boilerplate_detector.py
#
# Post-migration pass that identifies migrated content (green text) that
# exactly matches boilerplate from the destination template, and recolors
# those matching sentences blue so reviewers can see which content was
# copied verbatim from the template without customization.
#
# Why this is useful:
#   Programs often migrate content that still contains the original template
#   boilerplate text unchanged. Blue highlighting makes this immediately
#   visible in the output document, prompting the program to customize
#   that content for their specific needs.
#
# How it works:
#   1. Extract all text from the destination template and split into sentences
#   2. Open the output document and walk every paragraph
#   3. For any paragraph where all runs are green (migrated content), check
#      if the paragraph's full text exactly matches a boilerplate sentence
#   4. If matched, recolor that paragraph's runs from green to blue
#
# Sentence boundary definition:
#   A sentence ends at ". " (period followed by a space) or at a line break.
#   This covers most prose in SEP documents. Numbered lists and bullet points
#   that do not end with periods are matched as full-line units.
#
# Match type:
#   Exact match only (100% identical after stripping leading/trailing
#   whitespace). No fuzzy matching — only verbatim boilerplate reuse is flagged.
#
# This runs automatically after every migration with no user action required.
# =============================================================================

import sys
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor


# The green color used for all migrated content — defined in content_migrator.py
# and duplicated here to avoid a circular import. Both must match exactly.
MIGRATED_CONTENT_COLOR     = RGBColor(0x00, 0xB0, 0x50)   # #00B050 — green
MIGRATED_CONTENT_COLOR_HEX = "00B050"                      # Hex used in XML w:color val

# The blue color applied to migrated content that matches boilerplate verbatim.
# Using a clear mid-blue that is visually distinct from both the green
# migrated text and the black original text.
BOILERPLATE_MATCH_COLOR = RGBColor(0x00, 0x70, 0xC0)   # #0070C0 — blue


class BoilerplateDetector:
    """
    Identifies migrated content that exactly matches destination template
    boilerplate and recolors it blue in the output document.

    Usage:
        detector = BoilerplateDetector(
            template_path = "path/to/SEP_template_v2.docx",
            output_path   = "path/to/migrated_output.docx",
        )
        match_count = detector.run()
    """

    def __init__(self, template_path: str, output_path: str):
        # Path to the original destination template — the source of boilerplate text
        self.template_path = Path(template_path)

        # Path to the migrated output document — this file is modified in place
        self.output_path = Path(output_path)

    # =========================================================================
    # Public API
    # =========================================================================

    def run(self) -> int:
        """
        Execute the boilerplate detection pass on the output document.

        Opens the output file, scans all green paragraphs for exact matches
        against the boilerplate sentence set, recolors matches to blue,
        and saves the file back to the same path.

        Returns the number of paragraphs recolored blue.
        """
        # Build the set of boilerplate sentences from the template
        boilerplate_sentences = self._extract_boilerplate_sentences()

        if not boilerplate_sentences:
            # Template has no text content — nothing to compare against
            return 0

        # Open the output document for modification
        output_document = Document(str(self.output_path))

        recolored_count = 0

        for paragraph in output_document.paragraphs:
            # Only inspect paragraphs where all runs are green (migrated content).
            # Paragraphs with mixed colors (some green, some black) are original
            # destination content and should not be modified.
            if not self._paragraph_is_fully_green(paragraph):
                continue

            # Check if this paragraph's text exactly matches a boilerplate sentence
            paragraph_text = paragraph.text.strip()
            if not paragraph_text:
                continue

            if paragraph_text in boilerplate_sentences:
                # Exact match found — recolor all runs from green to blue
                self._recolor_paragraph_runs(paragraph, BOILERPLATE_MATCH_COLOR)
                recolored_count += 1

        if recolored_count > 0:
            # Only save if changes were made to avoid unnecessary file writes
            output_document.save(str(self.output_path))

        return recolored_count

    # =========================================================================
    # Private helpers
    # =========================================================================

    def _extract_boilerplate_sentences(self) -> set[str]:
        """
        Open the destination template and extract all unique sentences
        from its paragraphs.

        Sentence splitting:
          Each paragraph's text is split on ". " (period + space) to produce
          individual sentences. The paragraph text itself is also added as a
          whole unit to catch single-sentence paragraphs and list items that
          do not end with a period.

          Leading and trailing whitespace is stripped from each sentence.
          Empty strings are excluded.

        Returns a set of sentence strings for O(1) lookup during detection.
        """
        template_document = Document(str(self.template_path))
        boilerplate_sentences: set[str] = set()

        for paragraph in template_document.paragraphs:
            paragraph_text = paragraph.text.strip()
            if not paragraph_text:
                continue

            # Add the full paragraph text as one unit.
            # This handles list items, short headings, and single-sentence
            # paragraphs that do not end with a period.
            boilerplate_sentences.add(paragraph_text)

            # Also split on ". " to handle multi-sentence paragraphs.
            # We rejoin each split part with the period it was split on
            # so the stored sentence matches what appears in migrated content.
            parts = paragraph_text.split(". ")
            for index, part in enumerate(parts):
                part = part.strip()
                if not part:
                    continue
                # Re-add the period to all parts except the last one
                # (the last part either ends the paragraph or has no period)
                if index < len(parts) - 1:
                    part = part + "."
                boilerplate_sentences.add(part)

        return boilerplate_sentences

    def _paragraph_is_fully_green(self, paragraph) -> bool:
        """
        Return True if every text-bearing run in the paragraph carries the
        migrated content green color as an explicit XML attribute.

        Why XML instead of the python-docx API:
          run.font.color.rgb raises AttributeError when the run's color type
          is NONE (inherited) or THEME, which causes the API-based check to
          silently return False for valid green runs. Reading the <w:color
          w:val> attribute directly from the run's XML is always safe and
          matches exactly the hex value we wrote during migration.

        A paragraph with no runs (empty paragraph) returns False so we
        do not accidentally flag empty paragraphs as boilerplate matches.
        """
        text_bearing_runs = [run for run in paragraph.runs if run.text.strip()]

        if not text_bearing_runs:
            return False

        target_hex = MIGRATED_CONTENT_COLOR_HEX.upper()

        for run in text_bearing_runs:
            rpr = run._element.find(qn("w:rPr"))
            if rpr is None:
                return False   # No run properties → no explicit color set

            color_elem = rpr.find(qn("w:color"))
            if color_elem is None:
                return False   # No color element → color is inherited, not green

            run_color_hex = color_elem.get(qn("w:val"), "").upper()
            if run_color_hex != target_hex:
                return False   # Different explicit color (black, blue, auto, etc.)

        return True

    def _recolor_paragraph_runs(self, paragraph, new_color: RGBColor):
        """
        Set the font color of every run in the paragraph to new_color.

        Called when a green paragraph's text exactly matches a boilerplate
        sentence, converting it from green (migrated) to blue (boilerplate reuse).
        """
        for run in paragraph.runs:
            run.font.color.rgb = new_color

# =============================================================================
# core/boilerplate_detector.py
#
# Post-migration pass that identifies migrated content (green text) that
# duplicates the original boilerplate already present in the same destination
# section, and recolors those matching paragraphs blue.
#
# Design — section-scoped, sentence-level matching:
#
#   Section-scoped:
#     The template is parsed into a per-section sentence map:
#       { "1.1 Purpose": {"sentence A", "sentence B", ...}, ... }
#     When scanning the output document, migrated (green) paragraphs in a
#     section are only compared against sentences from THAT section in the
#     template — not against the entire template globally.
#
#     Why this matters: a sentence that appears in Section C of the template
#     is irrelevant when reviewing migrated content in Section B. Comparing
#     globally produces false positives and misses the real intent, which is
#     "does this migrated sentence duplicate what was already in this section?"
#
#   Sentence-level:
#     Each green paragraph's text is compared both as a whole unit and as
#     individual sentences (split on sentence-ending punctuation). A match
#     on any sentence is enough to flag the paragraph blue, because even one
#     verbatim boilerplate sentence in a migrated paragraph warrants review.
#
#   Bullet/list handling:
#     python-docx's paragraph.text returns only the stored text content —
#     the bullet character or list number is rendered by Word's list style
#     and is NOT part of the text string. This means a bullet-point paragraph
#     and a plain paragraph with identical text produce identical paragraph.text
#     values, so list items are matched naturally with no extra logic.
#
#   Whitespace normalization:
#     All text is normalized before comparison: runs of whitespace (spaces,
#     tabs, non-breaking spaces) are collapsed to a single space and leading /
#     trailing whitespace is stripped. This prevents invisible formatting
#     differences (extra spaces, tab characters, different line endings)
#     from causing valid matches to be missed.
#
# Match color:
#   Paragraphs whose text (or any contained sentence) verbatim matches the
#   section boilerplate are recolored from green to blue (#0070C0).
#   Paragraphs with no match stay green.
# =============================================================================

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor


MIGRATED_CONTENT_COLOR     = RGBColor(0x00, 0xB0, 0x50)   # #00B050 — green
MIGRATED_CONTENT_COLOR_HEX = "00B050"
BOILERPLATE_MATCH_COLOR    = RGBColor(0x00, 0x70, 0xC0)   # #0070C0 — blue

# Regex that matches one or more whitespace characters (including non-breaking
# spaces \xa0 and tabs) so all of them collapse to a single regular space.
_WHITESPACE_RE = re.compile(r'[\s\xa0]+')

# Regex used to split normalized text into individual sentences.
# Splits after any sentence-ending punctuation (. ! ?) that is followed by
# one or more whitespace characters. This handles "Sentence one. Sentence two."
# as well as exclamations and questions.
_SENTENCE_SPLIT_RE = re.compile(r'(?<=[.!?])\s+')


class BoilerplateDetector:
    """
    Recolors migrated (green) paragraphs blue when their text duplicates
    the original boilerplate from the same destination section.

    Usage:
        detector = BoilerplateDetector(
            template_path = "path/to/destination_template.docx",
            output_path   = "path/to/migrated_output.docx",
        )
        match_count = detector.run()
    """

    def __init__(self, template_path: str, output_path: str):
        self.template_path = Path(template_path)
        self.output_path   = Path(output_path)

    # =========================================================================
    # Public API
    # =========================================================================

    def run(self) -> int:
        """
        Execute the boilerplate detection pass and save the output document.

        Returns the number of paragraphs recolored blue.
        """
        # Build the per-section boilerplate map from the template
        section_boilerplate = self._extract_section_boilerplate()

        if not section_boilerplate:
            return 0

        output_document = Document(str(self.output_path))
        recolored_count = 0

        # Walk the output document, tracking which section we are currently in.
        # The section title resets every time we encounter a heading paragraph.
        # Migrated (green) paragraphs are compared only against the sentence set
        # for the section they belong to.
        current_section_title = None

        for paragraph in output_document.paragraphs:

            # ── Heading → update the active section ───────────────────────────
            if self._is_heading(paragraph):
                current_section_title = paragraph.text.strip()
                continue

            # ── Skip if no active section or paragraph is not green ───────────
            if current_section_title is None:
                continue
            if not self._paragraph_is_fully_green(paragraph):
                continue

            # ── Look up this section's boilerplate sentences ───────────────────
            section_sentences = section_boilerplate.get(current_section_title)
            if not section_sentences:
                continue

            # ── Compare paragraph against the section boilerplate ─────────────
            if self._matches_boilerplate(paragraph.text, section_sentences):
                self._recolor_paragraph_runs(paragraph, BOILERPLATE_MATCH_COLOR)
                recolored_count += 1

        if recolored_count > 0:
            output_document.save(str(self.output_path))

        return recolored_count

    # =========================================================================
    # Template parsing — build per-section sentence map
    # =========================================================================

    def _extract_section_boilerplate(self) -> dict[str, set[str]]:
        """
        Parse the destination template and build a mapping of:
            section heading title → set of normalized sentences

        Each non-heading paragraph contributes:
          - Its full normalized text as one entry
          - Each individual sentence (split on punctuation) as separate entries

        This means the lookup can match both whole-paragraph content and
        individual sentences that appear inside longer paragraphs.

        Bullet / list items are included automatically because paragraph.text
        returns the text content without any bullet character or list number.
        """
        template_document  = Document(str(self.template_path))
        section_boilerplate: dict[str, set[str]] = {}
        current_section = None

        for paragraph in template_document.paragraphs:

            if self._is_heading(paragraph):
                current_section = paragraph.text.strip()
                if current_section and current_section not in section_boilerplate:
                    section_boilerplate[current_section] = set()
                continue

            if current_section is None:
                continue

            normalized = _normalize(paragraph.text)
            if not normalized:
                continue

            # Add the whole paragraph as one lookup entry
            section_boilerplate[current_section].add(normalized)

            # Also add each individual sentence so that a migrated paragraph
            # whose text equals just one sentence from a multi-sentence
            # template paragraph is still detected.
            for sentence in _split_sentences(normalized):
                section_boilerplate[current_section].add(sentence)

        return section_boilerplate

    # =========================================================================
    # Matching
    # =========================================================================

    def _matches_boilerplate(
        self,
        paragraph_text:   str,
        section_sentences: set[str],
    ) -> bool:
        """
        Return True if the paragraph text (or any of its individual sentences)
        exactly matches an entry in the section's boilerplate sentence set.

        Two-pass check:
          Pass 1 — whole paragraph: catches paragraphs that are verbatim
                   copies of a full template paragraph.
          Pass 2 — sentence by sentence: catches paragraphs that contain
                   one or more boilerplate sentences mixed with new content.
                   A single matching sentence is enough to flag the paragraph,
                   because even partial verbatim reuse warrants user review.

        All comparisons use normalized text so whitespace differences do not
        cause valid matches to be missed.
        """
        normalized = _normalize(paragraph_text)
        if not normalized:
            return False

        # Pass 1 — full paragraph match
        if normalized in section_sentences:
            return True

        # Pass 2 — individual sentence match
        for sentence in _split_sentences(normalized):
            if sentence in section_sentences:
                return True

        return False

    # =========================================================================
    # Green paragraph detection
    # =========================================================================

    def _is_heading(self, paragraph) -> bool:
        """Return True if the paragraph uses a Heading style (any level)."""
        try:
            return paragraph.style.name.lower().startswith("heading")
        except AttributeError:
            return False

    def _paragraph_is_fully_green(self, paragraph) -> bool:
        """
        Return True if every text-bearing run in the paragraph carries the
        migrated content green color as an explicit XML w:color attribute.

        Reads the XML directly rather than using the python-docx color API
        because run.font.color.rgb raises AttributeError for inherited or
        theme-based colors, which would cause valid green runs to be skipped.
        """
        text_bearing_runs = [run for run in paragraph.runs if run.text.strip()]

        if not text_bearing_runs:
            return False

        target_hex = MIGRATED_CONTENT_COLOR_HEX.upper()

        for run in text_bearing_runs:
            rpr = run._element.find(qn("w:rPr"))
            if rpr is None:
                return False

            color_elem = rpr.find(qn("w:color"))
            if color_elem is None:
                return False

            if color_elem.get(qn("w:val"), "").upper() != target_hex:
                return False

        return True

    # =========================================================================
    # Recoloring
    # =========================================================================

    def _recolor_paragraph_runs(self, paragraph, new_color: RGBColor):
        """Set every run in the paragraph to new_color."""
        for run in paragraph.runs:
            run.font.color.rgb = new_color


# =============================================================================
# Module-level text helpers
# =============================================================================

def _normalize(text: str) -> str:
    """
    Collapse all whitespace (including tabs and non-breaking spaces) to a
    single space and strip leading / trailing whitespace.

    Applied to both template sentences and migrated paragraph text before
    any comparison so invisible formatting differences never cause a miss.
    """
    return _WHITESPACE_RE.sub(" ", text).strip()


def _split_sentences(text: str) -> list[str]:
    """
    Split normalized text into individual sentences.

    Splits on any sentence-ending punctuation (. ! ?) followed by whitespace.
    Each resulting part is stripped; empty parts are discarded.

    Examples:
      "Sentence one. Sentence two."  → ["Sentence one.", "Sentence two."]
      "Only one sentence"            → ["Only one sentence"]
      "First! Second? Third."        → ["First!", "Second?", "Third."]
    """
    parts = _SENTENCE_SPLIT_RE.split(text)
    return [p.strip() for p in parts if p.strip()]

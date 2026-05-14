# =============================================================================
# core/boilerplate_detector.py
#
# Post-migration pass that identifies migrated content (green text) that
# duplicates the original boilerplate already present in the same destination
# section, and recolors those sentences blue.
#
# Design — section-scoped, sentence-level matching and coloring:
#
#   Section-scoped:
#     The template is parsed into a per-section sentence map:
#       { "1.1 Purpose": {"sentence A", "sentence B", ...}, ... }
#     Migrated (green) paragraphs in a section are only compared against
#     sentences from THAT section in the template — not globally.
#
#   Sentence-level coloring (within a paragraph):
#     Each green paragraph is analyzed sentence by sentence. Sentences that
#     exactly match the section boilerplate are colored blue; sentences that
#     do not match stay green. Both colors can appear in the same paragraph.
#
#     Example — paragraph with 2 sentences:
#       "The contractor shall deliver reports. ← verbatim boilerplate  → blue
#        Deliverables are defined in SOW."     ← customised content    → green
#
#   How runs are colored:
#     Word stores paragraph text as a sequence of runs (<w:r> elements).
#     Runs can straddle sentence boundaries — e.g. "…performance. Probability
#     of an undesired event" may be one run even though it spans two sentences.
#     To handle this accurately we split any run that crosses a sentence
#     boundary in the XML before applying colors, so each resulting run
#     belongs entirely to one sentence. _split_and_color_run does this by
#     deep-copying the original run element (preserving all formatting) and
#     inserting the new elements immediately after the original.
#
#   Bullet / list handling:
#     paragraph.text returns only the stored text; bullet characters and list
#     numbers are rendered by the list style and are NOT in the text string.
#     A bullet-point paragraph and a plain paragraph with identical wording
#     produce identical paragraph.text values and match correctly.
#
#   Whitespace normalization:
#     All text is normalized before comparison (tabs, multiple spaces, and
#     non-breaking spaces collapsed to one space). The raw text and run
#     structure are never modified — normalization is only used for lookup.
# =============================================================================

import copy
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

# Boilerplate detection modes — passed as the `mode` argument to BoilerplateDetector.
#
#   SENTENCE  — color only the individual sentences that match boilerplate blue;
#               non-matching sentences in the same paragraph stay green.
#               Best for reviewing partially-customized paragraphs.
#
#   PARAGRAPH — color the entire paragraph blue if any sentence within it
#               matches boilerplate. Faster to scan at a glance but can flag
#               paragraphs that have been partially customized.
#
HIGHLIGHT_SENTENCES = "sentence"
HIGHLIGHT_PARAGRAPH = "paragraph"

# Collapses any run of whitespace (including tabs and non-breaking spaces)
# to a single regular space. Applied before every comparison.
_WHITESPACE_RE = re.compile(r'[\s\xa0]+')

# Splits text at sentence boundaries: matches whitespace that follows a
# sentence-ending character (. ! ?). The lookbehind keeps the punctuation
# attached to its sentence; the whitespace is consumed by the match so
# the next sentence span starts cleanly.
_SENTENCE_SPLIT_RE = re.compile(r'(?<=[.!?])\s+')


class BoilerplateDetector:
    """
    Recolors migrated (green) text blue at the sentence level when sentences
    verbatim duplicate the original boilerplate from the same destination section.

    Usage:
        detector = BoilerplateDetector(
            template_path = "path/to/destination_template.docx",
            output_path   = "path/to/migrated_output.docx",
        )
        match_count = detector.run()
    """

    def __init__(
        self,
        template_path: str,
        output_path:   str,
        mode:          str = HIGHLIGHT_SENTENCES,
    ):
        self.template_path = Path(template_path)
        self.output_path   = Path(output_path)
        # HIGHLIGHT_SENTENCES — color individual matching sentences blue
        # HIGHLIGHT_PARAGRAPH — color the whole paragraph blue if any sentence matches
        self.mode = mode

    # =========================================================================
    # Public API
    # =========================================================================

    def run(self) -> int:
        """
        Execute the boilerplate detection pass and save the output document.

        Walks the output document paragraph by paragraph, tracking the current
        section heading. For each fully-green paragraph, applies sentence-level
        coloring: matching sentences → blue, non-matching → stay green.

        Returns the total number of paragraphs where at least one sentence
        was recolored blue.
        """
        section_boilerplate = self._extract_section_boilerplate()
        if not section_boilerplate:
            return 0

        output_document       = Document(str(self.output_path))
        recolored_count       = 0
        current_section_title = None

        for paragraph in output_document.paragraphs:

            # Track the active section heading
            if self._is_heading(paragraph):
                current_section_title = paragraph.text.strip()
                continue

            if current_section_title is None:
                continue
            if not self._paragraph_is_fully_green(paragraph):
                continue

            section_sentences = section_boilerplate.get(current_section_title)
            if not section_sentences:
                continue

            if self.mode == HIGHLIGHT_PARAGRAPH:
                # Paragraph mode: whole paragraph turns blue if any sentence matches
                if self._any_sentence_matches(paragraph.text, section_sentences):
                    self._recolor_paragraph_runs(paragraph, BOILERPLATE_MATCH_COLOR)
                    recolored_count += 1
            else:
                # Sentence mode (default): color only the matching sentences
                if self._apply_sentence_colors(paragraph, section_sentences):
                    recolored_count += 1

        if recolored_count > 0:
            output_document.save(str(self.output_path))

        return recolored_count

    # =========================================================================
    # Template parsing
    # =========================================================================

    def _extract_section_boilerplate(self) -> dict[str, set[str]]:
        """
        Parse the destination template and build:
            { section_heading_title: set_of_normalized_sentences }

        Each non-heading paragraph contributes its full normalized text AND
        each individual sentence extracted from it. This lets the detector
        match both whole-paragraph copies and individual sentences that
        appear inside longer template paragraphs.

        Bullet / list items are included automatically because paragraph.text
        does not contain bullet characters or list numbers.
        """
        template_document   = Document(str(self.template_path))
        section_boilerplate: dict[str, set[str]] = {}
        current_section     = None

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

            section_boilerplate[current_section].add(normalized)
            for sentence in _split_sentences(normalized):
                section_boilerplate[current_section].add(sentence)

        return section_boilerplate

    # =========================================================================
    # Sentence-level coloring
    # =========================================================================

    def _any_sentence_matches(
        self,
        paragraph_text:    str,
        section_sentences: set[str],
    ) -> bool:
        """
        Return True if the full paragraph text OR any individual sentence
        within it exactly matches an entry in section_sentences.
        Used by HIGHLIGHT_PARAGRAPH mode.
        """
        normalized = _normalize(paragraph_text)
        if not normalized:
            return False
        if normalized in section_sentences:
            return True
        return any(s in section_sentences for s in _split_sentences(normalized))

    def _apply_sentence_colors(
        self,
        paragraph,
        section_sentences: set[str],
    ) -> bool:
        """
        Color each sentence in a green paragraph independently:
          - Sentences that match the section boilerplate → blue
          - Sentences that do not match                  → stay green

        Runs that straddle a sentence boundary are split in the XML at the
        boundary so each resulting segment is colored independently.

        Returns True if at least one sentence was colored blue.
        """
        color_map = self._build_sentence_color_map(
            paragraph.text,
            section_sentences,
        )

        if not color_map:
            return False

        any_blue = any(color is BOILERPLATE_MATCH_COLOR for _, _, color in color_map)
        if not any_blue:
            return False

        all_blue = all(color is BOILERPLATE_MATCH_COLOR for _, _, color in color_map)

        if all_blue:
            # Every sentence is boilerplate — simple whole-paragraph recolor
            for run in paragraph.runs:
                run.font.color.rgb = BOILERPLATE_MATCH_COLOR
            return True

        # Global positions where the color changes (start of each span after the first)
        boundaries = {span[0] for span in color_map[1:]}

        char_pos = 0
        for run in list(paragraph.runs):
            if not run.text:
                continue

            run_len = len(run.text)   # capture before any XML modification
            run_end = char_pos + run_len

            # Sentence boundaries that fall strictly inside this run (local offsets)
            inner = sorted(b - char_pos for b in boundaries if char_pos < b < run_end)

            if not inner:
                run.font.color.rgb = self._color_for_position(char_pos, color_map)
            else:
                self._split_and_color_run(run, inner, char_pos, color_map)

            char_pos = run_end

        return True

    def _split_and_color_run(
        self,
        run,
        split_positions: list[int],
        run_global_start: int,
        color_map: list[tuple[int, int, RGBColor]],
    ) -> None:
        """
        Split a run at local character offsets and color each segment.

        The original run element is modified in-place for the first segment.
        Deep-copies are inserted immediately after it for each subsequent
        segment, preserving all run formatting (bold, italic, font, etc.).
        """
        text = run.text
        cuts = [0] + split_positions + [len(text)]
        segments = [(cuts[i], cuts[i + 1]) for i in range(len(cuts) - 1)]

        original_elem = run._element
        # Clone before any modification so every copy has the original formatting
        clones = [copy.deepcopy(original_elem) for _ in segments[1:]]

        # Update original run to hold only the first segment
        _set_run_text(original_elem, text[segments[0][0] : segments[0][1]])
        _set_run_color(original_elem, self._color_for_position(run_global_start, color_map))

        # Insert remaining segments as new run elements after the original
        insert_after = original_elem
        for i, (seg_start, seg_end) in enumerate(segments[1:]):
            seg_color = self._color_for_position(run_global_start + seg_start, color_map)
            clone = clones[i]
            _set_run_text(clone, text[seg_start:seg_end])
            _set_run_color(clone, seg_color)
            insert_after.addnext(clone)
            insert_after = clone

    def _build_sentence_color_map(
        self,
        raw_text:         str,
        section_sentences: set[str],
    ) -> list[tuple[int, int, RGBColor]]:
        """
        Return a list of (start, end, color) spans covering the full raw_text.

        Each span corresponds to one sentence (including its trailing
        whitespace so spans are contiguous and cover the entire string).
        The color is BOILERPLATE_MATCH_COLOR if the normalized sentence text
        is in section_sentences, otherwise MIGRATED_CONTENT_COLOR.

        Normalization is applied only for the set lookup — raw_text positions
        are preserved so they map correctly onto run character offsets.
        """
        result   = []
        last_end = 0

        for match in _SENTENCE_SPLIT_RE.finditer(raw_text):
            sentence = _normalize(raw_text[last_end : match.start()])
            color    = (
                BOILERPLATE_MATCH_COLOR
                if sentence in section_sentences
                else MIGRATED_CONTENT_COLOR
            )
            result.append((last_end, match.end(), color))
            last_end = match.end()

        # Final sentence — no trailing split point
        if last_end < len(raw_text):
            sentence = _normalize(raw_text[last_end:])
            color    = (
                BOILERPLATE_MATCH_COLOR
                if sentence in section_sentences
                else MIGRATED_CONTENT_COLOR
            )
            result.append((last_end, len(raw_text), color))

        return result

    @staticmethod
    def _color_for_position(
        char_pos: int,
        color_map: list[tuple[int, int, RGBColor]],
    ) -> RGBColor:
        """
        Return the color assigned to the sentence that contains char_pos.
        Falls back to MIGRATED_CONTENT_COLOR (green) if no span matches.
        """
        for span_start, span_end, color in color_map:
            if span_start <= char_pos < span_end:
                return color
        return MIGRATED_CONTENT_COLOR

    # =========================================================================
    # Paragraph inspection helpers
    # =========================================================================

    def _is_heading(self, paragraph) -> bool:
        """Return True if the paragraph uses a Heading style (any level)."""
        try:
            return paragraph.style.name.lower().startswith("heading")
        except AttributeError:
            return False

    def _paragraph_is_fully_green(self, paragraph) -> bool:
        """
        Return True if every text-bearing run carries the migrated green color
        as an explicit XML w:color attribute.

        Reads XML directly rather than using font.color.rgb because the
        python-docx API raises AttributeError for inherited / theme colors,
        which would incorrectly exclude valid green paragraphs.
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

    def _recolor_paragraph_runs(self, paragraph, new_color: RGBColor):
        """Set every run in the paragraph to new_color."""
        for run in paragraph.runs:
            run.font.color.rgb = new_color


# =============================================================================
# Module-level text helpers
# =============================================================================

def _normalize(text: str) -> str:
    """
    Collapse all whitespace (tabs, multiple spaces, non-breaking spaces) to a
    single space and strip leading / trailing whitespace.
    Used only for comparison — never applied to the document content itself.
    """
    return _WHITESPACE_RE.sub(" ", text).strip()


def _split_sentences(text: str) -> list[str]:
    """
    Split normalized text into individual sentences on . ! ? boundaries.
    Empty parts are discarded.
    """
    return [p.strip() for p in _SENTENCE_SPLIT_RE.split(text) if p.strip()]


# =============================================================================
# Run XML helpers (used by _split_and_color_run)
# =============================================================================

_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _set_run_text(run_elem, text: str) -> None:
    """Update <w:t> on a run element, setting xml:space=preserve when needed."""
    t = run_elem.find(qn("w:t"))
    if t is None:
        return
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set(_XML_SPACE, "preserve")
    elif _XML_SPACE in t.attrib:
        del t.attrib[_XML_SPACE]


def _set_run_color(run_elem, color: RGBColor) -> None:
    """Set <w:color w:val> on a run element's existing <w:rPr>."""
    rpr = run_elem.find(qn("w:rPr"))
    if rpr is None:
        return
    color_elem = rpr.find(qn("w:color"))
    if color_elem is None:
        return
    color_elem.set(qn("w:val"), str(color))

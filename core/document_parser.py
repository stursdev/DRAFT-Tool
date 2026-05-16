# =============================================================================
# core/document_parser.py
#
# Reads a .docx file and produces a flat, ordered list of Section objects.
#
# Responsibility:
#   Open a Word document and extract every heading (at any level) along with
#   all content (paragraphs, tables, images) that belongs under each heading.
#
# Key behaviors:
#
#   1. Heading detection
#      Three strategies are tried in order:
#        a) Style name starts with "Heading" — covers all built-in heading styles.
#        b) Paragraph XML <w:outlineLvl> — covers custom-named styles that
#           explicitly set the outline level on each paragraph element.
#        c) Style inheritance chain (basedOn) — covers custom styles built on
#           top of a built-in heading style via Word's "Style based on" setting.
#           The outline level lives in the ancestor style definition (styles.xml),
#           not on each paragraph, so strategies (a) and (b) miss these.
#
#   2. Document order preservation
#      python-docx exposes paragraphs and tables as two separate flat lists,
#      which loses the interleaved order when a table sits between two
#      paragraphs. We recover the true order by walking the raw XML body
#      children directly and mapping each one back to the python-docx object.
#
#   3. Empty heading filtering
#      Blank headings (caused by pressing Enter while a Heading style is active)
#      are filtered out and never become Section objects. This prevents phantom
#      sections appearing in the mapping table.
#
#   4. Image detection
#      Paragraphs that contain inline images are tagged as BlockType.IMAGE so
#      the migrator can handle them differently (green border, centering).
# =============================================================================

import re
import sys
from pathlib import Path
from typing import Optional

# Add the project root to the Python path so sibling packages resolve correctly
# when this file is run directly or imported from any subdirectory.
# This is necessary because Python only adds the script's own directory to
# sys.path by default — not the parent project root.
sys.path.insert(0, str(Path(__file__).parent.parent))

from docx import Document
from docx.oxml.ns import qn

from models.sections import BlockType, ContentBlock, Section


# The maximum heading level we recognize. Word supports Heading 1–9 but
# documents rarely use beyond level 6. Capping at 6 matches common practice.
MAX_HEADING_LEVEL = 6

# The prefix used by Word for all built-in heading style names.
# "Heading 1", "Heading 2", etc. all start with this prefix (case-insensitive).
HEADING_STYLE_PREFIX = "heading"


class DocumentParser:
    """
    Parses a .docx file into a flat ordered list of Section objects.

    Each Section contains its heading metadata and all ContentBlocks
    (paragraphs, tables, images) that appear under that heading in the document.

    Usage:
        parser   = DocumentParser("path/to/file.docx")
        sections = parser.parse()
        document = parser.document   # Access the python-docx Document if needed
    """

    def __init__(self, filepath: str):
        # Store as a Path object for reliable cross-platform path handling
        self.filepath = Path(filepath)

        # Populated after parse() is called
        self.sections: list[Section] = []

        # The underlying python-docx Document — exposed via the .document
        # property so the migrator can access it after parsing
        self._document = None

    # =========================================================================
    # Public API
    # =========================================================================

    def parse(self) -> list[Section]:
        """
        Open the document and extract all non-empty sections in document order.

        Returns a flat list of Section objects. Sections with blank titles
        (empty headings) are silently filtered out.
        """
        self._document = Document(str(self.filepath))

        # Get all body elements (paragraphs and tables) in their true
        # interleaved document order — see _get_body_elements_in_order()
        ordered_body_elements = self._get_body_elements_in_order()

        sections: list[Section]       = []
        current_section: Optional[Section] = None

        # paragraph_index tracks position in the flat document.paragraphs list.
        # Only increments for paragraph elements (not tables), matching the
        # indexing used by python-docx's document.paragraphs list.
        paragraph_index = 0

        for element_type, element_object in ordered_body_elements:

            if element_type == "paragraph":
                paragraph   = element_object
                heading_level = self._get_heading_level(paragraph)

                if heading_level is not None:
                    # ── This paragraph is a heading ───────────────────────────
                    title = paragraph.text.strip()

                    # Skip blank headings — these are created when a user presses
                    # Enter while a Heading style is active, leaving a stray
                    # heading paragraph with no text content
                    if title:
                        current_section = Section(
                            title=title,
                            heading_level=heading_level,
                            paragraph_index=paragraph_index,
                        )
                        sections.append(current_section)
                    # If title is blank, current_section stays unchanged so
                    # content after the blank heading still attaches to the
                    # last real section

                else:
                    # ── Regular content paragraph ──────────────────────────────
                    # Only collect paragraphs that belong to a section
                    if current_section is not None:
                        block_type = (
                            BlockType.IMAGE
                            if self._paragraph_contains_image(paragraph)
                            else BlockType.PARAGRAPH
                        )
                        current_section.content_blocks.append(
                            ContentBlock(
                                block_type=block_type,
                                raw_object=paragraph,
                            )
                        )

                paragraph_index += 1

            elif element_type == "table":
                # Tables attach to the current section as TABLE blocks
                if current_section is not None:
                    current_section.content_blocks.append(
                        ContentBlock(
                            block_type=BlockType.TABLE,
                            raw_object=element_object,
                        )
                    )

        self.sections = sections
        return sections

    @property
    def document(self):
        """
        The underlying python-docx Document object.
        Only available after parse() has been called.
        Used by ContentMigrator which needs direct document access.
        """
        return self._document

    # =========================================================================
    # Private helpers
    # =========================================================================

    def _get_body_elements_in_order(self) -> list[tuple[str, object]]:
        """
        Return all paragraphs and tables from the document body in the exact
        order they appear in the document.

        Why this is necessary:
          python-docx provides document.paragraphs and document.tables as two
          separate flat lists. Using these lists directly loses the interleaved
          order — for example, if a table appears between two paragraphs, you
          cannot determine its position relative to those paragraphs from the
          lists alone.

          We recover the true order by iterating the raw XML children of the
          document body element (<w:body>), matching each child to its
          corresponding python-docx object using parallel iteration.

        Returns a list of (type_string, python_docx_object) tuples where
        type_string is either "paragraph" or "table".
        """
        ordered_elements = []

        # Parallel iterators for the python-docx object lists
        paragraph_iterator = iter(self._document.paragraphs)
        table_iterator     = iter(self._document.tables)

        # XML tag names for paragraph and table elements in the Word namespace
        paragraph_xml_tag = qn("w:p")    # <w:p> — paragraph element
        table_xml_tag     = qn("w:tbl")  # <w:tbl> — table element

        for xml_child in self._document.element.body:
            if xml_child.tag == paragraph_xml_tag:
                try:
                    ordered_elements.append(("paragraph", next(paragraph_iterator)))
                except StopIteration:
                    pass   # Should not occur in a well-formed document

            elif xml_child.tag == table_xml_tag:
                try:
                    ordered_elements.append(("table", next(table_iterator)))
                except StopIteration:
                    pass

        return ordered_elements

    def _get_heading_level(self, paragraph) -> Optional[int]:
        """
        Determine if a paragraph is a heading and return its level (1–6).
        Returns None if the paragraph is not a heading.

        Three detection strategies are used in priority order:

        Strategy 1 — Style name check:
          Word's built-in heading styles are named "Heading 1", "Heading 2",
          etc. We check if the paragraph style name starts with "heading"
          (case-insensitive) and extract the trailing number.

        Strategy 2 — Paragraph-level XML outline level:
          Some documents use custom style names (not "Heading N") but still
          set the paragraph's outline level directly in the paragraph XML.
          Word uses the outline level to include paragraphs in the Navigation
          pane and Table of Contents. Stored as:
            <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
          where val="0" means Heading 1, val="1" means Heading 2, etc.
          (0-indexed in XML, so we add 1 to convert to 1-indexed level)

        Strategy 3 — Style inheritance chain:
          Custom styles can be built on top of built-in heading styles via
          Word's "Style based on" setting (stored as <w:basedOn> in styles.xml).
          When this is done correctly the outline level lives in the ancestor
          style definition, not on each individual paragraph element — so
          Strategy 2 misses it. We walk the basedOn chain looking for either
          a "Heading N" ancestor name or an outlineLvl defined in the style
          definition itself.
        """
        style_name = paragraph.style.name.lower() if paragraph.style else ""

        # Strategy 1: Check the paragraph style name
        if style_name.startswith(HEADING_STYLE_PREFIX):
            match = re.search(r"(\d+)$", style_name)
            if match:
                return int(match.group(1))

        # Strategy 2: Check the XML outline level on the paragraph element
        paragraph_properties = paragraph._element.find(qn("w:pPr"))
        if paragraph_properties is not None:
            outline_level_element = paragraph_properties.find(qn("w:outlineLvl"))
            if outline_level_element is not None:
                raw_value = outline_level_element.get(qn("w:val"))
                if raw_value is not None:
                    # XML uses 0-indexed levels; add 1 for 1-indexed heading level
                    heading_level = int(raw_value) + 1
                    if 1 <= heading_level <= MAX_HEADING_LEVEL:
                        return heading_level

        # Strategy 3: Walk the style inheritance chain
        return self._heading_level_from_style_chain(paragraph.style)

    def _heading_level_from_style_chain(self, style) -> Optional[int]:
        """
        Walk the style's basedOn inheritance chain looking for evidence that
        this style is semantically a heading.

        At each step two things are checked:
          1. Whether the ancestor style name starts with "heading" — catches
             custom styles that inherit from a built-in Heading N style and
             whose name does not itself start with "heading".
          2. Whether the ancestor style definition in styles.xml sets
             <w:outlineLvl> — catches styles whose level is defined once on
             the style rather than repeated on every paragraph element.

        A visited-ID set guards against pathological circular basedOn chains
        (not valid OOXML, but defensively handled). The loop terminates
        naturally when base_style returns None (top of the chain).
        """
        if style is None:
            return None

        visited: set[int] = set()
        current = style

        while current is not None:
            if id(current) in visited:
                break
            visited.add(id(current))

            # Check the ancestor style name
            ancestor_name = current.name.lower() if current.name else ""
            if ancestor_name.startswith(HEADING_STYLE_PREFIX):
                match = re.search(r"(\d+)$", ancestor_name)
                if match:
                    return int(match.group(1))

            # Check outlineLvl in the style definition (styles.xml).
            # This is separate from the paragraph-element check in Strategy 2 —
            # the level is defined once on the style, not on every paragraph.
            try:
                style_pPr = current.element.find(qn("w:pPr"))
                if style_pPr is not None:
                    outline_elem = style_pPr.find(qn("w:outlineLvl"))
                    if outline_elem is not None:
                        raw_val = outline_elem.get(qn("w:val"))
                        if raw_val is not None:
                            level = int(raw_val) + 1
                            if 1 <= level <= MAX_HEADING_LEVEL:
                                return level
            except (AttributeError, ValueError):
                pass

            current = current.base_style

        return None   # Not a heading by any strategy

    def _paragraph_contains_image(self, paragraph) -> bool:
        """
        Return True if this paragraph contains at least one inline image.

        Word stores inline images inside <w:drawing> elements within runs.
        We check each run's XML element for a <w:drawing> child to detect
        whether the paragraph is an image paragraph vs a text paragraph.
        """
        drawing_tag = qn("w:drawing")
        for run in paragraph.runs:
            if run._element.find(drawing_tag) is not None:
                return True
        return False

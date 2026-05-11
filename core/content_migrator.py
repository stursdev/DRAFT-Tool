# =============================================================================
# core/content_migrator.py
#
# Writes migrated content from the source document into a copy of the
# destination template, producing the final output .docx file.
#
# Migration modes (set independently per row in the mapping table):
#
#   APPEND  — migrated content is inserted AFTER  any existing boilerplate
#   PREPEND — migrated content is inserted BEFORE any existing boilerplate
#   REPLACE — existing boilerplate is deleted; migrated content replaces it
#
# All migrated content (paragraphs, table text, images) is colored green
# (RGB #00B050) so the user can identify it against the destination boilerplate.
#
# Replace exclusivity rule:
#   If two source sections are both mapped to the same destination, only the
#   first one (in table order) may use REPLACE mode. The second is forced to
#   APPEND to prevent the first section's migrated content from being overwritten.
#   This rule is enforced in the mapping table UI before migration runs, but we
#   also enforce it here defensively.
#
# Design rules:
#   - Destination heading title and level are always preserved in the output
#   - Unmapped and Skipped rows produce no output (silently excluded)
#   - Multiple sources mapped to the same destination stack in table order
#   - Sub-sections always migrate with their parent (no option to separate)
# =============================================================================

import copy
import io
import os
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from docx import Document
from docx.oxml     import OxmlElement
from docx.oxml.ns  import qn
from docx.shared   import Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from models.sections import (
    ContentBlock,
    BlockType,
    MappingResult,
    MappingStatus,
    MigrationMode,
    Section,
)


# The green color applied to all migrated content.
# RGB #00B050 is Microsoft Word's standard "Green" palette color.
MIGRATED_CONTENT_COLOR     = RGBColor(0x00, 0xB0, 0x50)
MIGRATED_CONTENT_COLOR_HEX = "00B050"   # Hex string for direct XML attribute use


# =============================================================================
# MigrationReport
#
# Returned by ContentMigrator.run() to summarise what happened during the
# migration. Displayed in the completion dialog and passed to the report
# exporter and boilerplate detector.
# =============================================================================
@dataclass
class MigrationReport:
    # Absolute path to the output .docx file that was written
    output_path: str

    # Human-readable "source title → destination title" strings for each
    # section that was successfully migrated
    migrated_sections: list[str] = field(default_factory=list)

    # Titles of source sections the user explicitly skipped
    skipped_sections: list[str] = field(default_factory=list)

    # Titles of source sections that had no destination assigned
    unmapped_sections: list[str] = field(default_factory=list)

    # Error descriptions for any section that failed during migration
    errors: list[str] = field(default_factory=list)

    @property
    def total_migrated(self) -> int:
        return len(self.migrated_sections)


# =============================================================================
# SectionBoundary
#
# Holds stable XML element references for the boundaries of one section in
# the working document. We use XML element references rather than paragraph
# index numbers because inserting content into the document shifts all
# subsequent paragraph indices, making index-based references unreliable
# after the first insertion.
# =============================================================================
@dataclass
class SectionBoundary:
    # The python-docx Paragraph object for the section heading
    heading_paragraph: object

    # The lxml XML element of the section heading
    # Used for direct XML manipulation (inserting/removing body children)
    heading_element: object

    # The lxml XML element of the next heading of equal or higher importance,
    # or None if this is the last section in the document.
    # Content between heading_element and next_boundary_element belongs to
    # this section.
    next_boundary_element: Optional[object]

    # The Word heading level (1–6) of this section
    heading_level: int


# =============================================================================
# ContentMigrator
# =============================================================================
class ContentMigrator:
    """
    Executes the document migration based on the finalized mapping decisions.

    Usage:
        migrator = ContentMigrator(
            mapping_results = results,
            source_document = source_parser.document,
            dest_document   = dest_parser.document,
            dest_sections   = dest_parser.sections,
            output_path     = "output/migrated.docx",
        )
        report = migrator.run()
    """

    def __init__(
        self,
        mapping_results: list[MappingResult],
        source_document,                        # python-docx Document (source)
        dest_document,                          # python-docx Document (template)
        dest_sections:   list[Section],
        output_path:     str,
    ):
        self.mapping_results = mapping_results
        self.source_document = source_document
        self.dest_document   = dest_document
        self.dest_sections   = dest_sections
        self.output_path     = Path(output_path)

        # The working copy of the destination template.
        # We never modify the original template — all changes go to this copy.
        self._working_document = None

        # Cache of source-document numId → working-document numId mappings.
        # Built lazily as list paragraphs are encountered during migration.
        # Each unique source numId is registered into the working document's
        # numbering.xml exactly once; subsequent paragraphs sharing that numId
        # look up the already-assigned destination numId from this dict.
        self._source_to_dest_num_id: dict[int, int] = {}

        self.report = MigrationReport(output_path=str(output_path))

    # =========================================================================
    # Public API
    # =========================================================================

    def run(self) -> MigrationReport:
        """
        Execute the full migration and write the output .docx file.

        Steps:
          1. Create a working copy of the destination template
          2. Group mapping results by destination section
          3. Index section boundaries in the working document
          4. Process each destination section
          5. Save the output file

        Returns a MigrationReport describing what was migrated, skipped,
        unmapped, or errored.
        """
        self._working_document = self._create_working_copy()

        # Group source sections by their destination title so we can process
        # all sources for one destination together in table order
        sources_per_destination = self._group_sources_by_destination()

        # Build a stable index of section boundaries using XML element references
        section_boundary_index = self._build_section_boundary_index()

        # Process each destination section that has at least one source mapped to it
        for dest_title, source_pairs in sources_per_destination.items():
            try:
                self._process_destination_section(
                    dest_title,
                    source_pairs,
                    section_boundary_index,
                )
                for source_section, _ in source_pairs:
                    self.report.migrated_sections.append(
                        f"{source_section.title} → {dest_title}"
                    )
            except Exception as error:
                self.report.errors.append(
                    f"Error migrating into '{dest_title}': {error}"
                )

        # Write the completed working document to the output path
        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        self._working_document.save(str(self.output_path))

        return self.report

    # =========================================================================
    # Setup helpers
    # =========================================================================

    def _create_working_copy(self):
        """
        Create and return an in-memory copy of the destination template.

        We save the template to a temporary file and reload it so that all
        modifications go to the copy, never touching the original template.
        The temporary file is deleted immediately after reloading.
        """
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_file:
            temp_path = temp_file.name

        self.dest_document.save(temp_path)
        working_copy = Document(temp_path)
        os.unlink(temp_path)

        return working_copy

    def _group_sources_by_destination(self) -> dict[str, list[tuple]]:
        """
        Group mapping results by destination section title.

        Returns a dict where:
          key   = destination section title string
          value = ordered list of (source_section, migration_mode) tuples

        Unmapped and Skipped rows are recorded in the report and excluded
        from the returned dict — they produce no output.

        The order within each list matches the mapping table row order
        (top to bottom), which is the order content will be inserted.
        This matters for the Replace exclusivity rule.
        """
        sources_per_destination: dict[str, list[tuple]] = {}

        for result in self.mapping_results:

            if result.status == MappingStatus.UNMAPPED or result.dest_section is None:
                self.report.unmapped_sections.append(result.source_section.title)
                continue

            if result.status == MappingStatus.SKIPPED:
                self.report.skipped_sections.append(result.source_section.title)
                continue

            destination_title = result.dest_section.title
            migration_mode    = result.migration_mode or MigrationMode.APPEND

            if destination_title not in sources_per_destination:
                sources_per_destination[destination_title] = []

            sources_per_destination[destination_title].append(
                (result.source_section, migration_mode)
            )

        return sources_per_destination

    def _build_section_boundary_index(self) -> dict[str, SectionBoundary]:
        """
        Walk the working document and build a SectionBoundary for every heading.

        Returns a dict keyed by heading title string. Each value is a
        SectionBoundary with stable XML element references that remain valid
        even as content is inserted into or removed from the document.

        Why XML element references instead of paragraph indices:
          Inserting a paragraph into the document shifts all subsequent
          paragraph indices. An index that was correct before an insertion
          points to the wrong paragraph after it. XML element references
          (lxml Element objects) are stable identity references — they point
          to the same node regardless of insertions elsewhere in the document.

        If the same heading title appears more than once, only the first
        occurrence is indexed (rare in practice for well-formed documents).
        """
        import re as _re
        boundary_index: dict[str, SectionBoundary] = {}

        # Collect all heading paragraphs with their levels
        heading_entries: list[tuple] = []   # (paragraph, heading_level)

        for paragraph in self._working_document.paragraphs:
            style_name = paragraph.style.name.lower() if paragraph.style else ""
            if style_name.startswith("heading"):
                match = _re.search(r"(\d+)$", style_name)
                if match and paragraph.text.strip():
                    heading_level = int(match.group(1))
                    heading_entries.append((paragraph, heading_level))

        # For each heading, find the XML element where its section ends.
        # A section ends at the next heading of equal or higher importance
        # (lower or equal heading level number).
        for position, (paragraph, heading_level) in enumerate(heading_entries):

            next_boundary_element = None
            for later_position in range(position + 1, len(heading_entries)):
                later_paragraph, later_level = heading_entries[later_position]
                if later_level <= heading_level:
                    next_boundary_element = later_paragraph._element
                    break

            title = paragraph.text.strip()
            if title not in boundary_index:
                boundary_index[title] = SectionBoundary(
                    heading_paragraph=paragraph,
                    heading_element=paragraph._element,
                    next_boundary_element=next_boundary_element,
                    heading_level=heading_level,
                )

        return boundary_index

    # =========================================================================
    # Section processing
    # =========================================================================

    def _process_destination_section(
        self,
        destination_title:     str,
        source_pairs:          list[tuple],
        section_boundary_index: dict[str, SectionBoundary],
    ):
        """
        Migrate content from one or more source sections into a single
        destination section in the working document.

        Replace exclusivity:
          Only the first source in source_pairs that specifies REPLACE may
          actually replace the boilerplate. All subsequent sources are treated
          as APPEND regardless of their mode, preventing the first source's
          migrated content from being overwritten.
        """
        if destination_title not in section_boundary_index:
            raise ValueError(
                f"Section '{destination_title}' not found in the working document."
            )

        boundary = section_boundary_index[destination_title]

        # Track whether the boilerplate has already been replaced so we can
        # enforce the Replace exclusivity rule for subsequent sources
        boilerplate_has_been_replaced = False

        for source_section, requested_mode in source_pairs:

            # Resolve the effective mode — Replace is blocked if boilerplate
            # was already replaced by an earlier source in this group
            effective_mode = self._resolve_effective_mode(
                requested_mode,
                boilerplate_has_been_replaced,
            )

            self._apply_migration_mode(
                effective_mode,
                boundary,
                source_section,
            )

            if effective_mode == MigrationMode.REPLACE:
                boilerplate_has_been_replaced = True

    def _resolve_effective_mode(
        self,
        requested_mode:               MigrationMode,
        boilerplate_has_been_replaced: bool,
    ) -> MigrationMode:
        """
        Return the mode that will actually be applied to the current source.

        If the user requested REPLACE but the boilerplate was already replaced
        by an earlier source, we fall back to APPEND to prevent data loss.
        All other cases use the requested mode unchanged.
        """
        if requested_mode == MigrationMode.REPLACE and boilerplate_has_been_replaced:
            return MigrationMode.APPEND
        return requested_mode

    def _apply_migration_mode(
        self,
        mode:           MigrationMode,
        boundary:       SectionBoundary,
        source_section: Section,
    ):
        """
        Insert the source section's content into the working document
        according to the specified migration mode.

        REPLACE: Clear the section's existing content, then insert source content
                 immediately after the heading.
        PREPEND: Insert source content immediately after the heading, before
                 any existing boilerplate.
        APPEND:  Insert source content after all existing content in the section.
        """
        if mode == MigrationMode.REPLACE:
            self._clear_section_content(boundary)
            self._insert_content_blocks(
                blocks=source_section.content_blocks,
                insert_after_paragraph=boundary.heading_paragraph,
            )

        elif mode == MigrationMode.PREPEND:
            # Insert immediately after the heading so it appears before boilerplate
            self._insert_content_blocks(
                blocks=source_section.content_blocks,
                insert_after_paragraph=boundary.heading_paragraph,
            )

        elif mode == MigrationMode.APPEND:
            last_paragraph = self._find_last_paragraph_in_section(boundary)
            self._insert_content_blocks(
                blocks=source_section.content_blocks,
                insert_after_paragraph=last_paragraph,
            )

    # =========================================================================
    # XML body manipulation helpers
    # =========================================================================

    def _clear_section_content(self, boundary: SectionBoundary):
        """
        Remove all body child elements that belong to this section.

        Removes everything between the heading element (exclusive) and the
        next boundary element (exclusive). Works at the XML level so it
        handles both paragraphs and tables within the section.
        """
        document_body  = self._working_document.element.body
        elements_to_remove = []
        collecting     = False

        for xml_child in document_body:
            if xml_child is boundary.heading_element:
                collecting = True
                continue
            if boundary.next_boundary_element is not None:
                if xml_child is boundary.next_boundary_element:
                    break
            if collecting:
                elements_to_remove.append(xml_child)

        for element in elements_to_remove:
            document_body.remove(element)

    def _find_last_paragraph_in_section(self, boundary: SectionBoundary):
        """
        Return the last paragraph in this section for use as the APPEND
        insertion point.

        Walks the document body from the heading element forward, collecting
        all elements in this section, and returns the last one. If the section
        is empty (heading only), returns the heading paragraph itself.
        """
        document_body  = self._working_document.element.body
        last_element   = boundary.heading_element
        collecting     = False

        for xml_child in document_body:
            if xml_child is boundary.heading_element:
                collecting = True
                continue
            if boundary.next_boundary_element is not None:
                if xml_child is boundary.next_boundary_element:
                    break
            if collecting:
                last_element = xml_child

        # Convert the XML element back to a python-docx Paragraph if possible
        last_paragraph = self._find_paragraph_by_element(last_element)

        # If the last element is not a paragraph (e.g. it is a table),
        # add a spacer paragraph after it to use as the insertion point
        if last_paragraph is None:
            spacer_paragraph = self._working_document.add_paragraph("")
            last_element.addnext(spacer_paragraph._element)
            return spacer_paragraph

        return last_paragraph

    def _find_paragraph_by_element(self, target_element) -> Optional[object]:
        """
        Find the python-docx Paragraph object whose underlying lxml element
        matches the given target element by object identity.

        Returns None if no match is found (e.g. the element is a table).
        """
        for paragraph in self._working_document.paragraphs:
            if paragraph._element is target_element:
                return paragraph
        return None

    # =========================================================================
    # Content block insertion
    # =========================================================================

    def _insert_content_blocks(
        self,
        blocks:                  list[ContentBlock],
        insert_after_paragraph,
    ):
        """
        Insert a list of content blocks into the working document after the
        given paragraph. Handles PARAGRAPH, TABLE, and IMAGE block types.

        All content is rendered in green (MIGRATED_CONTENT_COLOR).

        Uses addnext() on the XML element level to insert each new element
        immediately after the current last element. Each insertion advances
        the insertion point so blocks end up in their original document order.
        """
        current_last_paragraph = insert_after_paragraph

        for block in blocks:
            if block.block_type == BlockType.PARAGRAPH:
                new_paragraph = self._copy_paragraph_with_green_text(block.raw_object)
                current_last_paragraph._element.addnext(new_paragraph._element)
                current_last_paragraph = new_paragraph

            elif block.block_type == BlockType.TABLE:
                copied_table_element = self._copy_table_with_green_text_and_borders(
                    block.raw_object
                )
                current_last_paragraph._element.addnext(copied_table_element)
                # Word requires a paragraph after every table to remain valid
                spacer = self._working_document.add_paragraph("")
                copied_table_element.addnext(spacer._element)
                current_last_paragraph = spacer

            elif block.block_type == BlockType.IMAGE:
                # An image paragraph may produce two paragraphs:
                # one for the image (with green border) and one for any caption
                # text (green text, no border). We insert them in sequence.
                image_paragraphs = self._copy_image_paragraph_with_border(
                    block.raw_object
                )
                for image_paragraph in image_paragraphs:
                    if image_paragraph is not None:
                        current_last_paragraph._element.addnext(
                            image_paragraph._element
                        )
                        current_last_paragraph = image_paragraph

    # =========================================================================
    # Paragraph copying
    # =========================================================================

    def _copy_paragraph_with_green_text(self, source_paragraph) -> object:
        """
        Create a new paragraph in the working document that mirrors the source.

        Copies:
          - Paragraph style (if it exists in the destination document)
          - Paragraph-level formatting (alignment, spacing, indentation)
          - Each run's text and character formatting (bold, italic, underline,
            font name, font size)

        All text runs are colored green (MIGRATED_CONTENT_COLOR) to mark
        this content as migrated rather than original boilerplate.
        """
        new_paragraph = self._working_document.add_paragraph()

        # Copy the paragraph style. If the style name does not exist in the
        # destination document's style sheet, skip silently — the default
        # paragraph style will be used instead.
        try:
            new_paragraph.style = self._working_document.styles[
                source_paragraph.style.name
            ]
        except (KeyError, AttributeError):
            pass

        # Copy paragraph-level formatting properties
        source_format = source_paragraph.paragraph_format
        dest_format   = new_paragraph.paragraph_format

        if source_format.alignment is not None:
            dest_format.alignment = source_format.alignment
        if source_format.space_before is not None:
            dest_format.space_before = source_format.space_before
        if source_format.space_after is not None:
            dest_format.space_after = source_format.space_after

        # Only copy left_indent and first_line_indent for list paragraphs.
        # List paragraphs need these values to pair correctly with the
        # abstractNum indentation we copy via _copy_list_numbering().
        # For non-list paragraphs we skip them entirely, letting the
        # destination style supply the indentation — this prevents source
        # document margin offsets from making migrated text appear pushed in.
        source_pPr = source_paragraph._element.find(qn("w:pPr"))
        is_list_paragraph = (
            source_pPr is not None
            and source_pPr.find(qn("w:numPr")) is not None
        )

        if is_list_paragraph:
            if source_format.left_indent is not None:
                dest_format.left_indent = source_format.left_indent
            # first_line_indent is negative for hanging-indent list items (the
            # number or bullet hangs to the left of the wrapped body text).
            # Without this, wrapped list items collapse to the left margin.
            if source_format.first_line_indent is not None:
                dest_format.first_line_indent = source_format.first_line_indent

        # ── List numbering (numPr) ─────────────────────────────────────────────
        # Word list items carry a <w:numPr> element in their <w:pPr> block.
        # numPr contains:
        #   <w:ilvl>  — the list indent level (0 = top, 1 = first sub-level, …)
        #   <w:numId> — an ID that points to a numbering definition in
        #               numbering.xml, which holds the bullet character,
        #               number format (1/a/i/•), and per-level indentation.
        #
        # Without copying numPr, Word treats the paragraph as a plain (non-list)
        # paragraph — no bullet, no number, and the indent looks unexplained.
        #
        # Because numId values are document-specific, we register each source
        # numId into the working document's numbering.xml (copying the underlying
        # abstractNum definition) and get back a working-document numId to use.
        self._copy_list_numbering(source_paragraph, new_paragraph)

        # Copy each run individually to preserve inline character formatting
        for source_run in source_paragraph.runs:
            new_run = new_paragraph.add_run(source_run.text)

            new_run.bold      = source_run.bold
            new_run.italic    = source_run.italic
            new_run.underline = source_run.underline

            if source_run.font.name:
                new_run.font.name = source_run.font.name
            if source_run.font.size:
                new_run.font.size = source_run.font.size

            # Override the text color to green regardless of the source color
            new_run.font.color.rgb = MIGRATED_CONTENT_COLOR

        return new_paragraph

    # =========================================================================
    # List numbering helpers
    #
    # Word's list system is built on two XML parts:
    #
    #   numbering.xml / <w:numbering>
    #     <w:abstractNum w:abstractNumId="0">   — defines the format for each
    #       <w:lvl w:ilvl="0">                    level: bullet char or number
    #         <w:numFmt w:val="bullet"/>           style, indent values, font.
    #         <w:lvlText w:val="•"/>             One abstractNum can have up to
    #         <w:ind w:left="720" …/>             9 levels (ilvl 0–8).
    #       </w:lvl>
    #       …
    #     </w:abstractNum>
    #
    #     <w:num w:numId="1">                   — a concrete instance that
    #       <w:abstractNumId w:val="0"/>          references one abstractNum.
    #     </w:num>                                Multiple <w:num> elements can
    #                                             share the same abstractNum,
    #   Each list paragraph in the body:          allowing independent restart
    #     <w:pPr>                                 tracking per list.
    #       <w:numPr>
    #         <w:ilvl  w:val="0"/>  ← list level
    #         <w:numId w:val="1"/>  ← references a <w:num>
    #       </w:numPr>
    #     </w:pPr>
    #
    # The migration problem:
    #   numId values are document-local integers. numId=1 in the source document
    #   may not exist in the destination, or may point to a completely different
    #   list format. We must:
    #     1. Copy the <w:abstractNum> definition into the destination's numbering.xml
    #        with a fresh abstractNumId that doesn't clash with existing ones.
    #     2. Add a new <w:num> in the destination referencing that abstractNum,
    #        with a fresh numId.
    #     3. Rewrite the paragraph's <w:numPr><w:numId> to use the new dest numId.
    #
    #   Each unique source numId is registered exactly once (lazy + cached in
    #   self._source_to_dest_num_id). Paragraphs that share a numId (i.e. items
    #   in the same list) are all remapped to the same destination numId, so they
    #   continue to belong to one continuous list in the output.
    # =========================================================================

    def _copy_list_numbering(self, source_paragraph, dest_paragraph):
        """
        If source_paragraph is a list item, copy its numbering definition into
        the working document and attach a correctly remapped <w:numPr> to
        dest_paragraph.

        Does nothing and returns silently if:
          - The source paragraph has no <w:numPr> (it is not a list item).
          - The source document has no numbering.xml part.
          - Any XML access fails (bad document structure, unsupported version).
        """
        source_pPr = source_paragraph._element.find(qn("w:pPr"))
        if source_pPr is None:
            return

        source_numPr = source_pPr.find(qn("w:numPr"))
        if source_numPr is None:
            return   # Paragraph is not a list item

        source_ilvl_elem  = source_numPr.find(qn("w:ilvl"))
        source_numId_elem = source_numPr.find(qn("w:numId"))
        if source_ilvl_elem is None or source_numId_elem is None:
            return

        source_num_id = int(source_numId_elem.get(qn("w:val"), 0))
        list_level    = source_ilvl_elem.get(qn("w:val"), "0")

        # Register (or look up) the source numId in the working document
        dest_num_id = self._register_source_num_id(source_num_id)
        if dest_num_id is None:
            return   # Registration failed — leave paragraph without numPr

        # Write the remapped <w:numPr> onto the destination paragraph
        dest_pPr = dest_paragraph._element.find(qn("w:pPr"))
        if dest_pPr is None:
            dest_pPr = OxmlElement("w:pPr")
            dest_paragraph._element.insert(0, dest_pPr)

        # Remove any stale numPr that may have been set by the style copy
        for stale_numPr in dest_pPr.findall(qn("w:numPr")):
            dest_pPr.remove(stale_numPr)

        new_numPr = OxmlElement("w:numPr")

        ilvl_elem = OxmlElement("w:ilvl")
        ilvl_elem.set(qn("w:val"), list_level)
        new_numPr.append(ilvl_elem)

        numId_elem = OxmlElement("w:numId")
        numId_elem.set(qn("w:val"), str(dest_num_id))
        new_numPr.append(numId_elem)

        dest_pPr.append(new_numPr)

    def _register_source_num_id(self, source_num_id: int) -> Optional[int]:
        """
        Ensure the numbering definition for source_num_id exists in the working
        document, and return the working document's numId for that definition.

        On first call for a given source_num_id:
          1. Locate the <w:num> element in the source document's numbering.xml.
          2. Locate the <w:abstractNum> it references.
          3. Deep-copy the <w:abstractNum> into the working document's
             numbering.xml with a fresh, non-conflicting abstractNumId.
          4. Add a new <w:num> element referencing the copied abstractNum.
          5. Cache and return the new numId.

        On subsequent calls for the same source_num_id, return the cached value
        immediately so each unique numbering definition is registered only once.
        Returns None if registration fails (missing parts, malformed XML).
        """
        # Return cached mapping if already registered
        if source_num_id in self._source_to_dest_num_id:
            return self._source_to_dest_num_id[source_num_id]

        source_numbering_elem = self._get_numbering_element(self.source_document)
        if source_numbering_elem is None:
            return None

        dest_numbering_elem = self._get_or_create_numbering_element()
        if dest_numbering_elem is None:
            return None

        try:
            return self._copy_num_definition_to_dest(
                source_numbering_elem,
                dest_numbering_elem,
                source_num_id,
            )
        except Exception:
            return None   # Malformed numbering XML — skip gracefully

    def _get_numbering_element(self, document):
        """
        Return the root <w:numbering> lxml element from the given document,
        or None if the document has no numbering.xml part.
        """
        try:
            from docx.opc.constants import RELATIONSHIP_TYPE as RT
            numbering_part = document.part.part_related_by(RT.NUMBERING)
            return numbering_part.element
        except (KeyError, AttributeError, Exception):
            return None

    def _get_or_create_numbering_element(self):
        """
        Return the root <w:numbering> lxml element from the working document,
        creating the numbering part from scratch if it does not yet exist.

        The destination template may have no lists at all (and therefore no
        numbering.xml). We create a minimal numbering part so that we can
        register definitions copied from the source document.
        """
        from docx.opc.constants import RELATIONSHIP_TYPE as RT

        # Try to access the existing numbering part first
        try:
            numbering_part = self._working_document.part.part_related_by(RT.NUMBERING)
            return numbering_part.element
        except (KeyError, AttributeError):
            pass

        # No numbering part exists — create a minimal one using python-docx internals
        try:
            from docx.parts.numbering import NumberingPart
            numbering_part = NumberingPart.new()
            self._working_document.part.relate_to(numbering_part, RT.NUMBERING)
            return numbering_part.element
        except Exception:
            return None

    def _copy_num_definition_to_dest(
        self,
        source_numbering_elem,
        dest_numbering_elem,
        source_num_id: int,
    ) -> Optional[int]:
        """
        Copy one <w:num> definition (and its referenced <w:abstractNum>) from
        the source numbering element into the destination numbering element.

        Assigns fresh IDs to avoid conflicts with definitions already present
        in the destination document. Updates self._source_to_dest_num_id with
        the new mapping and returns the destination numId.

        Returns None if the source does not contain the requested numId.
        """
        # ── Locate the source <w:num> for this numId ──────────────────────────
        source_num_elem = None
        for num_elem in source_numbering_elem.findall(qn("w:num")):
            if int(num_elem.get(qn("w:numId"), -1)) == source_num_id:
                source_num_elem = num_elem
                break

        if source_num_elem is None:
            return None

        # ── Locate the source <w:abstractNum> it references ───────────────────
        abstract_num_id_ref = source_num_elem.find(qn("w:abstractNumId"))
        if abstract_num_id_ref is None:
            return None

        source_abstract_num_id = int(abstract_num_id_ref.get(qn("w:val"), -1))

        source_abstract_num_elem = None
        for abs_elem in source_numbering_elem.findall(qn("w:abstractNum")):
            if int(abs_elem.get(qn("w:abstractNumId"), -1)) == source_abstract_num_id:
                source_abstract_num_elem = abs_elem
                break

        if source_abstract_num_elem is None:
            return None

        # ── Calculate next available IDs in the destination ───────────────────
        existing_abstract_ids = [
            int(e.get(qn("w:abstractNumId"), 0))
            for e in dest_numbering_elem.findall(qn("w:abstractNum"))
        ]
        existing_num_ids = [
            int(e.get(qn("w:numId"), 0))
            for e in dest_numbering_elem.findall(qn("w:num"))
        ]

        new_abstract_num_id = (max(existing_abstract_ids) + 1) if existing_abstract_ids else 0
        new_num_id          = (max(existing_num_ids) + 1)      if existing_num_ids      else 1

        # ── Copy <w:abstractNum> with the new ID ──────────────────────────────
        copied_abstract_num = copy.deepcopy(source_abstract_num_elem)
        copied_abstract_num.set(qn("w:abstractNumId"), str(new_abstract_num_id))

        # Remove cross-document style links — these reference style names that
        # may not exist in the destination, which would cause Word to error.
        for style_link in copied_abstract_num.findall(qn("w:styleLink")):
            copied_abstract_num.remove(style_link)
        for num_style_link in copied_abstract_num.findall(qn("w:numStyleLink")):
            copied_abstract_num.remove(num_style_link)

        # OOXML schema requires all <w:abstractNum> elements to precede all
        # <w:num> elements in the numbering XML. Insert before the first <w:num>.
        first_num_in_dest = dest_numbering_elem.find(qn("w:num"))
        if first_num_in_dest is not None:
            first_num_in_dest.addprevious(copied_abstract_num)
        else:
            dest_numbering_elem.append(copied_abstract_num)

        # ── Copy <w:num> with updated IDs ─────────────────────────────────────
        copied_num = copy.deepcopy(source_num_elem)
        copied_num.set(qn("w:numId"), str(new_num_id))

        # Rewrite the abstractNumId reference to point to our freshly copied one
        abstract_ref_in_copy = copied_num.find(qn("w:abstractNumId"))
        if abstract_ref_in_copy is not None:
            abstract_ref_in_copy.set(qn("w:val"), str(new_abstract_num_id))

        dest_numbering_elem.append(copied_num)

        # Cache the mapping and return
        self._source_to_dest_num_id[source_num_id] = new_num_id
        return new_num_id

    # =========================================================================
    # Table copying
    # =========================================================================

    def _copy_table_with_green_text_and_borders(self, source_table) -> object:
        """
        Deep-copy a Word table and apply green text and explicit cell borders.

        Why deep copy at the XML level:
          python-docx does not provide a method to copy a table as a whole.
          Using lxml's copy.deepcopy() on the underlying XML element (<w:tbl>)
          is the most reliable approach — it preserves merged cells, column
          widths, cell shading, and all inline formatting.

        Green text:
          We walk all <w:r> (run) elements inside the copied table XML and
          set or replace the <w:color> element in each run's properties.

        Explicit borders:
          Word table borders are often defined in the document's style sheet
          rather than inline on each cell. When copying the table to a new
          document the style sheet may differ, making borders invisible.
          We apply borders directly to every cell's <w:tcBorders> element
          to guarantee they appear in the output regardless of styles.

        Returns the copied lxml XML element (not a python-docx Table object).
        We insert it directly into the body XML via addnext().
        """
        copied_table_element = copy.deepcopy(source_table._tbl)

        # ── Reset table indentation to zero ───────────────────────────────────
        # The deep copy preserves the source table's <w:tblInd> element verbatim.
        # If the source table was indented (e.g. it sat inside a list context or
        # was manually indented), that non-zero indent value carries into the
        # destination, making the table appear pushed in from the left margin.
        # We normalise it to 0 so the table aligns with the destination's margins.
        tbl_properties = copied_table_element.find(qn("w:tblPr"))
        if tbl_properties is not None:
            for existing_tbl_ind in tbl_properties.findall(qn("w:tblInd")):
                tbl_properties.remove(existing_tbl_ind)
            # Explicitly write tblInd=0 so Word doesn't inherit a style-based indent
            tbl_ind_element = OxmlElement("w:tblInd")
            tbl_ind_element.set(qn("w:w"),    "0")
            tbl_ind_element.set(qn("w:type"), "dxa")
            tbl_properties.append(tbl_ind_element)

        # ── Apply green text to every run inside the table ────────────────────
        for run_element in copied_table_element.iter(qn("w:r")):
            run_properties = run_element.find(qn("w:rPr"))
            if run_properties is None:
                run_properties = OxmlElement("w:rPr")
                run_element.insert(0, run_properties)

            # Remove any existing color element before adding our green one
            for existing_color in run_properties.findall(qn("w:color")):
                run_properties.remove(existing_color)

            green_color_element = OxmlElement("w:color")
            green_color_element.set(qn("w:val"), MIGRATED_CONTENT_COLOR_HEX)
            run_properties.append(green_color_element)

        # ── Apply explicit borders to every table cell ────────────────────────
        for cell_element in copied_table_element.iter(qn("w:tc")):
            cell_properties = cell_element.find(qn("w:tcPr"))
            if cell_properties is None:
                cell_properties = OxmlElement("w:tcPr")
                cell_element.insert(0, cell_properties)

            # Remove any existing border definition to avoid conflicts
            for existing_borders in cell_properties.findall(qn("w:tcBorders")):
                cell_properties.remove(existing_borders)

            # Create fresh border definitions for all four sides of the cell
            cell_borders = OxmlElement("w:tcBorders")
            for border_side in ("top", "left", "bottom", "right"):
                border_element = OxmlElement(f"w:{border_side}")
                border_element.set(qn("w:val"),   "single")   # Solid line style
                border_element.set(qn("w:sz"),    "4")        # 0.5pt thickness
                border_element.set(qn("w:space"), "0")        # No space between border and text
                border_element.set(qn("w:color"), "000000")   # Black border color
                cell_borders.append(border_element)

            cell_properties.append(cell_borders)

        return copied_table_element

    # =========================================================================
    # Image paragraph copying
    # =========================================================================

    def _copy_image_paragraph_with_border(self, source_paragraph) -> list:
        """
        Copy a paragraph containing inline images from the source document.

        Returns an ordered list of new paragraphs to insert:
          [image_paragraph, caption_paragraph]  (caption only if text exists)
          or just [image_paragraph] if no caption text was found.

        Two paragraphs are created instead of one to ensure the green border
        box surrounds only the image, not any accompanying caption text.
        Word's paragraph border applies to the entire paragraph — mixing
        image and caption text in one paragraph would wrap both in the border.

        Image paragraph:
          - Centered alignment
          - Green border box applied via XML paragraph border properties

        Caption paragraph (if caption text exists):
          - Centered alignment
          - Green text color
          - No border

        For each image, the original width is read from the <wp:extent>
        element (stored in EMU units; 914400 EMU = 1 inch) so the image
        dimensions are preserved in the output.
        """
        image_data_list: list[tuple] = []   # (image_stream, width_inches_or_None)
        caption_runs:    list[tuple] = []   # (text, is_bold, is_italic)

        for run in source_paragraph.runs:
            drawing_element = run._element.find(qn("w:drawing"))

            if drawing_element is None:
                # This run contains plain text — save it for the caption paragraph
                if run.text:
                    caption_runs.append((run.text, run.bold, run.italic))
                continue

            # Locate the image relationship ID on the <a:blip> element.
            # <a:blip r:embed="rId5"/> tells us which document relationship
            # holds the actual image file data.
            blip_element = drawing_element.find(".//" + qn("a:blip"))
            if blip_element is None:
                continue

            relationship_id = blip_element.get(qn("r:embed"))
            if not relationship_id:
                continue

            try:
                # Retrieve the image binary data from the source document's
                # relationship map
                image_part   = self.source_document.part.related_parts[relationship_id]
                image_stream = io.BytesIO(image_part.blob)

                # Read the original image width from the extent element.
                # The extent is in EMU (English Metric Units): 914400 EMU = 1 inch.
                extent_element = drawing_element.find(".//" + qn("wp:extent"))
                width_inches   = None
                if extent_element is not None:
                    cx_emu = int(extent_element.get("cx", 0))
                    if cx_emu > 0:
                        width_inches = cx_emu / 914400

                image_data_list.append((image_stream, width_inches))

            except Exception as error:
                # If the image cannot be extracted, record a visible placeholder
                caption_runs.append(
                    (f"[Image could not be migrated: {error}]", False, False)
                )

        result_paragraphs = []

        # ── Create one centered, bordered paragraph per image ──────────────────
        for image_stream, width_inches in image_data_list:
            image_paragraph           = self._working_document.add_paragraph()
            image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            picture_run = image_paragraph.add_run()
            if width_inches:
                picture_run.add_picture(image_stream, width=Inches(width_inches))
            else:
                picture_run.add_picture(image_stream)

            # Apply the green border box to the image paragraph only
            self._apply_green_border_to_paragraph(image_paragraph)
            result_paragraphs.append(image_paragraph)

        # ── Create a separate caption paragraph with green text, no border ─────
        # Keeping caption text in a separate paragraph ensures the green border
        # wraps only the image, not the caption text below it.
        if caption_runs:
            caption_paragraph           = self._working_document.add_paragraph()
            caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            for caption_text, is_bold, is_italic in caption_runs:
                caption_run                  = caption_paragraph.add_run(caption_text)
                caption_run.bold             = is_bold
                caption_run.italic           = is_italic
                caption_run.font.color.rgb   = MIGRATED_CONTENT_COLOR

            result_paragraphs.append(caption_paragraph)

        return result_paragraphs

    def _apply_green_border_to_paragraph(self, paragraph):
        """
        Add a solid green box border around the given paragraph using XML.

        python-docx does not expose paragraph borders through its Python API,
        so we manipulate the XML directly. Paragraph borders are defined in:
          <w:pPr><w:pBdr><w:top .../><w:left .../><w:bottom .../><w:right .../>

        This visually frames the migrated image so users can identify it as
        transferred content even without the green text color cue.
        """
        paragraph_properties = paragraph._element.find(qn("w:pPr"))
        if paragraph_properties is None:
            paragraph_properties = OxmlElement("w:pPr")
            paragraph._element.insert(0, paragraph_properties)

        # Remove any existing border definition before adding ours
        for existing_border in paragraph_properties.findall(qn("w:pBdr")):
            paragraph_properties.remove(existing_border)

        paragraph_border = OxmlElement("w:pBdr")
        for border_side in ("top", "left", "bottom", "right"):
            border_element = OxmlElement(f"w:{border_side}")
            border_element.set(qn("w:val"),   "single")
            border_element.set(qn("w:sz"),    "6")                   # 0.75pt thickness
            border_element.set(qn("w:space"), "4")                   # 4pt gap from content
            border_element.set(qn("w:color"), MIGRATED_CONTENT_COLOR_HEX)
            paragraph_border.append(border_element)

        paragraph_properties.append(paragraph_border)

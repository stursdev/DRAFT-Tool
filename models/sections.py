# =============================================================================
# models/sections.py
#
# Core data models for the Document Migration Tool.
#
# This file is the single source of truth for every data structure used
# across the application. All other modules import from here.
#
# Reading order:
#   1. BlockType        — what kind of content a block holds
#   2. MatchMethod      — how a mapping was determined
#   3. MigrationMode    — how content is inserted into the destination
#   4. MappingStatus    — the current state of a row in the mapping table
#   5. ContentBlock     — one unit of content inside a section
#   6. Section          — a heading and all content beneath it
#   7. MappingResult    — one row in the mapping table (source + destination pair)
#
# Design principle:
#   Every field that is used across more than one module is defined here as
#   either an Enum or a dataclass field — never as a raw string or magic number
#   scattered through the codebase. This makes the codebase searchable and
#   refactorable: change the value in one place and it updates everywhere.
# =============================================================================

from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


# =============================================================================
# BlockType
#
# Identifies what kind of content object is stored in a ContentBlock.
# The migrator uses this to decide how to copy each block into the output.
# =============================================================================
class BlockType(Enum):
    PARAGRAPH = "paragraph"   # A text paragraph (plain or formatted)
    TABLE     = "table"       # A Word table with rows and columns
    IMAGE     = "image"       # A paragraph whose runs contain inline images


# =============================================================================
# MatchMethod
#
# Records how the auto-mapper arrived at a mapping suggestion.
# Stored on MappingResult so developers and users can understand why a
# particular match was made.
#
# EXACT  — the normalized titles were identical character-for-character
# FUZZY  — rapidfuzz token_sort_ratio scored above the review threshold
# NONE   — no destination scored high enough; row starts as Unmapped
# MANUAL — the user overrode the auto-mapper's suggestion via the dropdown
# =============================================================================
class MatchMethod(Enum):
    EXACT  = "exact"
    FUZZY  = "fuzzy"
    NONE   = "none"
    MANUAL = "manual"


# =============================================================================
# MigrationMode
#
# Controls where the migrated content is inserted relative to any existing
# boilerplate content in the destination section.
#
# APPEND  — migrated content goes AFTER  the boilerplate (default for all rows)
# PREPEND — migrated content goes BEFORE the boilerplate
# REPLACE — the boilerplate is deleted and replaced by the migrated content
#
# All migrated content is colored green regardless of mode, so the user can
# identify it in the output document.
# =============================================================================
class MigrationMode(Enum):
    APPEND  = "append"
    PREPEND = "prepend"
    REPLACE = "replace"


# =============================================================================
# MappingStatus
#
# The current state of a single row in the mapping table UI.
# Drives row background color, filter radio button counts, and the status
# bar at the bottom of the window.
#
# AUTO     — auto-matched with high confidence (>= 90%)       shown in green
# REVIEW   — auto-matched with medium confidence (65–89%)     shown in yellow
# UNMAPPED — no destination assigned; will be skipped          shown in red
# MANUAL   — user manually changed the destination             shown in blue
# SKIPPED  — user explicitly chose to skip this section        shown in grey
# =============================================================================
class MappingStatus(Enum):
    AUTO     = "auto"
    REVIEW   = "review"
    UNMAPPED = "unmapped"
    MANUAL   = "manual"
    SKIPPED  = "skipped"


# =============================================================================
# ContentBlock
#
# Represents a single content element found inside a section.
#
# We store the original python-docx object (raw_object) rather than
# converting it to a string early, because the migrator needs access to the
# full formatting information (runs, fonts, XML structure) when copying.
#
# raw_object holds:
#   PARAGRAPH → docx.text.paragraph.Paragraph
#   TABLE     → docx.table.Table
#   IMAGE     → docx.text.paragraph.Paragraph (containing image runs)
# =============================================================================
@dataclass
class ContentBlock:
    # What kind of content this block holds — determines how it is copied
    block_type: BlockType

    # The original python-docx object, kept in its native form for faithful copying
    raw_object: object = None

    def __repr__(self) -> str:
        return f"ContentBlock({self.block_type.value})"


# =============================================================================
# Section
#
# Represents one heading at any level (1–6) and all the content directly
# beneath it, up to but not including the next heading of equal or higher
# importance (lower or equal level number).
#
# Example document structure mapped to Section objects:
#   "1. Introduction"     → Section(title="1. Introduction",   heading_level=1)
#   "1.1 Purpose"         → Section(title="1.1 Purpose",       heading_level=2)
#   "1.2 Scope"           → Section(title="1.2 Scope",         heading_level=2)
#   "2. Management"       → Section(title="2. Management",     heading_level=1)
#
# paragraph_index is the position of the heading paragraph in the flat
# document.paragraphs list, used by the migrator to locate where a section
# begins when clearing or inserting content.
# =============================================================================
@dataclass
class Section:
    # The heading text exactly as it appears in the document
    title: str

    # Word heading level: 1 = Heading 1, 2 = Heading 2, etc.
    heading_level: int

    # Position of this heading in document.paragraphs (used for boundary finding)
    paragraph_index: int

    # All content blocks belonging to this section in document order
    content_blocks: list[ContentBlock] = field(default_factory=list)

    # -------------------------------------------------------------------------
    # display_title
    #
    # Returns the title with leading spaces added to visually represent the
    # heading hierarchy in the mapping table and destination dropdown.
    #
    # Indentation: (heading_level - 1) groups of 3 spaces.
    # Level 1 → no indent   ("1. Introduction")
    # Level 2 → 3 spaces    ("   1.1 Purpose")
    # Level 3 → 6 spaces    ("      1.1.1 Detail")
    # -------------------------------------------------------------------------
    @property
    def display_title(self) -> str:
        indent = "   " * (self.heading_level - 1)
        return f"{indent}{self.title}"

    # -------------------------------------------------------------------------
    # normalized_title
    #
    # A cleaned version of the title used exclusively for fuzzy matching.
    # We remove section number prefixes and strip non-alphanumeric characters
    # so the matcher compares the meaningful words only.
    #
    # Example:
    #   "1.2 Configuration Management Plan" → "configuration management plan"
    #   "3.2.1 Old Config Process"          → "old config process"
    # -------------------------------------------------------------------------
    @property
    def normalized_title(self) -> str:
        # Remove leading numbering patterns like "1.", "1.2", "3.2.1 ", etc.
        text = re.sub(r"^\d+(\.\d+)*\.?\s*", "", self.title)
        # Strip any remaining non-alphanumeric characters except spaces
        text = re.sub(r"[^a-z0-9\s]", "", text.lower())
        return text.strip()

    def __repr__(self) -> str:
        return (
            f"Section(L{self.heading_level}: '{self.title}', "
            f"blocks={len(self.content_blocks)})"
        )


# =============================================================================
# MappingResult
#
# Represents one row in the mapping table UI.
# Pairs a source section with its assigned destination section and records
# all decisions made about that pairing.
#
# Change detection:
#   original_dest_section and original_status are set once by the auto-mapper
#   and never modified afterward. The UI updates dest_section and status as
#   the user makes changes. Comparing current vs. original tells us whether
#   the user changed anything from the auto-mapper's suggestion.
#
#   Example: auto-mapper suggested "1.2 Scope" → "1.2 Scope & Purpose" (AUTO).
#   User changes it to "2.1 Overview" → status becomes MANUAL.
#   User changes it back to "1.2 Scope & Purpose" → status reverts to AUTO
#   because current destination matches original_dest_section again.
# =============================================================================
@dataclass
class MappingResult:
    # The source section — always present, never None
    source_section: Section

    # The destination section currently assigned by the user or auto-mapper.
    # None means the row is Unmapped or Skipped.
    dest_section: Optional[Section] = None

    # Snapshot of the destination the auto-mapper originally suggested.
    # Set once at analysis time and never modified again.
    # Used to detect whether the user has changed the auto-mapped suggestion.
    original_dest_section: Optional[Section] = None

    # How confident the auto-mapper was in this match (0.0 to 1.0)
    confidence: float = 0.0

    # How this mapping was determined — see MatchMethod enum above
    match_method: MatchMethod = MatchMethod.NONE

    # Current mapping status — drives row color and filter counts
    status: MappingStatus = MappingStatus.UNMAPPED

    # The auto-mapper's original status — stored so we can restore it
    # if the user reverts their change back to the original destination.
    # For example: if auto-mapper gave AUTO status, user changes destination
    # (becomes MANUAL), then user changes back to original destination —
    # we restore AUTO rather than leaving it as MANUAL.
    original_status: MappingStatus = MappingStatus.UNMAPPED

    # How content should be inserted into the destination section.
    # None for Unmapped and Skipped rows since they produce no output.
    migration_mode: Optional[MigrationMode] = MigrationMode.APPEND

    # -------------------------------------------------------------------------
    # top_suggestions
    #
    # Alternative destination sections produced by the auto-mapper, stored as
    # (Section, confidence_float) pairs sorted by confidence descending.
    #
    # Populated once at analysis time and never modified afterward. The mapping
    # table uses this list to display quick-pick suggestion items at the very
    # top of each destination dropdown — above the full section list — so the
    # user can immediately see and select strong alternatives without scrolling.
    #
    # The primary match (dest_section) is excluded to avoid showing the same
    # section twice. Empty for rows where no candidates scored above the
    # minimum suggestion threshold (SUGGESTION_MIN_SCORE in auto_mapper.py).
    # -------------------------------------------------------------------------
    top_suggestions: list[tuple[Section, float]] = field(default_factory=list)

    # -------------------------------------------------------------------------
    # confidence_display
    # Human-readable confidence for the Conf. column in the table.
    # Shows "–" for unmapped/skipped rows where confidence is meaningless.
    # -------------------------------------------------------------------------
    @property
    def confidence_display(self) -> str:
        if self.status in (MappingStatus.UNMAPPED, MappingStatus.SKIPPED):
            return "–"
        return f"{int(self.confidence * 100)}%"

    # -------------------------------------------------------------------------
    # status_icon
    # Emoji displayed in the Status column of the mapping table.
    # Uses solid colored circles for reliable cross-platform rendering.
    # Note: ⚠️ was replaced with 🟡 because the warning emoji renders as a
    # monochrome glyph (invisible on dark backgrounds) on macOS.
    # -------------------------------------------------------------------------
    @property
    def status_icon(self) -> str:
        icon_map = {
            MappingStatus.AUTO:     "🟢",
            MappingStatus.REVIEW:   "🟡",
            MappingStatus.UNMAPPED: "🔴",
            MappingStatus.MANUAL:   "📝",
            MappingStatus.SKIPPED:  "⏭",
        }
        return icon_map.get(self.status, "?")

    # -------------------------------------------------------------------------
    # has_user_changed_destination
    #
    # Returns True if the user has changed the destination from what the
    # auto-mapper originally suggested. Used to correctly compute the
    # "Changed" filter count and decide whether to show MANUAL status.
    #
    # Comparison logic:
    #   - Both None (was unmapped, still unmapped)   → not changed
    #   - One is None, the other is not              → changed
    #   - Both have titles but the titles differ     → changed
    #   - Both have the same title                   → not changed
    # -------------------------------------------------------------------------
    @property
    def has_user_changed_destination(self) -> bool:
        current_title  = self.dest_section.title if self.dest_section else None
        original_title = (
            self.original_dest_section.title
            if self.original_dest_section else None
        )
        return current_title != original_title

    def __repr__(self) -> str:
        destination = self.dest_section.title if self.dest_section else "None"
        return (
            f"MappingResult('{self.source_section.title}' → '{destination}' "
            f"[{self.confidence_display}] {self.status.value})"
        )

# =============================================================================
# gui/mapping_table.py
#
# MappingTableWidget — the interactive section mapping table in Step 2.
#
# Responsibilities:
#   - Displays one row per source section with source title, destination
#     dropdown, mode dropdown, confidence score, and status icon
#   - Manages the filter radio buttons (All, Auto-Mapped, Needs Review,
#     Unmapped, Skipped, Changed)
#   - Handles user interaction with destination and mode dropdowns
#   - Enforces Replace mode exclusivity (only one source per destination
#     may use Replace — see _update_replace_availability() for details)
#   - Tracks changes against the auto-mapper's original suggestions
#   - Emits countChanged signal whenever any mapping changes so the parent
#     widget can update the status bar and acknowledgement checkbox
#
# Sentinel values:
#   The destination dropdown contains two special "sentinel" items at the top
#   in addition to the real destination sections. A sentinel is a placeholder
#   item that represents a non-section state rather than an actual destination.
#
#   We identify sentinels by storing a negative integer as their item data
#   (UserRole). Real section items store their index into dest_sections (>= 0).
#   Negative values cannot be confused with valid section indices, making the
#   distinction unambiguous.
#
#   UNMAPPED_SENTINEL_VALUE = -1  → "Unmapped"          (no destination yet)
#   SKIP_SENTINEL_VALUE     = -2  → "Skip this section" (explicit exclusion)
# =============================================================================

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from PyQt5.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QCompleter,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui  import QColor, QFont

from models.sections import (
    MappingResult,
    MappingStatus,
    MigrationMode,
    MatchMethod,
    Section,
)
from gui.widgets import IndentedComboDelegate, NoScrollComboBox, SectionLabel


# =============================================================================
# Table column indices
# Centralised here so a column reorder only requires changes in one place.
# =============================================================================
COLUMN_SOURCE      = 0
COLUMN_DESTINATION = 1
COLUMN_MODE        = 2
COLUMN_CONFIDENCE  = 3
COLUMN_STATUS      = 4

COLUMN_WIDTHS = {
    COLUMN_MODE:       110,
    COLUMN_CONFIDENCE:  52,
    COLUMN_STATUS:      60,
}

# =============================================================================
# Destination combo sentinel values
#
# Sentinels are special items in the destination dropdown that represent
# non-section states. They use negative integers as their stored data value
# so they can never be confused with real section indices (which are >= 0).
#
# UNMAPPED_SENTINEL_VALUE (-1):
#   The default state for source sections the auto-mapper could not match.
#   Selecting this means "I have not yet decided where this section goes."
#   These rows appear red in the table and are excluded from migration.
#
# SKIP_SENTINEL_VALUE (-2):
#   An explicit user decision to exclude this section from migration.
#   Unlike Unmapped (which means undecided), Skip means "I know I don't
#   want this section migrated." Both are excluded from migration but
#   are tracked separately in the report and status bar counts.
# =============================================================================
UNMAPPED_SENTINEL_VALUE = -1
SKIP_SENTINEL_VALUE     = -2

# Fixed positions of the sentinel and first separator items.
# These four slots are always present at the top of every destination combo.
# Suggestion items (0–3) and their trailing separator are inserted dynamically
# after COMBO_INDEX_SEPARATOR_SECOND, so the first real section index varies.
COMBO_INDEX_UNMAPPED          = 0
COMBO_INDEX_SEPARATOR_FIRST   = 1
COMBO_INDEX_SKIP              = 2
COMBO_INDEX_SEPARATOR_SECOND  = 3
# Index 4 onward: 0–3 suggestion items, optional separator, then full section list.
# Because this is dynamic, code must use combo.count() to track positions
# rather than relying on a fixed constant.

# Qt item-data roles used by the destination combo.
# Qt.UserRole     (default) → section index (int >= 0) or sentinel value (int < 0)
# Qt.UserRole + 1           → heading level (int) used by IndentedComboDelegate
# Qt.UserRole + 2           → True if this item is a suggestion, absent otherwise
#
# Using distinct roles keeps each piece of data independently readable without
# conflicting with each other or with Qt's own reserved roles.
COMBO_ITEM_ROLE_SECTION_INDEX = Qt.UserRole        # Section index or sentinel
COMBO_ITEM_ROLE_HEADING_LEVEL = Qt.UserRole + 1    # For IndentedComboDelegate indentation
COMBO_ITEM_ROLE_IS_SUGGESTION = Qt.UserRole + 2    # True marks a suggestion item

# =============================================================================
# Mode dropdown options
# Ordered list of (display_label, MigrationMode) pairs for the mode combo.
# =============================================================================
MODE_DROPDOWN_OPTIONS = [
    ("Append",  MigrationMode.APPEND),
    ("Prepend", MigrationMode.PREPEND),
    ("Replace", MigrationMode.REPLACE),
]

# =============================================================================
# Row background colors per mapping status
# Soft tints that are readable without being distracting.
# =============================================================================
ROW_COLORS = {
    MappingStatus.AUTO:     QColor("#f0fdf4"),   # Soft green
    MappingStatus.REVIEW:   QColor("#fffbeb"),   # Soft amber
    MappingStatus.UNMAPPED: QColor("#fef2f2"),   # Soft red
    MappingStatus.MANUAL:   QColor("#eff6ff"),   # Soft blue
    MappingStatus.SKIPPED:  QColor("#f3f4f6"),   # Soft grey
}


class MappingTableWidget(QWidget):
    """
    Composite widget containing the filter radio buttons and the mapping table.

    Emits countChanged whenever any mapping changes. The payload is a dict
    with keys: "all", "auto", "review", "unmapped", "skipped", "changed".
    The parent widget (MigrationTab) connects to this signal to update the
    status bar and the unmapped acknowledgement checkbox.
    """

    countChanged = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)

        # Set when populate() is called — cleared when reset() is called
        self._mapping_results:   list[MappingResult] = []
        self._dest_sections:     list[Section]       = []

        self._build_ui()

    # =========================================================================
    # UI Construction
    # =========================================================================

    def _build_ui(self):
        """Build the filter radio row and the mapping table."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        layout.addLayout(self._build_filter_row())
        layout.addWidget(self._build_table())

    def _build_filter_row(self) -> QHBoxLayout:
        """
        Build the row of radio buttons for filtering table rows.
        Each radio button shows a live count badge in its label.
        """
        filter_row = QHBoxLayout()
        filter_row.setSpacing(16)

        self._filter_button_group = QButtonGroup(self)

        # (display_label, filter_key, accessible_description)
        # The accessible description is announced by screen readers and explains
        # what each filter shows — important because the count badge "(n)" in the
        # label alone does not convey the filter's purpose to assistive technology.
        filter_definitions = [
            (
                "All",
                "all",
                "Show all source sections regardless of their mapping status",
            ),
            (
                "Auto-Mapped",
                "auto",
                "Show only sections that were automatically matched with high confidence",
            ),
            (
                "Needs Review",
                "review",
                "Show only sections matched with lower confidence that need manual verification",
            ),
            (
                "Unmapped",
                "unmapped",
                "Show only sections that have not been assigned a destination and will be skipped",
            ),
            (
                "Skipped",
                "skipped",
                "Show only sections that were explicitly marked to skip during migration",
            ),
            (
                "Changed",
                "changed",
                "Show only sections where the destination was manually changed from the auto-mapped suggestion",
            ),
        ]

        self._filter_radio_buttons: dict[str, QRadioButton] = {}

        for display_label, filter_key, accessible_description in filter_definitions:
            radio_button = QRadioButton(display_label)
            radio_button.setProperty("filter_key", filter_key)
            radio_button.setAccessibleDescription(accessible_description)
            radio_button.toggled.connect(self._on_filter_radio_toggled)
            self._filter_button_group.addButton(radio_button)
            filter_row.addWidget(radio_button)
            self._filter_radio_buttons[filter_key] = radio_button

        self._filter_radio_buttons["all"].setChecked(True)
        filter_row.addStretch()

        return filter_row

    def _build_table(self) -> QTableWidget:
        """Build and configure the mapping table widget."""
        self._table = QTableWidget()
        self._table.setColumnCount(5)
        self._table.setHorizontalHeaderLabels([
            "Source Section",
            "Destination Section",
            "Mode",
            "Match",
            "Status",
        ])

        header = self._table.horizontalHeader()
        header.setSectionResizeMode(COLUMN_SOURCE,      QHeaderView.Stretch)
        header.setSectionResizeMode(COLUMN_DESTINATION, QHeaderView.Stretch)
        header.setSectionResizeMode(COLUMN_MODE,        QHeaderView.Fixed)
        header.setSectionResizeMode(COLUMN_CONFIDENCE,  QHeaderView.Fixed)
        header.setSectionResizeMode(COLUMN_STATUS,      QHeaderView.Fixed)

        for column, width in COLUMN_WIDTHS.items():
            self._table.setColumnWidth(column, width)

        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().setVisible(False)
        self._table.setMinimumHeight(240)

        # ── Section 508 accessibility ─────────────────────────────────────────
        # Provide a programmatic name and description that screen readers
        # announce when the table receives focus via keyboard navigation.
        self._table.setAccessibleName("Section mapping table")
        self._table.setAccessibleDescription(
            "Each row maps a source document section to a destination section. "
            "Use the Destination column dropdown to assign or reassign a section. "
            "Use the Mode column dropdown to choose how content is inserted. "
            "Rows are colour-coded by status and the status column shows an icon."
        )

        return self._table

    # =========================================================================
    # Public API
    # =========================================================================

    def populate(
        self,
        mapping_results: list[MappingResult],
        dest_sections:   list[Section],
    ):
        """
        Fill the table with one row per MappingResult.
        Called once after the AnalyzeWorker completes.
        """
        self._mapping_results = mapping_results
        self._dest_sections   = dest_sections

        self._table.setRowCount(0)   # Clear any previously displayed rows

        for row_index, mapping_result in enumerate(mapping_results):
            self._insert_table_row(row_index, mapping_result)

        self._apply_active_filter()
        self._emit_current_counts()

    def get_results(self) -> list[MappingResult]:
        """Return the current list of MappingResults including user changes."""
        return self._mapping_results

    def reset(self):
        """Clear the table. Called when Analyze is clicked again."""
        self._table.setRowCount(0)
        self._mapping_results = []
        self._dest_sections   = []

    # =========================================================================
    # Row insertion
    # =========================================================================

    def _insert_table_row(self, row_index: int, mapping_result: MappingResult):
        """Insert one row into the table for the given MappingResult."""
        self._table.insertRow(row_index)

        # ── Source section title (indented by heading level) ──────────────────
        source_item = QTableWidgetItem(mapping_result.source_section.display_title)
        source_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)

        # Bold level-1 headings to visually reinforce the document hierarchy
        if mapping_result.source_section.heading_level == 1:
            bold_font = source_item.font()
            bold_font.setBold(True)
            source_item.setFont(bold_font)

        self._table.setItem(row_index, COLUMN_SOURCE, source_item)

        # ── Destination dropdown ───────────────────────────────────────────────
        dest_combo = self._build_destination_combo(mapping_result, row_index)
        self._table.setCellWidget(row_index, COLUMN_DESTINATION, dest_combo)

        # ── Mode dropdown ──────────────────────────────────────────────────────
        mode_combo = self._build_mode_combo(mapping_result, row_index)
        self._table.setCellWidget(row_index, COLUMN_MODE, mode_combo)

        # ── Confidence display ─────────────────────────────────────────────────
        confidence_item = QTableWidgetItem(mapping_result.confidence_display)
        confidence_item.setTextAlignment(Qt.AlignCenter)
        confidence_item.setFlags(Qt.ItemIsEnabled)
        self._table.setItem(row_index, COLUMN_CONFIDENCE, confidence_item)

        # ── Status icon ────────────────────────────────────────────────────────
        status_item = QTableWidgetItem(mapping_result.status_icon)
        status_item.setTextAlignment(Qt.AlignCenter)
        status_item.setFlags(Qt.ItemIsEnabled)
        self._table.setItem(row_index, COLUMN_STATUS, status_item)

        self._apply_row_background_color(row_index, mapping_result.status)

    # =========================================================================
    # Destination combo construction
    # =========================================================================

    def _build_destination_combo(
        self,
        mapping_result: MappingResult,
        row_index:      int,
    ) -> NoScrollComboBox:
        """
        Build the destination section dropdown for one mapping row.

        Dropdown structure (items with dynamic suggestion count):
          [0] "Unmapped"              ← sentinel: no destination assigned yet
          [1] ──────────────          ← separator
          [2] "⏭ Skip this..."       ← sentinel: explicit user exclusion
          [3] ──────────────          ← separator
          [4] "✨ Section A (95%)"   ← suggestion 1  (only present if
          [5] "✨ Section B (78%)"   ← suggestion 2    top_suggestions
          [6] "✨ Section C (67%)"   ← suggestion 3    is non-empty)
          [7] ──────────────          ← separator after suggestions
          [8+] 1. Introduction        ← full section list (always present)
             1.1 Purpose              ← indented sub-sections
             ...

        Item data roles (see COMBO_ITEM_ROLE_* constants):
          UserRole     → section index (int >= 0) for real sections,
                         or sentinel value (int < 0) for Unmapped / Skip.
          UserRole + 1 → heading level (int) for IndentedComboDelegate.
          UserRole + 2 → True for suggestion items, absent for all others.

        Suggestion items store the same section index as the matching
        full-list entry. Clicking a suggestion causes the handler to
        transparently switch the combo to the full-list entry so the
        combo header always displays the clean section title.
        """
        combo = NoScrollComboBox()
        combo.setItemDelegate(IndentedComboDelegate(combo))

        # ── Section 508: programmatic label for screen readers ────────────────
        # The label includes the source title so the user knows which row this
        # combo belongs to when navigating by keyboard.
        source_title = mapping_result.source_section.title
        combo.setAccessibleName(f"Destination for: {source_title}")
        combo.setAccessibleDescription(
            "Select the destination section where this source section's content "
            "will be migrated. Suggestions appear at the top of the list."
        )

        # ── Sentinel: Unmapped ────────────────────────────────────────────────
        combo.addItem("Unmapped", UNMAPPED_SENTINEL_VALUE)
        combo.setItemData(COMBO_INDEX_UNMAPPED, 0, COMBO_ITEM_ROLE_HEADING_LEVEL)

        combo.insertSeparator(COMBO_INDEX_SEPARATOR_FIRST)

        # ── Sentinel: Skip this section ───────────────────────────────────────
        combo.addItem("⏭  Skip this section", SKIP_SENTINEL_VALUE)
        combo.setItemData(COMBO_INDEX_SKIP, 0, COMBO_ITEM_ROLE_HEADING_LEVEL)

        combo.insertSeparator(COMBO_INDEX_SEPARATOR_SECOND)

        # ── Suggestion items (0–3, dynamic) ───────────────────────────────────
        # Inserted between the sentinels and the full section list. A trailing
        # separator is added only if at least one suggestion item was inserted.
        self._insert_suggestion_items(combo, mapping_result)

        # ── Full destination section list ──────────────────────────────────────
        # dest_section_index is the clean 0-based index into self._dest_sections.
        # We use enumerate() directly so the index is always correct, regardless
        # of how many sentinel or suggestion items precede this block.
        for dest_section_index, dest_section in enumerate(self._dest_sections):
            combo_item_position = combo.count()
            combo.addItem(dest_section.display_title, dest_section_index)
            combo.setItemData(
                combo_item_position,
                dest_section.heading_level,
                COMBO_ITEM_ROLE_HEADING_LEVEL,
            )
            # COMBO_ITEM_ROLE_IS_SUGGESTION is intentionally NOT set on full-list
            # items — itemData returns None for absent roles, which evaluates as
            # falsy and cleanly distinguishes full-list items from suggestions.

        # Set the initial selected item to match the auto-mapper's suggestion
        self._set_combo_initial_selection(combo, mapping_result)

        # Connect change signal — row_index captured in the closure so each
        # combo knows which row it belongs to when it fires.
        combo.currentIndexChanged.connect(
            lambda _signal_index, r=row_index: self._on_destination_changed(r)
        )

        self._make_combo_searchable(combo)

        return combo

    def _make_combo_searchable(self, combo: NoScrollComboBox) -> None:
        """
        Make the destination combo searchable by typing.

        Enables the line edit and attaches a QCompleter that filters on
        any substring (MatchContains, case-insensitive). The completer
        only searches clean section titles — not sentinels or suggestion
        items with the ✨ prefix — so results are unambiguous.

        On focus-out: if the text in the line edit does not match any
        combo item exactly, the selection silently reverts to the last
        confirmed valid index so invalid free-text is never committed.
        """
        combo.setEditable(True)
        combo.setInsertPolicy(NoScrollComboBox.NoInsert)
        combo.lineEdit().setPlaceholderText("Type to search…")
        combo.lineEdit().setStyleSheet(
            "QLineEdit { border: none; background: transparent; padding: 2px 4px; }"
        )

        # Completer uses only the clean section titles from the full list
        section_titles = [s.display_title for s in self._dest_sections]
        completer = QCompleter(section_titles, combo)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        combo.setCompleter(completer)

        # When the user picks a completion, find and set the matching combo index
        def on_completion_activated(text: str) -> None:
            for i in range(combo.count()):
                if combo.itemText(i) == text:
                    combo.setCurrentIndex(i)
                    return

        completer.activated.connect(on_completion_activated)

        # Track the last confirmed valid index so we can revert on bad input
        last_valid = [combo.currentIndex()]

        def on_index_changed(idx: int) -> None:
            if idx >= 0:
                last_valid[0] = idx

        combo.currentIndexChanged.connect(on_index_changed)

        def on_editing_finished() -> None:
            text = combo.currentText()
            for i in range(combo.count()):
                if combo.itemText(i) == text:
                    return  # text matches an existing item — nothing to do
            # No match — silently revert to the last valid selection
            combo.blockSignals(True)
            combo.setCurrentIndex(last_valid[0])
            combo.blockSignals(False)

        combo.lineEdit().editingFinished.connect(on_editing_finished)

    def _insert_suggestion_items(
        self,
        combo:          NoScrollComboBox,
        mapping_result: MappingResult,
    ):
        """
        Insert quick-pick suggestion items from the auto-mapper into the combo,
        just before the full section list.

        Each suggestion item is labelled "✨  <title>  (nn%)" and stores the
        same section index as the matching entry in the full list below.
        It is also marked with COMBO_ITEM_ROLE_IS_SUGGESTION = True so the
        change handler can detect it and transparently redirect the combo
        selection to the full-list entry (keeping the header label clean).

        If mapping_result.top_suggestions is empty this method returns without
        adding anything — no separator is inserted either.
        """
        if not mapping_result.top_suggestions:
            return

        # Pre-build a title-to-index lookup so we can find each suggestion
        # section's position in self._dest_sections in O(1).
        dest_index_by_title: dict[str, int] = {
            section.title: index
            for index, section in enumerate(self._dest_sections)
        }

        suggestions_successfully_added = 0

        for suggestion_section, suggestion_confidence in mapping_result.top_suggestions:
            dest_section_index = dest_index_by_title.get(suggestion_section.title)

            # Guard: the destination template may have changed since analysis,
            # so a suggestion section might no longer exist in self._dest_sections.
            if dest_section_index is None:
                continue

            confidence_pct   = int(suggestion_confidence * 100)
            suggestion_label = f"✨  {suggestion_section.title}  ({confidence_pct}%)"
            combo_position   = combo.count()

            combo.addItem(suggestion_label, dest_section_index)

            # Suggestions are always displayed flush-left (heading level 1 = 0 indent)
            # so they stand out as a distinct group above the indented full list.
            combo.setItemData(combo_position, 1, COMBO_ITEM_ROLE_HEADING_LEVEL)

            # Mark this item as a suggestion so _on_destination_changed can
            # detect it and redirect to the full-list entry before processing.
            combo.setItemData(combo_position, True, COMBO_ITEM_ROLE_IS_SUGGESTION)

            suggestions_successfully_added += 1

        # Insert a separator between the suggestions and the full section list
        # only if at least one suggestion was actually added to the combo.
        if suggestions_successfully_added > 0:
            combo.insertSeparator(combo.count())

    def _set_combo_initial_selection(
        self,
        combo:          NoScrollComboBox,
        mapping_result: MappingResult,
    ):
        """
        Set the combo's selected item to match the MappingResult's initial state.

        Selection logic:
          - UNMAPPED status or no dest_section → select "Unmapped" (index 0)
          - SKIPPED status                     → select "Skip" (index 2)
          - Has a dest_section                 → find and select that section

        Important: We skip suggestion items during the search and always select
        the matching entry in the full section list. This ensures the combo
        header (visible when the dropdown is closed) shows the clean section
        title rather than a "✨ Title (nn%)" suggestion label.
        """
        if (mapping_result.status == MappingStatus.UNMAPPED
                or mapping_result.dest_section is None):
            combo.setCurrentIndex(COMBO_INDEX_UNMAPPED)
            return

        if mapping_result.status == MappingStatus.SKIPPED:
            combo.setCurrentIndex(COMBO_INDEX_SKIP)
            return

        # Scan for the matching entry in the full section list.
        # We explicitly skip suggestion items (COMBO_ITEM_ROLE_IS_SUGGESTION = True)
        # so the initial selection always lands on the full-list entry.
        for combo_index in range(combo.count()):
            is_suggestion = combo.itemData(combo_index, COMBO_ITEM_ROLE_IS_SUGGESTION)
            if is_suggestion:
                continue   # Never use a suggestion item as the initial selection

            stored_value = combo.itemData(combo_index, COMBO_ITEM_ROLE_SECTION_INDEX)
            # Real section items store non-negative integer indices
            if isinstance(stored_value, int) and stored_value >= 0:
                section_at_index = self._dest_sections[stored_value]
                if section_at_index.title == mapping_result.dest_section.title:
                    combo.setCurrentIndex(combo_index)
                    return

        # Fallback: if the section was not found in the full list, show Unmapped
        combo.setCurrentIndex(COMBO_INDEX_UNMAPPED)

    # =========================================================================
    # Mode combo construction
    # =========================================================================

    def _build_mode_combo(
        self,
        mapping_result: MappingResult,
        row_index:      int,
    ) -> NoScrollComboBox:
        """
        Build the migration mode dropdown for one mapping row.

        For rows that have a destination (AUTO, REVIEW, MANUAL):
          Shows Append / Prepend / Replace options, defaulting to Append.

        For rows without a destination (UNMAPPED, SKIPPED):
          Shows a disabled "–" placeholder. Mode is irrelevant for these rows
          since they produce no output during migration.
        """
        combo = NoScrollComboBox()

        # ── Section 508: programmatic label for screen readers ────────────────
        source_title = mapping_result.source_section.title
        combo.setAccessibleName(f"Migration mode for: {source_title}")
        combo.setAccessibleDescription(
            "Choose whether migrated content is appended after, prepended before, "
            "or replaces the destination section's existing content."
        )

        row_has_destination = (
            mapping_result.dest_section is not None
            and mapping_result.status not in (
                MappingStatus.UNMAPPED,
                MappingStatus.SKIPPED,
            )
        )

        if not row_has_destination:
            combo.addItem("–")
            combo.setEnabled(False)
            combo.setStyleSheet("color: #9ca3af;")
            return combo

        # Add the three mode options with MigrationMode enum as item data
        for display_label, migration_mode in MODE_DROPDOWN_OPTIONS:
            combo.addItem(display_label, migration_mode)

        # Set the current mode to match the result (defaults to APPEND)
        current_mode = mapping_result.migration_mode or MigrationMode.APPEND
        for combo_index in range(combo.count()):
            if combo.itemData(combo_index) == current_mode:
                combo.setCurrentIndex(combo_index)
                break

        combo.currentIndexChanged.connect(
            lambda _index, r=row_index: self._on_mode_changed(r)
        )

        return combo

    # =========================================================================
    # Destination change handler
    # =========================================================================

    def _on_destination_changed(self, row_index: int):
        """
        Called when the user changes the destination dropdown for a row.

        Sequence of actions:
          1. If the user selected a suggestion item, transparently redirect
             the combo to the matching full-list entry (so the header shows
             the clean section title, not the "✨ Title (nn%)" label).
          2. Read the now-current value from the combo.
          3. Update the MappingResult's dest_section and status.
          4. Rebuild the mode combo (it may need to enable or disable).
          5. Update Replace availability across all rows.
          6. Refresh the row's visual appearance.
          7. Re-apply the active filter (row visibility may change).
          8. Emit updated counts to the parent widget.
        """
        mapping_result = self._mapping_results[row_index]
        dest_combo     = self._table.cellWidget(row_index, COLUMN_DESTINATION)

        if dest_combo is None:
            return

        # ── Suggestion redirect ───────────────────────────────────────────────
        # If the user clicked a suggestion item, switch the combo to the
        # matching full-list entry before reading the selected value.
        # This keeps the closed-combo header clean (no "✨ ... %" visible).
        is_suggestion = dest_combo.currentData(COMBO_ITEM_ROLE_IS_SUGGESTION)
        if is_suggestion:
            self._redirect_suggestion_to_full_list_item(dest_combo)

        # Read the selection after any redirect — now guaranteed to be the
        # full-list entry (or a sentinel) rather than a suggestion item.
        selected_value = dest_combo.currentData(COMBO_ITEM_ROLE_SECTION_INDEX)

        # Separators store no data (None) — ignore accidental separator selection
        if selected_value is None:
            return

        if selected_value == UNMAPPED_SENTINEL_VALUE:
            self._apply_unmapped_state(mapping_result)

        elif selected_value == SKIP_SENTINEL_VALUE:
            self._apply_skipped_state(mapping_result)

        else:
            # selected_value is a valid 0-based index into self._dest_sections
            self._apply_destination_selection(mapping_result, selected_value)

        # Rebuild mode combo to reflect the new mapped / unmapped state
        self._rebuild_mode_combo(row_index, mapping_result)

        # Re-evaluate which rows may use Replace for their destination
        self._update_replace_availability()

        # Refresh status icon, confidence text, and row background color
        self._refresh_row_visuals(row_index, mapping_result)

        # Re-apply filter so this row's visibility updates if its status changed
        self._apply_active_filter()

        # Notify the parent widget of updated status counts
        self._emit_current_counts()

    def _redirect_suggestion_to_full_list_item(self, dest_combo: NoScrollComboBox):
        """
        Silently switch a combo from its currently selected suggestion item to
        the matching entry in the full section list.

        Why this is needed:
          Suggestion items display "✨ Section Title  (nn%)". If we left the
          suggestion selected, the closed combo header would show that label
          instead of the plain section title — which is confusing and cluttered.
          By switching to the full-list entry (which has the clean title), the
          header always looks correct regardless of how the selection was made.

        blockSignals(True/False) suppresses the currentIndexChanged signal during
        the programmatic switch. Without this, the switch would re-enter
        _on_destination_changed, causing an infinite loop.
        """
        # Read the section index stored on the currently selected suggestion item.
        # Both the suggestion item and its full-list counterpart store the same index.
        suggestion_section_index = dest_combo.currentData(COMBO_ITEM_ROLE_SECTION_INDEX)

        for combo_index in range(dest_combo.count()):
            # We want a non-suggestion item with the same section index
            is_suggestion  = dest_combo.itemData(combo_index, COMBO_ITEM_ROLE_IS_SUGGESTION)
            section_index  = dest_combo.itemData(combo_index, COMBO_ITEM_ROLE_SECTION_INDEX)

            if not is_suggestion and section_index == suggestion_section_index:
                dest_combo.blockSignals(True)
                dest_combo.setCurrentIndex(combo_index)
                dest_combo.blockSignals(False)
                return
        # If no matching full-list item was found (should not happen in normal
        # use), the combo remains on the suggestion item and the handler will
        # still process it correctly via the section index stored on the item.

    def _apply_unmapped_state(self, mapping_result: MappingResult):
        """Update a MappingResult to reflect the user selecting 'Unmapped'."""
        mapping_result.dest_section   = None
        mapping_result.migration_mode = None
        mapping_result.status         = MappingStatus.UNMAPPED

    def _apply_skipped_state(self, mapping_result: MappingResult):
        """Update a MappingResult to reflect the user selecting 'Skip this section'."""
        mapping_result.dest_section   = None
        mapping_result.migration_mode = None
        mapping_result.status         = MappingStatus.SKIPPED

    def _apply_destination_selection(
        self,
        mapping_result:     MappingResult,
        dest_section_index: int,
    ):
        """
        Update a MappingResult when the user selects a real destination section.

        Status logic:
          - If the selected destination matches the original auto-mapped
            destination, restore the original auto-mapped status (AUTO or REVIEW)
            rather than marking it as MANUAL. This correctly handles the case
            where the user changes their mind and reverts to the original.
          - If the selected destination differs from the original, mark as MANUAL.
        """
        if not (0 <= dest_section_index < len(self._dest_sections)):
            return

        selected_destination          = self._dest_sections[dest_section_index]
        mapping_result.dest_section   = selected_destination

        # Restore migration mode if it was cleared when the row was unmapped/skipped
        if mapping_result.migration_mode is None:
            mapping_result.migration_mode = MigrationMode.APPEND

        # Determine whether this is a revert or a new change
        if mapping_result.has_user_changed_destination:
            mapping_result.status = MappingStatus.MANUAL
        else:
            # User reverted to the original auto-mapped destination
            mapping_result.status = mapping_result.original_status

    # =========================================================================
    # Mode change handler
    # =========================================================================

    def _on_mode_changed(self, row_index: int):
        """
        Called when the user changes the mode dropdown for a row.

        Stores the new MigrationMode on the MappingResult and re-evaluates
        Replace availability across all rows, since changing one row's mode
        can free up or block Replace on other rows targeting the same destination.
        """
        mapping_result = self._mapping_results[row_index]
        mode_combo     = self._table.cellWidget(row_index, COLUMN_MODE)

        if mode_combo is None or not mode_combo.isEnabled():
            return

        selected_mode = mode_combo.currentData(Qt.UserRole)
        if isinstance(selected_mode, MigrationMode):
            mapping_result.migration_mode = selected_mode

        self._update_replace_availability()

    # =========================================================================
    # Replace exclusivity — split into two clearly named methods
    # =========================================================================

    def _update_replace_availability(self):
        """
        Enforce the Replace mode exclusivity rule across all table rows.

        Rule:
          For each destination section, only the FIRST row in table order that
          maps to it may use Replace. All other rows targeting the same
          destination have their Replace option greyed out immediately — before
          the user opens the mode dropdown — to prevent data loss.

        Data loss scenario prevented:
          Row A (Append → Dest X): appends Row A content after boilerplate.
          Row B (Replace → Dest X): clears the section — deleting Row A's
          already-migrated content — then inserts Row B content.
          Result: Row A content is silently lost. We prevent this by ensuring
          only the topmost row for a given destination can ever select Replace.

        This method coordinates by:
          1. Finding the Replace owner for each destination
          2. Applying the correct enabled/disabled state to each mode combo
        """
        replace_owner_by_destination = self._find_replace_owners()

        for row_index, mapping_result in enumerate(self._mapping_results):
            mode_combo = self._table.cellWidget(row_index, COLUMN_MODE)
            if mode_combo is None or not mode_combo.isEnabled():
                continue
            if mapping_result.dest_section is None:
                continue

            self._apply_replace_availability_to_row(
                row_index,
                mapping_result,
                mode_combo,
                replace_owner_by_destination,
            )

    def _find_replace_owners(self) -> dict[str, int]:
        """
        Build a map of which row is the sole permitted Replace user for each
        destination section.

        Returns a dict where:
          key   = destination section title
          value = row_index of the FIRST row (in table order) that maps to
                  that destination, regardless of its current mode selection

        Why ownership is position-based, not selection-based:
          The previous design awarded ownership to the first row that actively
          selected Replace. This created a two-step exploit:
            Step 1: User changes row B's destination to Dest X (now two rows
                    map to Dest X). Because nobody has selected Replace yet,
                    the ownership dict is empty. _apply_replace_availability_to_row
                    sees owner=None for every row, so Replace stays enabled on
                    both rows — including row B, which should be blocked.
            Step 2: User selects Replace on row B. Row B now becomes the owner,
                    and row A (the correct first-row owner) gets its Replace
                    greyed out retroactively. The wrong row owns Replace.

          The fix: ownership is assigned purely by table position. The first row
          that MAPS to a destination owns Replace for it — whether or not Replace
          is currently selected. All later rows sharing that destination have
          Replace greyed out from the moment they are assigned, before the user
          ever opens the mode dropdown.
        """
        replace_owner_by_destination: dict[str, int] = {}

        for row_index, mapping_result in enumerate(self._mapping_results):
            if mapping_result.dest_section is not None:
                destination_title = mapping_result.dest_section.title
                # Record only the first occurrence — later rows are not owners
                if destination_title not in replace_owner_by_destination:
                    replace_owner_by_destination[destination_title] = row_index

        return replace_owner_by_destination

    def _apply_replace_availability_to_row(
        self,
        row_index:                    int,
        mapping_result:               MappingResult,
        mode_combo:                   NoScrollComboBox,
        replace_owner_by_destination: dict[str, int],
    ):
        """
        Enable or grey out the Replace option in one row's mode combo.

        If this row is the Replace owner for its destination (or no owner
        exists yet), Replace remains enabled.

        If another row owns Replace for this destination, the Replace option
        is greyed out in this row's combo. If this row currently has Replace
        selected, it is automatically switched to Append.
        """
        destination_title  = mapping_result.dest_section.title
        owner_row_index    = replace_owner_by_destination.get(destination_title)
        this_row_owns_replace = (
            owner_row_index is None or owner_row_index == row_index
        )

        # Find the Replace item's index in this combo
        replace_item_index = None
        for combo_index in range(mode_combo.count()):
            if mode_combo.itemData(combo_index) == MigrationMode.REPLACE:
                replace_item_index = combo_index
                break

        if replace_item_index is None:
            return

        combo_model       = mode_combo.model()
        replace_model_index = combo_model.index(replace_item_index, 0)

        if this_row_owns_replace:
            # Re-enable Replace — remove the grey color override
            combo_model.setData(replace_model_index, None, Qt.ForegroundRole)
        else:
            # Grey out Replace — this destination already has an owner
            from PyQt5.QtGui import QColor as _QColor
            combo_model.setData(
                replace_model_index,
                _QColor("#9ca3af"),
                Qt.ForegroundRole,
            )
            # If this row currently has Replace selected, demote it to Append
            if mapping_result.migration_mode == MigrationMode.REPLACE:
                mapping_result.migration_mode = MigrationMode.APPEND
                for combo_index in range(mode_combo.count()):
                    if mode_combo.itemData(combo_index) == MigrationMode.APPEND:
                        mode_combo.blockSignals(True)
                        mode_combo.setCurrentIndex(combo_index)
                        mode_combo.blockSignals(False)
                        break

    # =========================================================================
    # Mode combo rebuild
    # =========================================================================

    def _rebuild_mode_combo(self, row_index: int, mapping_result: MappingResult):
        """
        Replace the mode combo for a row after its destination changes.

        When a row transitions between mapped and unmapped (or vice versa),
        the mode combo needs to be rebuilt from scratch because the enabled
        state changes entirely — from a real dropdown to a greyed-out "–"
        placeholder or back again.
        """
        old_combo = self._table.cellWidget(row_index, COLUMN_MODE)
        if old_combo:
            old_combo.blockSignals(True)

        new_combo = self._build_mode_combo(mapping_result, row_index)
        self._table.setCellWidget(row_index, COLUMN_MODE, new_combo)

    # =========================================================================
    # Row visual refresh
    # =========================================================================

    def _refresh_row_visuals(self, row_index: int, mapping_result: MappingResult):
        """
        Update the status icon, confidence display, and background color
        for a single row after its mapping changes.
        """
        status_item = self._table.item(row_index, COLUMN_STATUS)
        if status_item:
            status_item.setText(mapping_result.status_icon)

        confidence_item = self._table.item(row_index, COLUMN_CONFIDENCE)
        if confidence_item:
            confidence_item.setText(mapping_result.confidence_display)

        self._apply_row_background_color(row_index, mapping_result.status)

    def _apply_row_background_color(
        self,
        row_index: int,
        status:    MappingStatus,
    ):
        """Paint the row background with the color matching the mapping status."""
        background_color = ROW_COLORS.get(status, QColor("#ffffff"))

        for column in range(self._table.columnCount()):
            cell_item = self._table.item(row_index, column)
            if cell_item:
                cell_item.setBackground(background_color)

    # =========================================================================
    # Filter logic
    # =========================================================================

    def _on_filter_radio_toggled(self):
        """Called when any filter radio button is toggled."""
        self._apply_active_filter()

    def _apply_active_filter(self):
        """
        Show or hide table rows based on the currently selected filter.
        Rows are hidden (not removed) so their data is fully preserved.
        """
        active_filter_key = "all"
        for key, radio_button in self._filter_radio_buttons.items():
            if radio_button.isChecked():
                active_filter_key = key
                break

        for row_index, mapping_result in enumerate(self._mapping_results):
            row_is_visible = self._row_matches_filter(
                mapping_result, active_filter_key
            )
            self._table.setRowHidden(row_index, not row_is_visible)

    def _row_matches_filter(
        self,
        mapping_result:   MappingResult,
        active_filter_key: str,
    ) -> bool:
        """
        Return True if this row should be visible for the given filter key.
        """
        filter_conditions = {
            "all":      lambda r: True,
            "auto":     lambda r: r.status == MappingStatus.AUTO,
            "review":   lambda r: r.status == MappingStatus.REVIEW,
            "unmapped": lambda r: r.status == MappingStatus.UNMAPPED,
            "skipped":  lambda r: r.status == MappingStatus.SKIPPED,
            "changed":  lambda r: r.status == MappingStatus.MANUAL,
        }
        condition = filter_conditions.get(active_filter_key, lambda r: True)
        return condition(mapping_result)

    # =========================================================================
    # Count tracking and signal emission
    # =========================================================================

    def _get_current_status_counts(self) -> dict:
        """
        Count how many rows are in each mapping status category.
        Returns a dict with keys matching the filter radio button keys.
        """
        counts = {
            "all":      len(self._mapping_results),
            "auto":     0,
            "review":   0,
            "unmapped": 0,
            "skipped":  0,
            "changed":  0,
        }

        for result in self._mapping_results:
            if result.status == MappingStatus.AUTO:
                counts["auto"]     += 1
            elif result.status == MappingStatus.REVIEW:
                counts["review"]   += 1
            elif result.status == MappingStatus.UNMAPPED:
                counts["unmapped"] += 1
            elif result.status == MappingStatus.SKIPPED:
                counts["skipped"]  += 1
            elif result.status == MappingStatus.MANUAL:
                counts["changed"]  += 1

        return counts

    def _emit_current_counts(self):
        """
        Update filter radio button labels with live counts and emit
        countChanged so the parent widget can update the status bar
        and acknowledgement checkbox.
        """
        counts = self._get_current_status_counts()

        # Update each radio button label to show the current count
        display_labels = {
            "all":      "All",
            "auto":     "Auto-Mapped",
            "review":   "Needs Review",
            "unmapped": "Unmapped",
            "skipped":  "Skipped",
            "changed":  "Changed",
        }
        for key, radio_button in self._filter_radio_buttons.items():
            radio_button.setText(f"{display_labels[key]} ({counts[key]})")

        self.countChanged.emit(counts)

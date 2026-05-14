# =============================================================================
# gui/migration_tab.py
#
# MigrationTab — the "Document Migration" tab content.
#
# Contains the three-step migration workflow:
#   Step 1 — Load Documents      (file pickers + Analyze button)
#   Step 2 — Review Mappings     (mapping table, always unlocked after analyze)
#   Step 3 — Migration Options   (output settings + Run Migration button)
#
# Step unlock rules:
#   Step 2 — unlocks after Analyze completes successfully
#   Step 3 — unlocks after Step 2 with a ⚠ caution icon (not a 🔒 lock)
#             because the user CAN migrate even with unmapped sections
#
# Migrate button rules:
#   - If no unmapped sections exist → Migrate button is enabled directly
#   - If unmapped sections exist → user must check the acknowledgement
#     checkbox before the Migrate button enables. The checkbox resets
#     automatically whenever the mapping counts change after being checked.
#
# Worker threads (AnalyzeWorker, MigrateWorker) are defined in core/workers.py
# and imported here. They are application logic, not UI code.
# =============================================================================

import json
import os
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from PyQt5.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)
from PyQt5.QtCore import Qt

from core.audit_log          import MigrationReportExporter
from core.boilerplate_detector import HIGHLIGHT_SENTENCES, HIGHLIGHT_PARAGRAPH
from core.content_migrator   import ContentMigrator
from core.document_parser    import DocumentParser
from core.workers            import AnalyzeWorker, MigrateWorker
from models.sections         import MappingResult, MappingStatus, MigrationMode, Section
from gui.mapping_table     import MappingTableWidget
from gui.widgets           import SectionLabel


class MigrationTab(QWidget):
    """
    The full Document Migration UI — Steps 1, 2, and 3 in a single scrollable
    panel. Intended to be hosted inside a QTabWidget left tab bar.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        # ── Application state ─────────────────────────────────────────────────
        self._source_file_path:  str = ""
        self._dest_file_path:    str = ""
        self._dest_sections:     list[Section]       = []
        self._mapping_results:   list[MappingResult] = []

        # Parser objects kept after analysis so the migrator can re-parse
        # the same files on migration without re-opening the file picker
        self._source_parser: DocumentParser = None
        self._dest_parser:   DocumentParser = None

        # Whether the user has checked the unmapped acknowledgement checkbox.
        # Reset to False whenever mapping counts change after being checked.
        self._user_acknowledged_unmapped_sections = False

        # The most recent status counts dict received from the mapping table
        self._current_status_counts: dict = {}

        self._build_ui()
        self._apply_styles()

    # =========================================================================
    # UI Construction
    # =========================================================================

    def _build_ui(self):
        """Assemble the three steps and the status bar."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 8)
        main_layout.setSpacing(10)

        main_layout.addWidget(self._build_step1_group())

        self._step2_group = self._build_step2_group()
        main_layout.addWidget(self._step2_group)

        self._step3_group = self._build_step3_group()
        main_layout.addWidget(self._step3_group)

        main_layout.addWidget(self._build_status_bar())

        self._set_step2_enabled(False)
        self._set_step3_enabled(False)

    # -------------------------------------------------------------------------
    # Step 1 — Load Documents
    # -------------------------------------------------------------------------

    def _build_step1_group(self) -> QGroupBox:
        group = QGroupBox("  STEP 1 · LOAD DOCUMENTS")
        group.setObjectName("step_group")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(12, 16, 12, 12)
        layout.setSpacing(6)

        layout.addWidget(SectionLabel(
            "Select your existing document (source) and the new template "
            "(destination), then click Analyze Documents to detect sections "
            "and auto-map them."
        ))

        file_grid = QGridLayout()
        file_grid.setSpacing(8)

        # ── Source file row ────────────────────────────────────────────────────
        source_label = QLabel("Source Document")
        file_grid.addWidget(source_label, 0, 0)

        self._source_path_input = QLineEdit()
        self._source_path_input.setPlaceholderText(
            "Path to existing document (source)…"
        )
        self._source_path_input.setReadOnly(True)
        self._source_path_input.setAccessibleName("Source document path")
        self._source_path_input.setAccessibleDescription(
            "Read-only field showing the path to the selected source document"
        )
        file_grid.addWidget(self._source_path_input, 0, 1)

        source_browse_button = QPushButton("📂")
        source_browse_button.setFixedWidth(36)
        source_browse_button.setToolTip("Open a file picker to select the source document (.docx)")
        source_browse_button.setAccessibleName("Browse for source document")
        source_browse_button.setAccessibleDescription(
            "Open a file browser to choose the existing document whose content will be migrated"
        )
        source_browse_button.clicked.connect(self._on_browse_source_clicked)
        file_grid.addWidget(source_browse_button, 0, 2)

        # ── Destination file row ───────────────────────────────────────────────
        dest_label = QLabel("Destination Template")
        file_grid.addWidget(dest_label, 1, 0)

        self._dest_path_input = QLineEdit()
        self._dest_path_input.setPlaceholderText(
            "Path to destination template…"
        )
        self._dest_path_input.setReadOnly(True)
        self._dest_path_input.setAccessibleName("Destination template path")
        self._dest_path_input.setAccessibleDescription(
            "Read-only field showing the path to the selected destination template"
        )
        file_grid.addWidget(self._dest_path_input, 1, 1)

        dest_browse_button = QPushButton("📂")
        dest_browse_button.setFixedWidth(36)
        dest_browse_button.setToolTip("Open a file picker to select the destination template (.docx)")
        dest_browse_button.setAccessibleName("Browse for destination template")
        dest_browse_button.setAccessibleDescription(
            "Open a file browser to choose the template document that will receive migrated content"
        )
        dest_browse_button.clicked.connect(self._on_browse_dest_clicked)
        file_grid.addWidget(dest_browse_button, 1, 2)

        file_grid.setColumnStretch(1, 1)
        layout.addLayout(file_grid)

        # ── Analyze button row ─────────────────────────────────────────────────
        # Layout: [stretch] [status label] [Analyze button]
        # The stretch pushes both the label and the button to the right edge,
        # so the status text appears immediately to the left of the button —
        # mirroring the layout of the Run Migration row in Step 3.
        analyze_row = QHBoxLayout()
        analyze_row.addStretch()
        self._analyze_status_label = QLabel("")
        self._analyze_status_label.setObjectName("status_label")
        self._analyze_status_label.setAccessibleName("Analysis status")
        analyze_row.addWidget(self._analyze_status_label)

        self._analyze_button = QPushButton("  Analyze Documents →")
        self._analyze_button.setObjectName("primary_btn")
        self._analyze_button.setEnabled(False)
        self._analyze_button.setToolTip(
            "Parse both documents, detect section headings, and auto-map them"
        )
        self._analyze_button.setAccessibleName("Analyze documents")
        self._analyze_button.setAccessibleDescription(
            "Parse the source and destination documents, detect all section headings, "
            "and automatically match source sections to destination sections"
        )
        self._analyze_button.clicked.connect(self._on_analyze_clicked)
        analyze_row.addWidget(self._analyze_button)

        layout.addLayout(analyze_row)
        return group

    # -------------------------------------------------------------------------
    # Step 2 — Review Section Mappings
    # -------------------------------------------------------------------------

    def _build_step2_group(self) -> QGroupBox:
        group = QGroupBox("  STEP 2 · REVIEW SECTION MAPPINGS   🔒")
        group.setObjectName("step_group")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(12, 16, 12, 12)
        layout.setSpacing(6)

        layout.addWidget(SectionLabel(
            "Review auto-detected section mappings. Use the Destination "
            "dropdown to reassign or skip sections — the top suggestions "
            "appear first, labelled with their confidence score. Use the "
            "Mode dropdown to choose how content is inserted. Red rows are "
            "unmapped and will be skipped unless assigned a destination."
        ))

        # ── Profile toolbar ────────────────────────────────────────────────────
        # Save Profile captures the current mapping state to a JSON file so it
        # can be reused for future migrations between the same document types.
        # Load Profile reads a previously saved file and applies those mappings
        # to the current table, using the auto-mapper result as a fallback for
        # any source sections not present in the profile.
        profile_toolbar = QHBoxLayout()
        profile_toolbar.setSpacing(8)

        self._save_profile_button = QPushButton("💾  Save Profile")
        self._save_profile_button.setToolTip(
            "Save the current section mappings as a reusable JSON profile"
        )
        self._save_profile_button.setAccessibleName("Save mapping profile")
        self._save_profile_button.setAccessibleDescription(
            "Save the current destination assignments and migration modes "
            "to a JSON file that can be reloaded for future migrations"
        )
        self._save_profile_button.clicked.connect(self._on_save_profile_clicked)
        profile_toolbar.addWidget(self._save_profile_button)

        self._load_profile_button = QPushButton("📂  Load Profile")
        self._load_profile_button.setToolTip(
            "Load a previously saved mapping profile from a JSON file"
        )
        self._load_profile_button.setAccessibleName("Load mapping profile")
        self._load_profile_button.setAccessibleDescription(
            "Load a previously saved JSON profile and apply its destination "
            "assignments to the current mapping table. Sections not in the "
            "profile keep their auto-mapped result."
        )
        self._load_profile_button.clicked.connect(self._on_load_profile_clicked)
        profile_toolbar.addWidget(self._load_profile_button)

        profile_toolbar.addStretch()
        layout.addLayout(profile_toolbar)

        # ── Mapping table ──────────────────────────────────────────────────────
        self._mapping_table_widget = MappingTableWidget()
        self._mapping_table_widget.countChanged.connect(self._on_mapping_counts_changed)
        layout.addWidget(self._mapping_table_widget)

        return group

    # -------------------------------------------------------------------------
    # Step 3 — Migration Options
    # -------------------------------------------------------------------------

    def _build_step3_group(self) -> QGroupBox:
        group = QGroupBox("  STEP 3 · MIGRATION OPTIONS   🔒")
        group.setObjectName("step_group")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(12, 16, 12, 12)
        layout.setSpacing(8)

        layout.addWidget(SectionLabel(
            "Configure output settings and run the migration. "
            "Per-section mode (Append / Prepend / Replace) is set in the "
            "table above. All migrated content will appear in green. "
            "Content matching template boilerplate will be highlighted blue."
        ))

        # ── Unmapped acknowledgement checkbox ──────────────────────────────────
        # Shown only when unmapped sections exist. The user must check this
        # box to confirm they understand those sections will be skipped.
        # The checkbox text is updated dynamically with the live unmapped count.
        # It is hidden (not just disabled) when the count is zero so it does
        # not take up visual space unnecessarily.
        self._unmapped_acknowledgement_checkbox = QCheckBox("")
        self._unmapped_acknowledgement_checkbox.setObjectName("ack_checkbox")
        self._unmapped_acknowledgement_checkbox.setVisible(False)
        self._unmapped_acknowledgement_checkbox.setAccessibleName(
            "Acknowledge unmapped sections"
        )
        self._unmapped_acknowledgement_checkbox.setAccessibleDescription(
            "Check this box to confirm you understand that unmapped sections "
            "will not be included in the migration output"
        )
        self._unmapped_acknowledgement_checkbox.stateChanged.connect(
            self._on_acknowledgement_checkbox_changed
        )
        layout.addWidget(self._unmapped_acknowledgement_checkbox)

        # ── Separator ─────────────────────────────────────────────────────────
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        layout.addWidget(separator)

        # ── Output settings grid ───────────────────────────────────────────────
        output_grid = QGridLayout()
        output_grid.setSpacing(8)

        self._export_report_checkbox = QCheckBox(
            "Export migration report (.csv) alongside the output file"
        )
        self._export_report_checkbox.setChecked(True)
        self._export_report_checkbox.setAccessibleName("Export migration report")
        self._export_report_checkbox.setAccessibleDescription(
            "When checked, a CSV audit report listing every mapping decision "
            "is saved alongside the output document"
        )
        output_grid.addWidget(self._export_report_checkbox, 0, 0, 1, 4)

        # ── Boilerplate highlighting mode ──────────────────────────────────────
        # Label acts as a visible group heading and is referenced by the radio
        # buttons' accessible descriptions so screen readers have full context.
        boilerplate_label = QLabel("Template Match Highlighting")
        boilerplate_label.setObjectName("boilerplate_mode_label")
        output_grid.addWidget(boilerplate_label, 1, 0, 1, 4)

        self._boilerplate_sentence_radio = QRadioButton("Matching sentences only")
        self._boilerplate_sentence_radio.setChecked(True)   # Default — most precise
        self._boilerplate_sentence_radio.setToolTip(
            "Color only the individual sentences that verbatim match the template "
            "boilerplate blue. Sentences that have been customized stay green, "
            "even if they are in the same paragraph as a matching sentence."
        )
        self._boilerplate_sentence_radio.setAccessibleName(
            "Highlight matching sentences only"
        )
        self._boilerplate_sentence_radio.setAccessibleDescription(
            "When selected, only the exact sentences that match the destination "
            "template boilerplate are colored blue. Other sentences in the same "
            "paragraph remain green, making partial customization clearly visible."
        )

        self._boilerplate_paragraph_radio = QRadioButton(
            "Entire paragraph if any sentence matches"
        )
        self._boilerplate_paragraph_radio.setToolTip(
            "Color the entire paragraph blue if it contains any sentence that "
            "verbatim matches the template boilerplate. Easier to scan at a "
            "glance but may flag paragraphs that have been partially customized."
        )
        self._boilerplate_paragraph_radio.setAccessibleName(
            "Highlight entire paragraph if any sentence matches"
        )
        self._boilerplate_paragraph_radio.setAccessibleDescription(
            "When selected, the entire paragraph is colored blue if any sentence "
            "within it matches the destination template boilerplate. Use this for "
            "a quick visual scan when per-sentence precision is not needed."
        )

        # QButtonGroup enforces mutual exclusivity and gives assistive technology
        # a single logical group to navigate rather than two independent buttons.
        self._boilerplate_mode_group = QButtonGroup(self)
        self._boilerplate_mode_group.addButton(
            self._boilerplate_sentence_radio, 0
        )
        self._boilerplate_mode_group.addButton(
            self._boilerplate_paragraph_radio, 1
        )

        boilerplate_radio_row = QHBoxLayout()
        boilerplate_radio_row.setSpacing(16)
        boilerplate_radio_row.addWidget(self._boilerplate_sentence_radio)
        boilerplate_radio_row.addWidget(self._boilerplate_paragraph_radio)
        boilerplate_radio_row.addStretch()
        output_grid.addLayout(boilerplate_radio_row, 2, 0, 1, 4)

        # Thin separator before the file output fields
        boilerplate_separator = QFrame()
        boilerplate_separator.setFrameShape(QFrame.HLine)
        boilerplate_separator.setFrameShadow(QFrame.Sunken)
        output_grid.addWidget(boilerplate_separator, 3, 0, 1, 4)

        # Output folder
        output_grid.addWidget(QLabel("Output Folder"), 4, 0)
        self._output_folder_input = QLineEdit()
        self._output_folder_input.setPlaceholderText("Select output folder…")
        self._output_folder_input.setAccessibleName("Output folder path")
        self._output_folder_input.setAccessibleDescription(
            "The folder where the migrated document and optional report will be saved"
        )
        output_grid.addWidget(self._output_folder_input, 4, 1, 1, 2)

        folder_browse_button = QPushButton("📂")
        folder_browse_button.setFixedWidth(36)
        folder_browse_button.setToolTip("Open a folder picker to choose the output location")
        folder_browse_button.setAccessibleName("Browse for output folder")
        folder_browse_button.setAccessibleDescription(
            "Open a folder browser to choose where the migrated document will be saved"
        )
        folder_browse_button.clicked.connect(self._on_browse_output_folder_clicked)
        output_grid.addWidget(folder_browse_button, 4, 3)

        # Output file name
        output_grid.addWidget(QLabel("File Name"), 5, 0)
        filename_row = QHBoxLayout()
        self._output_filename_input = QLineEdit()
        self._output_filename_input.setPlaceholderText("output_filename")
        self._output_filename_input.setAccessibleName("Output file name")
        self._output_filename_input.setAccessibleDescription(
            "The name of the output file without the .docx extension"
        )
        filename_row.addWidget(self._output_filename_input)
        filename_row.addWidget(QLabel(".docx"))
        output_grid.addLayout(filename_row, 5, 1, 1, 3)

        output_grid.setColumnStretch(1, 1)
        layout.addLayout(output_grid)

        # ── Run button row ─────────────────────────────────────────────────────
        run_row = QHBoxLayout()
        run_row.addStretch()

        self._migration_progress_bar = QProgressBar()
        self._migration_progress_bar.setVisible(False)
        self._migration_progress_bar.setMaximumWidth(180)
        self._migration_progress_bar.setRange(0, 0)   # Indeterminate spinner
        self._migration_progress_bar.setAccessibleName("Migration progress")
        run_row.addWidget(self._migration_progress_bar)

        self._migration_status_label = QLabel("")
        self._migration_status_label.setObjectName("status_label")
        self._migration_status_label.setAccessibleName("Migration status")
        run_row.addWidget(self._migration_status_label)

        self._run_migration_button = QPushButton("  ▶  Run Migration")
        self._run_migration_button.setObjectName("primary_btn")
        self._run_migration_button.setEnabled(False)
        self._run_migration_button.setToolTip(
            "Start the migration using the current mapping table settings"
        )
        self._run_migration_button.setAccessibleName("Run migration")
        self._run_migration_button.setAccessibleDescription(
            "Start the migration process. Mapped sections will be copied to the "
            "destination document. Unmapped and skipped sections are excluded."
        )
        self._run_migration_button.clicked.connect(self._on_run_migration_clicked)
        run_row.addWidget(self._run_migration_button)

        layout.addLayout(run_row)
        return group

    # -------------------------------------------------------------------------
    # Status bar
    # -------------------------------------------------------------------------

    def _build_status_bar(self) -> QFrame:
        """
        Slim dark bar at the bottom showing live status counts.

        macOS fix: QLabel ignores parent background-color via Qt stylesheets
        due to macOS's Cocoa native rendering layer overriding the Qt paint
        path. Setting WA_StyledBackground forces Qt to paint the background
        before Cocoa can override it. Setting background-color: transparent
        on the labels makes them inherit the frame's dark color instead of
        rendering white.
        """
        status_bar_frame = QFrame()
        status_bar_frame.setObjectName("status_bar")
        status_bar_frame.setFixedHeight(32)
        status_bar_frame.setAttribute(Qt.WA_StyledBackground, True)

        bar_layout = QHBoxLayout(status_bar_frame)
        bar_layout.setContentsMargins(12, 0, 12, 0)
        bar_layout.setSpacing(0)

        self._status_label_auto     = QLabel("🟢  – Auto-Mapped")
        self._status_label_review   = QLabel("🟡  – Needs Review")
        self._status_label_unmapped = QLabel("🔴  – Unmapped")
        self._status_label_skipped  = QLabel("⏭  – Skipped")
        self._status_label_changed  = QLabel("📝  – Changed")

        all_status_labels = [
            self._status_label_auto,
            self._status_label_review,
            self._status_label_unmapped,
            self._status_label_skipped,
            self._status_label_changed,
        ]

        for status_label in all_status_labels:
            status_label.setObjectName("sb_label")
            status_label.setAttribute(Qt.WA_StyledBackground, True)
            bar_layout.addWidget(status_label)

            divider = QFrame()
            divider.setFrameShape(QFrame.VLine)
            divider.setFrameShadow(QFrame.Sunken)
            divider.setAttribute(Qt.WA_StyledBackground, True)
            divider.setStyleSheet("background-color: #475569; max-width: 1px;")
            bar_layout.addWidget(divider)

        bar_layout.addStretch()
        return status_bar_frame

    # =========================================================================
    # Stylesheet
    # =========================================================================

    def _apply_styles(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f2f5;
                font-family: Segoe UI, Arial, sans-serif;
                font-size: 12px;
            }
            QGroupBox#step_group {
                font-weight: bold;
                font-size: 11px;
                color: #1e293b;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                margin-top: 6px;
                background-color: #ffffff;
            }
            QGroupBox#step_group:disabled {
                color: #94a3b8;
                background-color: #f8fafc;
                border-color: #e2e8f0;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 2px 8px;
                left: 10px;
            }
            QPushButton#primary_btn {
                background-color: #2563eb;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 7px 20px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton#primary_btn:hover    { background-color: #1d4ed8; }
            QPushButton#primary_btn:disabled {
                background-color: #bfdbfe;
                color: #eff6ff;
            }
            QPushButton {
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                padding: 4px 8px;
                background-color: #ffffff;
            }
            QPushButton:hover { background-color: #f1f5f9; }
            QLineEdit {
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                padding: 4px 8px;
                background-color: #ffffff;
            }
            QCheckBox#ack_checkbox {
                color: #b45309;
                font-weight: bold;
                font-size: 11px;
                background-color: #fef3c7;
                padding: 6px 10px;
                border: 1px solid #fcd34d;
                border-radius: 4px;
            }
            QTableWidget {
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                gridline-color: #e2e8f0;
                background-color: #ffffff;
                alternate-background-color: #f8fafc;
            }
            QHeaderView::section {
                background-color: #f1f5f9;
                border: none;
                border-bottom: 1px solid #cbd5e1;
                padding: 5px 8px;
                font-weight: bold;
                font-size: 11px;
                color: #374151;
            }
            QFrame#status_bar {
                background-color: #1e293b;
                border-top: 1px solid #334155;
            }
            QLabel#sb_label {
                color: #e2e8f0;
                font-size: 11px;
                padding: 0 14px;
                background-color: transparent;
            }
            QLabel#status_label {
                color: #64748b;
                font-size: 11px;
                font-style: italic;
            }
            QRadioButton { font-size: 11px; spacing: 4px; }
            QComboBox {
                border: 1px solid #cbd5e1;
                border-radius: 3px;
                padding: 2px 6px;
                font-size: 11px;
                background-color: #ffffff;
                min-height: 22px;
            }

            /* ================================================================
               Section 508 — Keyboard focus indicators
               Every interactive element must display a clearly visible focus
               ring when navigated to by keyboard (Tab / Shift-Tab / arrow keys).
               WCAG 2.1 SC 2.4.7 requires that the keyboard focus indicator
               is always visible. We use a 2 px solid blue outline to meet the
               3:1 contrast requirement against the white widget backgrounds.
               ================================================================ */

            QPushButton:focus {
                border: 2px solid #1d4ed8;
                outline: none;
            }
            QPushButton#primary_btn:focus {
                /* Primary buttons have a blue background, so use a lighter
                   contrasting ring instead of a darker one. */
                border: 2px solid #93c5fd;
                outline: none;
            }
            QLineEdit:focus {
                border: 2px solid #2563eb;
            }
            QCheckBox:focus {
                outline: 2px solid #2563eb;
                outline-offset: 2px;
            }
            QRadioButton:focus {
                outline: 2px solid #2563eb;
                outline-offset: 2px;
            }
            QComboBox:focus {
                border: 2px solid #2563eb;
            }
        """)

    # =========================================================================
    # Step lock / unlock helpers
    # =========================================================================

    def _set_step2_enabled(self, is_enabled: bool):
        """Enable or disable Step 2 and update its title icon."""
        self._step2_group.setEnabled(is_enabled)
        icon  = "✅" if is_enabled else "🔒"
        self._step2_group.setTitle(
            f"  STEP 2 · REVIEW SECTION MAPPINGS   {icon}"
        )

    def _set_step3_enabled(self, is_enabled: bool):
        """
        Enable or disable Step 3 and update its title icon.
        Uses ⚠ (caution) rather than 🔒 (lock) when enabled, because Step 3
        is accessible even when unmapped sections exist — the user just needs
        to acknowledge them via the checkbox.
        """
        self._step3_group.setEnabled(is_enabled)
        icon = "⚠" if is_enabled else "🔒"
        self._step3_group.setTitle(
            f"  STEP 3 · MIGRATION OPTIONS   {icon}"
        )

    # =========================================================================
    # File browser slots
    # =========================================================================

    def _on_browse_source_clicked(self):
        """Open a file dialog for the source document."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Source Document", "", "Word Documents (*.docx)"
        )
        if file_path:
            self._source_file_path = file_path
            self._source_path_input.setText(file_path)
            # Pre-fill the output folder with the source file's directory
            self._output_folder_input.setText(str(Path(file_path).parent))
            self._auto_populate_output_filename()
            self._update_analyze_button_state()

    def _on_browse_dest_clicked(self):
        """Open a file dialog for the destination template."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Destination Template", "", "Word Documents (*.docx)"
        )
        if file_path:
            self._dest_file_path = file_path
            self._dest_path_input.setText(file_path)
            self._update_analyze_button_state()

    def _on_browse_output_folder_clicked(self):
        """Open a folder dialog for the output location."""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder",
            self._output_folder_input.text() or str(Path.home()),
        )
        if folder_path:
            self._output_folder_input.setText(folder_path)

    def _update_analyze_button_state(self):
        """Enable Analyze only when both file paths are set."""
        self._analyze_button.setEnabled(
            bool(self._source_file_path) and bool(self._dest_file_path)
        )

    def _auto_populate_output_filename(self):
        """Pre-fill the output filename from the source file stem + date."""
        if self._source_file_path:
            source_stem = Path(self._source_file_path).stem
            date_string = datetime.now().strftime("%Y%m%d")
            self._output_filename_input.setText(
                f"{source_stem}_migrated_{date_string}"
            )

    # =========================================================================
    # Analyze
    # =========================================================================

    def _on_analyze_clicked(self):
        """Start the background document analysis."""
        self._analyze_button.setEnabled(False)
        self._analyze_status_label.setText("Analyzing documents…")

        # Reset previously displayed results
        self._mapping_table_widget.reset()
        self._set_step2_enabled(False)
        self._set_step3_enabled(False)
        self._run_migration_button.setEnabled(False)

        self._analyze_worker = AnalyzeWorker(
            self._source_file_path,
            self._dest_file_path,
        )
        self._analyze_worker.progress.connect(self._analyze_status_label.setText)
        self._analyze_worker.finished.connect(self._on_analyze_finished)
        self._analyze_worker.error.connect(self._on_analyze_error)
        self._analyze_worker.start()

    def _on_analyze_finished(
        self,
        source_sections,
        dest_sections,
        mapping_results,
        source_parser,
        dest_parser,
    ):
        """Handle successful analysis — populate the table and unlock steps."""
        self._dest_sections   = dest_sections
        self._mapping_results = mapping_results
        self._source_parser   = source_parser
        self._dest_parser     = dest_parser

        self._analyze_status_label.setText(
            f"Found {len(source_sections)} source sections and "
            f"{len(dest_sections)} destination sections."
        )
        self._analyze_button.setEnabled(True)

        self._mapping_table_widget.populate(mapping_results, dest_sections)
        self._set_step2_enabled(True)
        self._set_step3_enabled(True)

        # Trigger initial acknowledgement checkbox state
        unmapped_count = self._current_status_counts.get("unmapped", 0)
        self._update_acknowledgement_visibility(unmapped_count)

    def _on_analyze_error(self, error_message: str):
        """Handle analysis failure."""
        self._analyze_status_label.setText("Analysis failed.")
        self._analyze_button.setEnabled(True)
        QMessageBox.critical(
            self,
            "Analysis Error",
            f"Failed to analyze documents:\n\n{error_message}",
        )

    # =========================================================================
    # Mapping Profiles — Save and Load
    #
    # A mapping profile is a JSON file that records the current state of the
    # mapping table so it can be restored in future sessions. This is especially
    # useful when the same pair of document types (e.g. old template → new
    # template) will be migrated repeatedly — the user configures the mappings
    # once, saves the profile, and loads it for each subsequent migration.
    #
    # Profile format (version 1):
    #   {
    #     "version": 1,
    #     "created_at": "2026-05-11T14:30:00",
    #     "mappings": [
    #       {
    #         "source_title": "1. Introduction",
    #         "dest_title":   "1. Overview",
    #         "migration_mode": "append"
    #       },
    #       {
    #         "source_title": "2. Scope",
    #         "dest_title":   null,
    #         "migration_mode": null
    #       }
    #     ]
    #   }
    #
    # Matching on load is done by exact source title (case-sensitive).
    # If a profile entry's source title is not found in the current document,
    # or if the recorded destination title is not found in the current template,
    # the row is left unchanged (auto-mapper fallback applies).
    # =========================================================================

    def _on_save_profile_clicked(self):
        """
        Save the current mapping table state to a user-chosen JSON file.

        Called when the user clicks the "💾 Save Profile" button in Step 2.
        """
        current_results = self._mapping_table_widget.get_results()

        if not current_results:
            QMessageBox.warning(
                self,
                "No Mappings to Save",
                "Please analyze documents first before saving a profile.",
            )
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Mapping Profile",
            "",
            "Mapping Profiles (*.json)",
        )

        if not save_path:
            return   # User cancelled the file dialog

        profile_data = self._build_profile_dict(current_results)

        try:
            with open(save_path, "w", encoding="utf-8") as profile_file:
                json.dump(profile_data, profile_file, indent=2, ensure_ascii=False)

            QMessageBox.information(
                self,
                "Profile Saved",
                f"Mapping profile saved successfully.\n\nFile: {save_path}",
            )

        except Exception as save_error:
            QMessageBox.critical(
                self,
                "Save Error",
                f"Could not write the profile file:\n\n{save_error}",
            )

    def _build_profile_dict(self, mapping_results: list[MappingResult]) -> dict:
        """
        Serialize the current mapping table state into a plain Python dict
        ready for JSON serialization.

        We store the plain .title (not .display_title) so matching on load is
        not affected by heading-level indentation spaces. Titles are unique
        identifiers within a document, so exact title matching is reliable.

        Each entry records:
          source_title   — used to find the corresponding row on load
          dest_title     — the chosen destination, or null if unmapped/skipped
          migration_mode — the mode enum value string, or null if not applicable
        """
        mapping_entries = []

        for result in mapping_results:
            entry = {
                "source_title": result.source_section.title,
                "dest_title": (
                    result.dest_section.title
                    if result.dest_section is not None
                    else None
                ),
                "migration_mode": (
                    result.migration_mode.value
                    if result.migration_mode is not None
                    else None
                ),
            }
            mapping_entries.append(entry)

        return {
            "version":    1,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "mappings":   mapping_entries,
        }

    def _on_load_profile_clicked(self):
        """
        Load a saved mapping profile JSON file and apply it to the current table.

        Called when the user clicks the "📂 Load Profile" button in Step 2.
        After applying the profile the table is repopulated so all visual
        state (row colours, status icons, mode combos) reflects the changes.
        """
        current_results = self._mapping_table_widget.get_results()

        if not current_results:
            QMessageBox.warning(
                self,
                "No Mappings Loaded",
                "Please analyze documents first before loading a profile.",
            )
            return

        load_path, _ = QFileDialog.getOpenFileName(
            self,
            "Load Mapping Profile",
            "",
            "Mapping Profiles (*.json)",
        )

        if not load_path:
            return   # User cancelled the file dialog

        try:
            with open(load_path, "r", encoding="utf-8") as profile_file:
                profile_data = json.load(profile_file)

        except Exception as read_error:
            QMessageBox.critical(
                self,
                "Load Error",
                f"Could not read the profile file:\n\n{read_error}",
            )
            return

        # Basic format validation before attempting to apply
        if not isinstance(profile_data, dict) or "mappings" not in profile_data:
            QMessageBox.critical(
                self,
                "Invalid Profile",
                "This file does not appear to be a valid mapping profile.\n\n"
                "Expected a JSON object with a 'mappings' key.",
            )
            return

        applied_count, skipped_count = self._apply_profile(
            profile_data["mappings"],
            current_results,
        )

        # Repopulate the table so every row reflects the profile changes.
        # This rebuilds all combos (with correct initial selections and
        # suggestion items) and refreshes all status icons and row colours.
        self._mapping_table_widget.populate(current_results, self._dest_sections)

        QMessageBox.information(
            self,
            "Profile Loaded",
            f"Profile applied successfully.\n\n"
            f"  • {applied_count} section(s) updated from the profile.\n"
            f"  • {skipped_count} section(s) kept their auto-mapped result\n"
            f"    (source title not found in the profile, or the recorded\n"
            f"     destination section no longer exists in the template).",
        )

    def _apply_profile(
        self,
        profile_mappings: list[dict],
        mapping_results:  list[MappingResult],
    ) -> tuple[int, int]:
        """
        Apply the saved profile entries to the current list of MappingResult
        objects, modifying each result in place.

        Matching strategy:
          Profile entries are matched to current results by source section title
          (exact, case-sensitive). This is reliable because section titles are
          unique within a document and the source document has not changed.

        Destination resolution:
          The saved dest_title is looked up by exact title in self._dest_sections.
          If not found (the destination template has changed since the profile
          was saved), the row is left unchanged and counted as skipped.

        Status restoration:
          After setting dest_section, we check has_user_changed_destination:
            - If the profile destination matches the original auto-mapped one,
              we restore the original auto status (AUTO or REVIEW) rather than
              marking the row as MANUAL — preserving the auto-mapper's confidence
              classification.
            - If the destination differs from the original, the row is MANUAL.

        Returns:
          (applied_count, skipped_count)
          applied_count — rows successfully updated from the profile
          skipped_count — rows with no profile match, or an unresolvable destination
        """
        # Build a quick lookup from source title → profile entry
        profile_entry_by_source_title: dict[str, dict] = {
            entry["source_title"]: entry
            for entry in profile_mappings
            if isinstance(entry.get("source_title"), str)
        }

        # Build a quick lookup from destination section title → Section object
        dest_section_by_title: dict[str, Section] = {
            section.title: section
            for section in self._dest_sections
        }

        applied_count = 0
        skipped_count = 0

        for result in mapping_results:
            source_title  = result.source_section.title
            profile_entry = profile_entry_by_source_title.get(source_title)

            if profile_entry is None:
                # This source section was not present in the profile — leave
                # the row as the auto-mapper set it and count it as skipped.
                skipped_count += 1
                continue

            saved_dest_title  = profile_entry.get("dest_title")
            saved_mode_value  = profile_entry.get("migration_mode")

            if saved_dest_title is None:
                # Profile recorded this section as unmapped or skipped.
                # Treat it as unmapped — the user can manually skip if desired.
                result.dest_section   = None
                result.migration_mode = None
                result.status         = MappingStatus.UNMAPPED
                applied_count += 1
                continue

            # Resolve the saved destination title to a Section object
            dest_section = dest_section_by_title.get(saved_dest_title)

            if dest_section is None:
                # The destination section from the profile no longer exists in
                # the current template — fall back to the auto-mapper result.
                skipped_count += 1
                continue

            # Apply the resolved destination and let the change-detection
            # logic determine the correct status
            result.dest_section = dest_section

            if result.has_user_changed_destination:
                # Profile destination differs from the original auto-mapped one
                result.status = MappingStatus.MANUAL
            else:
                # Profile destination matches the original — restore auto status
                result.status = result.original_status

            # Apply the saved migration mode; fall back to APPEND if the
            # saved value string is no longer a valid MigrationMode member
            if saved_mode_value:
                try:
                    result.migration_mode = MigrationMode(saved_mode_value)
                except ValueError:
                    result.migration_mode = MigrationMode.APPEND
            else:
                result.migration_mode = MigrationMode.APPEND

            applied_count += 1

        return applied_count, skipped_count

    # =========================================================================
    # Mapping count changes
    # =========================================================================

    def _on_mapping_counts_changed(self, counts: dict):
        """
        Called by MappingTableWidget whenever any mapping changes.
        Updates the status bar labels, the acknowledgement checkbox,
        and the Migrate button state.

        If the user had already checked the acknowledgement checkbox and then
        changes a mapping (which updates the unmapped count), we reset the
        checkbox — they must re-read and re-acknowledge the new count.
        """
        self._current_status_counts = counts

        # Update status bar labels with live counts
        self._status_label_auto.setText(
            f"🟢  {counts.get('auto', 0)} Auto-Mapped"
        )
        self._status_label_review.setText(
            f"🟡  {counts.get('review', 0)} Needs Review"
        )
        self._status_label_unmapped.setText(
            f"🔴  {counts.get('unmapped', 0)} Unmapped"
        )
        self._status_label_skipped.setText(
            f"⏭  {counts.get('skipped', 0)} Skipped"
        )
        self._status_label_changed.setText(
            f"📝  {counts.get('changed', 0)} Changed"
        )

        # If the user already checked the acknowledgement and then changed
        # mappings, reset it so they must re-acknowledge the new count
        if self._user_acknowledged_unmapped_sections:
            self._user_acknowledged_unmapped_sections = False
            self._unmapped_acknowledgement_checkbox.blockSignals(True)
            self._unmapped_acknowledgement_checkbox.setChecked(False)
            self._unmapped_acknowledgement_checkbox.blockSignals(False)

        self._update_acknowledgement_visibility(counts.get("unmapped", 0))

    def _update_acknowledgement_visibility(self, unmapped_count: int):
        """
        Show or hide the acknowledgement checkbox based on unmapped count,
        then update the Migrate button state.

        When unmapped_count is 0:
          The checkbox is hidden (not just disabled) and the Migrate button
          enables directly — no acknowledgement needed.

        When unmapped_count > 0:
          The checkbox is shown with the live count in its text. The Migrate
          button stays disabled until the user checks the box.
        """
        if unmapped_count == 0:
            self._unmapped_acknowledgement_checkbox.setVisible(False)
        else:
            self._unmapped_acknowledgement_checkbox.setText(
                f"  You have {unmapped_count} unmapped section(s). "
                f"Check this box to acknowledge that they will be skipped "
                f"during migration."
            )
            self._unmapped_acknowledgement_checkbox.setVisible(True)

        self._update_migrate_button_state()

    def _update_migrate_button_state(self):
        """
        Enable or disable the Migrate button based on current conditions.

        The button enables when ALL of the following are true:
          - Step 3 is unlocked (analysis has completed)
          - Either: no unmapped sections exist
          - Or:     unmapped sections exist AND the user has checked the
                    acknowledgement checkbox
        """
        step3_is_active  = self._step3_group.isEnabled()
        unmapped_count   = self._current_status_counts.get("unmapped", 0)

        if not step3_is_active:
            self._run_migration_button.setEnabled(False)
            return

        if unmapped_count == 0:
            self._run_migration_button.setEnabled(True)
        else:
            self._run_migration_button.setEnabled(
                self._user_acknowledged_unmapped_sections
            )

    def _on_acknowledgement_checkbox_changed(self, checkbox_state: int):
        """
        Called when the unmapped acknowledgement checkbox is checked or unchecked.
        Updates the stored acknowledgement flag and the Migrate button state.
        """
        self._user_acknowledged_unmapped_sections = (checkbox_state == Qt.Checked)
        self._update_migrate_button_state()

    # =========================================================================
    # Migration
    # =========================================================================

    def _on_run_migration_clicked(self):
        """Validate output settings and start the background migration."""

        output_folder   = self._output_folder_input.text().strip()
        output_filename = self._output_filename_input.text().strip()

        if not output_folder:
            QMessageBox.warning(self, "Missing Output Folder",
                                "Please select an output folder.")
            return
        if not output_filename:
            QMessageBox.warning(self, "Missing File Name",
                                "Please enter an output file name.")
            return

        output_file_path = str(Path(output_folder) / f"{output_filename}.docx")

        if Path(output_file_path).exists():
            response = QMessageBox.question(
                self,
                "File Already Exists",
                f"'{output_filename}.docx' already exists.\nOverwrite it?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if response != QMessageBox.Yes:
                return

        # Re-parse both documents to get fresh python-docx Document objects.
        # The parser objects from analysis may have stale state if the user
        # ran analysis multiple times.
        try:
            source_parser = DocumentParser(self._source_file_path)
            source_parser.parse()
            dest_parser = DocumentParser(self._dest_file_path)
            dest_parser.parse()
        except Exception as error:
            QMessageBox.critical(self, "Parse Error",
                                 f"Could not open documents:\n{error}")
            return

        current_mapping_results = self._mapping_table_widget.get_results()

        migrator = ContentMigrator(
            mapping_results = current_mapping_results,
            source_document = source_parser.document,
            dest_document   = dest_parser.document,
            dest_sections   = self._dest_sections,
            output_path     = output_file_path,
        )

        report_csv_path = str(
            Path(output_folder) / f"{output_filename}_migration_report.csv"
        )
        report_exporter = MigrationReportExporter(current_mapping_results)

        self._run_migration_button.setEnabled(False)
        self._migration_progress_bar.setVisible(True)
        self._migration_status_label.setText("Running migration…")

        boilerplate_mode = (
            HIGHLIGHT_SENTENCES
            if self._boilerplate_sentence_radio.isChecked()
            else HIGHLIGHT_PARAGRAPH
        )

        self._migrate_worker = MigrateWorker(
            migrator=migrator,
            report_exporter=report_exporter,
            export_report=self._export_report_checkbox.isChecked(),
            report_output_path=report_csv_path,
            template_path=self._dest_file_path,
            boilerplate_mode=boilerplate_mode,
        )
        self._migrate_worker.progress.connect(self._migration_status_label.setText)
        self._migrate_worker.finished.connect(self._on_migration_finished)
        self._migrate_worker.error.connect(self._on_migration_error)
        self._migrate_worker.start()

    def _on_migration_finished(self, migration_report):
        """Display the migration completion summary dialog."""
        self._migration_progress_bar.setVisible(False)
        self._run_migration_button.setEnabled(True)
        self._migration_status_label.setText("Migration complete.")

        summary_lines = [
            "✅  Migration complete!\n",
            f"Sections migrated:   {migration_report.total_migrated}",
            f"Sections skipped:    {len(migration_report.skipped_sections)}",
            f"Sections unmapped:   {len(migration_report.unmapped_sections)}",
        ]

        if migration_report.errors:
            summary_lines.append(
                f"\n⚠️  {len(migration_report.errors)} error(s) occurred:"
            )
            for error_text in migration_report.errors[:5]:
                summary_lines.append(f"   • {error_text}")

        summary_lines.append(f"\nOutput saved to:\n{migration_report.output_path}")

        dialog = QMessageBox(self)
        dialog.setWindowTitle("Migration Complete")
        dialog.setIcon(
            QMessageBox.Information
            if not migration_report.errors
            else QMessageBox.Warning
        )
        dialog.setText("\n".join(summary_lines))

        open_file_button = dialog.addButton("Open Output File", QMessageBox.AcceptRole)
        dialog.addButton("Close", QMessageBox.RejectRole)
        dialog.exec_()

        if dialog.clickedButton() is open_file_button:
            self._open_file_in_os(migration_report.output_path)

    def _on_migration_error(self, error_message: str):
        """Handle migration failure."""
        self._migration_progress_bar.setVisible(False)
        self._run_migration_button.setEnabled(True)
        self._migration_status_label.setText("Migration failed.")
        QMessageBox.critical(
            self, "Migration Error",
            f"Migration failed:\n\n{error_message}",
        )

    def _open_file_in_os(self, file_path: str):
        """Open a file using the operating system's default application."""
        import platform
        import subprocess
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", file_path])
            else:
                subprocess.Popen(["xdg-open", file_path])
        except Exception:
            pass   # If the open fails the user still has the path from the dialog

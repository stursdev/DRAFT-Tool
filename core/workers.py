# =============================================================================
# core/workers.py
#
# Background worker threads for the Document Migration Tool.
#
# Why background threads:
#   Document parsing, auto-mapping, and migration all involve file I/O and
#   CPU-bound processing that can take several seconds on large documents.
#   Running these operations on the main thread would freeze the PyQt5 GUI
#   for the duration, making the application appear unresponsive.
#
#   PyQt5's QThread allows us to run these operations on a separate thread
#   while the main thread continues to process UI events. Workers communicate
#   results back to the UI via Qt signals, which are thread-safe by design.
#
# Workers defined here:
#   AnalyzeWorker  — parses both documents and runs the auto-mapper
#   MigrateWorker  — runs the migration engine, then the boilerplate detector
#
# These classes contain no UI logic. They belong in core/ rather than gui/
# because they orchestrate business logic, not presentation.
# =============================================================================

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from PyQt5.QtCore import QThread, pyqtSignal

from core.auto_mapper          import AutoMapper
from core.boilerplate_detector import BoilerplateDetector
from core.content_migrator     import ContentMigrator, MigrationReport
from core.document_parser      import DocumentParser
from core.audit_log            import MigrationReportExporter
from models.sections           import MappingResult


# =============================================================================
# AnalyzeWorker
#
# Parses the source and destination documents and runs the auto-mapper.
# Emits the results back to the UI via the finished signal.
# =============================================================================
class AnalyzeWorker(QThread):
    """
    Background thread for document analysis.

    Signals:
        finished — emitted on success with parsed sections, mapping results,
                   and parser objects. The UI uses these to populate the table.
        error    — emitted with an error message string if anything fails.
        progress — emitted with human-readable status text during processing
                   so the UI can display progress to the user.
    """

    # Signal payloads:
    #   source_sections  : list[Section]
    #   dest_sections    : list[Section]
    #   mapping_results  : list[MappingResult]
    #   source_parser    : DocumentParser (needed later by MigrateWorker)
    #   dest_parser      : DocumentParser (needed later by MigrateWorker)
    finished = pyqtSignal(list, list, list, object, object)
    error    = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, source_path: str, dest_path: str):
        super().__init__()
        self.source_path = source_path
        self.dest_path   = dest_path

    def run(self):
        """
        Execute analysis on the background thread.
        Called automatically by QThread when start() is invoked.
        """
        try:
            self.progress.emit("Parsing source document…")
            source_parser    = DocumentParser(self.source_path)
            source_sections  = source_parser.parse()

            self.progress.emit("Parsing destination template…")
            dest_parser      = DocumentParser(self.dest_path)
            dest_sections    = dest_parser.parse()

            self.progress.emit("Running auto-mapper…")
            mapper          = AutoMapper(source_sections, dest_sections)
            mapping_results = mapper.run()

            self.finished.emit(
                source_sections,
                dest_sections,
                mapping_results,
                source_parser,
                dest_parser,
            )

        except Exception as error:
            self.error.emit(str(error))


# =============================================================================
# MigrateWorker
#
# Runs the migration engine on the background thread, then automatically
# runs the boilerplate detector as a post-processing pass.
# Optionally exports the audit log CSV.
# =============================================================================
class MigrateWorker(QThread):
    """
    Background thread for document migration.

    Runs in sequence:
      1. ContentMigrator  — writes migrated content to the output file
      2. BoilerplateDetector — recolors boilerplate-matching sentences blue
      3. MigrationReportExporter — writes the CSV audit log (if enabled)

    Signals:
        finished         — emitted with the MigrationReport on success
        error            — emitted with an error message string on failure
        progress         — emitted with status text during processing
        boilerplate_done — emitted with the count of blue-highlighted sentences
    """

    finished         = pyqtSignal(object)   # MigrationReport
    error            = pyqtSignal(str)
    progress         = pyqtSignal(str)
    boilerplate_done = pyqtSignal(int)      # Number of sentences recolored blue

    def __init__(
        self,
        migrator:           ContentMigrator,
        report_exporter:    MigrationReportExporter,
        export_report:      bool,
        report_output_path: str,
        template_path:      str,
    ):
        super().__init__()
        self.migrator            = migrator
        self.report_exporter     = report_exporter
        self.export_report       = export_report
        self.report_output_path  = report_output_path

        # The original destination template path — needed by BoilerplateDetector
        # to extract boilerplate sentences for comparison
        self.template_path = template_path

    def run(self):
        """
        Execute migration, boilerplate detection, and optional report export
        on the background thread.
        """
        try:
            # ── Step 1: Migrate content ───────────────────────────────────────
            self.progress.emit("Migrating content…")
            migration_report = self.migrator.run()

            # ── Step 2: Boilerplate detection (automatic post-processing) ──────
            # Runs on the output file that was just written. Recolors any
            # green migrated sentences that exactly match template boilerplate
            # from green to blue.
            self.progress.emit("Detecting boilerplate matches…")
            detector = BoilerplateDetector(
                template_path=self.template_path,
                output_path=migration_report.output_path,
            )
            boilerplate_match_count = detector.run()
            self.boilerplate_done.emit(boilerplate_match_count)

            # ── Step 3: Export report CSV (if enabled) ────────────────────────
            if self.export_report:
                self.progress.emit("Exporting migration report…")
                self.report_exporter.export(self.report_output_path)

            self.finished.emit(migration_report)

        except Exception as error:
            self.error.emit(str(error))

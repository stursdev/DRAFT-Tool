# =============================================================================
# core/audit_log.py
#
# Exports a CSV migration report summarising what happened during migration.
#
# The report contains exactly the information programs need for traceability:
#   - Which source sections were migrated and where they went
#   - Which sections were skipped or left unmapped
#   - What migration mode was used for each section
#
# Columns in the output CSV:
#   source_section      — title of the source section
#   destination_section — title of the destination section, or
#                         "Unmapped" / "Skipped" if no destination was assigned
#   migration_mode      — Append, Prepend, Replace, or None
#   outcome             — Migrated, Skipped, or Unmapped
#
# The report is written using polars for reliable, dependency-free CSV output.
# =============================================================================

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import polars as pl

from models.sections import MappingResult, MappingStatus


class MigrationReportExporter:
    """
    Builds and exports a CSV migration report from a list of MappingResults.

    The report intentionally excludes confidence scores, timestamps, content
    type details, and other diagnostic fields — the goal is a clean, readable
    record of mapping decisions that programs can attach as documentation.

    Usage:
        exporter = MigrationReportExporter(mapping_results)
        exporter.export("output/migration_report.csv")
    """

    def __init__(self, mapping_results: list[MappingResult]):
        self.mapping_results = mapping_results

    def build_dataframe(self) -> pl.DataFrame:
        """
        Build a polars DataFrame with one row per MappingResult.

        Outcome values:
          "Migrated" — the section had a destination and will appear in output
          "Skipped"  — the user explicitly chose to skip this section
          "Unmapped" — no destination was assigned; section was excluded
        """
        rows = []

        for result in self.mapping_results:

            # Determine the destination label for the report
            if result.dest_section is not None:
                destination_label = result.dest_section.title
            elif result.status == MappingStatus.SKIPPED:
                destination_label = "Skipped"
            else:
                destination_label = "Unmapped"

            # Determine the outcome label
            if result.status == MappingStatus.SKIPPED:
                outcome = "Skipped"
            elif result.dest_section is None:
                outcome = "Unmapped"
            else:
                outcome = "Migrated"

            # Determine the migration mode label
            if result.migration_mode is not None:
                mode_label = result.migration_mode.value.capitalize()
            else:
                mode_label = "None"

            rows.append({
                "source_section":      result.source_section.title,
                "destination_section": destination_label,
                "migration_mode":      mode_label,
                "outcome":             outcome,
            })

        return pl.DataFrame(rows)

    def export(self, output_path: str) -> str:
        """
        Write the migration report to a CSV file.

        Creates the output directory if it does not already exist.
        Returns the path of the file that was written.
        """
        dataframe   = self.build_dataframe()
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        dataframe.write_csv(str(output_file))
        return str(output_file)

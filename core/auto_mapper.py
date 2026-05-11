# =============================================================================
# core/auto_mapper.py
#
# Automatically matches source sections to destination sections and produces
# an initial list of MappingResult objects for display in the mapping table.
#
# Matching algorithm — two steps per source section:
#
#   Step 1 — Score every destination:
#     All destination sections are scored against the source at once.
#     Exact normalized-title matches receive score 1.0 (maximum).
#     All other candidates are scored with rapidfuzz token_sort_ratio,
#     which sorts words alphabetically before comparing — handling word-order
#     differences well. Scores are divided by 100 to produce 0.0–1.0 values.
#     The result is a list sorted descending by score.
#
#   Step 2 — Classify and build the result:
#     The top-scoring candidate determines the row's primary status:
#       >= AUTO_THRESHOLD   (0.90) → AUTO    — accepted automatically
#       >= REVIEW_THRESHOLD (0.65) → REVIEW  — user should verify
#       <  REVIEW_THRESHOLD        → UNMAPPED — no primary match assigned
#
#     Top-N alternatives (above SUGGESTION_MIN_SCORE) are stored on the result
#     as top_suggestions so the destination dropdown can surface them as quick-
#     pick items at the top of the list.
#
# Suggestion scoring:
#   For matched rows  — suggestions are the top N candidates after the primary,
#                       scored above SUGGESTION_MIN_SCORE.
#   For unmapped rows — suggestions are the top N candidates overall (the row
#                       has no primary match, so nothing is excluded), also
#                       above SUGGESTION_MIN_SCORE. This gives users actionable
#                       guidance even for sections the auto-mapper couldn't
#                       confidently resolve.
#
# Original state snapshot:
#   Both original_dest_section and original_status are set at analysis time
#   and never modified afterward. The UI compares these against the live
#   dest_section / status to detect user changes without needing to track
#   every dropdown interaction.
# =============================================================================

import sys
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from rapidfuzz import fuzz

from models.sections import (
    MappingResult,
    MappingStatus,
    MigrationMode,
    MatchMethod,
    Section,
)


# -----------------------------------------------------------------------------
# Module-level thresholds
# Defined here so they are visible at the module level for documentation and
# can be overridden per-instance via the AutoMapper constructor.
# -----------------------------------------------------------------------------

# Matches at or above this score are auto-accepted without user review.
AUTO_THRESHOLD = 0.90

# Matches at or above this score (but below AUTO_THRESHOLD) are surfaced for
# user review. Matches below this score are treated as unmapped.
REVIEW_THRESHOLD = 0.65

# Minimum score required for a candidate to appear as a suggestion in the
# destination dropdown. Set lower than REVIEW_THRESHOLD so that unmapped
# sections — whose best candidate scored below 0.65 — can still show their
# closest options as suggestions rather than offering no guidance at all.
SUGGESTION_MIN_SCORE = 0.30

# Number of suggestion items to show at the top of each destination dropdown.
# Three gives the user meaningful alternatives without overwhelming the list.
DEFAULT_TOP_N_SUGGESTIONS = 3


class AutoMapper:
    """
    Matches each source section to the best available destination section.

    Usage:
        mapper  = AutoMapper(source_sections, destination_sections)
        results = mapper.run()   # Returns list[MappingResult], one per source
    """

    def __init__(
        self,
        source_sections:      list[Section],
        destination_sections: list[Section],
        auto_threshold:       float = AUTO_THRESHOLD,
        review_threshold:     float = REVIEW_THRESHOLD,
        suggestion_min_score: float = SUGGESTION_MIN_SCORE,
        top_n_suggestions:    int   = DEFAULT_TOP_N_SUGGESTIONS,
    ):
        self.source_sections      = source_sections
        self.destination_sections = destination_sections
        self.auto_threshold       = auto_threshold
        self.review_threshold     = review_threshold
        self.suggestion_min_score = suggestion_min_score
        self.top_n_suggestions    = top_n_suggestions

    # =========================================================================
    # Public API
    # =========================================================================

    def run(self) -> list[MappingResult]:
        """
        Run the matching algorithm for every source section.

        Returns one MappingResult per source section in source document order.
        Each result has original_dest_section and original_status set to the
        auto-mapper's suggestion — these are never changed after this point
        and serve as the baseline for detecting user edits later.
        Each result also has top_suggestions populated with the best
        alternative destinations, for use by the destination dropdown UI.
        """
        return [
            self._match_source_to_destination(source_section)
            for source_section in self.source_sections
        ]

    # =========================================================================
    # Primary matching logic
    # =========================================================================

    def _match_source_to_destination(self, source_section: Section) -> MappingResult:
        """
        Find the best matching destination for a single source section and
        collect alternative suggestions for the destination dropdown.

        The method scores all destination sections at once (rather than
        stopping early on the first good match) so that the suggestion list
        can be populated from the full ranked set of candidates.
        """
        source_normalized_title = source_section.normalized_title

        # A section with no meaningful title text (only punctuation, numbers,
        # or whitespace after normalization) cannot produce useful matches.
        if not source_normalized_title:
            return self._build_unmapped_result(
                source_section,
                top_suggestions=[],
            )

        # Score every destination section and get them sorted best-first
        all_scored_candidates = self._score_all_destinations(source_normalized_title)

        if not all_scored_candidates:
            return self._build_unmapped_result(
                source_section,
                top_suggestions=[],
            )

        best_section, best_score, best_match_method = all_scored_candidates[0]

        # ── Classify the primary match ────────────────────────────────────────
        if best_score >= self.auto_threshold:
            primary_status = MappingStatus.AUTO

        elif best_score >= self.review_threshold:
            primary_status = MappingStatus.REVIEW

        else:
            # Best score is below the meaningful match threshold — unmapped.
            # Still collect suggestions so the dropdown can offer guidance:
            # even if no confident match exists, the top candidates (above
            # SUGGESTION_MIN_SCORE) help the user make an informed decision.
            unmapped_suggestions = self._collect_top_suggestions(
                all_scored_candidates,
                exclude_first=False,   # No primary match to exclude
            )
            return self._build_unmapped_result(
                source_section,
                top_suggestions=unmapped_suggestions,
            )

        # ── Collect alternative suggestions for matched rows ──────────────────
        # The primary match (index 0) is excluded because it will already be
        # shown as the initially selected value in the destination dropdown.
        # Showing it again in the suggestions section would be redundant.
        alternative_suggestions = self._collect_top_suggestions(
            all_scored_candidates,
            exclude_first=True,   # Skip index 0 — that is the primary match
        )

        return self._build_matched_result(
            source_section=source_section,
            destination_section=best_section,
            confidence=best_score,
            match_method=best_match_method,
            status=primary_status,
            top_suggestions=alternative_suggestions,
        )

    # =========================================================================
    # Scoring helpers
    # =========================================================================

    def _score_all_destinations(
        self,
        source_normalized_title: str,
    ) -> list[tuple[Section, float, MatchMethod]]:
        """
        Score every destination section against the given normalized source title.

        Returns a list of (Section, score, MatchMethod) triples, sorted
        descending by score so index 0 is always the best candidate.

        Destination sections whose normalized title is empty are skipped —
        they would produce meaningless scores and pollute the suggestion list.
        """
        scored_candidates: list[tuple[Section, float, MatchMethod]] = []

        for destination_section in self.destination_sections:
            destination_normalized_title = destination_section.normalized_title

            if not destination_normalized_title:
                continue

            if source_normalized_title == destination_normalized_title:
                # Exact normalized match — maximum score, no fuzzy calculation needed.
                # We do NOT return early here (unlike the original single-pass design)
                # because we need to score ALL destinations to build a complete
                # suggestion list for the dropdown.
                scored_candidates.append(
                    (destination_section, 1.0, MatchMethod.EXACT)
                )
            else:
                # token_sort_ratio sorts words before comparing, so word-order
                # differences do not artificially deflate the score.
                # Example: "Configuration Management Plan" vs "Plan for
                # Configuration Management" → high score despite different order.
                fuzzy_score = fuzz.token_sort_ratio(
                    source_normalized_title,
                    destination_normalized_title,
                ) / 100.0

                scored_candidates.append(
                    (destination_section, fuzzy_score, MatchMethod.FUZZY)
                )

        # Sort descending so the best match is always at index 0
        scored_candidates.sort(key=lambda candidate: candidate[1], reverse=True)

        return scored_candidates

    def _collect_top_suggestions(
        self,
        all_scored_candidates: list[tuple[Section, float, MatchMethod]],
        exclude_first:         bool,
    ) -> list[tuple[Section, float]]:
        """
        Extract the top N suggestions from the sorted candidate list.

        exclude_first:
            True  — skip index 0 (the primary match for matched rows).
                    Using True prevents the primary destination from also
                    appearing in the suggestions list.
            False — include all candidates from index 0 onward
                    (used for unmapped rows that have no primary match).

        Only candidates scoring at or above self.suggestion_min_score are
        included. Because the list is already sorted descending, we can break
        as soon as a score drops below the floor — nothing after that qualifies.

        Returns a list of (Section, confidence_float) pairs already sorted
        descending by confidence, ready to insert into the combo dropdown.
        """
        start_index = 1 if exclude_first else 0
        suggestions: list[tuple[Section, float]] = []

        for destination_section, score, _ in all_scored_candidates[start_index:]:
            if score < self.suggestion_min_score:
                break   # Already below floor — no more qualifiers exist

            if len(suggestions) >= self.top_n_suggestions:
                break   # Collected enough suggestions — stop scanning

            suggestions.append((destination_section, score))

        return suggestions

    # =========================================================================
    # Result builders
    # =========================================================================

    def _build_matched_result(
        self,
        source_section:      Section,
        destination_section: Section,
        confidence:          float,
        match_method:        MatchMethod,
        status:              MappingStatus,
        top_suggestions:     list[tuple[Section, float]],
    ) -> MappingResult:
        """
        Build a MappingResult for a successfully matched source-destination pair.

        Both dest_section and original_dest_section are set to the same
        destination. The UI will update dest_section when the user makes
        changes; original_dest_section is the frozen baseline for comparison.

        Both status and original_status are set to the auto-determined status.
        original_status allows the UI to restore the correct status if the
        user reverts their changes back to the original destination.
        """
        return MappingResult(
            source_section=source_section,
            dest_section=destination_section,
            original_dest_section=destination_section,   # Frozen — never modified
            confidence=confidence,
            match_method=match_method,
            status=status,
            original_status=status,                       # Frozen — never modified
            migration_mode=MigrationMode.APPEND,          # Default for all mapped rows
            top_suggestions=top_suggestions,
        )

    def _build_unmapped_result(
        self,
        source_section:  Section,
        top_suggestions: list[tuple[Section, float]],
    ) -> MappingResult:
        """
        Build a MappingResult for a source section that could not be matched.

        Both dest_section and original_dest_section are None — there was no
        confident match from the start, so there is nothing to compare against.
        Migration mode is None because unmapped sections produce no output.

        top_suggestions may still contain candidates if any scored above
        SUGGESTION_MIN_SCORE — these help the user manually assign a destination.
        """
        return MappingResult(
            source_section=source_section,
            dest_section=None,
            original_dest_section=None,
            confidence=0.0,
            match_method=MatchMethod.NONE,
            status=MappingStatus.UNMAPPED,
            original_status=MappingStatus.UNMAPPED,
            migration_mode=None,
            top_suggestions=top_suggestions,
        )

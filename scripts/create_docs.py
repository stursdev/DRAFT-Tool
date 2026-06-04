"""
Generate DRAFT Tool Technical Documentation (.docx) and User Guide (.pptx).
Run with: python3 scripts/create_docs.py
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor as PPTXColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

OUTPUT_DIR = Path(__file__).parent.parent / "docs"
OUTPUT_DIR.mkdir(exist_ok=True)


# =============================================================================
# Helpers
# =============================================================================

def add_heading(doc, text, level):
    h = doc.add_heading(text, level=level)
    return h

def add_body(doc, text):
    p = doc.add_paragraph(text)
    p.style = doc.styles["Normal"]
    return p

def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(text, style="List Bullet")
    p.paragraph_format.left_indent = Pt(18 * (level + 1))
    return p

def add_subbullet(doc, text):
    return add_bullet(doc, text, level=1)

def add_code(doc, text):
    p = doc.add_paragraph(text)
    p.style = doc.styles["Normal"]
    for run in p.runs:
        run.font.name = "Courier New"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x37, 0x41, 0x51)
    # Shade background
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F1F5F9")
    pPr = p._element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p._element.insert(0, pPr)
    pPr.append(shd)
    return p

def add_table_row(table, cells):
    row = table.add_row()
    for i, text in enumerate(cells):
        row.cells[i].text = text
    return row

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


# =============================================================================
# TECHNICAL DOCUMENTATION
# =============================================================================

def build_technical_doc():
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ── Title page ────────────────────────────────────────────────────────────
    title = doc.add_heading("DRAFT Tool", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph("Technical Reference Documentation")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.size = Pt(14)
    sub.runs[0].font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    doc.add_paragraph("")
    version = doc.add_paragraph("Version 2.0  ·  Document Migration Tool")
    version.alignment = WD_ALIGN_PARAGRAPH.CENTER
    version.runs[0].font.size = Pt(10)
    version.runs[0].font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)

    doc.add_page_break()

    # ── 1. Overview ───────────────────────────────────────────────────────────
    add_heading(doc, "1. Overview", 1)
    add_body(doc,
        "DRAFT Tool (Document Review and Formatting Transfer Tool) is a desktop "
        "application that automates the migration of content from a legacy Word "
        "document into a new document template. It parses both documents, "
        "automatically matches sections by heading similarity, lets the user review "
        "and adjust the mappings, and then writes a merged output file with migrated "
        "content colored green and template-matching boilerplate highlighted blue."
    )

    add_heading(doc, "1.1 Purpose", 2)
    add_body(doc,
        "Programs regularly transition between document templates when standards, "
        "formats, or contract requirements change. Manually copying section content "
        "is error-prone and time-consuming. DRAFT Tool automates the section-matching "
        "step and provides a review interface so analysts can verify or override each "
        "mapping before committing to output."
    )

    add_heading(doc, "1.2 Technology Stack", 2)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    set_cell_bg(tbl.rows[0].cells[0], "1E3A5F")
    set_cell_bg(tbl.rows[0].cells[1], "1E3A5F")
    tbl.rows[0].cells[0].text = "Component"
    tbl.rows[0].cells[1].text = "Technology / Library"
    for cell in tbl.rows[0].cells:
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for row_data in [
        ("GUI Framework",     "PyQt5 — cross-platform desktop UI"),
        ("Document I/O",      "python-docx — read and write .docx files via OOXML"),
        ("Fuzzy Matching",    "rapidfuzz — token-sort-ratio string similarity"),
        ("Report Export",     "polars — fast DataFrame and CSV writing"),
        ("Python Version",    "Python 3.9+"),
        ("Packaging",         "PyInstaller — self-contained executable"),
    ]:
        add_table_row(tbl, row_data)

    doc.add_paragraph("")

    # ── 2. Project Structure ──────────────────────────────────────────────────
    add_heading(doc, "2. Project Structure", 1)
    add_body(doc, "The repository is organized into four top-level packages:")

    tbl2 = doc.add_table(rows=1, cols=2)
    tbl2.style = "Table Grid"
    set_cell_bg(tbl2.rows[0].cells[0], "1E3A5F")
    set_cell_bg(tbl2.rows[0].cells[1], "1E3A5F")
    tbl2.rows[0].cells[0].text = "Path"
    tbl2.rows[0].cells[1].text = "Responsibility"
    for cell in tbl2.rows[0].cells:
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for row_data in [
        ("main.py",                      "Entry point — creates QApplication and MainWindow"),
        ("models/sections.py",           "All shared data models and enums"),
        ("core/document_parser.py",      "Reads .docx files into Section objects"),
        ("core/auto_mapper.py",          "Matches source sections to destination sections"),
        ("core/content_migrator.py",     "Writes migrated content into the output document"),
        ("core/boilerplate_detector.py", "Recolors boilerplate-matching text blue"),
        ("core/audit_log.py",            "Exports migration report as CSV"),
        ("core/workers.py",              "Background QThread workers for long operations"),
        ("gui/main_window.py",           "Top-level QMainWindow shell with sidebar tab bar"),
        ("gui/migration_tab.py",         "Three-step migration workflow UI (QSplitter layout)"),
        ("gui/mapping_table.py",         "Interactive section mapping table widget"),
        ("gui/mapping_visualizer.py",    "Side panel — live source→destination diagram"),
        ("gui/widgets.py",               "Reusable custom widgets (combo, delegate, label)"),
        ("assets/",                      "Application icon files"),
        ("docs/PROJECT_CONTEXT.txt",     "Full project context for resuming development"),
    ]:
        add_table_row(tbl2, row_data)

    doc.add_paragraph("")

    # ── 3. Data Models ────────────────────────────────────────────────────────
    add_heading(doc, "3. Data Models  —  models/sections.py", 1)
    add_body(doc,
        "This file is the single source of truth for all data structures. Every "
        "other module imports from here. No magic strings or raw integers are used "
        "for state that crosses module boundaries."
    )

    add_heading(doc, "3.1 Enumerations", 2)

    for enum_name, desc, members in [
        ("BlockType", "Identifies the content type stored in a ContentBlock.", [
            ("PARAGRAPH", "A plain or formatted text paragraph"),
            ("TABLE",     "A Word table (rows and columns)"),
            ("IMAGE",     "A paragraph whose runs contain inline images"),
        ]),
        ("MatchMethod", "Records how the auto-mapper determined a mapping.", [
            ("EXACT",  "Normalized titles were identical"),
            ("FUZZY",  "rapidfuzz token_sort_ratio scored above threshold"),
            ("NONE",   "No destination scored high enough"),
            ("MANUAL", "User overrode the auto-mapper via the dropdown"),
        ]),
        ("MigrationMode", "Controls where content is inserted relative to boilerplate.", [
            ("APPEND",  "Content inserted after existing boilerplate (default)"),
            ("PREPEND", "Content inserted before existing boilerplate"),
            ("REPLACE", "Boilerplate deleted; migrated content replaces it"),
        ]),
        ("MappingStatus", "Current state of a row in the mapping table.", [
            ("AUTO",     "Auto-matched with high confidence (≥ 90%) — green row"),
            ("REVIEW",   "Auto-matched with medium confidence (65–89%) — yellow row"),
            ("UNMAPPED", "No destination assigned; section excluded — red row"),
            ("MANUAL",   "User manually changed the destination — blue row"),
            ("SKIPPED",  "User explicitly excluded this section — grey row"),
        ]),
    ]:
        add_heading(doc, f"  {enum_name}", 3)
        add_body(doc, desc)
        for val, d in members:
            add_bullet(doc, f"{val} — {d}")

    add_heading(doc, "3.2 ContentBlock", 2)
    add_body(doc,
        "Represents one element of content inside a section. Stores the "
        "original python-docx object (paragraph or table) rather than converting "
        "to a string early, so the migrator has full access to formatting when "
        "copying."
    )
    add_bullet(doc, "block_type: BlockType — what kind of content this holds")
    add_bullet(doc, "raw_object — the original python-docx Paragraph or Table object")

    add_heading(doc, "3.3 Section", 2)
    add_body(doc,
        "Represents one heading at any level (1–6) and all content directly "
        "beneath it. Produced by DocumentParser.parse()."
    )
    add_bullet(doc, "title: str — heading text exactly as it appears in the document")
    add_bullet(doc, "heading_level: int — Word heading level (1 = Heading 1, etc.)")
    add_bullet(doc, "paragraph_index: int — position of the heading in document.paragraphs")
    add_bullet(doc, "content_blocks: list[ContentBlock] — all content under this heading")
    add_bullet(doc, "display_title (property) — title with leading spaces for UI indentation")
    add_bullet(doc, "normalized_title (property) — stripped of numbers and punctuation; used for fuzzy matching only")

    add_heading(doc, "3.4 MappingResult", 2)
    add_body(doc,
        "Represents one row in the mapping table. Pairs a source section with "
        "its assigned destination. Tracks both the current user state and the "
        "original auto-mapper suggestion so change detection works correctly."
    )
    add_bullet(doc, "source_section: Section — always present")
    add_bullet(doc, "dest_section: Optional[Section] — current assigned destination (None = unmapped/skipped)")
    add_bullet(doc, "original_dest_section: Optional[Section] — frozen auto-mapper suggestion; never modified")
    add_bullet(doc, "confidence: float — 0.0–1.0 match score")
    add_bullet(doc, "match_method: MatchMethod — how the match was determined")
    add_bullet(doc, "status: MappingStatus — current row state")
    add_bullet(doc, "original_status: MappingStatus — frozen auto-mapper status; used to restore status on revert")
    add_bullet(doc, "migration_mode: Optional[MigrationMode] — how content is inserted (None for unmapped/skipped)")
    add_bullet(doc, "top_suggestions: list[tuple[Section, float]] — alternative destinations for the dropdown")
    add_bullet(doc, "has_user_changed_destination (property) — True if current dest differs from original")
    add_bullet(doc, "confidence_display (property) — formatted string for the Match column (e.g. '87%' or '–')")
    add_bullet(doc, "status_icon (property) — emoji for the Status column")

    # ── 4. Core Modules ───────────────────────────────────────────────────────
    add_heading(doc, "4. Core Modules", 1)

    add_heading(doc, "4.1 DocumentParser  —  core/document_parser.py", 2)
    add_body(doc,
        "Reads a .docx file and produces a flat ordered list of Section objects. "
        "Uses three heading detection strategies in priority order, preserves "
        "paragraph-table interleave order by walking raw XML, filters blank "
        "headings, and tags image paragraphs."
    )

    add_heading(doc, "  Constructor", 3)
    add_bullet(doc, "DocumentParser(filepath: str)")

    add_heading(doc, "  Heading Detection — Three Strategies", 3)
    add_body(doc,
        "Strategies are tried in order for every paragraph. The first one that "
        "returns a level wins; if all three return None the paragraph is not a heading."
    )
    for strategy, desc in [
        ("Strategy 1 — Style name prefix",
         "paragraph.style.name.lower().startswith('heading'). Extracts the trailing "
         "digit as the level. Catches all built-in Word heading styles: Heading 1, "
         "Heading 2, etc."),
        ("Strategy 2 — Paragraph XML <w:outlineLvl>",
         "Reads the <w:pPr><w:outlineLvl w:val='N'> element directly from the "
         "paragraph XML. val is 0-indexed; add 1 for the level. Catches custom-named "
         "styles where each paragraph has its outline level set explicitly."),
        ("Strategy 3 — Style inheritance chain (basedOn walk)",
         "_heading_level_from_style_chain() walks current.base_style repeatedly "
         "(guarded against circular chains via a visited-id set). At each ancestor "
         "checks both the ancestor style name and its style-definition <w:outlineLvl>. "
         "Catches custom styles like P_Heading_1_numbered that inherit from Heading 1 "
         "via Word's 'Style based on' setting but have neither a 'heading' name nor "
         "per-paragraph outline level. Previously a gap that caused entire documents "
         "to appear sectionless when custom heading styles were used."),
    ]:
        add_bullet(doc, strategy)
        add_subbullet(doc, desc)

    add_heading(doc, "  Key Methods", 3)
    for meth, desc in [
        ("parse() → list[Section]",
         "Opens the document, walks all body elements in document order, and returns "
         "a list of Section objects. Must be called before accessing .document."),
        ("document (property)",
         "Returns the underlying python-docx Document object. Available after parse(). "
         "The ContentMigrator reads this to access the source document directly."),
        ("_get_body_elements_in_order()",
         "Returns all paragraphs and tables in their true interleaved order by iterating "
         "the raw XML body children and mapping each to its python-docx object. Necessary "
         "because python-docx exposes paragraphs and tables as separate flat lists."),
        ("_get_heading_level(paragraph) → Optional[int]",
         "Runs the three detection strategies in sequence. Returns the heading level "
         "(1–6) from the first strategy that matches, or None for non-heading paragraphs."),
        ("_heading_level_from_style_chain(style) → Optional[int]",
         "Strategy 3 implementation. Walks the basedOn ancestry chain checking each "
         "ancestor style's name and pPr/outlineLvl. Uses a visited-id set to guard "
         "against malformed circular basedOn references."),
        ("_paragraph_contains_image(paragraph) → bool",
         "Checks each run for a <w:drawing> XML child element to detect inline images."),
    ]:
        add_bullet(doc, f"{meth}")
        add_subbullet(doc, desc)

    add_heading(doc, "4.2 AutoMapper  —  core/auto_mapper.py", 2)
    add_body(doc,
        "Matches each source section to the best available destination section "
        "using two-step scoring: exact normalized title match (score 1.0) or "
        "rapidfuzz token_sort_ratio. Produces a list of MappingResult objects "
        "with confidence scores, top suggestions, and frozen original state."
    )

    add_heading(doc, "  Thresholds", 3)
    add_bullet(doc, "AUTO_THRESHOLD = 0.90 — matches at or above are accepted automatically (green)")
    add_bullet(doc, "REVIEW_THRESHOLD = 0.65 — matches in range [0.65, 0.90) need user review (yellow)")
    add_bullet(doc, "SUGGESTION_MIN_SCORE = 0.30 — minimum score for a candidate to appear in the dropdown suggestions")
    add_bullet(doc, "DEFAULT_TOP_N_SUGGESTIONS = 3 — maximum number of suggestions shown per row")

    add_heading(doc, "  Key Methods", 3)
    for meth, desc in [
        ("run() → list[MappingResult]",
         "Entry point. Runs _match_source_to_destination for every source section "
         "and returns results in source document order."),
        ("_match_source_to_destination(source_section)",
         "Scores all destination sections, classifies the best match (AUTO/REVIEW/UNMAPPED), "
         "collects top-N alternative suggestions, and builds the MappingResult."),
        ("_score_all_destinations(normalized_title)",
         "Scores every destination against the source title. Exact matches receive 1.0. "
         "All others use fuzz.token_sort_ratio / 100. Returns sorted list (best first). "
         "Scores all destinations even after an exact match so suggestion list is complete."),
        ("_collect_top_suggestions(candidates, exclude_first)",
         "Extracts the top N suggestions above SUGGESTION_MIN_SCORE. exclude_first=True "
         "skips the primary match for matched rows; False includes all for unmapped rows."),
        ("_build_matched_result(...) / _build_unmapped_result(...)",
         "Construct MappingResult with both current and frozen-original fields set."),
    ]:
        add_bullet(doc, f"{meth}")
        add_subbullet(doc, desc)

    add_heading(doc, "4.3 ContentMigrator  —  core/content_migrator.py", 2)
    add_body(doc,
        "Executes the migration: creates a working copy of the destination template, "
        "groups source sections by their destination, finds section boundaries in the "
        "working document using stable XML element references, then inserts content "
        "according to each row's migration mode. All inserted content is colored green "
        "(RGB #00B050)."
    )

    add_heading(doc, "  Key Internal Structures", 3)
    add_bullet(doc, "MigrationReport — dataclass returned by run(); contains lists of migrated, skipped, and unmapped section titles plus any errors")
    add_bullet(doc, "SectionBoundary — holds XML element references (not paragraph indices) for a section's heading and next boundary; stable across insertions")
    add_bullet(doc, "_source_to_dest_num_id — dict caching source numId → dest numId; ensures each list numbering definition is copied once")

    add_heading(doc, "  Key Methods", 3)
    for meth, desc in [
        ("run() → MigrationReport",
         "Creates working copy, groups sources by destination, builds section boundary "
         "index, processes each destination, saves output file."),
        ("_create_working_copy()",
         "Saves destination template to a temp file and reloads it so all modifications "
         "go to the copy, never touching the original."),
        ("_build_section_boundary_index()",
         "Walks the working document headings and builds a dict of title → SectionBoundary. "
         "Uses XML element references (stable) rather than paragraph indices (unstable after insertions)."),
        ("_process_destination_section(dest_title, source_pairs, boundary_index)",
         "Migrates one or more source sections into a single destination. Enforces Replace "
         "exclusivity: only the first source in a group may use REPLACE; subsequent "
         "sources are demoted to APPEND."),
        ("_clear_section_content(boundary)",
         "Removes all XML elements between the heading element and the next boundary element. "
         "Works at the XML level so it handles both paragraphs and tables."),
        ("_copy_paragraph_with_green_text(source_paragraph)",
         "Creates a new paragraph copying style, paragraph format, run-level formatting "
         "(bold, italic, underline, font name/size), and list numbering. All runs are "
         "colored green."),
        ("_copy_list_numbering(source_paragraph, dest_paragraph)",
         "Copies the <w:numPr> (list numbering reference) from source to destination, "
         "registering the underlying abstractNum definition into the destination's "
         "numbering.xml with fresh non-conflicting IDs."),
        ("_copy_table_with_green_text_and_borders(source_table)",
         "Deep-copies the table XML element, sets all run colors to green, resets table "
         "indentation to zero, and applies explicit black borders to every cell so they "
         "render correctly regardless of the destination template's style sheet."),
        ("_copy_image_paragraph_with_border(source_paragraph)",
         "Extracts inline images from runs via <a:blip r:embed> relationship IDs, reads "
         "original width from <wp:extent cx>, creates a centered bordered image paragraph "
         "plus a separate green-text caption paragraph."),
    ]:
        add_bullet(doc, f"{meth}")
        add_subbullet(doc, desc)

    add_heading(doc, "4.4 BoilerplateDetector  —  core/boilerplate_detector.py", 2)
    add_body(doc,
        "Post-migration pass that recolors migrated (green) sentences blue when they "
        "verbatim match boilerplate already present in the same destination section of "
        "the template. Runs automatically after every migration."
    )

    add_heading(doc, "  Detection Modes", 3)
    add_bullet(doc, "HIGHLIGHT_SENTENCES (default) — color only the individual matching sentences blue; non-matching sentences in the same paragraph stay green")
    add_bullet(doc, "HIGHLIGHT_PARAGRAPH — color the entire paragraph blue if any sentence within it matches; faster to scan but may flag partially-customized paragraphs")

    add_heading(doc, "  Key Design Points", 3)
    add_bullet(doc, "Section-scoped matching — template sentences are indexed per section; green text in section X is only compared against sentences from section X in the template")
    add_bullet(doc, "Run splitting — Word stores text spanning sentence boundaries in one XML run. _split_and_color_run() splits a run at sentence boundaries by deep-copying the original element and inserting new elements after it, preserving all formatting")
    add_bullet(doc, "Whitespace normalization — all text is normalized (tabs/multiple spaces/non-breaking spaces → single space) before comparison; raw text is never modified")
    add_bullet(doc, "Separator guard — Qt separator items return empty string from itemText(); empty text always reverts to last valid selection")

    add_heading(doc, "  Key Methods", 3)
    for meth, desc in [
        ("run() → int",
         "Extracts template boilerplate, walks output document paragraphs, applies "
         "coloring to fully-green paragraphs, saves if any changes were made. Returns "
         "count of paragraphs where at least one sentence was recolored blue."),
        ("_extract_section_boilerplate() → dict[str, set[str]]",
         "Parses the template and builds {section_title: set_of_normalized_sentences}. "
         "Adds both the full paragraph text and each individual sentence so both "
         "whole-paragraph and partial-match detection work."),
        ("_apply_sentence_colors(paragraph, section_sentences) → bool",
         "Builds a character-position color map, then walks runs applying colors. "
         "Runs that straddle sentence boundaries are split in XML before coloring."),
        ("_split_and_color_run(run, split_positions, run_global_start, color_map)",
         "Modifies the original run element for the first segment, deep-copies it for "
         "subsequent segments, inserts them after the original in document order."),
    ]:
        add_bullet(doc, f"{meth}")
        add_subbullet(doc, desc)

    add_heading(doc, "4.5 MigrationReportExporter  —  core/audit_log.py", 2)
    add_body(doc,
        "Exports a CSV audit report using polars. Contains one row per source section "
        "with columns: source_section, destination_section, migration_mode, outcome. "
        "Written alongside the output .docx when the 'Export migration report' checkbox "
        "is checked."
    )
    add_bullet(doc, "build_dataframe() → polars.DataFrame — builds the data; used separately if the DataFrame is needed before writing")
    add_bullet(doc, "export(output_path: str) → str — writes CSV, creates parent directories if needed, returns the written path")

    add_heading(doc, "4.6 Workers  —  core/workers.py", 2)
    add_body(doc,
        "Background QThread workers that keep long-running operations off the main "
        "thread so the GUI stays responsive. Both workers communicate results back to "
        "the UI exclusively via Qt signals (thread-safe by design)."
    )

    add_heading(doc, "  AnalyzeWorker(QThread)", 3)
    add_body(doc,
        "Parses both documents and runs the auto-mapper. Emits finished(source_sections, "
        "dest_sections, mapping_results, source_parser, dest_parser) on success. "
        "Emits error(str) on failure. Emits progress(str) as status text during each step."
    )

    add_heading(doc, "  MigrateWorker(QThread)", 3)
    add_body(doc,
        "Runs three steps in sequence: ContentMigrator → BoilerplateDetector → "
        "MigrationReportExporter (if enabled). Emits finished(MigrationReport), "
        "error(str), progress(str), and boilerplate_done(int) (count of blue-highlighted "
        "paragraphs)."
    )

    # ── 5. GUI Modules ────────────────────────────────────────────────────────
    add_heading(doc, "5. GUI Modules", 1)

    add_heading(doc, "5.1 main_window.py", 2)
    add_body(doc,
        "Defines the top-level QMainWindow shell. Contains two classes:"
    )

    add_heading(doc, "  HorizontalTabBar(QTabBar)", 3)
    add_body(doc,
        "Custom tab bar that draws each tab as a vertically stacked icon + label "
        "layout (icon centered above text). Qt's default West-position tab bar "
        "rotates text 90 degrees; this override draws text horizontally regardless "
        "of tab position. Uses setShape(RoundedWest) to lay tabs top-to-bottom. "
        "Tab button dimensions: 110×82 px. Dark navy color scheme with a blue "
        "left-edge accent strip on the selected tab."
    )

    add_heading(doc, "  MainWindow(QMainWindow)", 3)
    add_body(doc,
        "Hosts the QTabWidget with HorizontalTabBar. Default window size is "
        "screen-percentage based: 88% of screen width × 90% of screen height, "
        "capped at screen minus 20/40 px to keep a sliver of desktop visible. "
        "Hard minimum: 860×600 px (reduced from the earlier 960×860 to accommodate "
        "13-14 inch laptop displays). Centers itself on the primary screen. "
        "Loads app_icon.ico from assets/. Currently contains one tab (Document "
        "Migration). Adding a new tab requires only one addTab() call in _build_ui()."
    )

    add_heading(doc, "5.2 migration_tab.py", 2)
    add_body(doc,
        "The full Document Migration workflow — three step groups plus a status bar. "
        "Uses a QSplitter(Horizontal) layout so the Mapping Visualization side panel "
        "can share space with the main content. The status bar sits outside the "
        "splitter so it always spans the full window width."
    )

    add_heading(doc, "  Layout Structure", 3)
    add_body(doc, "Outer QVBoxLayout contains:")
    add_bullet(doc, "QSplitter (Horizontal, handleWidth=1, both panes non-collapsible)")
    add_subbullet(doc, "Left pane — content_widget with Steps 1, 2, 3 in a QVBoxLayout")
    add_subbullet(doc, "Right pane — MappingVisualizerWidget (hidden by default)")
    add_bullet(doc, "Status bar QFrame (32px, dark navy) — outside the splitter, always full-width")

    add_heading(doc, "  State Fields", 3)
    for field, desc in [
        ("_source_file_path / _dest_file_path", "Paths selected by the user in Step 1"),
        ("_dest_sections", "List of destination Section objects from the last analysis"),
        ("_mapping_results", "Current list of MappingResult objects"),
        ("_source_parser / _dest_parser", "Kept after analysis; re-parsed fresh at migration time"),
        ("_user_acknowledged_unmapped_sections", "Flag for the unmapped acknowledgement checkbox"),
        ("_current_status_counts", "Latest count dict from MappingTableWidget.countChanged"),
        ("_visualizer_panel", "MappingVisualizerWidget instance in the splitter's right pane"),
        ("_visualize_button", "Checkable QPushButton that toggles the visualizer panel"),
    ]:
        add_bullet(doc, f"{field} — {desc}")

    add_heading(doc, "  Step Lock Logic", 3)
    add_body(doc,
        "Step 2 and Step 3 start locked (disabled). Step 2 unlocks when analysis "
        "completes. Step 3 unlocks at the same time but shows a ⚠ caution icon "
        "(not 🔒) because migration is allowed even with unmapped sections — the "
        "user just needs to check the acknowledgement checkbox."
    )

    add_heading(doc, "  Mapping Profiles", 3)
    add_body(doc,
        "Save Profile captures the current mapping table state to a JSON file (version 1 "
        "format: {version, created_at, mappings: [{source_title, dest_title, migration_mode}]}). "
        "Load Profile applies a saved file to the current table. Matching is by exact "
        "source_title. If the recorded destination no longer exists in the current template, "
        "the row is left unchanged. After applying a profile the table is repopulated from "
        "scratch so all visual state updates correctly."
    )

    add_heading(doc, "  Visualize Mapping Toggle", 3)
    add_body(doc,
        "The '🗺 Visualize Mapping' checkable button is placed below the mapping table "
        "in Step 2. When clicked on: seeds the visualizer panel with current results, "
        "makes it visible, sets the splitter to a 60/40 split. When clicked off: hides "
        "the panel. Every mapping change (countChanged signal) pushes a live update to "
        "the panel when it is open. Locking Step 2 resets the toggle and hides the panel."
    )

    add_heading(doc, "  Migration Flow", 3)
    add_body(doc,
        "On Run Migration: validates output folder/filename, checks for file overwrite, "
        "re-parses both documents fresh, builds ContentMigrator and MigrationReportExporter, "
        "starts MigrateWorker. On completion shows a summary dialog with counts of migrated/"
        "skipped/unmapped sections and an 'Open Output File' button that opens the file "
        "in the OS default application."
    )

    add_heading(doc, "5.3 mapping_table.py", 2)
    add_body(doc,
        "Composite widget containing filter radio buttons and the five-column mapping "
        "table. Emits countChanged(dict) whenever any mapping changes."
    )

    add_heading(doc, "  Table Columns", 3)
    for col, desc in [
        ("Source Section (stretch)",     "Source section title; level-1 headings are bold"),
        ("Destination Section (stretch)", "Searchable, editable combo with sentinels, suggestions, and full section list"),
        ("Mode (110 px fixed)",          "Migration mode combo: Append / Prepend / Replace"),
        ("Match (52 px fixed)",          "Confidence score display (e.g. '87%') or '–'"),
        ("Status (60 px fixed)",         "Status icon emoji"),
    ]:
        add_bullet(doc, f"{col} — {desc}")

    add_heading(doc, "  Destination Combo Structure", 3)
    add_body(doc, "Each destination combo contains items in this fixed order:")
    for item in [
        "[0] 'Unmapped' (sentinel value −1)",
        "[1] Separator",
        "[2] '⏭  Skip this section' (sentinel value −2)",
        "[3] Separator",
        "[4..N] ✨ Suggestion items (0–3, dynamically inserted with section index as data)",
        "[N+1] Separator (only if suggestions exist)",
        "[N+2..end] Full destination section list (section index as data, heading level for indentation)",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "  Replace Exclusivity Rule", 3)
    add_body(doc,
        "For each destination section, only the first row (in table order) that maps "
        "to it may use Replace mode. All other rows targeting the same destination have "
        "Replace greyed out. Ownership is position-based (first mapper owns Replace) to "
        "prevent the exploit where changing a second row's destination before selecting "
        "Replace would award ownership to the wrong row."
    )

    add_heading(doc, "  _SearchLineEdit(QLineEdit)", 3)
    add_body(doc,
        "Custom line edit installed on each destination combo via setLineEdit(). "
        "On focus-out, reverts the displayed text to the last confirmed valid selection "
        "if the current text does not match any combo item. Critical bug handled: "
        "Qt separator items return empty string from itemText() — the revert check "
        "only runs for non-empty text to prevent separators from being treated as "
        "valid selections and suppressing the revert."
    )

    add_heading(doc, "5.4 mapping_visualizer.py", 2)
    add_body(doc,
        "Side panel that renders a live diagram of the current section mappings. "
        "Toggled open/closed by the 'Visualize Mapping' button in Step 2. "
        "Displays ALL sections from both documents — not just mapped ones — "
        "so the user can see the full picture of what is and is not connected."
    )

    add_heading(doc, "  MappingVisualizerWidget(QWidget)", 3)
    add_body(doc,
        "Container widget. Layout: QScrollArea (stretch=1) + pinned legend bar (32px). "
        "The legend bar has a light #f8fafc background and shows the color key "
        "(Mapped / Review / Unmapped / Skipped) centered horizontally. It is always "
        "visible regardless of scroll position because it sits outside the scroll area."
    )
    add_bullet(doc, "update_mappings(results, dest_sections) — public API; pushes new data to the canvas and triggers repaint")

    add_heading(doc, "  MappingCanvas(QWidget) — custom painted", 3)
    add_body(doc,
        "Draws source boxes on the left (~42% of canvas width) and destination boxes "
        "on the right (~42%), with bezier connector curves in the middle. Both columns "
        "stack all sections in document order. resizeEvent triggers _recalculate() so "
        "the layout adapts when the splitter is dragged."
    )

    add_heading(doc, "  Source Box Colors", 3)
    for status, fill, border in [
        ("AUTO / MANUAL", "#DCFCE7 fill", "#16A34A border (green)"),
        ("REVIEW",        "#FEF3C7 fill", "#D97706 border (amber)"),
        ("UNMAPPED",      "#FEF2F2 fill", "#DC2626 border (red)"),
        ("SKIPPED",       "#EFF6FF fill", "#2563EB border (blue)"),
    ]:
        add_bullet(doc, f"{status} — {fill}, {border}")

    add_heading(doc, "  Destination Box Colors", 3)
    add_body(doc,
        "Destination boxes use only two colors — red and amber are not used because "
        "a destination section is either reachable or not:"
    )
    add_bullet(doc, "≥1 active source maps to it — green fill (#DCFCE7), green border (#16A34A)")
    add_bullet(doc, "Nothing maps to it — blue fill (#EFF6FF), blue border (#2563EB)")
    add_body(doc,
        "'Active' means the source's status is not UNMAPPED or SKIPPED — those "
        "statuses produce no migration output."
    )

    add_heading(doc, "  Connector Lines", 3)
    add_bullet(doc, "Cubic bezier curves from source box right-edge midpoint to destination box left-edge midpoint")
    add_bullet(doc, "Color = _STATUS_BORDER[result.status] (green, amber, red, or blue)")
    add_bullet(doc, "Only drawn for results where dest_section is not None AND status is not UNMAPPED or SKIPPED")
    add_bullet(doc, "Line width 1.6px, Qt.RoundCap style")

    add_heading(doc, "  _recalculate()", 3)
    add_body(doc,
        "Called on set_data() and resizeEvent. Computes all box positions. "
        "Source boxes: stacked top-to-bottom in self._results order. "
        "Destination boxes: all sections from self._dest_sections stacked top-to-bottom "
        "in document order. Box heights calculated via QFontMetrics.boundingRect() "
        "with Qt.TextWordWrap so long titles wrap correctly. "
        "Calls setMinimumHeight(total_content_height) so the QScrollArea knows the "
        "scrollable extent."
    )

    add_heading(doc, "5.5 widgets.py", 2)

    for widget, desc in [
        ("NoScrollComboBox(QComboBox)",
         "Overrides wheelEvent to ignore mouse wheel scrolling when the dropdown is "
         "closed. Prevents accidental value changes when the user scrolls the mapping "
         "table. Tracks popup state via a boolean flag set in showPopup/hidePopup."),
        ("IndentedComboDelegate(QStyledItemDelegate)",
         "Draws each destination combo item with left indentation proportional to "
         "heading level (14 px per level beyond level 1). Reads heading level from "
         "Qt.UserRole + 1 on each item. Sentinel items have level 0 (no indent)."),
        ("SectionLabel(QLabel)",
         "Small italic instruction label displayed below each step group box title. "
         "Word-wrapping enabled; uses 11 px dark grey italic text."),
    ]:
        add_heading(doc, f"  {widget}", 3)
        add_body(doc, desc)

    # ── 6. Data Flow ──────────────────────────────────────────────────────────
    add_heading(doc, "6. End-to-End Data Flow", 1)
    add_body(doc, "A complete migration follows these steps in sequence:")

    steps = [
        ("Step 1 — File Selection",
         "User selects source and destination .docx files via file pickers. "
         "Output folder is pre-filled from source directory. Output filename is "
         "pre-filled as '{source_stem}_migrated_{YYYYMMDD}'."),
        ("Step 2 — Analysis (AnalyzeWorker)",
         "DocumentParser reads both files. For the source: extracts Section objects "
         "with content blocks. For the destination: extracts Section objects only "
         "(no content needed — the destination boilerplate stays in place). "
         "AutoMapper scores every source-destination pair and produces MappingResult "
         "objects with status, confidence, and top suggestions."),
        ("Step 3 — Review (MappingTableWidget)",
         "Results populate the table. User can reassign destinations (combo), "
         "change modes (Append/Prepend/Replace), filter by status, save/load profiles. "
         "countChanged fires on every change; status bar updates live."),
        ("Step 4 — Migration (MigrateWorker → ContentMigrator)",
         "Documents are re-parsed fresh. ContentMigrator creates a working copy of "
         "the destination template, groups sources by destination, finds section "
         "boundaries via XML element references, then for each destination: clears "
         "content (Replace), or finds insertion point (Append/Prepend), and inserts "
         "all mapped source content in green."),
        ("Step 5 — Boilerplate Detection (BoilerplateDetector)",
         "Runs on the output file automatically. Walks every green paragraph. If any "
         "sentence verbatim matches the template boilerplate for that section, it is "
         "recolored blue. Runs straddling sentence boundaries are split in XML first."),
        ("Step 6 — Report Export (MigrationReportExporter, optional)",
         "If 'Export migration report' is checked, a CSV is written alongside the "
         "output file with one row per source section recording the destination, "
         "mode, and outcome."),
    ]
    for title, desc in steps:
        add_bullet(doc, title)
        add_subbullet(doc, desc)

    # ── 7. Configuration & Thresholds ─────────────────────────────────────────
    add_heading(doc, "7. Tunable Constants", 1)
    add_body(doc,
        "The following constants control matching behavior. Change them in their "
        "respective files if the default thresholds do not suit the document corpus."
    )

    tbl3 = doc.add_table(rows=1, cols=3)
    tbl3.style = "Table Grid"
    set_cell_bg(tbl3.rows[0].cells[0], "1E3A5F")
    set_cell_bg(tbl3.rows[0].cells[1], "1E3A5F")
    set_cell_bg(tbl3.rows[0].cells[2], "1E3A5F")
    for i, hdr in enumerate(["Constant", "Default", "Effect"]):
        tbl3.rows[0].cells[i].text = hdr
        tbl3.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        tbl3.rows[0].cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for row_data in [
        ("AUTO_THRESHOLD",           "0.90", "Scores ≥ this are auto-accepted (green)"),
        ("REVIEW_THRESHOLD",         "0.65", "Scores in [0.65, 0.90) need review (yellow)"),
        ("SUGGESTION_MIN_SCORE",     "0.30", "Minimum score to appear as a dropdown suggestion"),
        ("DEFAULT_TOP_N_SUGGESTIONS","3",    "Max suggestions shown per destination combo"),
        ("MAX_HEADING_LEVEL",        "6",    "Deepest heading level the parser will detect"),
        ("INDENT_PIXELS_PER_LEVEL",  "14",   "Extra left indent per heading level in combo dropdown"),
        ("MINIMUM_WINDOW_WIDTH",     "860",  "Hard minimum window width (px)"),
        ("MINIMUM_WINDOW_HEIGHT",    "600",  "Hard minimum window height (px)"),
        ("TAB_BUTTON_WIDTH",         "110",  "Width of sidebar tab buttons (px)"),
        ("TAB_BUTTON_HEIGHT",        "82",   "Height of sidebar tab buttons (px)"),
        ("_FONT_SIZE (visualizer)",  "10 pt","Box text and header font size in visualizer"),
        ("_BOX_GAP (visualizer)",    "8 px", "Vertical gap between boxes in visualizer"),
        ("_LINE_WIDTH (visualizer)", "1.6 px","Connector line stroke width in visualizer"),
    ]:
        add_table_row(tbl3, row_data)

    doc.add_paragraph("")

    # ── 8. Known Issues & Gotchas ─────────────────────────────────────────────
    add_heading(doc, "8. Known Issues and Maintenance Notes", 1)

    for title, detail in [
        ("Qt separator itemText() returns empty string",
         "QComboBox separator items return '' from itemText(i). Any loop checking "
         "item text for validity must guard against empty text or separators will "
         "be incorrectly treated as valid selections. See _SearchLineEdit._revert_if_invalid()."),
        ("setCurrentIndex() is a no-op when index unchanged",
         "After the line edit text is cleared programmatically without changing "
         "the combo's stored index, calling setCurrentIndex(same_idx) does nothing. "
         "Use setEditText() to force the visual text to update."),
        ("Qt setLineEdit() Python GC ownership",
         "setLineEdit() gives Qt C++ ownership of the line edit. Store a Python "
         "reference on the combo (combo._search_line_edit = le) to prevent the "
         "Python wrapper from being garbage-collected while Qt holds only the C++ object."),
        ("Custom heading styles — resolved by Strategy 3",
         "Documents using styles like P_Heading_1_numbered that inherit from built-in "
         "Heading styles via Word's 'Style based on' setting are now fully detected by "
         "the basedOn chain walk (Strategy 3 in _get_heading_level). This was a gap "
         "that previously caused entire documents with custom heading styles to appear "
         "sectionless."),
        ("Same destination heading appearing twice",
         "_build_section_boundary_index() indexes only the first occurrence of each "
         "heading title. Documents with duplicate heading titles will only have "
         "content inserted into the first matching section."),
        ("numId cross-document conflict",
         "List numbering definitions use document-local integer IDs. ContentMigrator "
         "copies abstractNum definitions into the destination's numbering.xml with "
         "fresh non-conflicting IDs. Malformed numbering XML is caught and silently "
         "skipped — the paragraph migrates without list formatting."),
        ("Visualizer legend alignment",
         "The color-key legend for the visualizer panel must live inside the "
         "MappingVisualizerWidget (below the QScrollArea), NOT in the main status bar. "
         "The status bar sits outside the QSplitter; the visualizer panel is inside it. "
         "Placing the legend in the status bar creates a height mismatch because the "
         "two widgets have no shared Y reference when content heights differ."),
        ("PyInstaller --onefile startup time",
         "--onefile extracts all bundled files to a temp directory on every launch, "
         "causing approximately 10 second cold-start times. Use --onedir for faster "
         "startup at the cost of distributing a folder rather than a single file."),
    ]:
        add_heading(doc, f"  {title}", 3)
        add_body(doc, detail)

    # ── 9. Packaging ──────────────────────────────────────────────────────────
    add_heading(doc, "9. Packaging with PyInstaller", 1)
    add_body(doc,
        "All Python packages installed at build time are bundled into the executable "
        "or distribution folder by PyInstaller. The application does not download or "
        "install any packages at runtime — all imports are resolved from the bundle."
    )
    add_bullet(doc, "Entry point: main.py")
    add_bullet(doc, "Include assets/ folder using the --add-data flag")
    add_bullet(doc, "Icon: assets/app_icon.ico (contains 16, 24, 32, 48, 64, 128, 256 px sizes)")
    add_bullet(doc, "Hidden imports may be needed for: docx, rapidfuzz, polars, PyQt5")
    add_body(doc, "Example one-file build command:")
    add_code(doc,
        "pyinstaller --onefile --windowed --icon=assets/app_icon.ico \\\n"
        "  --add-data 'assets:assets' main.py"
    )

    out = OUTPUT_DIR / "DRAFT_Tool_Technical_Documentation.docx"
    doc.save(str(out))
    print(f"Saved: {out}")
    return out


# =============================================================================
# USER GUIDE POWERPOINT
# =============================================================================

def rgb(r, g, b):
    return PPTXColor(r, g, b)

NAVY   = rgb(0x1E, 0x29, 0x3B)
BLUE   = rgb(0x25, 0x63, 0xEB)
LTBLUE = rgb(0xDB, 0xEA, 0xFE)
WHITE  = rgb(0xFF, 0xFF, 0xFF)
GREY   = rgb(0x64, 0x74, 0x8B)
LGREY  = rgb(0xF1, 0xF5, 0xF9)
GREEN  = rgb(0x00, 0xB0, 0x50)
AMBER  = rgb(0xD9, 0x77, 0x06)
RED    = rgb(0xDC, 0x26, 0x26)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def add_slide(prs, layout_idx=6):
    layout = prs.slide_layouts[layout_idx]
    return prs.slides.add_slide(layout)


def txt_box(slide, text, l, t, w, h, size=18, bold=False, color=None, align=PP_ALIGN.LEFT, wrap=True):
    txb = slide.shapes.add_textbox(l, t, w, h)
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return txb


def filled_rect(slide, l, t, w, h, fill_color):
    from pptx.util import Emu
    shape = slide.shapes.add_shape(
        pptx.enum.shapes.MSO_SHAPE_TYPE.AUTO_SHAPE if False else 1,  # MSO_SHAPE.RECTANGLE
        l, t, w, h
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def title_slide(prs):
    slide = add_slide(prs, 0)
    # Full-bleed navy background
    bg = filled_rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)

    # Blue accent bar left
    filled_rect(slide, 0, 0, Inches(0.12), SLIDE_H, BLUE)

    # Title
    txt_box(slide, "DRAFT Tool", Inches(1), Inches(1.8), Inches(11), Inches(1.2),
            size=54, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Subtitle
    txt_box(slide, "Document Review and Formatting Transfer Tool",
            Inches(1), Inches(3.1), Inches(11), Inches(0.7),
            size=20, bold=False, color=rgb(0xCB, 0xD5, 0xE1), align=PP_ALIGN.LEFT)

    # Tagline
    txt_box(slide, "Automate content migration between Word document templates",
            Inches(1), Inches(3.9), Inches(10), Inches(0.6),
            size=14, color=rgb(0x94, 0xA3, 0xB8), align=PP_ALIGN.LEFT)

    # User Guide label bottom right
    txt_box(slide, "User Guide", Inches(10), Inches(6.7), Inches(3), Inches(0.5),
            size=11, color=GREY, align=PP_ALIGN.RIGHT)


def section_divider(prs, number, title):
    slide = add_slide(prs, 0)
    filled_rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    filled_rect(slide, 0, 0, Inches(0.12), SLIDE_H, BLUE)
    txt_box(slide, number, Inches(1), Inches(2.2), Inches(1.5), Inches(1.2),
            size=72, bold=True, color=BLUE)
    txt_box(slide, title, Inches(2.5), Inches(2.6), Inches(10), Inches(1.2),
            size=36, bold=True, color=WHITE)


def content_slide(prs, title, bullets, notes=""):
    slide = add_slide(prs, 0)
    filled_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)

    # Header bar
    filled_rect(slide, 0, 0, SLIDE_W, Inches(1.1), NAVY)
    txt_box(slide, title, Inches(0.4), Inches(0.15), Inches(12), Inches(0.8),
            size=22, bold=True, color=WHITE)

    # Accent line under header
    filled_rect(slide, 0, Inches(1.1), SLIDE_W, Inches(0.04), BLUE)

    # Bullet content
    y = Inches(1.3)
    for bullet in bullets:
        if bullet.startswith("##"):  # Sub-bullet
            text = bullet[2:].strip()
            filled_rect(slide, Inches(0.7), y + Inches(0.08), Inches(0.06), Inches(0.06), BLUE)
            txt_box(slide, text, Inches(0.95), y, Inches(11.8), Inches(0.42),
                    size=13, color=GREY)
            y += Inches(0.42)
        elif bullet.startswith("!"):  # Highlight box
            text = bullet[1:].strip()
            box = filled_rect(slide, Inches(0.4), y, Inches(12.5), Inches(0.5), LTBLUE)
            txt_box(slide, text, Inches(0.6), y + Inches(0.04), Inches(12), Inches(0.42),
                    size=13, bold=True, color=BLUE)
            y += Inches(0.55)
        else:
            # Main bullet
            filled_rect(slide, Inches(0.4), y + Inches(0.14), Inches(0.12), Inches(0.12), BLUE)
            txt_box(slide, bullet, Inches(0.7), y, Inches(12.1), Inches(0.48),
                    size=15, color=NAVY)
            y += Inches(0.5)

    return slide


def two_col_slide(prs, title, left_title, left_items, right_title, right_items):
    slide = add_slide(prs, 0)
    filled_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    filled_rect(slide, 0, 0, SLIDE_W, Inches(1.1), NAVY)
    txt_box(slide, title, Inches(0.4), Inches(0.15), Inches(12), Inches(0.8),
            size=22, bold=True, color=WHITE)
    filled_rect(slide, 0, Inches(1.1), SLIDE_W, Inches(0.04), BLUE)

    # Left column
    filled_rect(slide, Inches(0.3), Inches(1.25), Inches(6.1), Inches(5.9), LGREY)
    txt_box(slide, left_title, Inches(0.45), Inches(1.35), Inches(5.8), Inches(0.45),
            size=14, bold=True, color=NAVY)
    filled_rect(slide, Inches(0.45), Inches(1.82), Inches(5.8), Inches(0.03), BLUE)
    y = Inches(1.95)
    for item in left_items:
        filled_rect(slide, Inches(0.55), y + Inches(0.10), Inches(0.1), Inches(0.1), GREEN)
        txt_box(slide, item, Inches(0.78), y, Inches(5.4), Inches(0.42), size=12, color=NAVY)
        y += Inches(0.44)

    # Right column
    filled_rect(slide, Inches(6.9), Inches(1.25), Inches(6.1), Inches(5.9), LGREY)
    txt_box(slide, right_title, Inches(7.05), Inches(1.35), Inches(5.8), Inches(0.45),
            size=14, bold=True, color=NAVY)
    filled_rect(slide, Inches(7.05), Inches(1.82), Inches(5.8), Inches(0.03), BLUE)
    y = Inches(1.95)
    for item in right_items:
        filled_rect(slide, Inches(7.15), y + Inches(0.10), Inches(0.1), Inches(0.1), BLUE)
        txt_box(slide, item, Inches(7.38), y, Inches(5.4), Inches(0.42), size=12, color=NAVY)
        y += Inches(0.44)

    return slide


def color_legend_slide(prs):
    slide = add_slide(prs, 0)
    filled_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    filled_rect(slide, 0, 0, SLIDE_W, Inches(1.1), NAVY)
    txt_box(slide, "Understanding the Output Document", Inches(0.4), Inches(0.15),
            Inches(12), Inches(0.8), size=22, bold=True, color=WHITE)
    filled_rect(slide, 0, Inches(1.1), SLIDE_W, Inches(0.04), BLUE)

    intro = ("After migration, your output document uses two text colors to help "
             "you quickly identify what is original template content versus what was migrated.")
    txt_box(slide, intro, Inches(0.5), Inches(1.25), Inches(12.3), Inches(0.6),
            size=14, color=NAVY)

    entries = [
        (GREEN,          "Green Text — Migrated Content",
         "All content copied from your source document appears in green. "
         "This makes it easy to distinguish newly migrated text from the "
         "original template boilerplate."),
        (rgb(0,0x70,0xC0), "Blue Text — Template Boilerplate Match",
         "Migrated sentences that are verbatim identical to boilerplate "
         "already in the destination section are recolored blue automatically. "
         "Blue text means 'this was in the template too — no new content here.'"),
        (rgb(0,0,0),     "Black Text — Original Template Content",
         "Text that was already in the destination template and was not "
         "touched by migration remains black. This is your baseline boilerplate."),
    ]

    y = Inches(2.05)
    for color, title, desc in entries:
        filled_rect(slide, Inches(0.4), y, Inches(0.35), Inches(0.9), color)
        txt_box(slide, title, Inches(0.9), y, Inches(11.8), Inches(0.4),
                size=15, bold=True, color=color)
        txt_box(slide, desc, Inches(0.9), y + Inches(0.38), Inches(11.5), Inches(0.5),
                size=12, color=GREY)
        y += Inches(1.1)


def status_legend_slide(prs):
    slide = add_slide(prs, 0)
    filled_rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)
    filled_rect(slide, 0, 0, SLIDE_W, Inches(1.1), NAVY)
    txt_box(slide, "Mapping Status Colors at a Glance", Inches(0.4), Inches(0.15),
            Inches(12), Inches(0.8), size=22, bold=True, color=WHITE)
    filled_rect(slide, 0, Inches(1.1), SLIDE_W, Inches(0.04), BLUE)

    rows = [
        (rgb(0xDC,0xFC,0xE7), GREEN,          "🟢 Auto-Mapped",
         "Section was matched automatically with ≥ 90% confidence. Safe to proceed."),
        (rgb(0xFF,0xFB,0xEB), AMBER,          "🟡 Needs Review",
         "Section was matched with 65–89% confidence. Verify the destination is correct before migrating."),
        (rgb(0xFE,0xF2,0xF2), RED,            "🔴 Unmapped",
         "No destination was found or assigned. This section will be skipped unless you assign one."),
        (rgb(0xEF,0xF6,0xFF), BLUE,           "📝 Changed",
         "You manually changed the destination from the auto-mapper's suggestion."),
        (rgb(0xF3,0xF4,0xF6), GREY,           "⏭  Skipped",
         "You explicitly chose to exclude this section from migration."),
    ]

    y = Inches(1.3)
    for bg, accent, label, desc in rows:
        filled_rect(slide, Inches(0.3), y, Inches(12.7), Inches(0.95), bg)
        filled_rect(slide, Inches(0.3), y, Inches(0.12), Inches(0.95), accent)
        txt_box(slide, label, Inches(0.55), y + Inches(0.05), Inches(3.2), Inches(0.42),
                size=14, bold=True, color=accent)
        txt_box(slide, desc, Inches(3.9), y + Inches(0.08), Inches(9), Inches(0.8),
                size=13, color=NAVY)
        y += Inches(1.02)


def build_user_guide():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    # ── Slide 1: Title ────────────────────────────────────────────────────────
    title_slide(prs)

    # ── Slide 2: What is DRAFT Tool ───────────────────────────────────────────
    content_slide(prs, "What is DRAFT Tool?", [
        "DRAFT Tool automates content migration between Word document templates",
        "## When a program needs to move content from an old document format to a new template, DRAFT Tool does the heavy lifting",
        "It reads both documents, automatically identifies matching sections, and produces a merged output file",
        "The tool is designed for programs that migrate documents repeatedly — set up the mappings once, save them as a profile, and reuse",
        "!  All migrated content is color-coded so reviewers can instantly see what moved and what stayed",
        "Output document uses green (migrated content) and blue (boilerplate matches) for quick review",
        "A CSV audit report is generated automatically so you have a full record of every mapping decision",
    ])

    # ── Slide 3: Key Features ─────────────────────────────────────────────────
    two_col_slide(prs, "Key Features",
        "Automation",
        [
            "Auto-matches sections by heading similarity",
            "Confidence scores for every match",
            "Top 3 alternative suggestions per row",
            "Green / blue color coding in output",
            "Boilerplate detection runs automatically",
            "CSV audit report exported automatically",
        ],
        "Control",
        [
            "Override any auto-mapped destination",
            "Searchable destination dropdown",
            "Per-section Append / Prepend / Replace mode",
            "Filter table by status (Auto, Review, Unmapped…)",
            "Save & load mapping profiles (.json)",
            "Unmapped acknowledgement safety gate",
        ]
    )

    # ── Slide 4: The Three-Step Workflow ──────────────────────────────────────
    content_slide(prs, "The Three-Step Workflow", [
        "Step 1 — Load Documents",
        "## Select your source document (.docx) and the destination template (.docx)",
        "## Click 'Analyze Documents' — the tool parses both files and auto-maps sections",
        "Step 2 — Review Section Mappings",
        "## Check each row: green rows are confidently mapped, yellow rows need verification",
        "## Use the Destination dropdown to reassign sections; use Mode to choose Append/Prepend/Replace",
        "## Save the profile for reuse on future migrations between the same document types",
        "Step 3 — Migration Options",
        "## Choose output folder and filename",
        "## Select boilerplate highlighting mode (sentence-level or paragraph-level)",
        "## Click 'Run Migration' to produce the output document",
    ])

    # ── Slide 5: Step 1 Detail ────────────────────────────────────────────────
    content_slide(prs, "Step 1 — Loading Documents", [
        "Click the folder icon (📂) next to Source Document to select your existing .docx",
        "## This is the document whose content you want to migrate",
        "Click the folder icon next to Destination Template to select the new template .docx",
        "## This is the target document structure you want content moved into",
        "The output folder and filename are pre-filled automatically",
        "## Output folder: same folder as the source document",
        "## Output filename: {source_name}_migrated_{YYYYMMDD}.docx",
        "Click 'Analyze Documents →' — this is enabled only when both files are selected",
        "!  Analysis runs in the background — the UI stays responsive while parsing large files",
    ])

    # ── Slide 6: Step 2 Detail ────────────────────────────────────────────────
    content_slide(prs, "Step 2 — Reviewing Section Mappings", [
        "The table shows one row per source section with its auto-mapped destination",
        "Match % column shows confidence: 100% = exact title match, 90%+ = auto-accepted",
        "Status icons: 🟢 Auto-Mapped  🟡 Needs Review  🔴 Unmapped  📝 Changed  ⏭ Skipped",
        "To reassign a destination: click the Destination dropdown and type to search, or scroll",
        "## The top 3 suggestions appear first with confidence percentages (✨ markers)",
        "## Select 'Unmapped' to clear a mapping; select '⏭ Skip' to intentionally exclude",
        "To change insertion mode: use the Mode dropdown (Append / Prepend / Replace)",
        "## Append: migrated content goes after template boilerplate (default)",
        "## Prepend: migrated content goes before template boilerplate",
        "## Replace: template boilerplate is deleted and replaced by migrated content",
    ])

    # ── Slide 7: Filter & Profile ─────────────────────────────────────────────
    content_slide(prs, "Filtering and Profiles", [
        "Filter radio buttons above the table let you focus on specific groups",
        "## All — show every section (default view)",
        "## Auto-Mapped — rows accepted automatically at high confidence",
        "## Needs Review — rows matched at medium confidence that need human verification",
        "## Unmapped — sections with no destination; these are skipped unless assigned",
        "## Changed — rows where you manually overrode the auto-mapper's suggestion",
        "Save Profile (💾) — saves the current mapping table to a .json file",
        "## Useful when the same source/destination template pair is used repeatedly",
        "Load Profile (📂) — applies a saved profile to the current table",
        "## Rows with no matching entry in the profile keep their auto-mapped result",
        "!  Profiles match by exact section title — the source document must not have changed",
    ])

    # ── Slide 8: Step 3 Detail ────────────────────────────────────────────────
    content_slide(prs, "Step 3 — Migration Options", [
        "Export migration report (.csv) — checked by default; writes a CSV alongside the output",
        "## CSV contains: source section, destination, mode, and outcome (Migrated/Skipped/Unmapped)",
        "Template Match Highlighting — choose how boilerplate matches are colored blue",
        "## Matching sentences only: precise — only exact matching sentences turn blue",
        "## Entire paragraph: quick scan — full paragraph turns blue if any sentence matches",
        "Output Folder — where the output .docx (and optional .csv) will be saved",
        "File Name — output filename without the .docx extension",
        "!  If unmapped sections exist, you must check the acknowledgement checkbox before Run Migration enables",
        "After clicking 'Run Migration' a completion dialog shows the counts and lets you open the file directly",
    ])

    # ── Slide 9: Output Color Guide ───────────────────────────────────────────
    color_legend_slide(prs)

    # ── Slide 10: Status Row Colors ───────────────────────────────────────────
    status_legend_slide(prs)

    # ── Slide 11: Tips & Best Practices ──────────────────────────────────────
    content_slide(prs, "Tips and Best Practices", [
        "Review all yellow (Needs Review) rows before running migration — these are 65–89% matches",
        "Use the 'Needs Review' filter to focus only on rows that need attention",
        "If a section has no good match, set it to 'Unmapped' rather than a wrong destination",
        "Use Replace mode carefully — it deletes the destination boilerplate entirely",
        "## Only the first source mapped to a destination can use Replace; subsequent ones are forced to Append",
        "Save a Profile after you've configured mappings for a document pair",
        "## Future migrations of the same document types can load the profile and be done in minutes",
        "Blue text in the output means the content was already in the template — it can often be deleted",
        "!  Always review the output document before distribution — the tool assists but does not replace human review",
    ])

    # ── Slide 12: Troubleshooting ─────────────────────────────────────────────
    content_slide(prs, "Troubleshooting", [
        "Analyze button is greyed out",
        "## Both a source document and destination template must be selected first",
        "Sections are not being detected",
        "## Ensure headings use Word's built-in Heading styles (Heading 1, Heading 2, etc.)",
        "## Custom style names must have the XML outline level set to be recognized",
        "All sections show as Unmapped",
        "## Section titles in the source and destination may be very different — try Load Profile if you have one",
        "## Lower AUTO_THRESHOLD in auto_mapper.py if you want more aggressive automatic matching",
        "Run Migration button stays disabled",
        "## If unmapped sections exist, check the acknowledgement checkbox to enable the button",
        "Output file already exists warning",
        "## Choose a different filename or confirm overwrite in the dialog",
        "!  If the application crashes or an error dialog appears, note the error text and contact your tool administrator",
    ])

    # ── Slide 13: Glossary ────────────────────────────────────────────────────
    content_slide(prs, "Glossary", [
        "Source Document — the existing .docx whose content is being migrated",
        "Destination Template — the new .docx template receiving migrated content",
        "Section — a heading and all content beneath it until the next same-level heading",
        "Confidence / Match % — how similar the source and destination section titles are (0–100%)",
        "Auto-Mapped — a match accepted automatically because confidence is ≥ 90%",
        "Boilerplate — standard text already present in the destination template",
        "Append — insert migrated content after existing boilerplate",
        "Prepend — insert migrated content before existing boilerplate",
        "Replace — delete boilerplate and substitute with migrated content",
        "Profile — a saved JSON file recording the mapping configuration for reuse",
        "Migration Report — a CSV listing every section's outcome (Migrated/Skipped/Unmapped)",
    ])

    out = OUTPUT_DIR / "DRAFT_Tool_User_Guide.pptx"
    prs.save(str(out))
    print(f"Saved: {out}")
    return out


# =============================================================================
# Main
# =============================================================================

if __name__ == "__main__":
    print("Building Technical Documentation…")
    build_technical_doc()

    print("Building User Guide PowerPoint…")
    build_user_guide()

    print("Done.")

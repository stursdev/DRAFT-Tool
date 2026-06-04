# =============================================================================
# gui/mapping_visualizer.py
#
# MappingVisualizerWidget — side panel that renders a live visual of the
# current section mappings.
#
# Layout inside the panel:
#   ┌──────────────────────────────────────────────────────┐
#   │  Header: "Mapping Visualization"                     │
#   ├──────────────────────────────────────────────────────┤
#   │  QScrollArea                                         │
#   │  ┌─────────────┐  bezier lines  ┌─────────────────┐ │
#   │  │ Source boxes │ ─────────────▶ │ Destination     │ │
#   │  │ (all source  │               │ boxes (unique   │ │
#   │  │  sections)   │               │ mapped dests)   │ │
#   │  └─────────────┘               └─────────────────┘ │
#   └──────────────────────────────────────────────────────┘
#
# Color coding (box border + connector line):
#   AUTO / MANUAL → green   (#16A34A)   mapped with confidence
#   REVIEW        → amber   (#D97706)   mapped but needs verification
#   UNMAPPED      → red     (#DC2626)   no destination assigned
#   SKIPPED       → blue    (#2563EB)   intentionally excluded
#
# Destination boxes are neutral (slate) — the color is carried by the
# source box and its connecting line, not the destination box.
#
# Live updates:
#   Call update_mappings(results) to push new data. The canvas recalculates
#   all positions and repaints in the same call, so changes made in the
#   mapping table reflect immediately while the panel is open.
# =============================================================================

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from PyQt5.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)
from PyQt5.QtCore  import Qt, QRect, QPointF, QSize
from PyQt5.QtGui   import (
    QBrush,
    QColor,
    QFont,
    QFontMetrics,
    QPainter,
    QPainterPath,
    QPen,
)

from models.sections import MappingResult, MappingStatus


# =============================================================================
# Color constants
# =============================================================================

# Border of source box + connector line color per status
_STATUS_BORDER = {
    MappingStatus.AUTO:     QColor("#16A34A"),
    MappingStatus.REVIEW:   QColor("#D97706"),
    MappingStatus.UNMAPPED: QColor("#DC2626"),
    MappingStatus.MANUAL:   QColor("#16A34A"),
    MappingStatus.SKIPPED:  QColor("#2563EB"),
}

# Fill of source box per status
_STATUS_FILL = {
    MappingStatus.AUTO:     QColor("#DCFCE7"),
    MappingStatus.REVIEW:   QColor("#FEF3C7"),
    MappingStatus.UNMAPPED: QColor("#FEF2F2"),
    MappingStatus.MANUAL:   QColor("#DCFCE7"),
    MappingStatus.SKIPPED:  QColor("#EFF6FF"),
}

# Destination boxes — neutral slate regardless of which sources map to them
_DEST_FILL   = QColor("#F1F5F9")
_DEST_BORDER = QColor("#475569")

# Canvas background
_BG_COLOR    = QColor("#FFFFFF")


# =============================================================================
# Layout constants
# =============================================================================

_FONT_SIZE       = 9     # pt — keeps text compact in the narrow panel
_BOX_PADDING     = 8     # px — inner horizontal and vertical padding
_BOX_GAP         = 8     # px — vertical gap between consecutive boxes
_BOX_MIN_HEIGHT  = 32    # px — minimum box height even for very short titles
_BOX_RADIUS      = 5     # px — rounded corner radius
_LINE_WIDTH      = 1.6   # px — connector line stroke width
_HEADER_HEIGHT   = 14    # px — reserved space for the "Source / Destination" labels
_TOP_MARGIN      = 6     # px — space above the first box
_SIDE_MARGIN     = 8     # px — canvas left / right margin


# =============================================================================
# MappingCanvas
# =============================================================================

class MappingCanvas(QWidget):
    """
    Custom-painted widget that draws source boxes on the left, destination
    boxes on the right, and bezier connector lines between them.

    Sized to its content height so the parent QScrollArea can scroll
    vertically when the section count is large.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._results: list[MappingResult] = []
        self._font  = QFont("Segoe UI", _FONT_SIZE)
        self._fm    = QFontMetrics(self._font)
        self._total_height = 200

        # Calculated by _recalculate() — used in paintEvent
        self._src_boxes : list = []   # (QRect, MappingResult)
        self._dst_boxes : list = []   # (QRect, title: str)
        self._dst_index : dict = {}   # title → index in _dst_boxes

        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.setMinimumWidth(180)

    # =========================================================================
    # Public API
    # =========================================================================

    def set_data(self, results: list[MappingResult]):
        """Push a new set of mapping results and trigger a repaint."""
        self._results = results
        self._recalculate()
        self.update()

    # =========================================================================
    # Qt overrides
    # =========================================================================

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._recalculate()

    def sizeHint(self) -> QSize:
        return QSize(self.width(), self._total_height)

    def paintEvent(self, event):
        if not self._results:
            self._paint_empty(QPainter(self))
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.TextAntialiasing)

        # Background
        painter.fillRect(self.rect(), _BG_COLOR)

        # Column header labels
        self._paint_headers(painter)

        # Connector lines drawn first so boxes sit on top of line ends
        self._paint_connectors(painter)

        # Source boxes
        for rect, result in self._src_boxes:
            self._paint_box(
                painter, rect, result.source_section.title,
                _STATUS_FILL[result.status],
                _STATUS_BORDER[result.status],
                text_color=QColor("#111827"),
            )

        # Destination boxes
        for rect, title in self._dst_boxes:
            self._paint_box(
                painter, rect, title,
                _DEST_FILL, _DEST_BORDER,
                text_color=QColor("#1E293B"),
            )

        painter.end()

    # =========================================================================
    # Layout calculation
    # =========================================================================

    def _recalculate(self):
        """
        Compute the position and size of every source box and destination box.

        Source boxes:
          Stacked top-to-bottom in source document order, each sized to fit
          its title text at the available column width.

        Destination boxes:
          Only unique mapped destinations are shown. Each is vertically
          centered at the average Y-midpoint of all source boxes that map
          to it. Overlap between adjacent destination boxes is resolved by
          pushing later boxes downward.
        """
        if not self._results:
            self._src_boxes    = []
            self._dst_boxes    = []
            self._dst_index    = {}
            self._total_height = 200
            self.setMinimumHeight(self._total_height)
            return

        canvas_w = max(self.width(), 200)

        # Column geometry — source and destination each occupy ~40% of width.
        # The remaining ~20% in the middle holds the connector curves.
        col_w   = int((canvas_w - _SIDE_MARGIN * 2) * 0.42)
        src_x   = _SIDE_MARGIN
        dst_x   = canvas_w - _SIDE_MARGIN - col_w

        # ── Source boxes ──────────────────────────────────────────────────────
        src_boxes: list = []
        y = _HEADER_HEIGHT + _TOP_MARGIN

        for result in self._results:
            h    = self._box_height(result.source_section.title, col_w)
            rect = QRect(src_x, y, col_w, h)
            src_boxes.append((rect, result))
            y += h + _BOX_GAP

        src_bottom = y

        # ── Destination boxes ─────────────────────────────────────────────────
        # Build ordered list of unique destinations and the Y-midpoints of all
        # source boxes that connect to each one.
        dst_order   : list[str] = []   # unique dest titles in appearance order
        dst_seen    : set       = set()
        dst_src_mids: dict      = {}   # title → [source_box_Y_midpoints]
        dst_heights : dict      = {}   # title → calculated box height

        for i, (rect, result) in enumerate(src_boxes):
            if result.dest_section is None:
                continue
            title = result.dest_section.title
            if title not in dst_seen:
                dst_order.append(title)
                dst_seen.add(title)
                dst_src_mids[title] = []
                dst_heights[title]  = self._box_height(title, col_w)
            dst_src_mids[title].append(rect.top() + rect.height() // 2)

        # Assign initial Y for each destination — vertically centered on
        # the average midpoint of its connected sources.
        dst_ideal_y: dict = {}
        for title in dst_order:
            mids     = dst_src_mids[title]
            center_y = int(sum(mids) / len(mids))
            dst_ideal_y[title] = center_y - dst_heights[title] // 2

        # Sort by ideal Y, then push overlapping boxes down.
        sorted_titles = sorted(dst_order, key=lambda t: dst_ideal_y[t])
        resolved_y: dict = {}
        prev_bottom = _HEADER_HEIGHT + _TOP_MARGIN

        for title in sorted_titles:
            y_pos = max(dst_ideal_y[title], prev_bottom)
            resolved_y[title] = y_pos
            prev_bottom = y_pos + dst_heights[title] + _BOX_GAP

        # Build final dst_boxes list in original appearance order
        dst_boxes : list = []
        dst_index : dict = {}

        for title in dst_order:
            y_pos = resolved_y[title]
            h     = dst_heights[title]
            rect  = QRect(dst_x, y_pos, col_w, h)
            dst_index[title] = len(dst_boxes)
            dst_boxes.append((rect, title))

        dst_bottom = (
            max(rect.bottom() for rect, _ in dst_boxes) + 16
            if dst_boxes else 0
        )

        self._src_boxes    = src_boxes
        self._dst_boxes    = dst_boxes
        self._dst_index    = dst_index
        self._total_height = max(src_bottom, dst_bottom, 200)
        self.setMinimumHeight(self._total_height)

    def _box_height(self, text: str, box_width: int) -> int:
        """
        Return the pixel height needed to render text inside a box of the
        given width, accounting for word-wrap. Never smaller than _BOX_MIN_HEIGHT.
        """
        inner_w = max(box_width - _BOX_PADDING * 2, 20)
        bound   = self._fm.boundingRect(
            0, 0, inner_w, 9999,
            Qt.TextWordWrap | Qt.AlignLeft,
            text,
        )
        return max(_BOX_MIN_HEIGHT, bound.height() + _BOX_PADDING * 2)

    # =========================================================================
    # Painting helpers
    # =========================================================================

    def _paint_empty(self, painter: QPainter):
        painter.fillRect(self.rect(), _BG_COLOR)
        painter.setPen(QColor("#94A3B8"))
        painter.setFont(QFont("Segoe UI", 10))
        painter.drawText(
            self.rect(),
            Qt.AlignCenter | Qt.TextWordWrap,
            "No mappings to display.\nAnalyze documents first.",
        )
        painter.end()

    def _paint_headers(self, painter: QPainter):
        """Draw 'Source Document' and 'Destination' column labels."""
        canvas_w = self.width()
        col_w    = int((canvas_w - _SIDE_MARGIN * 2) * 0.42)

        hdr_font = QFont("Segoe UI", _FONT_SIZE - 1)
        hdr_font.setBold(True)
        painter.setFont(hdr_font)
        painter.setPen(QColor("#64748B"))

        painter.drawText(
            QRect(_SIDE_MARGIN, 0, col_w, _HEADER_HEIGHT),
            Qt.AlignLeft | Qt.AlignVCenter,
            "Source Document",
        )
        painter.drawText(
            QRect(canvas_w - _SIDE_MARGIN - col_w, 0, col_w, _HEADER_HEIGHT),
            Qt.AlignLeft | Qt.AlignVCenter,
            "Destination",
        )
        painter.setFont(self._font)

    def _paint_connectors(self, painter: QPainter):
        """
        Draw a bezier curve from each mapped source box's right edge to
        its destination box's left edge. Line color matches the mapping status.
        """
        canvas_w = self.width()
        col_w    = int((canvas_w - _SIDE_MARGIN * 2) * 0.42)
        src_right = _SIDE_MARGIN + col_w
        dst_left  = canvas_w - _SIDE_MARGIN - col_w

        for src_rect, result in self._src_boxes:
            if result.dest_section is None:
                continue
            idx = self._dst_index.get(result.dest_section.title)
            if idx is None:
                continue

            dst_rect, _ = self._dst_boxes[idx]
            color = _STATUS_BORDER[result.status]

            src_y = src_rect.top() + src_rect.height() // 2
            dst_y = dst_rect.top() + dst_rect.height() // 2

            pen = QPen(color, _LINE_WIDTH)
            pen.setCapStyle(Qt.RoundCap)
            painter.setPen(pen)

            # Cubic bezier: control points pull horizontally from each endpoint
            # so the curve bows gently in the middle rather than diagonal.
            ctrl_span = (dst_left - src_right) * 0.45
            path = QPainterPath()
            path.moveTo(QPointF(src_right, src_y))
            path.cubicTo(
                QPointF(src_right + ctrl_span, src_y),
                QPointF(dst_left  - ctrl_span, dst_y),
                QPointF(dst_left,              dst_y),
            )
            painter.drawPath(path)

    def _paint_box(
        self,
        painter:    QPainter,
        rect:       QRect,
        text:       str,
        fill:       QColor,
        border:     QColor,
        text_color: QColor,
    ):
        """Draw a single rounded-rectangle box with word-wrapped text."""
        painter.setBrush(QBrush(fill))
        painter.setPen(QPen(border, 1.2))
        painter.drawRoundedRect(rect, _BOX_RADIUS, _BOX_RADIUS)

        painter.setPen(text_color)
        inner = rect.adjusted(_BOX_PADDING, _BOX_PADDING // 2,
                               -_BOX_PADDING, -_BOX_PADDING // 2)
        painter.drawText(inner, Qt.TextWordWrap | Qt.AlignVCenter | Qt.AlignLeft, text)


# =============================================================================
# MappingVisualizerWidget
# =============================================================================

class MappingVisualizerWidget(QWidget):
    """
    Side panel containing the mapping visualization canvas.

    Displayed at roughly 40% of the application window width when the user
    clicks 'Visualize Mapping' in Step 2. Hidden when they click again.

    Public API:
        update_mappings(results: list[MappingResult]) — push fresh data
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()
        self._apply_styles()

    # =========================================================================
    # Public API
    # =========================================================================

    def update_mappings(self, results: list[MappingResult]):
        """Push updated mapping results to the canvas and repaint."""
        self._canvas.set_data(results)

    # =========================================================================
    # UI construction
    # =========================================================================

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # ── Header bar ────────────────────────────────────────────────────────
        header = QWidget()
        header.setObjectName("viz_header")
        header.setFixedHeight(40)
        header_row = QHBoxLayout(header)
        header_row.setContentsMargins(12, 0, 12, 0)

        title_label = QLabel("Mapping Visualization")
        title_label.setObjectName("viz_title")
        header_row.addWidget(title_label)
        header_row.addStretch()

        # Color key — compact legend
        for color_hex, label_text in [
            ("#16A34A", "Mapped"),
            ("#D97706", "Review"),
            ("#DC2626", "Unmapped"),
            ("#2563EB", "Skipped"),
        ]:
            dot = QLabel("●")
            dot.setStyleSheet(f"color: {color_hex}; font-size: 10px;")
            key_lbl = QLabel(label_text)
            key_lbl.setObjectName("viz_key_label")
            header_row.addWidget(dot)
            header_row.addWidget(key_lbl)

        layout.addWidget(header)

        # ── Scroll area ───────────────────────────────────────────────────────
        self._scroll = QScrollArea()
        self._scroll.setObjectName("viz_scroll")
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        self._canvas = MappingCanvas()
        self._scroll.setWidget(self._canvas)
        layout.addWidget(self._scroll, stretch=1)

    def _apply_styles(self):
        self.setStyleSheet("""
            QWidget#viz_header {
                background-color: #1e293b;
                border-left: 3px solid #3b82f6;
            }
            QLabel#viz_title {
                color: #f1f5f9;
                font-weight: bold;
                font-size: 12px;
                font-family: 'Segoe UI', Arial, sans-serif;
                padding-right: 16px;
            }
            QLabel#viz_key_label {
                color: #94a3b8;
                font-size: 10px;
                font-family: 'Segoe UI', Arial, sans-serif;
                padding-right: 8px;
            }
            QScrollArea#viz_scroll {
                border: none;
                border-left: 1px solid #cbd5e1;
                background-color: #ffffff;
            }
        """)

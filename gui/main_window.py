# =============================================================================
# gui/main_window.py
#
# The top-level application window.
#
# Contains:
#   HorizontalTabBar — custom QTabBar that draws tab text horizontally
#                      (Qt's default West-position tab bar rotates text 90
#                      degrees which is unreadable for multi-word labels)
#   MainWindow       — QMainWindow shell that hosts the left tab bar
#
# Adding a new tab in the future:
#   1. Create a new QWidget subclass in gui/your_new_tab.py
#   2. Import it in this file
#   3. Call self._tab_widget.addTab(YourNewTab(), "Label Text")
#   No other files need to change.
# =============================================================================

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QTabBar,
    QTabWidget,
)
from PyQt5.QtCore import Qt, QRect, QSize
from PyQt5.QtGui  import QColor, QFont, QPainter

from gui.migration_tab import MigrationTab


# =============================================================================
# HorizontalTabBar
#
# Problem being solved:
#   Qt's default QTabBar, when positioned on the West (left) side of a
#   QTabWidget, rotates each tab's content 90 degrees so the text reads
#   from bottom to top. For short labels this is tolerable but for labels
#   like "Document Migration" the text is unreadable and may be clipped.
#
# Solution:
#   Subclass QTabBar and override two methods:
#     tabSizeHint() — returns a fixed width/height for each tab button,
#                     wide enough to display the full label horizontally
#     paintEvent()  — draws each tab manually using QPainter with upright
#                     (0 degree rotation) text centered in the tab area
#
# This approach gives full control over tab appearance without the text
# rotation side effect of Qt's default West tab rendering.
# =============================================================================

# Fixed dimensions for every tab button
TAB_BUTTON_WIDTH  = 110   # Wide enough for "Document Migration" across two lines
TAB_BUTTON_HEIGHT = 56    # Tall enough for two lines of text

# Colors for the tab bar — dark navy theme
TAB_COLOR_BACKGROUND          = QColor("#1e293b")   # Default tab background
TAB_COLOR_BACKGROUND_SELECTED = QColor("#273549")   # Selected tab background
TAB_COLOR_TEXT_NORMAL         = QColor("#94a3b8")   # Default tab text
TAB_COLOR_TEXT_SELECTED       = QColor("#ffffff")   # Selected tab text
TAB_COLOR_ACCENT              = QColor("#3b82f6")   # Blue left-edge accent bar


class HorizontalTabBar(QTabBar):
    """
    A QTabBar that draws all tab labels horizontally regardless of tab position.
    Intended for use with QTabWidget.West tab position.
    """

    def tabSizeHint(self, tab_index: int) -> QSize:
        """
        Return the size for the tab button at the given index.
        All tabs use the same fixed width and height for visual consistency.
        """
        return QSize(TAB_BUTTON_WIDTH, TAB_BUTTON_HEIGHT)

    def paintEvent(self, paint_event):
        """
        Custom paint event — draws all tabs with horizontal text.

        For each tab:
          1. Fill the background color (differs for selected vs unselected)
          2. Draw a blue left-edge accent bar on the selected tab
          3. Draw the label text horizontally centered in the tab area
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        for tab_index in range(self.count()):
            tab_rect     = self.tabRect(tab_index)
            is_selected  = (self.currentIndex() == tab_index)

            # ── Background ─────────────────────────────────────────────────────
            background_color = (
                TAB_COLOR_BACKGROUND_SELECTED if is_selected
                else TAB_COLOR_BACKGROUND
            )
            painter.fillRect(tab_rect, background_color)

            # ── Left accent bar (selected tab only) ────────────────────────────
            if is_selected:
                accent_bar_rect = QRect(
                    tab_rect.left(),
                    tab_rect.top(),
                    3,               # 3px wide accent strip
                    tab_rect.height(),
                )
                painter.fillRect(accent_bar_rect, TAB_COLOR_ACCENT)

            # ── Label text ─────────────────────────────────────────────────────
            label_font = QFont("Segoe UI", 10 if is_selected else 9)
            label_font.setBold(is_selected)
            painter.setFont(label_font)
            painter.setPen(
                TAB_COLOR_TEXT_SELECTED if is_selected else TAB_COLOR_TEXT_NORMAL
            )

            # Inset the text area from the left accent bar
            text_rect = QRect(
                tab_rect.left() + 8,        # Leave room past the accent bar
                tab_rect.top() + 4,
                tab_rect.width() - 12,
                tab_rect.height() - 8,
            )
            painter.drawText(
                text_rect,
                Qt.AlignHCenter | Qt.AlignVCenter | Qt.TextWordWrap,
                self.tabText(tab_index),
            )

        painter.end()


# =============================================================================
# MainWindow
# =============================================================================
class MainWindow(QMainWindow):
    """
    Application shell — hosts the left horizontal-text tab bar and all tabs.
    """

    MINIMUM_WINDOW_WIDTH  = 960
    MINIMUM_WINDOW_HEIGHT = 760

    def __init__(self):
        super().__init__()
        self._configure_window()
        self._build_ui()

    def _configure_window(self):
        """Set window title, minimum size, and center on the primary screen."""
        self.setWindowTitle("Document Migration Tool")
        self.setMinimumSize(self.MINIMUM_WINDOW_WIDTH, self.MINIMUM_WINDOW_HEIGHT)
        self.resize(
            self.MINIMUM_WINDOW_WIDTH  + 40,
            self.MINIMUM_WINDOW_HEIGHT + 40,
        )

        # Center the window on the primary screen
        screen_geometry = QApplication.primaryScreen().geometry()
        center_x = (screen_geometry.width()  - self.width())  // 2
        center_y = (screen_geometry.height() - self.height()) // 2
        self.move(center_x, center_y)

    def _build_ui(self):
        """Create the tab widget with the custom horizontal tab bar."""
        self._tab_widget = QTabWidget()
        self._tab_widget.setTabPosition(QTabWidget.West)
        self._tab_widget.setTabBar(HorizontalTabBar())
        self._tab_widget.setObjectName("main_tabs")

        # ── Tab 1: Document Migration ──────────────────────────────────────────
        self._tab_widget.addTab(MigrationTab(), "📄 Document\nMigration")

        # ── Future tabs (add here without touching other files) ────────────────
        # from gui.some_other_tab import SomeOtherTab
        # self._tab_widget.addTab(SomeOtherTab(), "🔧 Other\nTool")

        self.setCentralWidget(self._tab_widget)
        self._apply_styles()

    def _apply_styles(self):
        """
        Apply styles to the window shell and tab bar container.
        Tab button painting (colors, text, accent bar) is handled entirely
        by HorizontalTabBar.paintEvent — we do not define QTabBar::tab rules
        here because they would conflict with the custom paint logic.
        """
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1e293b;
            }
            QTabWidget#main_tabs::pane {
                border: none;
                background-color: #f0f2f5;
            }
            QTabBar {
                background-color: #1e293b;
            }
            QTabBar::tab {
                background-color: transparent;
                border: none;
            }
        """)

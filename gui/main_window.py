# =============================================================================
# gui/main_window.py
#
# The top-level application window.
#
# Contains:
#   HorizontalTabBar — custom QTabBar that draws each tab as a stacked
#                      icon-above-label layout (icon centered, text below).
#                      Qt's default West-position tab bar rotates text 90
#                      degrees, which is unreadable for multi-word labels.
#   MainWindow       — QMainWindow shell that hosts the left tab bar.
#
# Adding a new tab:
#   1. Create a QWidget subclass in gui/your_new_tab.py
#   2. Import it here
#   3. Call self._tab_widget.addTab(YourTab(), QIcon(path), "Label\nText")
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
from PyQt5.QtGui  import QColor, QFont, QIcon, QPainter

from gui.migration_tab import MigrationTab


# Resolve the assets directory relative to this file so paths work whether
# the app is run from source or packaged with PyInstaller.
ASSETS_DIR = Path(__file__).parent.parent / "assets"


# =============================================================================
# HorizontalTabBar
#
# Draws each tab as: icon centered horizontally, label text below the icon.
# Both elements are centered within the fixed-size tab button.
#
# Uses Qt's built-in tabIcon(index) so icons are set the standard way via
# QTabWidget.addTab(widget, QIcon(...), "Label") — no custom API needed.
# =============================================================================

TAB_BUTTON_WIDTH  = 110   # px — wide enough for two-word labels
TAB_BUTTON_HEIGHT = 82    # px — icon(28) + gap(6) + text(~28) + padding(20)
TAB_ICON_SIZE     = 28    # px — square icon rendered in each tab

# Dark navy theme colors
TAB_COLOR_BACKGROUND          = QColor("#1e293b")
TAB_COLOR_BACKGROUND_SELECTED = QColor("#273549")
TAB_COLOR_TEXT_NORMAL         = QColor("#94a3b8")
TAB_COLOR_TEXT_SELECTED       = QColor("#ffffff")
TAB_COLOR_ACCENT              = QColor("#3b82f6")   # Blue left-edge accent strip


class HorizontalTabBar(QTabBar):
    """
    QTabBar that renders each tab as a vertically stacked icon + label,
    regardless of the tab widget's West/East/North/South position setting.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        # Force RoundedWest shape so Qt lays tabs out top-to-bottom.
        # Without this, the default RoundedNorth shape lays them left-to-right,
        # which causes both tabs to appear side by side instead of stacked.
        self.setShape(QTabBar.RoundedWest)

    def tabSizeHint(self, index: int) -> QSize:
        return QSize(TAB_BUTTON_WIDTH, TAB_BUTTON_HEIGHT)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)

        for index in range(self.count()):
            rect        = self.tabRect(index)
            is_selected = (self.currentIndex() == index)

            # ── Background ─────────────────────────────────────────────────────
            bg = TAB_COLOR_BACKGROUND_SELECTED if is_selected else TAB_COLOR_BACKGROUND
            painter.fillRect(rect, bg)

            # ── Left accent bar (selected tab only) ────────────────────────────
            if is_selected:
                painter.fillRect(
                    QRect(rect.left(), rect.top(), 3, rect.height()),
                    TAB_COLOR_ACCENT,
                )

            # ── Icon (centered, top portion) ───────────────────────────────────
            icon     = self.tabIcon(index)
            icon_top = rect.top() + 10   # pixels of top padding above the icon

            if not icon.isNull():
                # Slightly dim the icon when the tab is not active
                painter.setOpacity(1.0 if is_selected else 0.55)
                pixmap = icon.pixmap(QSize(TAB_ICON_SIZE, TAB_ICON_SIZE))
                icon_x = rect.left() + (rect.width() - TAB_ICON_SIZE) // 2
                painter.drawPixmap(icon_x, icon_top, pixmap)
                painter.setOpacity(1.0)
                text_top = icon_top + TAB_ICON_SIZE + 6   # gap between icon and text
            else:
                # No icon — start text a bit lower for visual balance
                text_top = icon_top + 4

            # ── Label (centered, below icon) ───────────────────────────────────
            label_font = QFont("Segoe UI", 9)
            label_font.setBold(is_selected)
            painter.setFont(label_font)
            painter.setPen(
                TAB_COLOR_TEXT_SELECTED if is_selected else TAB_COLOR_TEXT_NORMAL
            )

            text_rect = QRect(
                rect.left() + 4,
                text_top,
                rect.width() - 8,
                rect.bottom() - text_top - 4,
            )
            painter.drawText(
                text_rect,
                Qt.AlignHCenter | Qt.AlignTop | Qt.TextWordWrap,
                self.tabText(index),
            )

        painter.end()


# =============================================================================
# MainWindow
# =============================================================================
class MainWindow(QMainWindow):
    """
    Application shell — hosts the left sidebar tab bar and all content tabs.
    """

    MINIMUM_WINDOW_WIDTH  = 960
    MINIMUM_WINDOW_HEIGHT = 860

    def __init__(self):
        super().__init__()
        self._configure_window()
        self._build_ui()

    def _configure_window(self):
        """Set title, icon, minimum size, and center on the primary screen."""
        self.setWindowTitle("DRAFT Tool")
        self.setMinimumSize(self.MINIMUM_WINDOW_WIDTH, self.MINIMUM_WINDOW_HEIGHT)
        self.resize(
            self.MINIMUM_WINDOW_WIDTH  + 40,
            self.MINIMUM_WINDOW_HEIGHT + 60,
        )

        # Application icon — used by the OS for the window title bar, dock /
        # taskbar, and (when set on QApplication in main.py) the .exe icon.
        app_icon_path = ASSETS_DIR / "app_icon.ico"
        if app_icon_path.exists():
            self.setWindowIcon(QIcon(str(app_icon_path)))

        # Center the window on the primary screen
        screen = QApplication.primaryScreen().geometry()
        self.move(
            (screen.width()  - self.width())  // 2,
            (screen.height() - self.height()) // 2,
        )

    def _build_ui(self):
        """Create the tab widget with the custom icon-above-label tab bar."""
        self._tab_widget = QTabWidget()
        # setTabBar must come BEFORE setTabPosition. Qt sets the bar's shape
        # (RoundedWest) when setTabPosition is called — if the bar hasn't been
        # installed yet, it sets the shape on the default bar, not ours.
        self._tab_widget.setTabBar(HorizontalTabBar())
        self._tab_widget.setTabPosition(QTabWidget.West)
        self._tab_widget.setObjectName("main_tabs")

        # ── Tab 1: Document Migration ──────────────────────────────────────────
        self._tab_widget.addTab(
            MigrationTab(),
            self._load_icon("app_icon.ico"),
            "Document\nMigration",
        )

        # ── Future tabs (add here without touching other files) ────────────────
        # from gui.document_review_tab import DocumentReviewTab
        # self._tab_widget.addTab(DocumentReviewTab(), self._load_icon("genai.png"), "Document\nReview")

        self.setCentralWidget(self._tab_widget)
        self._apply_styles()

    @staticmethod
    def _load_icon(filename: str) -> QIcon:
        """
        Load a QIcon from the assets directory.
        Returns a null QIcon (renders as nothing) if the file does not exist,
        so a missing asset file never crashes the application.
        """
        path = ASSETS_DIR / filename
        if path.exists():
            return QIcon(str(path))
        return QIcon()

    def _apply_styles(self):
        """
        Stylesheet for the window shell and tab widget container.
        Tab button painting (colors, icons, text) is handled entirely by
        HorizontalTabBar.paintEvent — QTabBar::tab rules here are kept
        minimal to avoid conflicting with the custom paint logic.
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

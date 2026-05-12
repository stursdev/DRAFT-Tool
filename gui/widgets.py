# =============================================================================
# gui/widgets.py
#
# Reusable custom PyQt5 widgets shared across the application.
#
# Contents:
#   NoScrollComboBox      — QComboBox that ignores accidental mouse wheel
#                           scrolling when the dropdown is closed
#   IndentedComboDelegate — QStyledItemDelegate that draws dropdown items
#                           with indentation based on heading level
#   SectionLabel          — Small italic instruction label for step headers
# =============================================================================

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from PyQt5.QtWidgets import (
    QComboBox,
    QLabel,
    QSizePolicy,
    QStyledItemDelegate,
    QStyleOptionViewItem,
)
from PyQt5.QtCore import Qt, QRect, QSize
from PyQt5.QtGui  import QPainter, QFont


# =============================================================================
# NoScrollComboBox
#
# Problem being solved:
#   The default QComboBox changes its selected value when the user scrolls
#   the mouse wheel over it, even without opening the dropdown. In the mapping
#   table, this causes accidental destination or mode changes when the user
#   is simply scrolling to see more rows.
#
# Solution:
#   Override wheelEvent to do nothing when the dropdown popup is closed.
#   We track the popup state ourselves using a boolean flag, because PyQt5
#   does not expose the popup open/closed state through its public API.
#   The flag is set to True in showPopup() and False in hidePopup(), which
#   are called by Qt whenever the dropdown opens or closes.
# =============================================================================
class NoScrollComboBox(QComboBox):

    def __init__(self, parent=None):
        super().__init__(parent)

        # Tracks whether the dropdown list is currently open.
        # False by default — the dropdown starts closed.
        self._dropdown_is_open = False

    def showPopup(self):
        """Called by Qt when the user clicks to open the dropdown list."""
        self._dropdown_is_open = True
        super().showPopup()

    def hidePopup(self):
        """Called by Qt when the dropdown list closes (by selection or click-away)."""
        self._dropdown_is_open = False
        super().hidePopup()

    def wheelEvent(self, event):
        """
        Override the default mouse wheel behavior.

        When the dropdown is open, wheel scrolling through the list is
        expected behavior — allow it through to the default handler.

        When the dropdown is closed, ignore the wheel event entirely.
        This prevents accidental value changes while scrolling the table.
        """
        if self._dropdown_is_open:
            super().wheelEvent(event)
        else:
            event.ignore()


# =============================================================================
# IndentedComboDelegate
#
# A custom item delegate for the destination dropdown that draws each item
# with left indentation based on its Word heading level.
#
# Why a custom delegate is needed:
#   Qt's default combo box item rendering ignores custom indentation stored
#   in item data. We need to read the heading level from each item and
#   calculate the left margin ourselves during painting.
#
# How heading level is stored:
#   Each section item in the combo has its heading level stored under
#   Qt.UserRole + 1. We use Qt.UserRole + 1 (rather than Qt.UserRole) because
#   Qt.UserRole is already used to store the section index value on each item.
#   Using a different role avoids overwriting that value.
#
#   Special (sentinel) items like "Unmapped" and "Skip this section" have
#   their heading level stored as 0, which produces no indentation —
#   these items should stand out visually at the top of the list.
#
# Indentation:
#   Level 1 → 0 extra pixels  (no indent — top-level sections)
#   Level 2 → 14 pixels       (one level of indent)
#   Level 3 → 28 pixels       (two levels of indent)
#   etc.
# =============================================================================

# Pixels of indentation added per heading level step beyond level 1
INDENT_PIXELS_PER_LEVEL = 14


class IndentedComboDelegate(QStyledItemDelegate):

    def sizeHint(self, option: QStyleOptionViewItem, index) -> QSize:
        """Return a consistent row height for all items in the dropdown."""
        base_size = super().sizeHint(option, index)
        return QSize(base_size.width(), 24)

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index):
        """
        Draw one dropdown item with heading-level-appropriate indentation.

        Reads the heading level from Qt.UserRole + 1 on the item.
        Level 1 items get no extra indentation.
        Each additional level adds INDENT_PIXELS_PER_LEVEL pixels.
        """
        # Read the heading level stored on this item.
        # Returns 0 for sentinel items (Unmapped, Skip) and separators.
        item_heading_level = index.data(Qt.UserRole + 1) or 0

        # Calculate left indent: only levels 2+ get indented
        # max(0, ...) ensures negative levels (if they ever occur) are treated as 0
        left_indent_pixels = max(0, item_heading_level - 1) * INDENT_PIXELS_PER_LEVEL

        painter.save()

        # Draw the selection highlight or normal background
        if option.state & 0x0200:   # QStyle.State_Selected flag value
            painter.fillRect(option.rect, option.palette.highlight())
            painter.setPen(option.palette.highlightedText().color())
        else:
            painter.fillRect(option.rect, option.palette.base())
            painter.setPen(option.palette.text().color())

        # Draw the item text with the calculated left margin
        text_draw_rect = QRect(
            option.rect.left() + left_indent_pixels + 4,   # 4px base margin + indent
            option.rect.top(),
            option.rect.width() - left_indent_pixels - 4,
            option.rect.height(),
        )
        item_text = index.data(Qt.DisplayRole) or ""
        painter.drawText(
            text_draw_rect,
            Qt.AlignVCenter | Qt.TextSingleLine,
            item_text,
        )

        painter.restore()


# =============================================================================
# SectionLabel
#
# A small, styled instruction label displayed below each step group box title.
# Provides the user with concise guidance about what to do in each step
# without taking up much vertical space.
# =============================================================================
class SectionLabel(QLabel):

    def __init__(self, instruction_text: str, parent=None):
        super().__init__(instruction_text, parent)

        self.setStyleSheet(
            "color: #111827; font-size: 11px; font-style: italic;"
        )
        self.setWordWrap(True)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)

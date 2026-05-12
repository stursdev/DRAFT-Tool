# =============================================================================
# gui/document_review_tab.py
#
# Placeholder for the Document Review tab.
# This tab will house AI-assisted document review features.
# =============================================================================

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QLabel,
    QFrame,
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui  import QFont


class DocumentReviewTab(QWidget):
    """
    Document Review tab — placeholder for upcoming AI-assisted review features.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("document_review_tab")
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(12)

        card = QFrame()
        card.setObjectName("review_card")
        card.setFixedWidth(360)
        card_layout = QVBoxLayout(card)
        card_layout.setAlignment(Qt.AlignCenter)
        card_layout.setSpacing(10)
        card_layout.setContentsMargins(32, 32, 32, 32)

        title = QLabel("Document Review")
        title_font = QFont("Segoe UI", 16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        title.setObjectName("review_title")

        subtitle = QLabel("AI-assisted document review features are coming soon.")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setWordWrap(True)
        subtitle.setObjectName("review_subtitle")

        card_layout.addWidget(title)
        card_layout.addWidget(subtitle)
        layout.addWidget(card)

        self.setStyleSheet("""
            QWidget#document_review_tab {
                background-color: #f0f2f5;
            }
            QFrame#review_card {
                background-color: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
            }
            QLabel#review_title {
                color: #1e293b;
            }
            QLabel#review_subtitle {
                color: #475569;
                font-size: 13px;
            }
        """)

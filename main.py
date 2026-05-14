# =============================================================================
# main.py
#
# Entry point for the Document Migration Tool.
#
# Run this file to launch the application:
#   python main.py
#
# What this file does:
#   1. Adds the project root directory to Python's module search path so all
#      package imports (core/, gui/, models/) resolve correctly regardless of
#      which directory the user runs the script from.
#
#      Why this is necessary:
#        Python only adds the script's own directory to sys.path by default,
#        not its parent. Without this fix, running "python main.py" from the
#        project root works, but running it from another directory (or via
#        a shortcut/batch file) would raise ModuleNotFoundError for core,
#        gui, and models imports.
#
#   2. Enables high-DPI display scaling for sharp rendering on modern monitors.
#      These attributes must be set before creating the QApplication instance.
#
#   3. Creates the QApplication, shows the MainWindow, and starts the Qt
#      event loop. app.exec_() blocks until the user closes the window.
# =============================================================================

import sys
from pathlib import Path

# Add the project root directory to sys.path.
# Path(__file__).parent resolves to the directory containing main.py,
# which is the project root. We check before inserting to avoid duplicates
# if this module is somehow imported twice.
PROJECT_ROOT = str(Path(__file__).parent)
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore    import Qt
from PyQt5.QtGui     import QIcon

from gui.main_window import MainWindow


def main():
    # Enable crisp rendering on HiDPI/Retina displays (4K monitors, Mac Retina).
    # Must be called before QApplication is created.
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps,    True)

    # sys.argv is passed so Qt can process any command-line arguments it
    # recognises (e.g. -style, -platform flags)
    app = QApplication(sys.argv)
    app.setApplicationName("DRAFT Tool")
    app.setOrganizationName("SEP Tools")

    # Set the application-level icon. This is what the OS uses for the
    # taskbar/dock entry and what PyInstaller embeds as the .exe icon when
    # the app is packaged. Setting it on QApplication (not just the window)
    # ensures it propagates to all dialogs and child windows.
    icon_path = Path(__file__).parent / "assets" / "app_icon.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    window = MainWindow()
    window.show()

    # Start the Qt event loop. Blocks here until the window is closed.
    # sys.exit() passes the return code back to the OS.
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

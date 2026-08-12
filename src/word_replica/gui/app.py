from __future__ import annotations

import sys


def main() -> int:
    try:
        from PySide6.QtWidgets import QApplication
    except ImportError as exc:
        raise SystemExit("PySide6 is required for word-replica-gui; install the project dependencies first") from exc
    from word_replica.gui.controller import RebuildController
    from word_replica.gui.main_window import MainWindow
    app = QApplication(sys.argv)
    controller = RebuildController()
    window = MainWindow(controller)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

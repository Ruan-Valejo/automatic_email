import sys

from PySide6.QtWidgets import QApplication

from app.db import Base, engine
from app.ui.main_window import MainWindow


def init_db() -> None:
    Base.metadata.create_all(bind=engine)


if __name__ == "__main__":
    init_db()

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
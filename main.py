import sys
from PyQt5.QtWidgets import QApplication
from word_to_md_combined_refactored import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec_()) 
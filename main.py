import sys
from PyQt5.QtWidgets import QApplication
from src.interface import Janela
import src.style.stylesheet as stylesheet

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Janela()
    window.show()
    sys.exit(app.exec_())

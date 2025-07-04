import sys
from PyQt5.QtWidgets import QApplication, QSplashScreen
from PyQt5.QtCore import QTimer
from app.ui.main_window import App
from app.utils import criar_splash_pixmap

if __name__ == "__main__":
    app = QApplication(sys.argv)

    app.setStyleSheet("""
    QWidget {
        background-color: #FFFFFF;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-size: 14px;
        color: #616161;
    }
    QPushButton {
        background-color: #2A5C8A;
        color: white;
        border-radius: 6px;
        padding: 10px 20px;
        font-weight: 600;
    }
    QPushButton:hover {
        background-color: #1f4567;
    }
    QTableWidget {
        background-color: #FFFFFF;
        gridline-color: #E0E0E0;
        font-size: 13px;
        border: 1px solid #E0E0E0;
        alternate-background-color: #E3F2FD;
    }
    QHeaderView::section {
        background-color: #2A5C8A;
        color: white;
        padding: 6px;
        font-weight: 700;
        border: none;
    }
    QComboBox {
        padding: 6px;
        font-size: 13px;
    }
    """)

    splash = QSplashScreen(criar_splash_pixmap())
    splash.show()

    main_window = None  # variável global para manter a referência

    def abrir_janela():
        global main_window
        main_window = App()
        splash.finish(main_window)
        main_window.show()

    QTimer.singleShot(2000, abrir_janela)

    sys.exit(app.exec())

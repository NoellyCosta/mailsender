import os
import sys
from PyQt5.QtGui import QPixmap, QPainter, QColor, QFont, QLinearGradient
from PyQt5.QtCore import QRect, Qt


def resource_path(relative_path: str) -> str:
    """
    Retorna o caminho absoluto para um recurso (imagem, ícone etc.),
    funcionando tanto no modo desenvolvimento quanto no executável PyInstaller.

    :param relative_path: caminho relativo do recurso dentro do projeto
    :return: caminho absoluto para uso em runtime
    """
    base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    return os.path.join(base_path, relative_path)


def criar_splash_pixmap() -> QPixmap:
    """
    Cria um QPixmap com gradiente vertical e textos customizados
    para a tela de splash inicial do aplicativo.

    :return: QPixmap pronto para uso no QSplashScreen
    """
    largura, altura = 600, 360
    pixmap = QPixmap(largura, altura)
    pixmap.fill(Qt.transparent)  # Fundo transparente para evitar artefatos

    painter = QPainter(pixmap)
    grad = QLinearGradient(0, 0, 0, altura)
    grad.setColorAt(0.0, QColor("#a8c0ff"))
    grad.setColorAt(1.0, QColor("#ffffff"))
    painter.fillRect(0, 0, largura, altura, grad)

    # Título principal
    font_title = QFont("Segoe UI", 28, QFont.Bold)
    painter.setFont(font_title)
    painter.setPen(QColor("#2c3e50"))
    painter.drawText(QRect(0, 80, largura, 60), Qt.AlignHCenter | Qt.AlignVCenter, "Walmir Valença")

    # Subtítulo
    font_subtitle = QFont("Segoe UI", 20)
    painter.setFont(font_subtitle)
    painter.setPen(QColor("#4a6fa5"))
    painter.drawText(QRect(0, 140, largura, 40), Qt.AlignHCenter | Qt.AlignTop, "Advocacia")

    # Descrição/rodapé
    font_desc = QFont("Segoe UI", 14, QFont.StyleItalic)
    painter.setFont(font_desc)
    painter.setPen(QColor("#7f8c8d"))
    painter.drawText(QRect(0, 270, largura, 40), Qt.AlignHCenter, "Disparador de E-mails")

    painter.end()
    return pixmap

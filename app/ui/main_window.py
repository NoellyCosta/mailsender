import os
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QHBoxLayout, QTableWidget,
    QTableWidgetItem, QFileDialog, QMessageBox, QLabel, QProgressBar,
    QApplication, QHeaderView, QComboBox, QGroupBox
)
from PyQt5.QtCore import Qt, QTimer, QUrl
from PyQt5.QtGui import QIcon
from PyQt5.Qt import QDesktopServices

from app.email_sender import send_email
from app.utils import resource_path


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Disparador de Emails'
        self.df = pd.DataFrame()
        self.arquivo_excel = None
        self.logo_path = None
        self.header_path = None
        self.lote_size = 90
        self.pause_segundos = 300
        self.lote_atual = 0

        try:
            self.setWindowIcon(QIcon(resource_path('src/assets/app.ico')))
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

        self.setWindowTitle(self.title)
        self.initUI()

    def initUI(self):
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(15)

        btn_abrir = QPushButton('Abrir Planilha Excel')
        btn_abrir.clicked.connect(self.abrir_planilha)

        tipo_group = QGroupBox("Configuração do envio")
        tipo_layout = QVBoxLayout()

        label_tipo = QLabel("Selecione o tipo de e-mail:")
        self.combo_tipo_email = QComboBox()
        self.combo_tipo_email.addItem("Primeiro Contato", "primeiro_contato")
        self.combo_tipo_email.addItem("Negativa", "negativa")
        self.combo_tipo_email.currentIndexChanged.connect(self.remover_foco_tipo_email)

        tipo_layout.addWidget(label_tipo)
        tipo_layout.addWidget(self.combo_tipo_email)
        tipo_group.setLayout(tipo_layout)

        unidade_group = QGroupBox("Selecionar Unidade para todos")
        unidade_layout = QVBoxLayout()

        label_unidade = QLabel("Escolha a unidade para aplicar em todos:")
        self.combo_unidade = QComboBox()
        self.combo_unidade.addItems(["Arapiraca", "Santana do Ipanema", "Garanhuns", "Itapipoca"])
        self.combo_unidade.currentIndexChanged.connect(self.alterar_unidade_todos)

        unidade_layout.addWidget(label_unidade)
        unidade_layout.addWidget(self.combo_unidade)
        unidade_group.setLayout(unidade_layout)

        self.btn_header = QPushButton('Escolher Header')
        self.btn_header.clicked.connect(self.escolher_header)

        self.btn_assinatura = QPushButton('Escolher Assinatura')
        self.btn_assinatura.clicked.connect(self.escolher_assinatura)

        self.btn_salvar = QPushButton('Salvar Alterações')
        self.btn_salvar.clicked.connect(self.salvar_dados)

        self.btn_enviar_todos = QPushButton('Enviar E-mails (Todos)')
        self.btn_enviar_todos.clicked.connect(self.iniciar_envio_em_lotes)

        btn_suporte = QPushButton('Entrar em contato com o suporte')
        btn_suporte.clicked.connect(self.abrir_email_suporte)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.btn_salvar)
        button_layout.addWidget(self.btn_enviar_todos)
        button_layout.addWidget(btn_suporte)

        img_button_layout = QHBoxLayout()
        img_button_layout.addWidget(self.btn_header)
        img_button_layout.addWidget(self.btn_assinatura)

        self.table = QTableWidget()
        self.table.setMinimumHeight(250)  # Altura menor para a tabela
        self.table.cellChanged.connect(self.on_cell_changed)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)

        self.label_status = QLabel("")
        self.label_status.setVisible(False)

        assinatura = QLabel("By Noelly Costa")
        assinatura.setStyleSheet("color: gray; font-style: italic; font-size: 8pt;")
        self.label_version = QLabel("Versão 1.2.1")
        self.label_version.setStyleSheet("color: gray; font-size: 10px; padding-left: 10px;")

        footer_layout = QHBoxLayout()
        footer_layout.addStretch()
        footer_layout.addWidget(assinatura)
        footer_layout.addWidget(self.label_version)
        footer_layout.addStretch()

        self.layout.addWidget(btn_abrir)
        self.layout.addWidget(tipo_group)
        self.layout.addWidget(unidade_group)
        self.layout.addLayout(img_button_layout)
        self.layout.addWidget(self.table)
        self.layout.addLayout(button_layout)
        self.layout.addWidget(self.progress_bar)
        self.layout.addWidget(self.label_status)
        self.layout.addLayout(footer_layout)

        self.setLayout(self.layout)
        self.show()

    def remover_foco_tipo_email(self):
        self.combo_tipo_email.clearFocus()
        self.btn_salvar.setFocus()

    def alterar_unidade_todos(self):
        if self.df.empty:
            return
        unidade_selecionada = self.combo_unidade.currentText()

        self.df['Unidade'] = unidade_selecionada

        self.table.blockSignals(True)
        col_index = self.df.columns.get_loc('Unidade')
        for row in range(self.table.rowCount()):
            item = self.table.item(row, col_index)
            if item:
                item.setText(unidade_selecionada)
        self.table.blockSignals(False)

        self.combo_unidade.clearFocus()
        self.btn_salvar.setFocus()

    def abrir_email_suporte(self):
        email = QUrl("mailto:noelly1232@gmail.com?subject=Disparador%20de%20E-mails")
        QDesktopServices.openUrl(email)

    def abrir_planilha(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Selecionar Planilha", "", "Arquivos Excel (*.xlsx *.xls)")
        if file_path:
            self.arquivo_excel = file_path
            try:
                self.df = pd.read_excel(file_path)

                required_cols = {'Email', 'Nome', 'Vaga', 'Unidade'}
                if not required_cols.issubset(set(self.df.columns)):
                    QMessageBox.critical(self, "Erro", f"A planilha deve conter as colunas: {', '.join(required_cols)}")
                    self.df = pd.DataFrame()
                    self.atualizar_tabela()
                    return

                self.atualizar_tabela()
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Falha ao abrir a planilha: {e}")

    def escolher_header(self):
        path, _ = QFileDialog.getOpenFileName(self, "Escolher Header", "", "Imagens (*.png *.jpg *.jpeg *.bmp)")
        if path:
            self.header_path = path
            QMessageBox.information(self, "Header selecionado", f"Header alterado para:\n{path}")

    def escolher_assinatura(self):
        path, _ = QFileDialog.getOpenFileName(self, "Escolher Assinatura", "", "Imagens (*.png *.jpg *.jpeg *.bmp)")
        if path:
            self.logo_path = path
            QMessageBox.information(self, "Assinatura selecionada", f"Assinatura alterada para:\n{path}")

    def atualizar_tabela(self):
        if self.df.empty:
            self.table.clear()
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        self.table.blockSignals(True)
        self.table.setRowCount(len(self.df))
        self.table.setColumnCount(len(self.df.columns))
        self.table.setHorizontalHeaderLabels(self.df.columns.tolist())

        for i in range(len(self.df)):
            for j in range(len(self.df.columns)):
                item = QTableWidgetItem(str(self.df.iat[i, j]))
                col_name = self.df.columns[j]
                if col_name == "Vaga":
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                else:
                    item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.table.setItem(i, j, item)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)

        self.table.blockSignals(False)

    def salvar_dados(self):
        if self.arquivo_excel is None:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo carregado!")
            return

        for i in range(self.table.rowCount()):
            for j in range(self.table.columnCount()):
                col_name = self.df.columns[j]
                item = self.table.item(i, j)
                if item:
                    self.df.at[i, col_name] = item.text()

        try:
            self.df.to_excel(self.arquivo_excel, index=False)
            QMessageBox.information(self, "Sucesso", "Planilha salva com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao salvar a planilha: {e}")

    def iniciar_envio_em_lotes(self):
        if self.df.empty:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo carregado!")
            return

        self.salvar_dados()
        self.lote_atual = 0

        self.erros = []
        self.enviados = 0
        self.total = len(self.df)

        if self.logo_path and not os.path.isfile(self.logo_path):
            QMessageBox.warning(self, "Aviso", "Arquivo de assinatura não encontrado!")
            self.logo_path = None

        if self.header_path and not os.path.isfile(self.header_path):
            QMessageBox.warning(self, "Aviso", "Arquivo de header não encontrado!")
            self.header_path = None

        self.progress_bar.setMaximum(self.total)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)

        self.label_status.setText("Iniciando envio...")
        self.label_status.setVisible(True)

        self.btn_salvar.setEnabled(False)
        self.btn_enviar_todos.setEnabled(False)
        self.btn_header.setEnabled(False)
        self.btn_assinatura.setEnabled(False)

        self.timer = QTimer()
        self.timer.timeout.connect(self.processar_lote)
        self.timer.start(100)

    def processar_lote(self):
        self.timer.stop()

        lote_inicio = self.lote_atual
        lote_fim = min(lote_inicio + self.lote_size, self.total)

        tipo_email = self.combo_tipo_email.currentData()

        for i in range(lote_inicio, lote_fim):
            try:
                destinatario = self.df.at[i, 'Email']
                nome = self.df.at[i, 'Nome']
                vaga = self.df.at[i, 'Vaga']
                unidade = self.df.at[i, 'Unidade']

                sucesso = send_email(destinatario, nome, vaga, unidade,
                                     tipo_email=tipo_email,
                                     logo_path=self.logo_path, header_path=self.header_path)
                if sucesso:
                    self.enviados += 1
                else:
                    self.erros.append(destinatario)
            except Exception as e:
                self.erros.append(destinatario)
                print(f"Erro ao enviar email para {destinatario}: {e}")

            self.progress_bar.setValue(i + 1)
            self.label_status.setText(f"Enviando: {i + 1} / {self.total}")
            QApplication.processEvents()

        self.lote_atual += self.lote_size

        if self.lote_atual < self.total:
            self.label_status.setText("Lote enviado. Pausa de 5 minutos...")
            self.timer.start(self.pause_segundos * 1000)
        else:
            self.finalizar_envio()

    def finalizar_envio(self):
        self.label_status.setText("Envio finalizado.")

        msg = f"E-mails enviados com sucesso: {self.enviados}"
        if self.erros:
            msg += f"\nFalha ao enviar para: {', '.join(self.erros)}"

        QMessageBox.information(self, "Resultado do envio", msg)

        self.btn_salvar.setEnabled(True)
        self.btn_enviar_todos.setEnabled(True)
        self.btn_header.setEnabled(True)
        self.btn_assinatura.setEnabled(True)

    def on_cell_changed(self, row, column):
        if self.df.empty:
            return

        col_name = self.df.columns[column]
        novo_valor = self.table.item(row, column).text()

        self.table.blockSignals(True)

        if col_name == "Vaga":
            self.df['Vaga'] = novo_valor
            for r in range(self.table.rowCount()):
                item = self.table.item(r, column)
                if item:
                    item.setText(novo_valor)
        elif col_name == "Unidade":
            self.df.at[row, col_name] = novo_valor

        self.table.blockSignals(False)

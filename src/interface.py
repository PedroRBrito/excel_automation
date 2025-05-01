from cProfile import label
from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QComboBox,
    QCheckBox,
    QProgressBar,
    QTableView,
    QMessageBox,
    QFileDialog,
    QListWidget,
    QAbstractItemView,
    QLabel,
)
from src.models import PandasModel
import src.style.stylesheet as stylesheet
from src.file_manager import carregar_excel
from src.comparison_logic import comparar_planilhas
from src.file_manager import carregar_excel
from src.comparison_logic import comparar_planilhas


class Janela(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Comparador de Planilhas Excel")
        self.setStyleSheet("background-color: #F5F5F5;")
        self.resize(900, 600)
        self.setMinimumSize(700, 500)

        self.df1 = None
        self.df2 = None

        # Layout principal
        layout = QVBoxLayout()

        # Carregamento de arquivos
        layout_arquivos = QHBoxLayout()

        self.button1 = QPushButton("Abrir Primeiro Arquivo")
        self.button1.clicked.connect(self.abrir_arquivo1)
        self.button1.setStyleSheet(stylesheet.button_style())

        self.button2 = QPushButton("Abrir Segundo Arquivo")
        self.button2.clicked.connect(self.abrir_arquivo2)
        self.button2.setStyleSheet(stylesheet.button_style())

        layout_arquivos.addWidget(self.button1)
        layout_arquivos.addWidget(self.button2)

        # Label colunas
        layout_label_colunas = QVBoxLayout()

        label_colunas = QLabel("Escolha as colunas onde serão realizados os cálculos: ")
        layout_label_colunas.addWidget(label_colunas)

        # Seleção de colunas
        layout_colunas = QHBoxLayout()
        self.coluna1_combo = QComboBox()
        self.coluna1_combo.addItem("Escolha a coluna da primeira planilha")
        self.coluna1_combo.setStyleSheet(stylesheet.qcombobox_style())

        self.coluna2_combo = QComboBox()
        self.coluna2_combo.addItem("Escolha a coluna da segunda planilha")
        self.coluna2_combo.setStyleSheet(stylesheet.qcombobox_style())

        layout_colunas.addWidget(self.coluna1_combo)
        layout_colunas.addWidget(self.coluna2_combo)

        # Opções de cálculo
        layout_calculos = QHBoxLayout()
        layout_calculos.setSpacing(5)

        label_calculos = QLabel("Escolha os cálculos para comparação: ")
        layout_calculos.addWidget(label_calculos)

        self.checkbox_soma = QCheckBox("Soma")
        self.checkbox_soma.setStyleSheet(stylesheet.qcheckbox_style())

        self.checkbox_diferenca = QCheckBox("Diferença")
        self.checkbox_diferenca.setStyleSheet(stylesheet.qcheckbox_style())

        self.checkbox_multiplicacao = QCheckBox("Multiplicação")
        self.checkbox_multiplicacao.setStyleSheet(stylesheet.qcheckbox_style())

        self.checkbox_divisao = QCheckBox("Divisão")
        self.checkbox_divisao.setStyleSheet(stylesheet.qcheckbox_style())

        self.checkbox_media = QCheckBox("Média")
        self.checkbox_media.setStyleSheet(stylesheet.qcheckbox_style())

        self.checkbox_diferenca_percentual = QCheckBox("Diferença Percentual")
        self.checkbox_diferenca_percentual.setStyleSheet(stylesheet.qcheckbox_style())

        layout_calculos.addWidget(self.checkbox_soma)
        layout_calculos.addWidget(self.checkbox_diferenca)
        layout_calculos.addWidget(self.checkbox_multiplicacao)
        layout_calculos.addWidget(self.checkbox_divisao)
        layout_calculos.addWidget(self.checkbox_media)
        layout_calculos.addWidget(self.checkbox_diferenca_percentual)

        # Ações
        layout_acoes_comparar = QHBoxLayout()
        self.button_comparar = QPushButton("Comparar Planilhas")
        self.button_comparar.clicked.connect(self.executar_comparacao)
        self.button_comparar.setStyleSheet(stylesheet.button_style())

        layout_acoes_comparar.addWidget(self.button_comparar)

        # Tabelas
        layout_tabelas = QHBoxLayout()
        self.table_primeiro_arquivo = QTableView()
        self.table_primeiro_arquivo.setStyleSheet(stylesheet.qtableview_style())

        self.table_segundo_arquivo = QTableView()
        self.table_segundo_arquivo.setStyleSheet(stylesheet.qtableview_style())

        layout_tabelas.addWidget(self.table_primeiro_arquivo)
        layout_tabelas.addWidget(self.table_segundo_arquivo)

        # header
        layout_header = QHBoxLayout()

        label_arquivos = QLabel("Escolha os arquivos para comparação:")
        layout_header.addWidget(label_arquivos)

        layout_header.addStretch()

        self.button_reset = QPushButton("Resetar")
        self.button_reset.clicked.connect(self.resetar_campos)
        self.button_reset.setStyleSheet(stylesheet.button_style())

        layout_header.addWidget(self.button_reset)

        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet(stylesheet.progress_bar_style())

        # layout principal
        layout.addLayout(layout_header)
        layout.addLayout(layout_arquivos)
        layout.addLayout(layout_colunas)
        layout.addLayout(layout_calculos)
        layout.addLayout(layout_tabelas)
        layout.addWidget(self.progress_bar)
        layout.addLayout(layout_acoes_comparar)

        self.setLayout(layout)

    def resetar_campos(self):
        self.table_primeiro_arquivo.setModel(None)
        self.table_segundo_arquivo.setModel(None)
        self.df1 = None
        self.df2 = None
        self.coluna1_combo.clear()
        self.coluna1_combo.addItem("Escolha a coluna da primeira planilha")
        self.coluna2_combo.clear()
        self.coluna2_combo.addItem("Escolha a coluna da segunda planilha")
        self.progress_bar.setValue(0)

    def abrir_arquivo1(self):
        self.progress_bar.setValue(0)
        self.df1 = carregar_excel(self)
        if self.df1 is not None:
            self.coluna1_combo.clear()
            self.progress_bar.setValue(50)
            model = PandasModel(self.df1)
            self.table_primeiro_arquivo.setModel(model)
            self.table_primeiro_arquivo.setSortingEnabled(True)
            self.coluna1_combo.addItem("Escolha a coluna da primeira planilha")
            self.coluna1_combo.addItems(list(self.df1.columns))
            self.progress_bar.setValue(100)

    def abrir_arquivo2(self):
        self.progress_bar.setValue(0)
        self.df2 = carregar_excel(self)
        if self.df2 is not None:
            self.coluna2_combo.clear()
            self.progress_bar.setValue(50)
            model = PandasModel(self.df2)
            self.table_segundo_arquivo.setModel(model)
            self.table_segundo_arquivo.setSortingEnabled(True)
            self.coluna2_combo.addItem("Escolha a coluna da segunda planilha")
            self.coluna2_combo.addItems(list(self.df2.columns))
            self.progress_bar.setValue(100)

    def executar_comparacao(self):
        try:

            if self.df1 is None or self.df2 is None:
                QMessageBox.warning(
                    self, "Erro", "Carregue as duas planilhas antes de comparar."
                )
                return

            coluna1_combo = self.coluna1_combo.currentText()
            coluna2_combo = self.coluna2_combo.currentText()

            if (
                coluna1_combo == "Escolha a coluna da primeira planilha"
                or coluna2_combo == "Escolha a coluna da segunda planilha"
            ):
                QMessageBox.warning(
                    self, "Erro", "Selecione as colunas-chave para ambas as planilhas."
                )
                return

            checkboxes = {
                "soma": self.checkbox_soma.isChecked(),
                "diferenca": self.checkbox_diferenca.isChecked(),
                "multiplicacao": self.checkbox_multiplicacao.isChecked(),
                "divisao": self.checkbox_divisao.isChecked(),
                "media": self.checkbox_media.isChecked(),
                "diferenca_percentual": self.checkbox_diferenca_percentual.isChecked(),
            }

            self.comparado = comparar_planilhas(
                self.df1, self.df2, coluna1_combo, coluna2_combo, checkboxes
            )

            QMessageBox.information(
                self, "Sucesso", "Planilhas comparadas com sucesso!"
            )

            self.janela_comparacao = JanelaComparacao(self.comparado)
            self.janela_comparacao.show()

        except ValueError as e:
            QMessageBox.warning(self, "Erro", str(e))


class JanelaComparacao(QWidget):
    def __init__(self, df):
        super().__init__()
        self.setWindowTitle("Comparação")
        self.resize(600, 400)
        self.setMinimumSize(400, 200)

        self.df = df

        layout = QVBoxLayout()

        # Tabela preview da exportação
        self.table = QTableView()
        self.table.setModel(PandasModel(df))

        # Label Colunas Exportar
        label_colunas_exportar = QLabel(
            "Escolha as colunas que serão exportadas para o excel: "
        )

        # Colunas para exportar
        self.colunas_selecionadas_list = QListWidget()
        self.colunas_selecionadas_list.setSelectionMode(
            QAbstractItemView.MultiSelection
        )
        self.colunas_selecionadas_list.itemSelectionChanged.connect(
            self.atualizar_tabela
        )
        self.colunas_selecionadas_list.addItems(self.df.columns)

        self.button_exportar = QPushButton("Exportar")
        self.button_exportar.clicked.connect(self.exportar_excel)

        # layout Principal
        layout.addWidget(self.table)
        layout.addWidget(label_colunas_exportar)
        layout.addWidget(self.colunas_selecionadas_list)
        layout.addWidget(self.button_exportar)

        self.setLayout(layout)

    def atualizar_tabela(self):
        itens_selecionados = self.colunas_selecionadas_list.selectedItems()
        colunas_selecionadas = [item.text() for item in itens_selecionados]

        if colunas_selecionadas:
            df_filtrado = self.df[colunas_selecionadas]
            self.table.setModel(PandasModel(df_filtrado))
        else:
            self.table.setModel(PandasModel(self.df))

    def exportar_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Salvar arquivo", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        colunas_selecionadas = [
            item.text() for item in self.colunas_selecionadas_list.selectedItems()
        ]
        if not colunas_selecionadas:
            QMessageBox.warning(self, "Erro", "Selecione pelo menos uma coluna")
        if file_path:
            self.df[colunas_selecionadas].to_excel(file_path, index=False)
            QMessageBox.information(self, "Sucesso", "Arquivo salvo com sucesso.")

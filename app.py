from os import path
import sys
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QProgressBar,
    QTableView,
    QComboBox,
    QCheckBox,
    QListWidget,
    QAbstractItemView
)
from PyQt5.QtCore import QAbstractTableModel, Qt
import pandas as pd


class Janela(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Leitor de Arquivo")
        self.resize(800, 600)
        self.setMinimumSize(600, 400)

        self.df1 = None
        self.df2 = None

        layout = QVBoxLayout()
        layout_table = QHBoxLayout()

        self.table_primeiro_arquivo = QTableView()
        self.table_segundo_arquivo = QTableView()

        self.button_reset = QPushButton("Reset")
        self.button_reset.clicked.connect(self.resetar_campos)

        self.coluna1_combo = QComboBox()
        self.coluna1_combo.addItem("Escolha a coluna da primeira planilha")

        self.coluna2_combo = QComboBox()
        self.coluna2_combo.addItem("Escolha a coluna da segunda planilha")

        self.progress_bar = QProgressBar()

        self.button1 = QPushButton("Abrir primeiro arquivo")
        self.button1.clicked.connect(self.abrir_arquivo1)

        self.button2 = QPushButton("Abrir segundo arquivo")
        self.button2.clicked.connect(self.abrir_arquivo2)

        self.checkbox_soma = QCheckBox("Soma")
        self.checkbox_multiplicacao = QCheckBox("Multiplicação")
        self.checkbox_divisao = QCheckBox("Divisão")
        self.checkbox_media = QCheckBox("Média")
        self.checkbox_diferenca = QCheckBox("Diferença")
        self.checkbox_diferenca_percentual = QCheckBox("Diferença Percentual")

        self.button_comparar = QPushButton("Comparar planilhas")
        self.button_comparar.clicked.connect(self.comparar_planilhas)

        
        layout.addWidget(self.button_reset)
        layout.addWidget(self.button1)
        layout.addWidget(self.button2)
        layout.addWidget(self.coluna1_combo)
        layout.addWidget(self.coluna2_combo)
        layout.addWidget(self.checkbox_soma)
        layout.addWidget(self.checkbox_multiplicacao)
        layout.addWidget(self.checkbox_divisao)
        layout.addWidget(self.checkbox_media)
        layout.addWidget(self.checkbox_diferenca)
        layout.addWidget(self.checkbox_diferenca_percentual)
        layout.addWidget(self.progress_bar)
        layout.addLayout(layout_table)
        layout_table.addWidget(self.table_primeiro_arquivo)
        layout_table.addWidget(self.table_segundo_arquivo)
        layout.addWidget(self.button_comparar)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.addSpacing(10)

        self.progress_bar.setValue(0)

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
        self.df1 = self.carregar_excel()
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
        self.df2 = self.carregar_excel()
        if self.df2 is not None:
            self.coluna2_combo.clear()
            self.progress_bar.setValue(50)
            model = PandasModel(self.df2)
            self.table_segundo_arquivo.setModel(model)
            self.table_segundo_arquivo.setSortingEnabled(True)
            self.coluna2_combo.addItem("Escolha a coluna da segunda planilha")
            self.coluna2_combo.addItems(list(self.df2.columns))
            self.progress_bar.setValue(100)

    def carregar_excel(self):
        filters = "Excel (*.xlsx);;Todos os Arquivos (*)"
        file_path, filter_used = QFileDialog.getOpenFileName(
            self, "Abrir arquivo", "", filters
        )
        if not file_path:
            return None

        try:
            df = pd.read_excel(file_path)
            return df
        except Exception as e:
            QMessageBox.warning(self, "Erro", f"Erro ao carregar: {str(e)}")
            return None

    def comparar_planilhas(self):
        if self.df1 is None or self.df2 is None:
            QMessageBox.warning(
                self, "Erro", "Carregue as duas planilhas antes de comparar."
            )
            return

        coluna1 = f"{self.coluna1_combo.currentText()}_1"
        coluna2 = f"{self.coluna2_combo.currentText()}_2"

        if (
            self.coluna1_combo.currentText() == "Escolha a coluna da primeira planilha"
            or self.coluna2_combo.currentText() == "Escolha a coluna da segunda planilha"
        ):
            QMessageBox.warning(
                self, "Erro", "Por favor, seleciona as duas colunas para comparar"
            )
            return

        self.comparado = pd.merge(
            self.df1,
            self.df2,
            left_on=self.df1.columns[0],
            right_on=self.df2.columns[0],
            how="outer",
            suffixes=("_1", "_2"),
        )

        if self.checkbox_soma.isChecked():
            self.comparado["Soma"] = self.comparado[coluna1] + self.comparado[coluna2]
        if self.checkbox_diferenca.isChecked():
            self.comparado["Diferença"] = (
                self.comparado[coluna1] - self.comparado[coluna2]
            )
        if self.checkbox_multiplicacao.isChecked():
            self.comparado["Multiplicação"] = (
                self.comparado[coluna1] * self.comparado[coluna2]
            )
        if self.checkbox_divisao.isChecked():
            self.comparado["Divisão"] = (
                self.comparado[coluna1] / self.comparado[coluna2]
            )
        if self.checkbox_media.isChecked():
            self.comparado["Média"] = (
                self.comparado[coluna1] + self.comparado[coluna2]
            ) / 2
        if self.checkbox_diferenca_percentual.isChecked():
            self.comparado["Diferença percentual"] = (
                (self.comparado[coluna1] - self.comparado[coluna2])
                / self.comparado[coluna2]
            ) * 100

        self.janela_comparacao = JanelaComparacao(self.comparado)
        self.janela_comparacao.show()

    def conectar_exportar(self):
        self.exportar_excel(self.comparado)


class JanelaComparacao(QWidget):
    def __init__(self, df):
        super().__init__()
        self.setWindowTitle("Comparação")
        self.resize(600, 400)
        self.setMinimumSize(400, 200)

        self.df = df

        layout = QVBoxLayout()

        self.table = QTableView()
        self.table.setModel(PandasModel(df))


        self.colunas_selecionadas_list = QListWidget()
        self.colunas_selecionadas_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.colunas_selecionadas_list.itemSelectionChanged.connect(self.atualizar_tabela)
        self.colunas_selecionadas_list.addItems(self.df.columns)

        self.button_exportar = QPushButton("Exportar")
        self.button_exportar.clicked.connect(self.exportar_excel)

        layout.addWidget(self.table)
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
        colunas_selecionadas = [item.text() for item in self.colunas_selecionadas_list.selectedItems()]
        if not colunas_selecionadas:
            QMessageBox.warning(self, "Erro", "Selecione pelo menos uma coluna")
        if file_path:
            self.df[colunas_selecionadas].to_excel(file_path, index=False)
            QMessageBox.information(self, "Sucesso", "Arquivo salvo com sucesso.")


class PandasModel(QAbstractTableModel):
    def __init__(self, df):
        super().__init__()
        self.df = df

    def rowCount(self, parent=None):
        return self.df.shape[0]

    def columnCount(self, parent=None):
        return self.df.shape[1]

    def data(self, index, role=Qt.DisplayRole):  # type: ignore
        if index.isValid() and role == Qt.DisplayRole:  # type: ignore
            return str(self.df.iat[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):  # type: ignore
        if role != Qt.DisplayRole:  # type: ignore
            return None
        if orientation == Qt.Horizontal:  # type: ignore
            return str(self.df.columns[section])
        else:
            return str(self.df.index[section])

    def sort(self, column, order):  # type: ignore
        column_name = self.df.columns[column]
        ascending = order == Qt.AscendingOrder  # type: ignore
        self.df.sort_values(by=column_name, ascending=ascending, inplace=True)
        self.layoutChanged.emit()


app = QApplication(sys.argv)
window = Janela()
window.show()
sys.exit(app.exec_())

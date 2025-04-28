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
    QLineEdit,
    QCheckBox,
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

        self.coluna1_input = QLineEdit()
        self.coluna1_input.setPlaceholderText(
            "Coluna da primeira planilha para cálculo"
        )

        self.coluna2_input = QLineEdit()
        self.coluna2_input.setPlaceholderText("Coluna da segunda planilha para cálculo")

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

        layout.addLayout(layout_table)
        layout_table.addWidget(self.table_primeiro_arquivo)
        layout_table.addWidget(self.table_segundo_arquivo)
        layout.addWidget(self.coluna1_input)
        layout.addWidget(self.coluna2_input)
        layout.addWidget(self.checkbox_soma)
        layout.addWidget(self.checkbox_multiplicacao)
        layout.addWidget(self.checkbox_divisao)
        layout.addWidget(self.checkbox_media)
        layout.addWidget(self.checkbox_diferenca)
        layout.addWidget(self.checkbox_diferenca_percentual)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.button1)
        layout.addWidget(self.button2)
        layout.addWidget(self.button_comparar)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.addSpacing(10)

        self.setLayout(layout)

    def abrir_arquivo1(self):
        self.progress_bar.setValue(0)
        self.df1 = self.carregar_excel()
        if self.df1 is not None:
            self.progress_bar.setValue(50)
            model = PandasModel(self.df1)
            self.table_primeiro_arquivo.setModel(model)
            self.table_primeiro_arquivo.setSortingEnabled(True)
            self.progress_bar.setValue(100)

    def abrir_arquivo2(self):
        self.progress_bar.setValue(0)
        self.df2 = self.carregar_excel()
        if self.df2 is not None:
            self.progress_bar.setValue(50)
            model = PandasModel(self.df2)
            self.table_segundo_arquivo.setModel(model)
            self.table_segundo_arquivo.setSortingEnabled(True)
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

        coluna1 = f"{self.coluna1_input.text()}_1"
        coluna2 = f"{self.coluna2_input.text()}_2"

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

        table = QTableView()
        table.setModel(PandasModel(df))

        self.button = QPushButton("Exportar")
        self.button.clicked.connect(self.exportar_excel)

        layout.addWidget(table)
        layout.addWidget(self.button)
        self.setLayout(layout)

    def exportar_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Salvar arquivo", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            self.df.to_excel(file_path, index=False)
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

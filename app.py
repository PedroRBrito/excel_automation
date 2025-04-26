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
)
from PyQt5.QtCore import QAbstractTableModel, QObject, Qt
from click import progressbar
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

        self.progress_bar = QProgressBar()

        self.button1 = QPushButton("Abrir arquivo")
        self.button1.clicked.connect(self.abrir_arquivo1)

        self.button2 = QPushButton("Abrir segundo arquivo")
        self.button2.clicked.connect(self.abrir_arquivo2)

        self.button_comparar = QPushButton("Comparar planilhas")
        self.button_comparar.clicked.connect(self.comparar_planilhas)

        layout.addLayout(layout_table)
        layout_table.addWidget(self.table_primeiro_arquivo)
        layout_table.addWidget(self.table_segundo_arquivo)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.button1)
        layout.addWidget(self.button2)
        layout.addWidget(self.button_comparar)
        self.setLayout(layout)

    def abrir_arquivo1(self):
        self.df1 = self.carregar_excel()
        if self.df1 is not None:
            model = PandasModel(self.df1)
            self.table_primeiro_arquivo.setModel(model)
            self.table_primeiro_arquivo.setSortingEnabled(True)
            QMessageBox.information(
                self, "Arquivo 1", f"{len(self.df1)} linhas carregadas"
            )

    def abrir_arquivo2(self):
        self.df2 = self.carregar_excel()
        if self.df2 is not None:
            model = PandasModel(self.df2)
            self.table_segundo_arquivo.setModel(model)
            self.table_segundo_arquivo.setSortingEnabled(True)
            QMessageBox.information(
                self, "Arquivo 2", f"{len(self.df2)} linhas carregadas!"
            )

    def carregar_excel(self):
        filters = "Excel (*.xlsx);;Todos os Arquivos (*)"
        file_path, filter_used = QFileDialog.getOpenFileName(
            self, "Abrir arquivo", "", filters
        )
        if not file_path:
            return None

        try:
            self.progress_bar.setValue(0)
            self.progress_bar.setMaximum(100)
            df = pd.read_excel(file_path)
            self.progress_bar.setValue(100)
            return df
        except Exception as e:
            QMessageBox.warning(self, "Erro", f"Erro ao carregar: {str(e)}")
            return None

    def comparar_planilhas(self):
        if self.df1 is None or self.df2 is None:
            QMessageBox.warning(
                self, "Erro", "Carregue as duas planilhas antes de comparar."
            )

        df1 = self.df1.rename(columns={"Valor": "Realizado"}) # type: ignore
        df2 = self.df2.rename(columns={"Valor": "Previsto"}) # type: ignore

        self.comparado = pd.merge(df1, df2, on="categoria", how="outer")

        self.comparado["Diferença"] = self.comparado["Realizado"] - self.comparado["Previsto"]
        self.comparado["% Diferença"] = (
            self.comparado["Diferença"] / self.comparado["Previsto"]
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

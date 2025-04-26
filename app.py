from os import path
import sys
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
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

        self.table = QTableView()
        layout.addWidget(self.table)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        self.button1 = QPushButton("Abrir arquivo")
        self.button1.clicked.connect(self.abrir_arquivo1)

        self.button2 = QPushButton("Abrir segundo arquivo")
        self.button2.clicked.connect(self.abrir_arquivo2)

        self.button_comparar = QPushButton("Comparar planilhas")
        self.button_comparar.clicked.connect(self.comparar_planilhas)

        self.button_exportar = QPushButton("Exportar planilha")
        self.button_exportar.clicked.connect(self.conectar_exportar)

        layout.addWidget(self.button1)
        layout.addWidget(self.button2)
        layout.addWidget(self.button_comparar)
        layout.addWidget(self.button_exportar)
        self.setLayout(layout)

    def abrir_arquivo1(self):
        self.df1 = self.carregar_excel()
        if self.df1 is not None:
            model = PandasModel(self.df1)
            self.table.setModel(model)
            self.table.setSortingEnabled(True)
            QMessageBox.information(
                self, "Arquivo 1", f"{len(self.df1)} linhas carregadas"
            )

    def abrir_arquivo2(self):
        self.df2 = self.carregar_excel()
        if self.df2 is not None:
            print(self.df2)
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

        model = PandasModel(self.comparado)
        self.table.setModel(model)
        self.table.setSortingEnabled(True)

    def exportar_excel(self, comparado):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Salvar arquivo", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            comparado.to_excel(file_path, index=False)
            QMessageBox.information(self, "Sucesso", "Arquivo salvo com sucesso.")

    def conectar_exportar(self):
        self.exportar_excel(self.comparado)


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

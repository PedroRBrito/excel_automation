import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox


def carregar_excel(parent):
    filters = "Excel (*.xlsx);;Todos os Arquivos (*)"
    file_path, _ = QFileDialog.getOpenFileName(parent, "Abrir arquivo", "", filters)

    if not file_path:
        return None

    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        QMessageBox.warning(parent, "Erro", f"Erro ao carregar o arquivo: {str(e)}")
        return None
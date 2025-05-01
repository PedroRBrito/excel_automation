from PyQt5.QtCore import QAbstractTableModel, Qt


class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole): # type: ignore
        if index.isValid() and role == Qt.DisplayRole: # type: ignore
            return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole): # type: ignore
        if role == Qt.DisplayRole: # type: ignore
            if orientation == Qt.Horizontal: # type: ignore
                return self._data.columns[section]
            if orientation == Qt.Vertical: # type: ignore
                return str(self._data.index[section])
        return None
def button_style():
    return """
    QPushButton {
        background-color: #0033A0;
        color: white;
        padding: 6px 12px;
        border: none;
        border-radius: 5px;
    }
    QPushButton:hover {
        background-color: #002080;
    }
"""


def progress_bar_style():
    return """
QProgressBar {
    border: 1px solid #666;
    border-radius: 5px;
    background-color: #CCCCCC;
    text-align: center;
    color: #0033A0;
    font: bold 10pt "Segoe UI";
    padding: 0px;
}

QProgressBar::chunk {
    background-color: white;
    width: 10px;
    margin: 0.5px;
}
"""


def qlistwidget_style():
    return """
QListWidget {
    background-color: #FFFFFF;
    border: 1px solid #666;
    padding: 5px;
    color: #000;
}
QListWidget::item:selected {
    background-color: #0033A0;
    color: white;
}
"""


def qcheckbox_style():
    return """
QCheckBox {
    color: #000000;
    spacing: 5px;
}
QCheckBox::indicator {
    width: 18px;
    height: 18px;
}
QCheckBox::indicator:checked {
    image: url(src/style/check_checkbox.png);
}
QCheckBox::indicator:unchecked {
    image: url(src/style/uncheck_checkbox.png);
}
"""


def qcombobox_style():
    return """
QComboBox {
    background-color: #FFFFFF;
    border: 1px solid #666;
    padding: 5px;
    color: #000000;
}
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 20px;
    border-left: 1px solid #666;
}
QComboBox::down-arrow {
    image: url(:/qt-project.org/styles/commonstyle/images/arrowdown-16.png);
}
QComboBox QAbstractItemView {
    background-color: #FFFFFF;
    selection-background-color: #0033A0;
    selection-color: white;
}
"""


def qtableview_style():
    return """
QTableView {
    background-color: #FFFFFF;
    gridline-color: #CCCCCC;
    color: #000000;
    selection-background-color: #0033A0;
    selection-color: white;
    border: 1px solid #666;
}
QHeaderView::section {
    background-color: #0033A0;
    color: white;
    padding: 4px;
    border: 1px solid #666;
}
"""

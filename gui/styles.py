from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QStyledItemDelegate


class AlignCenterDelegate(QStyledItemDelegate):
    """Устанавливает положение значений в столбце таблицы по центру."""
    def initStyleOption(self, option, index):
        super(AlignCenterDelegate, self).initStyleOption(option, index)
        option.displayAlignment = Qt.AlignCenter


class AlignLeftDelegate(QStyledItemDelegate):
    """Устанавливает положение значений в столбце таблицы слева."""
    def initStyleOption(self, option, index):
        super(AlignLeftDelegate, self).initStyleOption(option, index)
        option.displayAlignment = (Qt.AlignLeft | Qt.AlignVCenter)


rx_tx_white_stylesheet = """
    border-radius: 6px;
    min-height: 12px;
    max-height: 12px;
    min-width: 12px;
    max-width: 12px;
    background-color: white;
"""

rx_stylesheet = """
    border-radius: 6px;
    min-height: 12px;
    max-height: 12px;
    min-width: 12px;
    max-width: 12px;
    background-color: green;
"""
tx_stylesheet = """
    border-radius: 6px;
    min-height: 12px;
    max-height: 12px;
    min-width: 12px;
    max-width: 12px;
    background-color: red;
"""

kl_adress_font = 'font: bold 13px'
font_bold_11 = 'font: bold 11px'

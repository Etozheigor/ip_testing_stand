from datetime import timedelta

from PyQt5.QtCore import QDate, Qt, QTimer
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import (QCheckBox, QComboBox, QDateEdit, QDialog,
                             QHBoxLayout, QLabel, QPushButton, QVBoxLayout)

from db.utils import (combobox_remembered_text, delete_remembered_combobox,
                      is_checked_checkbox, remember_checkbox,
                      remember_combobox)
from gui.constants import REPORT_FOR_PERIOD, TESTS_FOR_PERIOD
from modbus.utils import get_ports


class SelectPortDialog(QDialog):
    """Диалоговое окно для выбора порта."""
    def __init__(self):
        super().__init__()
        self.setModal(True)
        self.setWindowTitle('Выбор порта')
        self.setWindowFlags(Qt.WindowTitleHint)
        self.setFixedSize(200, 200)
        self.layout_dialog = QVBoxLayout()
        self.setLayout(self.layout_dialog)

        # конфигурация комбобокса
        self.select_port = QComboBox()
        self.ports_list = get_ports()
        self.select_port.addItems(self.ports_list)
        if is_checked_checkbox():
            if combobox_remembered_text() in self.ports_list:
                self.select_port.setCurrentIndex(self.ports_list.index(combobox_remembered_text()))
            else:
                self.select_port.setCurrentIndex(self.ports_list.index('Не выбран'))
        else:
            self.select_port.setCurrentIndex(self.ports_list.index('Не выбран'))
        self.layout_dialog.addWidget(self.select_port)

        # конфигурация чекбокса
        self.layout_checkbox = QHBoxLayout()
        self.checkbox_remember_port = QCheckBox()
        self.checkbox_remember_port.setChecked(is_checked_checkbox())
        self.lablel_remember_port = QLabel('Запомнить порт')
        self.layout_checkbox.addWidget(self.checkbox_remember_port)
        self.layout_checkbox.addWidget(self.lablel_remember_port)
        self.layout_checkbox.setAlignment(Qt.AlignCenter | Qt.AlignBottom)
        self.layout_dialog.addLayout(self.layout_checkbox)
        self.layout_dialog.addSpacing(20)

        # конфигурация кнопки Ок
        self.bnt_ok = QPushButton('Ок')
        self.bnt_ok.clicked.connect(self.ok_click)
        self.layout_btn_ok = QHBoxLayout()
        self.layout_btn_ok.addWidget(self.bnt_ok)
        self.layout_btn_ok.setAlignment(Qt.AlignHCenter)
        self.layout_dialog.addLayout(self.layout_btn_ok)

        # конфигурация таймеров и флагов
        self.selected_port = 'Не выбран'
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_ports)
        self.timer.start(1000)

    def update_ports(self):
        """Обновляет список доступных портов."""
        new_ports_list = get_ports()
        if self.ports_list != new_ports_list:
            self.ports_list = new_ports_list
            self.select_port.clear()
            self.select_port.addItems(new_ports_list)

    def ok_click(self):
        """Применяет выбранный порт и закрывает диалог."""
        is_checked = self.checkbox_remember_port.isChecked()
        remember_checkbox(is_checked)
        combobox_text = self.select_port.currentText()
        if is_checked:
            remember_combobox(combobox_text)
        else:
            delete_remembered_combobox()
        self.selected_port = combobox_text
        self.close()


class SelectReportPeriodDialog(QDialog):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.setModal(True)
        self.setWindowTitle('Выберите период')
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setFixedSize(230, 100)
        self.layout_dialog = QVBoxLayout()
        self.setLayout(self.layout_dialog)

        self.layout_filter_checkbox = QVBoxLayout()
        self.layout_filter = QHBoxLayout()
        self.filter_date_from = QDateEdit(QDate.currentDate())
        self.filter_date_from.setDisplayFormat('yyyy-MM-dd')
        self.filter_date_from.setCalendarPopup(True)
        self.filter_date_to = QDateEdit(QDate.currentDate())
        self.filter_date_to.setDisplayFormat('yyyy-MM-dd')
        self.filter_date_to.setCalendarPopup(True)
        self.layout_filter.addWidget(QLabel('c'))
        self.layout_filter.addWidget(self.filter_date_from)
        self.layout_filter.addWidget(QLabel('по'))
        self.layout_filter.addWidget(self.filter_date_to)

        self.layout_filter_checkbox.addLayout(self.layout_filter)
        self.layout_dialog.addLayout(self.layout_filter_checkbox)

        self.layout_ok_btn = QHBoxLayout()
        self.btn_ok = QPushButton('Применить')
        self.btn_ok.clicked.connect(self.apply)
        self.btn_cancel = QPushButton('Отмена')
        self.btn_cancel.clicked.connect(self.close)
        self.layout_ok_btn.addWidget(self.btn_ok)
        self.layout_ok_btn.addWidget(self.btn_cancel)
        self.layout_dialog.addLayout(self.layout_ok_btn)

    def apply(self):
        date_from = self.filter_date_from.date().toPyDate()
        date_to = self.filter_date_to.date().toPyDate()
        date_to += timedelta(days=1)
        self.close()
        self.parent.save_table(table=REPORT_FOR_PERIOD, period=(date_from, date_to))


class SelectTestPeriodDialog(SelectReportPeriodDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.layout_error_checkbox = QHBoxLayout()
        self.label_error = QLabel('Выводить только ошибки:')
        self.checkbox_error = QCheckBox()
        self.layout_error_checkbox.addWidget(self.label_error)
        self.layout_error_checkbox.addWidget(self.checkbox_error)
        self.layout_filter_checkbox.addLayout(self.layout_error_checkbox)

    def apply(self):
        date_from = self.filter_date_from.date().toPyDate()
        date_to = self.filter_date_to.date().toPyDate()
        date_to += timedelta(days=1)
        self.close()
        self.parent.save_table(table=TESTS_FOR_PERIOD, period=(date_from, date_to), is_error_flag=self.checkbox_error.isChecked())


class AboutProgramDialog(QDialog):
    """Диалоговое окно для пункта меню 'О программе'"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle('О программе')
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setFixedSize(270, 300)
        self.layout_dialog = QVBoxLayout()
        self.setLayout(self.layout_dialog)

        self.label_picture = QLabel()
        self.label_description = QLabel(
            'Описание: Программа для сбора и анализа\n проверок ИП')
        self.label_description.setAlignment(Qt.AlignCenter)
        self.label_version = QLabel('Версия: 1.3')
        self.label_version.setAlignment(Qt.AlignCenter)
        self.label_website = QLabel(
            'Сайт компании: <a href="***"> *** </a>')
        self.label_support_email = QLabel(
            'Техническая поддержка: ***')
        self.label_telephone = QLabel('Телефон офиса: ***')
        self.label_telephone.setAlignment(Qt.AlignCenter)
        self.label_website.setOpenExternalLinks(True)
        self.label_website.setAlignment(Qt.AlignCenter)

        self.layout_dialog.addWidget(self.label_picture)
        self.layout_dialog.addWidget(self.label_description)
        self.layout_dialog.addWidget(self.label_version)
        self.layout_dialog.addWidget(self.label_website)
        self.layout_dialog.addWidget(self.label_support_email)
        self.layout_dialog.addWidget(self.label_telephone)

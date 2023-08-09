from datetime import datetime

from PyQt5.QtCore import Qt, QTime, QTimer
from PyQt5.QtWidgets import (QAbstractItemView, QBoxLayout, QCheckBox,
                             QComboBox, QGridLayout, QHBoxLayout, QLabel,
                             QLineEdit, QPushButton, QScrollArea, QSizePolicy,
                             QTableWidget, QVBoxLayout, QWidget)

from db.models import Report
from gui.styles import (AlignCenterDelegate, AlignLeftDelegate, font_bold_11,
                        kl_adress_font)


class StatisticWidget(QWidget):
    """Окно общей статистики."""
    def __init__(self):
        super().__init__()
        self.setFixedWidth(250)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet('background-color: white')
        self.setMinimumSize(280, 220)

        self.main_layout = QHBoxLayout()
        self.setLayout(self.main_layout)
        self.scrolled_area = QScrollArea(self)
        self.scrolled_area.setWidgetResizable(True)
        self.central_widget = QWidget()
        self.layout_statstic = QVBoxLayout()
        self.layout_statstic.setAlignment(Qt.AlignTop)
        self.central_widget.setLayout(self.layout_statstic)
        self.scrolled_area.setWidget(self.central_widget)
        self.main_layout.addWidget(self.scrolled_area)

        self.btn_start = QPushButton('Старт')
        self.btn_save_report = QPushButton('Сохранить отчет')
        self.btn_save_tests = QPushButton('Сохранить результаты тестов')
        self.label_all_devices = QLabel('    Всего тестов запущено: 0')
        self.label_all_devices.setStyleSheet(font_bold_11)
        self.label_success_all = QLabel('- успешно: 0')
        self.label_error_all = QLabel('- ошибка: 0')
        self.label_canceled_all = QLabel('- прервано: 0')

        self.label_success = QLabel('    Успешно:')
        self.label_success.setStyleSheet(font_bold_11)
        self.label_success_ipt = QLabel('- ИПТ-И: 0')
        self.label_success_ipt_s = QLabel('- ИПТ-СИ: 0')
        self.label_success_mks = QLabel('- МКС: 0')
        self.label_success_ipd = QLabel('- ИПД: 0')
        self.label_success_ipr = QLabel('- ИПР: 0')

        self.label_error = QLabel('    Ошибка:')
        self.label_error.setStyleSheet(font_bold_11)
        self.label_error_ipt = QLabel('- ИПТ-И: 0')
        self.label_error_ipt_s = QLabel('- ИПТ-СИ: 0')
        self.label_error_mks = QLabel('- МКС: 0')
        self.label_error_ipd = QLabel('- ИПД: 0')
        self.label_error_ipr = QLabel('- ИПР: 0')

        self.layout_statstic.addWidget(self.btn_start)
        self.layout_statstic.addWidget(self.btn_save_report)
        self.layout_statstic.addWidget(self.btn_save_tests)
        self.layout_statstic.addWidget(self.label_all_devices)
        self.layout_statstic.addWidget(self.label_success_all)
        self.layout_statstic.addWidget(self.label_error_all)
        self.layout_statstic.addWidget(self.label_canceled_all)
        self.layout_statstic.addSpacing(20)

        self.layout_statstic.addWidget(self.label_success)
        self.layout_statstic.addWidget(self.label_success_ipt)
        self.layout_statstic.addWidget(self.label_success_ipt_s)
        self.layout_statstic.addWidget(self.label_success_mks)
        self.layout_statstic.addWidget(self.label_success_ipd)
        self.layout_statstic.addWidget(self.label_success_ipr)
        self.layout_statstic.addSpacing(20)

        self.layout_statstic.addWidget(self.label_error)
        self.layout_statstic.addWidget(self.label_error_ipt)
        self.layout_statstic.addWidget(self.label_error_ipt_s)
        self.layout_statstic.addWidget(self.label_error_mks)
        self.layout_statstic.addWidget(self.label_error_ipd)
        self.layout_statstic.addWidget(self.label_error_ipr)

        self.statistics = Report(
            date_time=datetime.now(),
            success_all=0,
            interrupted_all=0,
            error_all=0,
            success_ipr=0,
            success_ipt=0,
            success_ipt_s=0,
            success_ipd=0,
            success_ipp=0,
            success_ipp_s=0,
            success_udp=0,
            success_mks=0,
            error_ipr=0,
            error_ipt=0,
            error_ipt_s=0,
            error_ipd=0,
            error_ipp=0,
            error_ipp_s=0,
            error_udp=0,
            error_mks=0
        )

    def update_statistics(self, statistics_difference):
        self.statistics.success_all += statistics_difference[0]
        self.statistics.error_all += statistics_difference[1]
        self.statistics.interrupted_all += statistics_difference[2]
        self.statistics.success_ipt += statistics_difference[3]
        self.statistics.success_ipt_s += statistics_difference[4]
        self.statistics.success_mks += statistics_difference[5]
        self.statistics.success_ipd += statistics_difference[6]
        self.statistics.success_ipr += statistics_difference[7]
        self.statistics.error_ipt += statistics_difference[8]
        self.statistics.error_ipt_s += statistics_difference[9]
        self.statistics.error_mks += statistics_difference[10]
        self.statistics.error_ipd += statistics_difference[11]
        self.statistics.error_ipr += statistics_difference[12]
        self.update_labels()

    def update_labels(self):
        all_devices = self.statistics.success_all + self.statistics.interrupted_all + self.statistics.error_all
        self.label_all_devices.setText(f'    Всего тестов запущено: {all_devices}')
        self.label_success_all.setText(f'- успешно: {self.statistics.success_all}')
        self.label_error_all.setText(f'- ошибка: {self.statistics.error_all}')
        self.label_canceled_all.setText(f'- прервано: {self.statistics.interrupted_all}')

        self.label_success_ipt.setText(f'- ИПТ-И: {self.statistics.success_ipt}')
        self.label_success_ipt_s.setText(f'- ИПТ-CИ: {self.statistics.success_ipt_s}')
        self.label_success_mks.setText(f'- МКС: {self.statistics.success_mks}')
        self.label_success_ipd.setText(f'- ИПД: {self.statistics.success_ipd}')
        self.label_success_ipr.setText(f'- ИПР: {self.statistics.success_ipr}')

        self.label_error_ipt.setText(f'- ИПТ-И: {self.statistics.error_ipt}')
        self.label_error_ipt_s.setText(f'- ИПТ-CИ: {self.statistics.error_ipt_s}')
        self.label_error_mks.setText(f'- МКС: {self.statistics.error_mks}')
        self.label_error_ipd.setText(f'- ИПД: {self.statistics.error_ipd}')
        self.label_error_ipr.setText(f'- ИПР: {self.statistics.error_ipr}')


class LayoutMainTableToolBar(QHBoxLayout):
    """Тулбар главной таблицы."""
    def __init__(self, layout_filters):
        super().__init__()
        self.layout_filters = layout_filters
        self.addLayout(self.layout_filters)
        # конфигурация кнопок
        self.layout_buttons = QHBoxLayout()
        self.btn_filter = QPushButton('Применить фильтр')
        self.btn_download_to_file = QPushButton('Экспорт в файл')
        self.btn_clean = QPushButton('Очистить')
        self.layout_buttons.addWidget(self.btn_filter)
        self.layout_buttons.addWidget(self.btn_download_to_file)
        self.layout_buttons.addWidget(self.btn_clean)
        self.addLayout(self.layout_buttons)


class LayoutJournalToolBar(LayoutMainTableToolBar):
    """Тулбар таблицы журнала."""
    pass


class LayoutFilterJournal(QHBoxLayout):
    """Слой фильтров журнала."""
    def __init__(self):
        super().__init__()
        self.addStretch(0)
        self.setDirection(QBoxLayout.RightToLeft)
        self.label_events_filter = QLabel('Фильтр событий:')
        self.label_kl = QLabel("КЛ")
        self.filter_kl = QLineEdit()
        self.filter_kl.setMaximumWidth(40)
        self.checkbox_kl = QCheckBox()
        self.label_event_code = QLabel("Код события")
        self.filter_event_code = QComboBox()
        self.filter_event_code.SizeAdjustPolicy(QComboBox.AdjustToContents)
        self.filter_event_code.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Minimum)
        self.checkbox_event_code = QCheckBox()
        self.addWidget(self.checkbox_event_code)
        self.addWidget(self.filter_event_code)
        self.addWidget(self.label_event_code)
        self.addSpacing(15)
        self.addWidget(self.checkbox_kl)
        self.addWidget(self.filter_kl)
        self.addWidget(self.label_kl)
        self.addSpacing(15)
        self.addWidget(self.label_events_filter)


class LayoutFilterMainTable(QHBoxLayout):
    """Слой фильтров главной таблицы."""
    def __init__(self):
        super().__init__()
        self.addStretch(0)
        self.setDirection(QBoxLayout.RightToLeft)
        self.label_events_filter = QLabel('Фильтр событий:')
        self.label_kl = QLabel("КЛ")
        self.filter_kl = QLineEdit()
        self.checkbox_kl = QCheckBox()

        self.label_test_number = QLabel("№ теста: ")
        self.filter_exit_code_1 = QLineEdit()
        self.label_exit_code_1 = QLabel("Код 1: ")
        self.filter_exit_code_2 = QLineEdit()
        self.label_exit_code_2 = QLabel("Код 2: ")
        self.filter_exit_code_3 = QLineEdit()
        self.checkbox_exit_code = QCheckBox()

        self.label_device_type = QLabel("Тип устройства")
        for filter in [
            self.filter_kl, self.filter_exit_code_1,
            self.filter_exit_code_2, self.filter_exit_code_3
        ]:
            filter.setMaximumWidth(40)
        self.filter_device_type = QComboBox()
        self.filter_device_type.SizeAdjustPolicy(QComboBox.AdjustToContents)
        self.filter_device_type.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Minimum)
        self.filter_device_type.addItems([
            'Не выбрано', 'ИПР', 'ИПТ', 'ИПТ-С', 'ИПД', 'ИПП', 'ИПП-С', 'УДП', 'МКВ-2',
            'МКВ-4', 'МКР-2', 'МКР-4', 'ИПР-В', 'МКВ2Р2', 'МКВ2А', 'МКО-В', 'МКО-С', 'МКС'
        ])
        self.checkbox_device_type = QCheckBox()
        self.addWidget(self.checkbox_device_type)
        self.addWidget(self.filter_device_type)
        self.addWidget(self.label_device_type)
        self.addSpacing(15)
        self.addWidget(self.checkbox_exit_code)
        self.addWidget(self.filter_exit_code_3)
        self.addWidget(self.label_exit_code_2)
        self.addWidget(self.filter_exit_code_2)
        self.addWidget(self.label_exit_code_1)
        self.addWidget(self.filter_exit_code_1)
        self.addWidget(self.label_test_number)
        self.addSpacing(15)
        self.addWidget(self.checkbox_kl)
        self.addWidget(self.filter_kl)
        self.addWidget(self.label_kl)
        self.addSpacing(15)
        self.addWidget(self.label_events_filter)


class JournalTable(QTableWidget):
    """Таблица журнала."""
    def __init__(self):
        super().__init__()
        self.setColumnCount(8)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setHorizontalHeaderLabels(
            ['№', 'Дата и время', 'КЛ', 'Номер входа', 'Адрес МА', 'Код', 'Доп. параметры', 'Описание'])
        self.horizontalHeader().setStretchLastSection(True)
        self.horizontalHeader().setVisible(True)
        for col in range(7):
            self.setColumnWidth(col, 40)
        self.setColumnWidth(1, 120)
        self.setColumnWidth(3, 90)
        self.setColumnWidth(4, 75)
        self.setColumnWidth(6, 107)
        self.table_center_alignment = AlignCenterDelegate(self)
        self.table_left_alignment = AlignLeftDelegate(self)
        self.setItemDelegate(self.table_center_alignment)  # выравниевание данных в ячейках по центру
        self.setItemDelegateForColumn(7, self.table_left_alignment)  # выравнивание данных в ячейках по левому краю


class MainTable(QTableWidget):
    """Главная таблица."""
    def __init__(self):
        super().__init__()
        self.setColumnCount(27)
        self.setRowCount(0)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setHorizontalHeaderLabels([
            '№', 'Время', 'КЛ', '№ теста', 'Код 1', 'Код 2',
            '1\nВремя\nготовности', '2\nТип', "2\nВерсия", "2\nДлит.\nСБ", "2\nТок\nСБ", "2\nСредн.\nток", "4\nТ\n(С)", "5\nUв",
            "5\nUc", "5\nUн", "7\nСв", "7\nСс", "7\nСн", "8\nТок\nзаряда", "9\nСв", "9\nСс", "9\nСн", "Резерв 1", "Резерв 2",
            "Описание ошибки", "Доп. параметр"
        ])
        self.horizontalHeader().setStretchLastSection(True)
        self.horizontalHeader().setVisible(True)
        for col in [0, 2, 10, 23]:
            self.setColumnWidth(col, 50)
        self.setColumnWidth(6, 92)
        for col in [1, 4, 5, 9]:
            self.setColumnWidth(col, 65)
        for col in [7]:
            self.setColumnWidth(col, 60)
        for col in [8]:
            self.setColumnWidth(col, 70)
        for col in [11, 19]:
            self.setColumnWidth(col, 67)
        for col in [3, 23, 24]:
            self.setColumnWidth(col, 77)
        self.setColumnWidth(25, 600)
        self.setColumnWidth(26, 200)
        self.table_center_alignment = AlignCenterDelegate(self)
        self.setItemDelegate(self.table_center_alignment)  # выравниевание данных в ячейках по центру
        self.table_left_alignment = AlignLeftDelegate(self)
        self.setItemDelegateForColumn(25, self.table_left_alignment)
        self.setItemDelegateForColumn(26, self.table_left_alignment)


class GraphWidget(QWidget):
    """Поле графов КЛ."""
    def __init__(self):
        super().__init__()
        self.main_layout = QHBoxLayout()
        self.setLayout(self.main_layout)
        self.scrolled_area = QScrollArea(self)
        self.scrolled_area.setWidgetResizable(True)
        self.central_widget = QWidget()
        self.layout_graph = QGridLayout()
        self.central_widget.setLayout(self.layout_graph)
        self.scrolled_area.setWidget(self.central_widget)
        self.main_layout.addWidget(self.scrolled_area)

        # конфигрурация виджета графа
        self.graph_kl_1 = KLGraphWidget(1)
        self.graph_kl_2 = KLGraphWidget(2)
        self.graph_kl_3 = KLGraphWidget(3)
        self.graph_kl_4 = KLGraphWidget(4)
        self.graph_kl_5 = KLGraphWidget(5)
        self.graph_kl_6 = KLGraphWidget(6)
        self.graph_kl_7 = KLGraphWidget(7)
        self.graph_kl_8 = KLGraphWidget(8)
        self.graph_kl_9 = KLGraphWidget(9)
        self.graph_kl_10 = KLGraphWidget(10)
        self.graph_kl_11 = KLGraphWidget(11)
        self.graph_kl_12 = KLGraphWidget(12)
        self.graph_kl_13 = KLGraphWidget(13)
        self.graph_kl_14 = KLGraphWidget(14)
        self.graph_kl_15 = KLGraphWidget(15)
        self.graph_kl_16 = KLGraphWidget(16)
        graphs = [
            self.graph_kl_1, self.graph_kl_2, self.graph_kl_3, self.graph_kl_4,
            self.graph_kl_5, self.graph_kl_6, self.graph_kl_7, self.graph_kl_8,
            self.graph_kl_9, self.graph_kl_10, self.graph_kl_11, self.graph_kl_12,
            self.graph_kl_13, self.graph_kl_14, self.graph_kl_15, self.graph_kl_16
        ]
        positions = [(i, j) for i in range(1, 5) for j in range(1, 5)]
        for position, graph in zip(positions, graphs):
            self.layout_graph.addWidget(graph, *position)


class KLGraphWidget(QWidget):
    """Граф КЛ."""
    def __init__(self, address):
        super().__init__()
        self.address = address
        self.statistics = Report(
            date_time=datetime.now(),
            success_all=0,
            interrupted_all=0,
            error_all=0,
            success_ipr=0,
            success_ipt=0,
            success_ipt_s=0,
            success_ipd=0,
            success_ipp=0,
            success_ipp_s=0,
            success_udp=0,
            success_mks=0,
            error_ipr=0,
            error_ipt=0,
            error_ipt_s=0,
            error_ipd=0,
            error_ipp=0,
            error_ipp_s=0,
            error_udp=0,
            error_mks=0
        )
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet('background-color: lightGray')

        self.layout_main = QVBoxLayout()
        self.layout_main.setContentsMargins(0, 0, 0, 0)
        self.setLayout(self.layout_main)
        self.layout_main.setAlignment(Qt.AlignTop)

        self.name_version_widget = QWidget()
        self.layout_name_version = QHBoxLayout()
        self.label_name = QLabel(f'Адрес: {self.address}')
        self.label_name.setStyleSheet(kl_adress_font)
        self.label_version = QLabel('версия:    ')
        self.layout_name_version.addWidget(self.label_name)
        self.layout_name_version.addWidget(self.label_version, alignment=Qt.AlignRight)
        self.name_version_widget.setLayout(self.layout_name_version)
        # self.name_version_widget.setStyleSheet('background-color: red')
        self.layout_main.addWidget(self.name_version_widget)

        self.layout_total_success_error = QHBoxLayout()
        self.layout_total_success_error.setContentsMargins(5, 0, 5, 0)
        self.layout_total = QVBoxLayout()
        self.layout_total.setAlignment(Qt.AlignTop)
        self.layout_success = QVBoxLayout()
        self.layout_error = QVBoxLayout()
        self.layout_total_success_error.addLayout(self.layout_total)
        self.layout_total_success_error.addLayout(self.layout_success)
        self.layout_total_success_error.addLayout(self.layout_error)
        self.layout_main.addLayout(self.layout_total_success_error)

        self.label_total_all = QLabel(f'Всего: {0}')
        self.label_total_all.setStyleSheet(font_bold_11)
        self.label_success_all = QLabel(f'Успешно: {0}')
        self.label_error_all = QLabel(f'Ошибка: {0}')
        self.label_cancel_all = QLabel(f'Прервано: {0}')
        self.layout_total.addWidget(self.label_total_all)
        self.layout_total.addWidget(self.label_success_all)
        self.layout_total.addWidget(self.label_error_all)
        self.layout_total.addWidget(self.label_cancel_all)

        self.label_total_success = QLabel('Успешно:')
        self.label_total_success.setStyleSheet(font_bold_11)
        self.label_success_ipt = QLabel('ИПТ-И: 0')
        self.label_success_ipt_s = QLabel('ИПТ-СИ: 0')
        self.label_success_mks = QLabel('МКС: 0')
        self.label_success_ipd = QLabel('ИПД: 0')
        self.label_success_ipr = QLabel('ИПР: 0')
        self.layout_success.addWidget(self.label_total_success)
        self.layout_success.addWidget(self.label_success_ipt)
        self.layout_success.addWidget(self.label_success_ipt_s)
        self.layout_success.addWidget(self.label_success_mks)
        self.layout_success.addWidget(self.label_success_ipd)
        self.layout_success.addWidget(self.label_success_ipr)

        self.label_total_error = QLabel('Ошибка:')
        self.label_total_error.setStyleSheet(font_bold_11)
        self.label_error_ipt = QLabel('ИПТ-И: 0')
        self.label_error_ipt_s = QLabel('ИПТ-СИ: 0')
        self.label_error_mks = QLabel('МКС: 0')
        self.label_error_ipd = QLabel('ИПД: 0')
        self.label_error_ipr = QLabel('ИПР: 0')
        self.layout_error.addWidget(self.label_total_error)
        self.layout_error.addWidget(self.label_error_ipt)
        self.layout_error.addWidget(self.label_error_ipt_s)
        self.layout_error.addWidget(self.label_error_mks)
        self.layout_error.addWidget(self.label_error_ipd)
        self.layout_error.addWidget(self.label_error_ipr)

        self.layout_test_number_and_timer = QHBoxLayout()
        self.layout_test_number_and_timer.setContentsMargins(5, 0, 5, 5)
        self.label_current_test = QLabel('Выполняемый тест: ')
        self.label_time = QLabel('00:00:00')
        self.layout_test_number_and_timer.addWidget(self.label_current_test, alignment=Qt.AlignLeft)
        self.layout_test_number_and_timer.addWidget(self.label_time, alignment=Qt.AlignRight)
        self.layout_main.addLayout(self.layout_test_number_and_timer)

        self.time = QTime(0, 0, 0)
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_time)
        self.is_timer_started = False
        self.kl_status = 'd'

    def get_statistics_difference(self, new_statistics):
        differences = []
        differences.append(new_statistics.success_all - self.statistics.success_all)
        differences.append(new_statistics.error_all - self.statistics.error_all)
        differences.append(new_statistics.interrupted_all - self.statistics.interrupted_all)
        differences.append(new_statistics.success_ipt - self.statistics.success_ipt)
        differences.append(new_statistics.success_ipt_s - self.statistics.success_ipt_s)
        differences.append(new_statistics.success_mks - self.statistics.success_mks)
        differences.append(new_statistics.success_ipd - self.statistics.success_ipd)
        differences.append(new_statistics.success_ipr - self.statistics.success_ipr)
        differences.append(new_statistics.error_ipt - self.statistics.error_ipt)
        differences.append(new_statistics.error_ipt_s - self.statistics.error_ipt_s)
        differences.append(new_statistics.error_mks - self.statistics.error_mks)
        differences.append(new_statistics.error_ipd - self.statistics.error_ipd)
        differences.append(new_statistics.error_ipr - self.statistics.error_ipr)
        for el in differences:
            if el < 0:
                return False
        else:
            return differences

    def update_statistics(self):
        """Обновляет статистику в графе КЛ."""
        all_devices = self.statistics.success_all + self.statistics.interrupted_all + self.statistics.error_all
        self.label_total_all.setText(f'Всего: {all_devices}')
        self.label_success_all.setText(f'Успешно: {self.statistics.success_all}')
        self.label_error_all.setText(f'Ошибка: {self.statistics.error_all}')
        self.label_cancel_all.setText(f'Прервано: {self.statistics.interrupted_all}')

        self.label_success_ipt.setText(f'ИПТ-И: {self.statistics.success_ipt}')
        self.label_success_ipt_s.setText(f'ИПТ-CИ: {self.statistics.success_ipt_s}')
        self.label_success_mks.setText(f'МКС: {self.statistics.success_mks}')
        self.label_success_ipd.setText(f'ИПД: {self.statistics.success_ipd}')
        self.label_success_ipr.setText(f'ИПР: {self.statistics.success_ipr}')

        self.label_error_ipt.setText(f'ИПТ-И: {self.statistics.error_ipt}')
        self.label_error_ipt_s.setText(f'ИПТ-CИ: {self.statistics.error_ipt_s}')
        self.label_error_mks.setText(f'МКС: {self.statistics.error_mks}')
        self.label_error_ipd.setText(f'ИПД: {self.statistics.error_ipd}')
        self.label_error_ipr.setText(f'ИПР: {self.statistics.error_ipr}')

    def manage_timer(self, status):
        if status == 'd':
            if self.is_timer_started:
                self.timer.stop()
                self.is_timer_started = False
            self.time = QTime(0, 0, 0)
            self.label_time.setText(self.time.toString('HH:mm:ss'))
        else:
            if status != self.kl_status:
                self.time = QTime(0, 0, 0)
                if not self.is_timer_started:
                    self.timer.start(1000)
                    self.is_timer_started = True
        self.kl_status = status

    def update_time(self):
        """Устанавливает значения таймера на статус-панели."""
        self.time = self.time.addSecs(1)
        self.label_time.setText(self.time.toString('HH:mm:ss'))

    def update_color(self, head, body):
        self.setStyleSheet(f'background-color: {body}')
        self.name_version_widget.setStyleSheet(f'background-color: {head}')

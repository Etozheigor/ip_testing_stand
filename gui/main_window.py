import ctypes.wintypes
import logging
import os.path
import traceback
from calendar import monthrange
from datetime import datetime, timedelta

import win32.lib.win32con as win32con
import win32gui
from PyQt5 import QtCore
from PyQt5.QtCore import Qt, QThread, QTime, QTimer, pyqtSlot
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QBoxLayout, QFileDialog, QFrame, QHBoxLayout,
                             QLabel, QMainWindow, QMessageBox, QPushButton,
                             QSizePolicy, QSplitter, QTabWidget, QVBoxLayout,
                             QWidget)
from sqlalchemy import not_

from db.models import Event, Test
from db.utils import (archive_test_db, clear_events_db, get_count_of_tests_in_db,
                      clear_tests_db_for_period, combobox_remembered_text,
                      create_saved_statistic, get_count_of_events_in_db,
                      get_events_codes, get_latest_report, get_save_file_path,
                      get_session, is_statistic_saved, remember_save_file_path,
                      save_events_to_db, save_report_to_db, save_test_to_db)
from gui.constants import (JOURNAL, REPORT_FOR_PERIOD, TESTS,
                           TESTS_FOR_PERIOD)
from gui.dialog_windows import (AboutProgramDialog, SelectPortDialog,
                                SelectReportPeriodDialog,
                                SelectTestPeriodDialog)
from gui.event_codes import get_codes_dictionary, get_error_dictionary
from gui.styles import rx_stylesheet, rx_tx_white_stylesheet, tx_stylesheet
from gui.threads import ReadingWorker
from gui.utils import (add_rows_journal_table, add_rows_main_table,
                       write_journal_to_file, write_report_for_period_to_file,
                       write_report_to_file, write_tests_results_to_file)
from gui.widgets import (GraphWidget, JournalTable, LayoutFilterJournal,
                         LayoutFilterMainTable, LayoutJournalToolBar,
                         LayoutMainTableToolBar, MainTable, StatisticWidget)
from modbus.utils import get_ports


class MainWindow(QMainWindow):
    """Главное окно программы."""
    def nativeEvent(self, eventType, message):
        """Конфигурация кастомного обработчика событий эвентлупа."""
        retval, result = super(MainWindow, self).nativeEvent(eventType, message)
        # плучение структуры типа MSG
        msg = ctypes.wintypes.MSG.from_address(message.__int__())
        if msg.message == win32con.WM_SYSCOMMAND:
            if msg.wParam == self.custom_menu_id:
                dialog = AboutProgramDialog()
                dialog.show()
                dialog.exec()
        return retval, result

    def __init__(self):
        super().__init__()
        # конфигурация главного окна и центрального виджета
        self.central_widget = QWidget(self)
        self.setGeometry(50, 50, 1300, 700)
        self.setWindowTitle('Стенд проверки ИП')
        self.setCentralWidget(self.central_widget)
        self.layout_main_window = QVBoxLayout()
        self.central_widget.setLayout(self.layout_main_window)
        # новый пункт меню:
        self.custom_menu_id = 6
        wind_id = self.winId()
        system_menu = win32gui.GetSystemMenu(wind_id, False)
        win32gui.AppendMenu(system_menu, win32con.MF_SEPARATOR, 5, '')
        win32gui.AppendMenu(system_menu, win32con.MF_STRING, self.custom_menu_id, 'О программе')

        # конфигурация фреймов для сплиттера
        self.frame_top = QFrame()
        sizePolicy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        self.frame_top.setSizePolicy(sizePolicy)
        self.frame_bot = QFrame()
        # конфигурация главной рабочей зоны
        self.layout_main_work_zone = QHBoxLayout()
        self.frame_top.setLayout(self.layout_main_work_zone)
        # конфигурация окна статистики
        self.statistic_widget = StatisticWidget()
        self.statistic_widget.btn_start.clicked.connect(self.manage_reading)
        self.statistic_widget.btn_save_report.clicked.connect(self.select_report_period)
        self.statistic_widget.btn_save_tests.clicked.connect(self.select_test_period)
        self.layout_main_work_zone.addWidget(self.statistic_widget)

        # конфигурация главной таблицы/графа
        # конфигурация виджета таблицы с тулбаром
        self.table_zone_widget = QWidget()
        self.layout_table_zone = QVBoxLayout()
        self.table_zone_widget.setLayout(self.layout_table_zone)
        # конфигурация тулбара главной таблицы
        self.main_table_toolbar = LayoutMainTableToolBar(LayoutFilterMainTable())
        self.layout_table_zone.addLayout(self.main_table_toolbar)
        self.main_table_toolbar.btn_download_to_file.clicked.connect((lambda: self.save_table(TESTS)))
        self.main_table_toolbar.btn_filter.clicked.connect(self.filter_main_table)
        self.main_table_toolbar.btn_clean.clicked.connect(self.clear_main_table)
        self.main_table_toolbar.layout_filters.filter_kl.textChanged.connect(self.update_main_table_filter)
        self.main_table_toolbar.layout_filters.filter_exit_code_1.textChanged.connect(self.update_main_table_filter)
        self.main_table_toolbar.layout_filters.filter_exit_code_2.textChanged.connect(self.update_main_table_filter)
        self.main_table_toolbar.layout_filters.filter_exit_code_3.textChanged.connect(self.update_main_table_filter)
        self.main_table_toolbar.layout_filters.filter_device_type.currentIndexChanged.connect(self.update_main_table_filter)
        self.main_table_toolbar.layout_filters.checkbox_kl.stateChanged.connect(self.update_main_table_filter)
        self.main_table_toolbar.layout_filters.checkbox_exit_code.stateChanged.connect(self.update_main_table_filter)
        self.main_table_toolbar.layout_filters.checkbox_device_type.stateChanged.connect(self.update_main_table_filter)

        # конфигурация таблицы
        self.layout_table = QHBoxLayout()
        self.main_table = MainTable()
        self.layout_table.addWidget(self.main_table)
        self.layout_table_zone.addLayout(self.layout_table)
        # конфигурация графа
        self.graph_widget = GraphWidget()
        # конфигурация вкладок
        self.tab_graph_table = QTabWidget()
        self.layout_main_work_zone.addWidget(self.tab_graph_table)
        self.tab_graph_table.addTab(self.graph_widget, 'Таблица КЛ (граф)')
        self.tab_graph_table.addTab(self.table_zone_widget, 'Таблица УА')

        # конфигурация зоны журнала
        self.layout_journal_zone = QVBoxLayout()
        self.frame_bot.setLayout(self.layout_journal_zone)

        # конфигурация тулбара журнала
        self.journal_toolbar = LayoutJournalToolBar(LayoutFilterJournal())
        self.journal_toolbar.btn_clean.clicked.connect(self.clear_journal)
        self.journal_toolbar.btn_filter.clicked.connect(self.filter_journal_table)
        self.journal_toolbar.btn_download_to_file.clicked.connect(lambda: self.save_table(JOURNAL))
        self.journal_toolbar.layout_filters.filter_event_code.currentIndexChanged.connect(self.update_journal_table_filter)
        self.journal_toolbar.layout_filters.filter_kl.textChanged.connect(self.update_journal_table_filter)
        self.journal_toolbar.layout_filters.checkbox_event_code.stateChanged.connect(self.update_journal_table_filter)
        self.journal_toolbar.layout_filters.checkbox_kl.stateChanged.connect(self.update_journal_table_filter)

        self.layout_journal_zone.addLayout(self.journal_toolbar)
        # конфигурация таблицы журнала
        self.layout_journal_table = QHBoxLayout()
        self.table_journal = JournalTable()
        self.table_journal.setSortingEnabled(True)
        self.layout_journal_table.addWidget(self.table_journal)
        self.layout_journal_zone.addLayout(self.layout_journal_table)

        # конфигурация сплиттера
        self.splitter = QSplitter(Qt.Vertical)
        self.splitter.addWidget(self.frame_top)
        self.splitter.addWidget(self.frame_bot)
        self.layout_main_window.addWidget(self.splitter)

        # конфигурация статус панели
        self.layout_status_bar = QHBoxLayout()
        self.layout_status_bar.addStretch(0)
        self.layout_status_bar.setDirection(QBoxLayout.RightToLeft)
        ports_list = get_ports()
        self.btn_select_port = QPushButton('Выбрать порт')
        self.btn_select_port.clicked.connect(self.select_port)
        self.label_current_port = QLabel()
        if combobox_remembered_text() in ports_list:
            self.label_current_port.setText(f': {combobox_remembered_text()}')
        else:
            self.label_current_port.setText(': Не выбран')

        self.label_tx = QLabel('Tx: ')  # конфигурация Rx Tx:
        self.tx = QLabel()
        self.tx.setStyleSheet(rx_tx_white_stylesheet)
        self.label_rx = QLabel('Rx: ')  # конфигурация Rx Tx:
        self.rx = QLabel()
        self.rx.setStyleSheet(rx_tx_white_stylesheet)

        self.label_trans_packets = QLabel('Передано: 0  ')
        self.label_rec_packets = QLabel('Принято: 0  ')
        self.label_session_time = QLabel('00:00:00')

        self.layout_status_bar.addWidget(self.label_session_time)
        self.layout_status_bar.addSpacing(30)
        self.layout_status_bar.addWidget(self.label_rec_packets)
        self.layout_status_bar.addSpacing(30)
        self.layout_status_bar.addWidget(self.label_trans_packets)
        self.layout_status_bar.addSpacing(30)
        self.layout_status_bar.addWidget(self.rx)
        self.layout_status_bar.addWidget(self.label_rx)
        self.layout_status_bar.addSpacing(30)
        self.layout_status_bar.addWidget(self.tx)
        self.layout_status_bar.addWidget(self.label_tx)
        self.layout_status_bar.addSpacing(30)
        self.layout_status_bar.addWidget(self.label_current_port)
        self.layout_status_bar.addWidget(self.btn_select_port)
        self.layout_main_window.addLayout(self.layout_status_bar)

        # переменные
        self.rx_color, self.tx_color = 'w', 'w'  # w - белый, g - зеленый, r - красный
        self.total_trans_packets_count = 0
        self.total_rec_packets_count = 0
        self.timer = QTimer()
        self.timer.timeout.connect(self.set_statusbar_time)
        self.statusbar_time = QTime(0, 0, 0)
        self.is_timer_started = False
        self.is_reading = False
        self.kl_graph_addresses_dict = {
            1: self.graph_widget.graph_kl_1, 2: self.graph_widget.graph_kl_2,
            3: self.graph_widget.graph_kl_3, 4: self.graph_widget.graph_kl_4,
            5: self.graph_widget.graph_kl_5, 6: self.graph_widget.graph_kl_6,
            7: self.graph_widget.graph_kl_7, 8: self.graph_widget.graph_kl_8,
            9: self.graph_widget.graph_kl_9, 10: self.graph_widget.graph_kl_10,
            11: self.graph_widget.graph_kl_11, 12: self.graph_widget.graph_kl_12,
            13: self.graph_widget.graph_kl_13, 14: self.graph_widget.graph_kl_14,
            15: self.graph_widget.graph_kl_15, 16: self.graph_widget.graph_kl_16,
        }

        self.check_messages_files()  # проверяем файл с кодами сообщений на корректность
        self.check_errors_files()  # проверяем файл с описаниеями ошибок на корректность
        self.check_latest_report_saved()
        self.save_month_statistic(REPORT_FOR_PERIOD)
        self.save_month_statistic(TESTS_FOR_PERIOD)
        clear_events_db()  # очищаем базу даннных от событий
        archive_test_db()  # архивируем старые тесты в БД
        self.start_reading()

    def closeEvent(self, event):
        """Выполняет вызов функции сохранения отчета перед закрытием главного окна."""
        msg = QMessageBox(self)
        msg.setWindowTitle("Выход")
        msg.setIcon(QMessageBox.Question)
        msg.setText("Вы действительно хотите завершить работу программы?")
        button_yes = msg.addButton("Да", QMessageBox.YesRole)
        button_no = msg.addButton("Нет", QMessageBox.RejectRole)
        msg.setDefaultButton(button_yes)
        msg.exec_()
        if msg.clickedButton() == button_yes:
            try:
                self.save_statistics_report()
            except Exception:
                event.ignore()
            event.accept()
        elif msg.clickedButton() == button_no:
            event.ignore()

    def select_port(self):
        """Выбор порта."""
        if self.is_reading:
            logging.warning('Ошибка. Попытка выбора порта во время чтения.')
            QMessageBox.warning(self, 'Ошибка', 'Остановите чтение перед тем, как изменить порт.')
            return
        modal_dialog = SelectPortDialog()
        modal_dialog.show()
        modal_dialog.exec_()
        selected_port = modal_dialog.selected_port
        self.label_current_port.setText(f': {selected_port}')
        logging.info(f'Подключение к порту {selected_port}')

    def manage_reading(self):
        if self.is_reading:
            self.stop_reading()
        else:
            self.start_reading()

    def start_reading(self):
        """Начало чтения всех КЛ."""
        if self.label_current_port.text()[2:] == 'Не выбран':
            logging.warning('Не выбран порт при старте чтения')
            QMessageBox.warning(self, 'Выберите порт', 'Необходимо выбрать порт перед стартом чтения')
            return
        logging.info('Старт чтения данных из стенда')
        self.is_reading = True
        self.statistic_widget.btn_start.setText('Стоп')

        self.reading_thread = QThread(parent=self)
        self.reading_widget = ReadingWorker(port=self.label_current_port.text()[2:], main=self.main_table, journal=self.table_journal)
        self.reading_widget.moveToThread(self.reading_thread)
        self.reading_widget.data_trans_rec.connect(self.set_rec_trans_data)
        self.reading_widget.tx_rx.connect(self.get_rx_tx_blink)
        self.reading_widget.manage_table_sortig.connect(self.manage_table_sorting)
        self.reading_widget.update_kl_graph_color.connect(self.update_kl_graph_color)
        self.reading_widget.update_timer.connect(self.update_kl_graph_timer)
        self.reading_widget.update_kl_tests_results.connect(self.update_kl_statistic)
        self.reading_widget.update_kl_version.connect(self.update_kl_version)
        self.reading_widget.update_current_test.connect(self.update_kl_current_test)
        self.reading_widget.save_tests.connect(self.save_tests)
        self.reading_widget.save_events.connect(self.save_events)
        self.reading_widget.port_connection_error.connect(self.reading_error)
        self.reading_thread.started.connect(self.reading_widget.run)

        if not self.is_timer_started:  # запускаем таймер если он еще не запущен
            self.timer.start(1000)
            self.is_timer_started = True
        self.reading_thread.start()

    def stop_reading(self):
        self.is_reading = False
        self.statistic_widget.btn_start.setText('Старт')
        self.reading_widget.is_cancelled = True
        self.reading_thread.quit()

    @pyqtSlot(bool)
    def reading_error(self, is_error):
        self.stop_reading()
        self.label_current_port.setText(': Не выбран')
        logging.warning('Ошибка. Потеряно соединение с портом.')
        QMessageBox.warning(self, 'Ошибка  в работе программы "Стенд проверки ИП"', 'Потеряно соединение с портом. Проверьте подключение и повторно начните чтение.')

    def fill_journal_table(self, events, is_filtering=False):
        """Заполняет таблицу журнала данными."""
        self.table_journal.setSortingEnabled(False)
        add_rows_journal_table(self.table_journal, events)
        if not is_filtering:
            self.update_event_code_filter()
        self.table_journal.setSortingEnabled(True)

    def fill_main_table(self, test):
        """Заполняет главную таблицу данными."""
        self.main_table.setSortingEnabled(False)
        add_rows_main_table(self.main_table, test)
        self.main_table.setSortingEnabled(True)

    @pyqtSlot(object)
    def save_tests(self, test):
        saved_test = save_test_to_db(test)
        self.fill_main_table([saved_test])

    @pyqtSlot(object)
    def save_events(self, events):
        if self.table_journal.rowCount() > 10000:
            self.clear_journal()
        saved_events = save_events_to_db(events)
        self.fill_journal_table(saved_events)

    @QtCore.pyqtSlot(int, str, str)
    def update_kl_graph_color(self, address, head, body):
        """Обновляет цвет графа КЛ."""
        graph = self.kl_graph_addresses_dict[address]
        graph.update_color(head, body)

    @QtCore.pyqtSlot(int, str)
    def update_kl_graph_timer(self, address, status):
        """Запускает или останавливает таймер в графе КЛ в зависимости от состояния КЛ."""
        kl = self.kl_graph_addresses_dict[address]
        kl.manage_timer(status)

    @QtCore.pyqtSlot(int, object)
    def update_kl_statistic(self, address, statistics):
        """Обновляет статистику графа КЛ."""
        kl = self.kl_graph_addresses_dict[address]
        statistics_difference = kl.get_statistics_difference(statistics)
        kl.statistics = statistics
        kl.update_statistics()
        if statistics_difference:
            self.update_total_kl_statistics(statistics_difference)

    def update_total_kl_statistics(self, statistics_difference):
        """Обновляет общую статистику по всем КЛ."""
        self.statistic_widget.update_statistics(statistics_difference)
        save_report_to_db(self.statistic_widget.statistics)

    @QtCore.pyqtSlot(int, int)
    def update_kl_version(self, address, version):
        """Обновляет версию КЛ."""
        kl = self.kl_graph_addresses_dict[address]
        kl.label_version.setText(f'версия: {version}')

    @QtCore.pyqtSlot(int, int)
    def update_kl_current_test(self, address, current_test):
        """Обновляет текущий тест КЛ."""
        kl = self.kl_graph_addresses_dict[address]
        kl.label_current_test.setText(f'Выполняемый тест: {current_test}')

    def set_statusbar_time(self):
        """Устанавливает значения таймера на статус-панели."""
        self.statusbar_time = self.statusbar_time.addSecs(1)
        self.label_session_time.setText(self.statusbar_time.toString('HH:mm:ss'))

    @QtCore.pyqtSlot(int, int)
    def set_rec_trans_data(self, trans, rec):
        """Обновляет количество переданных и полученных пакетов."""
        self.total_trans_packets_count += trans
        self.total_rec_packets_count += rec
        self.label_trans_packets.setText(f'Передано:  {self.total_trans_packets_count}')
        self.label_rec_packets.setText(f'Получено:  {self.total_rec_packets_count}')

    @QtCore.pyqtSlot(str)
    def get_rx_tx_blink(self, text):
        """Принимает сигнал Rx и Tx."""
        if text == 'Rx':
            if self.rx_color == 'w':
                self.rx_tx_blink('Rx_blink')
                QtCore.QTimer.singleShot(
                    500, lambda: self.rx_tx_blink('Rx_white'))
        elif text == 'Tx':
            if self.tx_color == 'w':
                self.rx_tx_blink('Tx_blink')
                QtCore.QTimer.singleShot(
                    500, lambda: self.rx_tx_blink('Tx_white'))

    def rx_tx_blink(self, text):
        """Запускает мигание индикаторов Rx и Tx."""
        if text == 'Rx_blink':
            self.rx.setStyleSheet(rx_stylesheet)
            self.rx_color = 'g'
        elif text == 'Tx_blink':
            self.tx.setStyleSheet(tx_stylesheet)
            self.tx_color = 'r'
        elif text == 'Rx_white':
            self.rx.setStyleSheet(rx_tx_white_stylesheet)
            self.rx_color = 'w'
        elif text == 'Tx_white':
            self.tx.setStyleSheet(rx_tx_white_stylesheet)
            self.tx_color = 'w'

    def clear_main_table(self):
        """Очищает главную таблицу и архивирует все тесты."""
        logging.info('Архивирование всех тестов из таблицы')
        archive_test_db()
        self.main_table.setRowCount(0)

    def clear_journal(self):
        """Очищает таблицу журанала и удаляет все события из БД."""
        logging.info('Удаление всех событий журнала из БД')
        clear_events_db()
        self.table_journal.setRowCount(0)

    @pyqtSlot(str, bool)
    def manage_table_sorting(self, table, is_enabled):
        """Включает/отключает сортировку."""
        if table == 'Journal':
            self.table_journal.setSortingEnabled(is_enabled)
        elif table == 'Test':
            self.main_table.setSortingEnabled(is_enabled)

    def update_event_code_filter(self):
        """Обновляет комбобокс с кодами событий."""
        self.journal_toolbar.layout_filters.filter_event_code.clear()  # обновляем данные в комбобоксе фильтрации кода события
        self.journal_toolbar.layout_filters.filter_event_code.addItems(get_events_codes())

    def select_report_period(self, table):
        modal_dialog = SelectReportPeriodDialog(self)
        modal_dialog.show()
        modal_dialog.exec_()

    def select_test_period(self, table):
        modal_dialog = SelectTestPeriodDialog(self)
        modal_dialog.show()
        modal_dialog.exec_()

    def save_table(self, table, period=None, is_error_flag=False):
        """Сохраняет события из базы данных в файл."""
        logging.info(f'Старт экспорта таблицы "{table}" в файл.')
        saved_file_path = get_save_file_path(table)
        if saved_file_path:
            filename, _ = QFileDialog.getSaveFileName(self, 'Выбор папки', f'{saved_file_path}', 'Excel files (*.xlsx)')  # 'Text files (*.txt)'
        else:
            filename, _ = QFileDialog.getSaveFileName(self, 'Выбор папки', '/', 'Excel files (*.xlsx)')
        if filename:
            try:
                if table == JOURNAL:
                    write_journal_to_file(filename)
                elif table == TESTS:
                    write_tests_results_to_file(filename)
                elif table == TESTS_FOR_PERIOD:
                    write_tests_results_to_file(filename, period, is_error_flag)
                elif table == REPORT_FOR_PERIOD:
                    write_report_for_period_to_file(filename, period)
                remember_save_file_path(filename, table)
                logging.info(f'Файл по адресу {filename} успешно создан')
                QMessageBox.information(self, 'Успешно', f'Файл по адресу {filename} успешно создан')
            except Exception as error:
                logging.error(f'Ошибка при создании файла по адресу {filename}.\n{traceback.format_exc()}')
                QMessageBox.information(self, 'Успешно', f'Ошибка при создании файла по адресу {filename}. Поверьте, не открыт ли этот файл другой программой.')

    def save_statistics_report(self, report=None):
        """Сохраняет отчет за день  в файл."""
        logging.info('Старт сохранения отчета за день.')
        if not report:
            report = self.statistic_widget.statistics
        try:
            filename = f"reports/{report.date_time.strftime('%d-%m-%Y')}.xlsx"
            write_report_to_file(filename, report)
            logging.info(f'Файл {filename} успешно создан')
        except Exception as error:
            logging.error(f'Ошибка при создании файла отчета.\n{traceback.format_exc()}')
            QMessageBox.information(self, 'Ошибка', f'Ошибка при создании отчета по адресу {filename}. Поверьте, не открыт ли этот файл другой программой. Ошибка: {error}')
            raise

    def filter_journal_table(self):
        """Фильтрует данные в таблице."""
        logging.info('Старт фильтрации журнала')
        self.table_journal.setSortingEnabled(False)
        session = get_session()
        result = session.query(Event)
        if (
            self.journal_toolbar.layout_filters.checkbox_kl.isChecked()
            and self.journal_toolbar.layout_filters.filter_kl.text() != ''
        ):
            result = result.filter(Event.kl == int(self.journal_toolbar.layout_filters.filter_kl.text()))

        if (
            self.journal_toolbar.layout_filters.checkbox_event_code.isChecked()
            and self.journal_toolbar.layout_filters.filter_event_code.count() > 0
        ):
            # разделяем строку "код события - событие" и берем только код
            code = int(self.journal_toolbar.layout_filters.filter_event_code.currentText().split(' ', 1)[0])
            result = result.filter(Event.event_code == code)

        self.table_journal.setRowCount(0)
        self.fill_journal_table(result.all(), is_filtering=True)
        self.journal_toolbar.btn_filter.setEnabled(False)
        self.table_journal.setSortingEnabled(True)
        logging.info('Фильтры журнала успешно применены')

    def filter_main_table(self):
        """Фильтрует данные в главной таблице."""
        logging.info('Старт фильтрации тестов')
        self.table_journal.setSortingEnabled(False)
        session = get_session()
        result = session.query(Test).filter(not_(Test.is_archived))

        if (
            self.main_table_toolbar.layout_filters.checkbox_kl.isChecked()
            and self.main_table_toolbar.layout_filters.filter_kl.text() != ''
        ):
            result = result.filter(Test.kl == int(self.main_table_toolbar.layout_filters.filter_kl.text()))

        if self.main_table_toolbar.layout_filters.checkbox_exit_code.isChecked():
            if self.main_table_toolbar.layout_filters.filter_exit_code_1.text() != '':
                result = result.filter(Test.exit_code_1 == int(self.main_table_toolbar.layout_filters.filter_exit_code_1.text()))
            if self.main_table_toolbar.layout_filters.filter_exit_code_2.text() != '':
                result = result.filter(Test.exit_code_2 == int(self.main_table_toolbar.layout_filters.filter_exit_code_2.text()))
            if self.main_table_toolbar.layout_filters.filter_exit_code_3.text() != '':
                result = result.filter(Test.exit_code_3 == int(self.main_table_toolbar.layout_filters.filter_exit_code_3.text()))

        if (
            self.main_table_toolbar.layout_filters.checkbox_device_type.isChecked()
            and self.main_table_toolbar.layout_filters.filter_device_type.currentText() != 'Не выбрано'
        ):
            result = result.filter(Test.device_type == self.main_table_toolbar.layout_filters.filter_device_type.currentText())

        self.main_table.setRowCount(0)
        self.fill_main_table(result.all())
        self.main_table_toolbar.btn_filter.setEnabled(False)
        self.main_table.setSortingEnabled(True)
        logging.info('Фильтры тестов успешно применены')

    def update_journal_table_filter(self):
        """Обновляет кнопку фильтрации журнала."""
        if get_count_of_events_in_db() > 0:
            self.journal_toolbar.btn_filter.setEnabled(True)

    def update_main_table_filter(self):
        """Обновляет кнопку фильтрации главной таблицы."""
        if get_count_of_tests_in_db() > 0:
            self.main_table_toolbar.btn_filter.setEnabled(True)

    def check_messages_files(self):
        """Проверяет на корректность файлы с кодами сообщений."""
        try:
            codes_dictionary = get_codes_dictionary()
            return codes_dictionary
        except Exception as error:
            logging.error(
                f'{error.message}. \n Устраните ошибку и перезапустите программу.')
            QMessageBox.information(
                self, 'Ошибка', f'{error.message}. \n Устраните ошибку и перезапустите программу.')
            self.close()

    def check_errors_files(self):
        """Проверяет на корректность файлы с кодами сообщений."""
        try:
            errors_dictionary = get_error_dictionary()
            return errors_dictionary
        except Exception as error:
            logging.error(
                f'{error.message}. \n Устраните ошибку и перезапустите программу.')
            QMessageBox.information(
                self, 'Ошибка', f'{error.message}. \n Устраните ошибку и перезапустите программу.')
            self.close()

    def check_latest_report_saved(self):
        """Проверяет, сохранен ли последний отчет в эксель файл."""
        report = get_latest_report()
        if not report:
            return
        date = report.date_time.strftime('%d-%m-%Y')
        if not os.path.isfile(f'reports/{date}.xlsx'):
            self.save_statistics_report(report)

    def save_month_statistic(self, document):
        current_month = datetime.now().month
        current_year = datetime.now().year
        if current_month == 1:
            prev_month = 12
            prev_year = current_year - 1
        else:
            prev_month = current_month - 1
            prev_year = current_year

        if len(str(prev_month)) == 1:
            str_prev_month = '0' + str(prev_month)
        else:
            str_prev_month = str(prev_month)

        if document == REPORT_FOR_PERIOD:
            filename = f'reports/Отчет {str_prev_month}-{prev_year}.xlsx'
        elif document == TESTS_FOR_PERIOD:
            filename = f'tests/Результаты тестов {str_prev_month}-{prev_year}.xlsx'

        if not is_statistic_saved(document, prev_month, prev_year):
            date_from = datetime(year=prev_year, month=prev_month, day=1)
            date_to = datetime(year=prev_year, month=prev_month, day=monthrange(prev_year, prev_month)[1])
            date_to += timedelta(days=1)
            period =(date_from, date_to)
            logging.info('Старт сохранения отчета за месяц.')
            try:
                if document == REPORT_FOR_PERIOD:
                    write_report_for_period_to_file(filename, period)
                elif document == TESTS_FOR_PERIOD:
                    write_tests_results_to_file(filename, period)
                logging.info(f'Файл {filename} успешно создан')
                if os.path.isfile(filename):
                    create_saved_statistic(document, prev_month, prev_year)
                    if document == TESTS_FOR_PERIOD:
                        clear_tests_db_for_period(period)
            except Exception as error:
                logging.error(f'Ошибка при создании файла отчета за месяц.\n{traceback.format_exc()}')
                QMessageBox.information(self, 'Ошибка', f'Ошибка при создании отчета по адресу {filename}. Поверьте, не открыт ли этот файл другой программой.')
                raise

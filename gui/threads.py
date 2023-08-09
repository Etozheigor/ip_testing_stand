import logging
import traceback
from time import sleep

from PyQt5.QtCore import QObject, pyqtSignal

from modbus.utils import (get_active_devices, get_client,
                          get_count_of_log_entries, get_events_from_registers,
                          get_kl_version, get_registers_values,
                          get_test_results, is_new_test_results, is_port_open,
                          registers_values_to_bin)


class ReadingWorker(QObject):
    """Воркер для потока чтения журнала и результатов тестов из всех КЛ."""
    data_trans_rec = pyqtSignal(int, int)
    tx_rx = pyqtSignal(str)
    manage_table_sortig = pyqtSignal(str, bool)
    update_kl_graph_color = pyqtSignal(int, str, str)
    update_kl_tests_results = pyqtSignal(int, object)
    update_kl_version = pyqtSignal(int, int)
    update_timer = pyqtSignal(int, str)  # p - pause, w - working, d- disconnected
    update_current_test = pyqtSignal(int, int)
    save_events = pyqtSignal(object)
    save_tests = pyqtSignal(object)
    exit_thread_after_cancel = pyqtSignal(bool)
    port_connection_error = pyqtSignal(bool)

    def __init__(self, port, main, journal):
        super().__init__()
        self.is_cancelled = False
        self.port = port
        self.active_devices = []
        self.main = main
        self.journal = journal

    def check_port(self):
        i = 0
        while i < 10:
            self.tx_rx.emit('Tx')
            self.data_trans_rec.emit(1, 0)
            if not is_port_open(self.port):
                sleep(1)
            else:
                self.tx_rx.emit('Rx')
                self.data_trans_rec.emit(0, 1)
                return
            i += 1
        self.port_connection_error.emit(True)
        self.is_cancelled = True

    def get_journal(self, address, client):
        """Получает журнал событий КЛ."""
        self.tx_rx.emit('Tx')
        self.data_trans_rec.emit(1, 0)
        count_of_log_entries = get_count_of_log_entries(self.port, client, address)[0]
        self.tx_rx.emit('Rx')
        self.data_trans_rec.emit(0, 1)

        while count_of_log_entries != 0:
            self.tx_rx.emit('Tx')
            self.data_trans_rec.emit(1, 0)
            registers_bin = registers_values_to_bin(
                get_registers_values(client, address, min(count_of_log_entries*2, 124)))
            self.tx_rx.emit('Rx')
            self.data_trans_rec.emit(0, 1)

            events = get_events_from_registers(registers_bin, address)
            self.save_events.emit(events)

            self.tx_rx.emit('Tx')
            self.data_trans_rec.emit(1, 0)
            count_of_log_entries = get_count_of_log_entries(self.port, client, address)[0]
            self.tx_rx.emit('Rx')
            self.data_trans_rec.emit(0, 1)

    def get_tests_results(self, address, client):
        """Получает результаты теста КЛ."""
        self.tx_rx.emit('Tx')
        self.data_trans_rec.emit(1, 0)
        kl_version = get_kl_version(client, address)
        self.tx_rx.emit('Rx')
        self.data_trans_rec.emit(0, 1)
        self.update_kl_version.emit(address, kl_version)

        self.tx_rx.emit('Tx')
        self.data_trans_rec.emit(1, 0)
        status_code, is_new_tests = is_new_test_results(client, address)
        self.tx_rx.emit('Rx')
        self.data_trans_rec.emit(0, 1)
        
        if status_code == 0:
            head, body = 'lightGray', 'lightGray'
        if status_code == 1:
            head, body = 'lightGreen', 'lightGreen'
        elif status_code == 2:
            head, body = 'yellow', 'yellow'
        elif status_code == 3:
            head, body = 'lightGreen', 'yellow'
        elif status_code == 4:
            head, body = 'red', 'yellow'
        elif status_code == 5:
            head, body = 'purple', 'yellow'
        self.update_kl_graph_color.emit(address, head, body)
        if status_code == 1:
            self.update_timer.emit(address, 'w')
        elif status_code in [2, 3, 4, 5]:
            self.update_timer.emit(address, 'p')
        else:
            self.update_timer.emit(address, 'd')

        self.tx_rx.emit('Tx')
        self.data_trans_rec.emit(1, 0)
        test, total_statistic, current_test = get_test_results(client, address)
        self.data_trans_rec.emit(0, 1)
        self.tx_rx.emit('Rx')
        self.update_current_test.emit(address, current_test)
        self.update_kl_tests_results.emit(address, total_statistic)

        if is_new_tests:
            self.save_tests.emit(test)

    def read(self, address, client):
        """Начинает чтение данных из одного КЛ."""
        try:
            self.get_journal(address, client)
        except Exception as error:
            logging.error(
                f'Ошибка во время чтения журнала КЛ по адресу {address}. Код ошибки: {error}\n{traceback.format_exc()}')
            return
        try:
            self.get_tests_results(address, client)
        except Exception as error:
            logging.error(
                f'Ошибка во время чтения результатов тестов КЛ по адресу {address}. Код ошибки: {error}\n{traceback.format_exc()}')
            return

    def update_active_devices(self, client):
        """Обновляет список подключенных КЛ."""
        updated_active_devices = get_active_devices(client, self.tx_rx, self.data_trans_rec)
        new_devices = []

        for device in self.active_devices:  # Меняем цвет КЛ, если они были в прошлом списке, но нет в обновленном
            if device not in updated_active_devices:
                self.update_kl_graph_color.emit(device, 'lightGray', 'lightGray')
                self.update_timer.emit(device, 'd')

        for device in updated_active_devices:  # Обновляем список активных КЛ
            if device not in self.active_devices:
                new_devices.append(device)

        self.active_devices = updated_active_devices
        return new_devices

    def run(self):
        """Статует работу воркера в потоке."""
        self.tx_rx.emit('Tx')
        self.data_trans_rec.emit(1, 0)
        client = get_client(self.port)
        self.tx_rx.emit('Rx')
        self.data_trans_rec.emit(0, 1)

        self.active_devices.extend(get_active_devices(client, self.tx_rx, self.data_trans_rec))

        while True and not self.is_cancelled:
            for address in self.active_devices:
                self.read(address, client)
            new_devices = self.update_active_devices(client)
            if new_devices:
                for address in new_devices:
                    self.read(address, client)
            self.check_port()
        if self.is_cancelled:
            for address in self.active_devices:
                self.update_kl_graph_color.emit(address, 'lightGray', 'lightGray')
                self.update_timer.emit(address, 'd')
            client.close()
            return

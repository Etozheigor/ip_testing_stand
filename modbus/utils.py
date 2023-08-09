from datetime import datetime

import serial.tools.list_ports
from pymodbus.client.sync import ModbusSerialClient as ModbusClient
from serial import Serial

from db.models import Event, Report, Test
from gui.utils import codes_dictionary, errors_dictionary


def port_is_usable(port):
    """Определяет свободен ли порт."""
    try:
        Serial(port)
        return True
    except Exception:
        return False


def is_port_open(port):
    """Определяет открыт ли порт."""
    ports = [port.name for port in serial.tools.list_ports.comports()]
    try:
        ser = Serial(port=port)
        is_open = ser.isOpen()
    except Exception:
        is_open = False
    if port not in ports and not is_open:
        return False
    return True


def get_ports():
    """Возвращет список портов."""
    ports_list = ['Не выбран']
    for port in serial.tools.list_ports.comports():
        if port_is_usable(port.name):
            ports_list.append(port.name)
        else:
            ports_list.append(f'{port.name} (занят)')
    return ports_list


def get_client(port):
    """Возвращает клиент подключения к COM-порту по modbus."""
    client = ModbusClient(
        method='rtu', port=port, stopbits=2, bytesize=8, parity='N', baudrate=19200, timeout=0.1)
    return client


def get_count_of_log_entries(port, client, address):
    """Возвращает количество непрочитанных записей.
    При повторном обращении КЛ считает, что прошлый
    пакет даных успешно передан."""
    response = client.read_holding_registers(address=0, count=1, unit=address)
    count_of_log_entries = response.registers
    return count_of_log_entries


def get_registers_values(client, address, registers_count):
    """Возвращает значения регистров начиная с 1 в десятичном формате."""
    response = client.read_holding_registers(address=1, count=registers_count, unit=address)
    registers_values = response.registers
    return registers_values


def registers_values_to_bin(registers_values):
    """Принимает список из значений регистров в десятичном виде.
    Возвращает список из байтов регистров в двоичном виде.
    """
    bin_value_list = []
    for value in registers_values:
        value = str(bin(value))[2:]
        if len(value) < 16:
            value = '0' * (16-len(value)) + value
        bin_value_list.append(value[:8])
        bin_value_list.append(value[8:])
    return bin_value_list


def bin_str_to_dec(bin_str):
    """Преобразует строку из нулей и единиц в десятичное число."""
    sum = 0
    for i in range(0, len(bin_str)):
        sum += int(bin_str[i]) * (2**(len(bin_str) - 1 - i))
    return sum


def get_active_devices(client, tx_rx, data_trans_rec):
    """Возвращает список поключенных КЛ с адресами 1-16."""
    active_adresses = []
    for address in range(1, 17):
        try:
            tx_rx.emit('Tx')
            data_trans_rec.emit(1, 0)
            response = client.read_holding_registers(address=976, count=1, unit=address)
            if response.registers[0] == 101:
                active_adresses.append(address)
                tx_rx.emit('Rx')
                data_trans_rec.emit(0, 1)
        except Exception as error:
            continue
    return active_adresses


def get_kl_version(client, address):
    """Возвращает версию КЛ."""
    response = client.read_holding_registers(address=977, count=1, unit=address)
    version = response.registers[0]
    return version


def get_event_params(event_bin, kl_address):
    """Преобразует двоичную запись события из журнала в его параметры в десятичном виде."""
    date_time = datetime.now()
    kl = kl_address
    entry_number = bin_str_to_dec(event_bin[0])
    au_address = bin_str_to_dec(event_bin[1])
    event_code = bin_str_to_dec(event_bin[3])
    addit_param = bin_str_to_dec(event_bin[2])
    return [date_time, kl, entry_number, au_address, event_code, addit_param]


def get_events_from_registers(registers, kl_address):
    """Преобразует значения из регистров в события Event и записывает в БД."""
    events_params = []
    for i in range(0, len(registers), 4):
        event_bin = registers[i:i+4]
        params = get_event_params(event_bin, kl_address)
        events_params.append(params)
    events = []
    for event_params in events_params:
        date_time, kl, entry_number, au_address, event_code, addit_param = event_params
        event = Event(
            number=1,
            date_time=date_time,
            kl=kl,
            entry_number=entry_number,
            au_address=au_address,
            event_code=event_code,
            addit_param=addit_param,
            description=codes_dictionary[event_code][0])
        events.append(event)
    return events


def is_new_test_results(client, address):
    """Запрашивает наличие новых резултатов тестов в КЛ.
    """
    response = client.read_holding_registers(address=20000, count=1, unit=address)
    response_bin = registers_values_to_bin(response.registers)

    status_code = bin_str_to_dec(response_bin[1])
    is_new_results = bin_str_to_dec(response_bin[0])

    return status_code, is_new_results


device_type_dict = {
    0: '',
    1: 'ИПР',
    2: 'ИПТ',
    3: 'ИПТ-С',
    4: 'ИПД',
    5: 'ИПП',
    6: 'ИПП-С',
    7: 'УДП',
    8: 'МКВ-2',
    9: 'МКВ-4',
    10: 'МКР-2',
    11: 'МКР-4',
    12: 'ИПР-В',
    13: 'МКВ2Р2',
    14: 'МКВ2А',
    15: 'МКО-В',
    16: 'МКО-С',
    17: 'МКС'
}


def get_test_results(client, address):
    """Возвращает результаты проведения теста из КЛ."""
    response = client.read_holding_registers(address=20001, count=40, unit=address)
    registers_values = response.registers
    successfull_checks_devices_count = registers_values[0]

    manually_interrupted_tests_count = registers_values[1]

    tests_completed_with_error_count = registers_values[2]

    exit_codes_bin = registers_values_to_bin([registers_values[3], registers_values[4]])
    exit_code_1 = bin_str_to_dec(exit_codes_bin[1])
    exit_code_2 = bin_str_to_dec(exit_codes_bin[0])
    exit_code_3 = bin_str_to_dec(exit_codes_bin[3])

    ready_time = registers_values[5]

    device_type_version_bin = registers_values_to_bin([registers_values[6]])
    device_type = device_type_dict[bin_str_to_dec(device_type_version_bin[1])]
    device_version = bin_str_to_dec(device_type_version_bin[0])

    sb_duration = registers_values[7]

    electricity_sb = registers_values[8]

    average_electricity = registers_values[9] / 10

    temperature = registers_values[10]

    Uv = registers_values[11] / 100
    Us = registers_values[12] / 100
    Un = registers_values[13] / 100

    Cv = registers_values[14] / 100
    Cs = registers_values[15] / 100
    Cn = registers_values[16] / 100

    charge_electricity_current = registers_values[17] / 10

    discharge_rate_upper_ionistor = registers_values[18] / 100
    discharge_rate_middle_ionistor = registers_values[19] / 100
    discharge_rate_lower_ionistor = registers_values[20] / 100

    success_count_ipr = registers_values[21]
    success_count_ipt = registers_values[22]
    success_count_ipt_s = registers_values[23]
    success_count_ipd = registers_values[24]
    success_count_ipp = registers_values[25]
    success_count_ipp_s = registers_values[26]
    success_count_udp = registers_values[27]
    success_count_mks = registers_values[28]

    error_count_ipr = registers_values[29]
    error_count_ipt = registers_values[30]
    error_count_ipt_s = registers_values[31]
    error_count_ipd = registers_values[32]
    error_count_ipp = registers_values[33]
    error_count_ipp_s = registers_values[34]
    error_count_udp = registers_values[35]
    error_count_mks = registers_values[36]

    current_test = registers_values[37]
    reserve_1 = registers_values[38]
    reserve_2 = registers_values[39]

    # если код 1.2 == 7.3 или 7.4 или 9.2 или 9.3
    # тогда к значению кода 3 нужно дописать нули, если его длина меньше 3 символов
    str_code_1_2 = f'{exit_code_1}{exit_code_2}'
    if str_code_1_2 in ('73', '74', '92', '93'):
        str_code_3 = '0' * (3 - len(str(exit_code_3))) + str(exit_code_3)
        full_exit_code = f'{exit_code_1}.{exit_code_2}.{str_code_3}'
    else:
        full_exit_code = f'{exit_code_1}.{exit_code_2}.{exit_code_3}'
    exit_code_1_2 = f'{exit_code_1}.{exit_code_2}.?'

    if errors_dictionary.get(full_exit_code):
        description_1, description_2 = errors_dictionary[full_exit_code]
    elif errors_dictionary.get(exit_code_1_2):
        description_1 = errors_dictionary[exit_code_1_2][0]
        description_2 = exit_code_3
    else:
        description_1, description_2 = '', ''

    number = 1
    time = datetime.now()
    test = Test(
        number=number,
        time=time,
        kl=address,
        exit_code_1=exit_code_1,
        exit_code_2=exit_code_2,
        exit_code_3=exit_code_3,
        ready_time=ready_time,
        device_type=device_type,
        device_version=device_version,
        sb_duration=sb_duration,
        electricity_sb=electricity_sb,
        average_electricity=average_electricity,
        temperature=temperature,
        Uv=Uv,
        Us=Us,
        Un=Un,
        Cv=Cv,
        Cs=Cs,
        Cn=Cn,
        charge_electricity_current=charge_electricity_current,
        discharge_rate_upper_ionistor=discharge_rate_upper_ionistor,
        discharge_rate_middle_ionistor=discharge_rate_middle_ionistor,
        discharge_rate_lower_ionistor=discharge_rate_lower_ionistor,
        reserve_1=reserve_1,
        reserve_2=reserve_2,
        description_1=description_1,
        description_2=description_2
    )

    total_tests_results = Report(
        date_time=datetime.now(),
        success_all=successfull_checks_devices_count,
        interrupted_all=manually_interrupted_tests_count,
        error_all=tests_completed_with_error_count,
        success_ipr=success_count_ipr,
        success_ipt=success_count_ipt,
        success_ipt_s=success_count_ipt_s,
        success_ipd=success_count_ipd,
        success_ipp=success_count_ipp,
        success_ipp_s=success_count_ipp_s,
        success_udp=success_count_udp,
        success_mks=success_count_mks,
        error_ipr=error_count_ipr,
        error_ipt=error_count_ipt,
        error_ipt_s=error_count_ipt_s,
        error_ipd=error_count_ipd,
        error_ipp=error_count_ipp,
        error_ipp_s=error_count_ipp_s,
        error_udp=error_count_udp,
        error_mks=error_count_mks
    )
    return test, total_tests_results, current_test

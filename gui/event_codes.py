import logging
import traceback

from PyQt5.QtGui import QColorConstants

from gui.exceptions import IncorrectMessageFileException


def get_codes_dictionary():
    """Возвращает словарь с кодами событий для КЛ."""
    logging.info('Попытка формирования словаря из файла KL_Messages.txt')
    codes_dictionary = {}
    txt_file = 'messages/KL_Messages.txt'
    try:
        open(txt_file)
    except Exception:
        logging.warning(f'Отсутствует файл "{txt_file}".\n {traceback.format_exc()}')
        raise IncorrectMessageFileException(f'Отсутствует файл "{txt_file}"')
    with open(txt_file, encoding='utf-8') as file:
        lines = file.readlines()
        if len(lines) == 0:
            logging.warning(f'файл "{txt_file}" пуст.')
            raise IncorrectMessageFileException(f'файл "{txt_file}" пуст. ')
        for line in lines:
            if len(line.split()) < 2:
                logging.warning(
                    f'Некорректные данные в файле "{txt_file}". Отсуствует часть параметров в строке "{line.strip()}".')
                raise IncorrectMessageFileException(
                    f'Некорректные данные в файле "{txt_file}". Отсуствует часть параметров в строке "{line.strip()}".')
            try:
                int(line.split()[0])
            except Exception:
                logging.warning(
                    f'В файле "{txt_file}" некорректный код сообщения в строке "{line.strip()}".\n {traceback.format_exc()}')
                raise IncorrectMessageFileException(
                    f'В файле "{txt_file}" некорректный код сообщения в строке "{line.strip()}".')
            codes = line.strip().split(' ', 2)  # инд. 0 - код ошибки; инд. 1 - цвет; инд. 2 - описание.
            if codes[1].lower() not in vars(QColorConstants.Svg).keys():
                logging.warning(
                    f'В файле "{txt_file}" rgb-код цвета в строке "{line.strip()}" отсуствует в программе.\n {traceback.format_exc()}')
                raise IncorrectMessageFileException(
                    f'В файле "{txt_file}" rgb-код цвета в строке "{line.strip()}" отсуствует в программе.')
            if len(codes) == 3:
                codes_dictionary[int(codes[0])] = (codes[2], codes[1])
            else:
                codes_dictionary[int(codes[0])] = ('', codes[1])
    return codes_dictionary


def get_error_dictionary():
    logging.info('Попытка формирования словаря из файла Stend_IP_error_codes.txt')
    errors_dictionary = {}
    errors_dictionary['0.0.0'] = ['', '']
    txt_file = 'messages/Stend_IP_error_codes.txt'
    try:
        open(txt_file)
    except Exception:
        logging.warning(f'Отсутствует файл "{txt_file}".\n {traceback.format_exc()}')
        raise IncorrectMessageFileException(f'Отсутствует файл "{txt_file}"')
    with open(txt_file, encoding='utf-8') as file:
        lines = file.readlines()
        if len(lines) == 0:
            logging.warning(f'файл "{txt_file}" пуст.')
            raise IncorrectMessageFileException(f'файл "{txt_file}" пуст. ')
        for line in lines:
            if len(line.strip().split('%')) < 3:
                logging.warning(
                    f'Некорректные данные в файле "{txt_file}". Отсуствует часть параметров в строке "{line.strip()}".')
                raise IncorrectMessageFileException(
                    f'Некорректные данные в файле "{txt_file}". Отсуствует часть параметров в строке "{line.strip()}".')
            code_1, code_2, code_3, desc_1_2, desc_3 = line.strip().split('%')
            if not code_1.isdigit() or not code_2.isdigit() or (not code_3.isdigit() and code_3 != '?'):
                logging.warning(f'Некооректный код ошибки в строке "{line}" файла "{txt_file}".')
                raise IncorrectMessageFileException(f'Некооректный код ошибки в строке "{line}" файла "{txt_file}".')
            code = f'{code_1}.{code_2}.{code_3}'
            errors_dictionary[code] = [desc_1_2, desc_3]
        return errors_dictionary

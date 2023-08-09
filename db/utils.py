from copy import deepcopy
from datetime import datetime

from sqlalchemy import create_engine, func, not_, or_
from sqlalchemy.orm import sessionmaker

from db.models import (CheckBox, Event, RememberedCombobox, Report,
                       SaveFilePath, StatisticSaved, Test)
from gui.event_codes import get_codes_dictionary

engine = create_engine('sqlite:///sqlite.db')
Session = sessionmaker(bind=engine, expire_on_commit=False)

try:
    codes_dictionary = get_codes_dictionary()
except Exception:
    pass


def get_session():
    """Получает объект сессии."""
    session = Session()
    return session


def get_count_of_events_in_db(session=get_session()):
    """Возвращает количество событий из журнала в базе данных."""
    events_count = session.query(Event).count()
    session.close()
    return events_count


def get_count_of_tests_in_db(session=get_session()):
    """Возвращает количество тестов из журнала в базе данных."""
    tests_count = session.query(Test).filter(not_(Test.is_archived)).count()
    session.close()
    return tests_count


def get_unarchived_tests(session=get_session()):
    """Возвращает все незаархивированные тесты."""
    tests = session.query(Test).filter(not_(Test.is_archived)).all()
    session.close()
    return tests


def get_tests_for_period(period, session=get_session()):
    tests = session.query(Test).filter(Test.time.between(period[0], period[1])).all()
    session.close()
    return tests


def get_error_tests_for_period(period, session=get_session()):
    tests = session.query(Test).filter(Test.time.between(period[0], period[1])).filter(or_(Test.exit_code_1.between(1, 19), Test.exit_code_1 >= 100)).all()
    session.close()
    return tests


def get_reports_for_period(period, session=get_session()):
    reports = session.query(Report).filter(Report.date_time.between(period[0], period[1])).all()
    session.close()
    return reports


def save_events_to_db(events, session=get_session()):
    """Сохраняет в БД события из пакета данных."""
    packet_events = []
    current_number = get_count_of_events_in_db() + 1
    for event in events:
        event.number = current_number
        session.add(event)
        packet_events.append(event)
        current_number += 1
    session.commit()
    session.close()
    return packet_events


def save_test_to_db(test, session=get_session()):
    """Сохраняет в БД тесты из пакета данных."""
    test.number = get_count_of_tests_in_db() + 1
    session.add(test)
    session.commit()
    session.close()
    return test


def save_report_to_db(report, session=get_session()):
    report.date_time = datetime.now()
    new_report = deepcopy(report)
    session.query(Report).filter(func.DATE(Report.date_time) == report.date_time.date()).delete()
    session.add(new_report)
    session.commit()
    session.close()


def clear_events_db(session=get_session()):
    """Удаляет все события журнала из базы данных."""
    session.query(Event).delete()
    session.commit()
    session.close()


def clear_tests_db_for_period(period, session=get_session()):
    """Удаляет все тесты за указанный период из базы данных."""
    session.query(Test).filter(Test.time.between(period[0], period[1])).delete()
    session.commit()
    session.close()


def archive_test_db(session=get_session()):
    """Архивирует результаты тестов находящихся в таблице."""
    session.query(Test).filter(Test.is_archived is not True).update({'is_archived': True})
    session.commit()
    session.close()


def get_latest_report(session=get_session()):
    """Возвращает из БД самый последний отчет"""
    report = session.query(Report).order_by(Report.date_time.desc()).limit(1).first()
    return report


def is_statistic_saved(document, month, year, session=get_session()):
    saved_statistic = session.query(
        StatisticSaved).filter(StatisticSaved.document == document).filter(StatisticSaved.year == year).filter(StatisticSaved.month == month).all()
    if not saved_statistic:
        return False
    else:
        return True


def create_saved_statistic(document, month, year, session=get_session()):
    saved_statistic = StatisticSaved(document=document, year=year, month=month)
    session.add(saved_statistic)
    session.commit()
    session.close()


def get_events_codes(session=get_session()):
    """Получает стороку типа 'код события - описание' для
    всех событий в таблице главного окна.
    Применяется для фильтрации по полю код события и описание."""
    codes_list = []
    result = session.query(Event.event_code.distinct(), Event.description).all()
    result.sort()
    if result:
        for event in result:
            codes_list.append(f'{event[0]} - {event[1]}')
    session.close()
    return codes_list


def remember_checkbox(is_checked, session=get_session()):
    """Записывает в БД положение чекбокса выбора порта."""
    if len(session.query(CheckBox).all()) == 0:
        checkbox = CheckBox(is_checked=is_checked)
        session.add(checkbox)
    else:
        checkbox_to_del = session.query(CheckBox).first()
        session.delete(checkbox_to_del)
        checkbox = CheckBox(is_checked=is_checked)
        session.add(checkbox)
    session.commit()
    session.close()


def is_checked_checkbox(session=get_session()):
    """Возвращает из БД положение чекбокса выбора порта."""
    if len(session.query(CheckBox).all()) == 0:
        session.close()
        return False
    else:
        session.close()
        return session.query(CheckBox).first().is_checked


def remember_combobox(current_text, session=get_session()):
    """Записывает в БД положение комбобокса выбора порта."""
    if len(session.query(RememberedCombobox).all()) == 0:
        combobox = RememberedCombobox(current_text=current_text)
        session.add(combobox)
    else:
        combobox_to_del = session.query(RememberedCombobox).first()
        session.delete(combobox_to_del)
        combobox = RememberedCombobox(current_text=current_text)
        session.add(combobox)
    session.commit()
    session.close()


def delete_remembered_combobox(session=get_session()):
    """Удаляет из БД значение комбобокса выбора порта."""
    if len(session.query(RememberedCombobox).all()) != 0:
        combobox_to_del = session.query(RememberedCombobox).first()
        session.delete(combobox_to_del)
        session.commit()
        session.close()


def combobox_remembered_text(session=get_session()):
    """Возвращает значение комбобокса выбора порта."""
    if len(session.query(RememberedCombobox).all()) == 0:
        session.close()
        return 'Не выбран'
    combobox_text = session.query(RememberedCombobox).first().current_text
    session.close()
    return combobox_text


def remember_save_file_path(path, table, session=get_session()):
    """Запоминает путь к сохраненному файлу."""
    if len(session.query(SaveFilePath).all()) == 0:
        save_file_path = SaveFilePath(path=path, table=table)
        session.add(save_file_path)
    else:
        save_file_path_to_del = session.query(SaveFilePath).filter(SaveFilePath.table == table).first()
        if save_file_path_to_del:
            session.delete(save_file_path_to_del)
        save_file_path = SaveFilePath(path=path, table=table)
        session.add(save_file_path)
    session.commit()
    session.close()


def get_save_file_path(table, session=get_session()):
    """Возвращает путь к сохраненному файлу."""
    if len(session.query(SaveFilePath).filter(SaveFilePath.table == table).all()) == 0:
        session.close()
        return
    file_path = session.query(SaveFilePath).filter(SaveFilePath.table == table).first().path
    session.close()
    return file_path

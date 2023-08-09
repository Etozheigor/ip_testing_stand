from sqlalchemy import (Boolean, Column, DateTime, Float, Integer, String,
                        create_engine)
from sqlalchemy.orm import Session, declarative_base

engine = create_engine('sqlite:///sqlite.db')
Base = declarative_base()


class Event(Base):
    """Модель события из журнала КЛ."""
    __tablename__ = 'event'
    id = Column(Integer, primary_key=True)
    number = Column(Integer)
    date_time = Column(DateTime)
    kl = Column(Integer)
    entry_number = Column(Integer)
    au_address = Column(Integer)
    event_code = Column(Integer)
    addit_param = Column(Integer)
    description = Column(String)

    def __repr__(self):
        return str(self.event_code)


class Test(Base):
    """Модель теста КЛ."""
    __tablename__ = 'test'
    id = Column(Integer, primary_key=True)
    is_archived = Column(Boolean, default=False)
    number = Column(Integer)
    time = Column(DateTime)
    kl = Column(Integer)
    exit_code_1 = Column(Integer)
    exit_code_2 = Column(Integer)
    exit_code_3 = Column(Integer)
    ready_time = Column(Integer)
    device_type = Column(String)
    device_version = Column(Integer)
    sb_duration = Column(Integer)
    electricity_sb = Column(Integer)
    average_electricity = Column(Float)
    temperature = Column(Integer)
    Uv = Column(Float)
    Us = Column(Float)
    Un = Column(Float)
    Cv = Column(Float)
    Cs = Column(Float)
    Cn = Column(Float)
    charge_electricity_current = Column(Float)
    discharge_rate_upper_ionistor = Column(Float)
    discharge_rate_middle_ionistor = Column(Float)
    discharge_rate_lower_ionistor = Column(Float)
    reserve_1 = Column(Integer)
    reserve_2 = Column(Integer)
    description_1 = Column(String)
    description_2 = Column(String)


class CheckBox(Base):
    """Модель положения чекбокса."""
    __tablename__ = 'checkbox'
    id = Column(Integer, primary_key=True)
    is_checked = Column(Boolean)


class RememberedCombobox(Base):
    """Модель значения комбобокса."""
    __tablename__ = 'remembered_combobox'
    id = Column(Integer, primary_key=True)
    current_text = Column(String)


class SaveFilePath(Base):
    """Модель пути к сохраненному файлу."""
    __tablename__ = 'save_file_path'
    id = Column(Integer, primary_key=True)
    table = Column(String)
    path = Column(String)

    def __repr__(self):
        return str(self.path)


class Report(Base):
    """Модель общей статистики тестов КЛ."""
    __tablename__ = 'report'
    id = Column(Integer, primary_key=True)
    date_time = Column(DateTime)
    is_archived = Column(Boolean, default=False)
    success_all = Column(Integer)
    interrupted_all = Column(Integer)
    error_all = Column(Integer)
    success_ipr = Column(Integer)
    success_ipt = Column(Integer)
    success_ipt_s = Column(Integer)
    success_ipd = Column(Integer)
    success_ipp = Column(Integer)
    success_ipp_s = Column(Integer)
    success_udp = Column(Integer)
    success_mks = Column(Integer)
    error_ipr = Column(Integer)
    error_ipt = Column(Integer)
    error_ipt_s = Column(Integer)
    error_ipd = Column(Integer)
    error_ipp = Column(Integer)
    error_ipp_s = Column(Integer)
    error_udp = Column(Integer)
    error_mks = Column(Integer)


class StatisticSaved(Base):
    """Хранит данные о том, были ли сохранены отчеты за каждый месяц."""
    __tablename__ = 'statistic'
    id = Column(Integer, primary_key=True)
    document = Column(String)
    year = Column(Integer)
    month = Column(Integer)


Base.metadata.create_all(engine)
session = Session(engine)

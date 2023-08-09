import logging
import sys

from PyQt5.QtCore import QLocale, QTranslator
from PyQt5.QtWidgets import QApplication, QMessageBox

from gui.main_window import MainWindow
from logs_config import config_logs

if __name__ == '__main__':
    import traceback

    def excepthook(exc_type, exc_value, exc_tb):
        """Выводит сообщение об ошибке в случае краша главного окна."""
        tb = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        QMessageBox.critical(
            None,
            "Критическая ошибка",
            f'Критическая ошибка в ходе выполнения программы. Перезапустите программу и повторите попытку.\nКод ошибки: {tb}'
        )
        QApplication.quit()

    config_logs()
    logging.info('Старт работы программы')

    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    sys.excepthook = excepthook
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())

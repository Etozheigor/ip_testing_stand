class IncorrectMessageFileException(Exception):
    """Кастомное исключение, если файлы с текстами сообщений заполнены некорректно."""
    def __init__(self, message):
        super().__init__()
        self.message = message

from docparser import DocxEnumTag


class Field:
    """
    Поле для сохранения значения.
    """
    columns: int = 1
    rows: int = 1
    owner: DocxEnumTag = None
    value = None

    def __call__(self, value: str):
        raise NotImplementedError


class LineField(Field):
    """
    Однострочное поле, множественные вызовы перезаписывают значение.
    """
    def __init__(self, columns: int, owner: DocxEnumTag):
        """
        :param columns: Колличество колонок занимаемое значением.
        :param owner: Тэг владелец значения.
        """
        self.columns = columns
        self.owner = owner
        self.rows = 1

    def __call__(self, value: str):
        self.value = value


class MultiField:
    """
    Многострочное поле, множественные вызовы дописывают значение.
    """
    def __init__(self, columns: int, owner: DocxEnumTag, rows: int = None):
        """
        :param columns: Колличество колонок занимаемое значением.
        :param owner: Тэг владелец значения.
        :param rows: Колличество строк занимаемое значением, если колличество неизвестно, то None.
        """
        self.rows = rows
        self.columns = columns
        self.owner = owner
        self.value: list[list[str]] = []

    def __call__(self, value: list[str]):
        self.value.append(value)


class UnsetFieldError(Exception):
    pass


class XlsxData:
    """
    Интерфейс доступа к данным xlsx документа.
    """
    def get_field(self, xlsx_field: str) -> Field | None:
        """
        Возвращает поле, соответствующую строке xlsx_field.

        :param xlsx_field: имя xlsx поля.
        :return: хранилище значения.
        """
        raise NotImplementedError

    def get_unset_fields(self) -> tuple[tuple[str, Field]]:
        """
        Возвращает не установленные поля данных.

        :return: название поля, поле
        """
        raise NotImplementedError

    def get(self, tag: DocxEnumTag) -> Field | None:
        """ Возвращает поле, соответствующую тэгу tag. """
        raise NotImplementedError

    def __iter__(self):
        raise NotImplementedError

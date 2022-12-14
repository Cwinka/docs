from enum import Enum
from typing import Iterator, Type, Iterable
from pathlib import Path


def raise_invalid_path(path: Path, throw: Type[Exception], *, exts: Iterable[str] = None):
    """
    Проверяет существует ли путь и является ли путь валидным, иначе поднимает
    исключение throw.

    :param path: путь до документа.
    :param throw: тип исключения.
    :param exts: разрешенные расширения (указываются с точкой).
    :return:
    """
    if not path.exists():
        raise throw(f'Путь "{path}" не существует.')
    if path.name in ('', '.', '..'):
        raise throw(f'Некорректный путь: {path}.')
    if exts is not None:
        if path.suffix not in exts:
            raise throw(f'Недопустимое расширение документа "{path}". Необходим документ в '
                        f'одном из форматов: [{", ".join(exts)}]')


class DocxEnumTag(Enum):
    """
    Список доступных тэгов для использования в шаблоне docx
    """
    KIND = 'KIND'
    AIM = 'AIM'
    GRADE = 'GRADE'
    FACULTY = 'FACULTY'
    GROUP = 'GROUP'
    STUDY_TYPE = 'STUDY_TYPE'
    SPECIALIZATION = 'SPECIALIZATION'
    PERIOD_YEARS = 'PERIOD_YEARS'
    PERIOD_DAYS = 'PERIOD_DAYS'
    PULPIT = 'PULPIT'
    DIRECTOR = 'DIRECTOR'
    DIRECTOR_NAME = 'DIRECTOR_NAME'
    STUDENTS = 'STUDENTS'
    TABLES = 'TABLES'


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


class MultiField(Field):
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

    def help_iter(self) -> Iterator[tuple[str, Field]]:
        """ Итератор, возвращающий кононичное название поля и само поле"""
        raise NotImplementedError

    def __iter__(self) -> Iterator[Field]:
        raise NotImplementedError

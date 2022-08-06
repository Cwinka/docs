from typing import Callable
from docx.document import Document as HintDocument
from docx.table import _Cell
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches
from docparser import DocxEnumTag


class Field:
    def __init__(self, columns: int, owner: DocxEnumTag, rows: int = None):
        """
        :param rows: Колличество строк занимаемое значением.
        :param columns: Колличество колонок занимаемое значением.
        :param owner: Тэг владелец значения.
        """
        self.value = None
        self.rows = rows
        self.columns = columns
        self.owner = owner

    def __call__(self, value: str):
        self.value = value


class UnsetFieldError(Exception):
    pass


class XlsxData:

    def get_field(self, xlsx_field: str) -> Field | None:
        """
        Возвращает функцию установки значения поля xlsx_field в аргумент класса.

        :param xlsx_field: имя xlsx поля.
        :return: функция установки, колличество строк которое занимают значения (если None, тогда значения идут до
                 следующего ряда с установленным полем A), колличество стольбцов для одного значения.
        """
        raise NotImplementedError

    def get_unset_fields(self) -> tuple[tuple[str, Field]]:
        """
        Возвращает не установленные поля данных. Неустановленные поля являются причиной
        завершения программы.

        :return:
        """
        raise NotImplementedError

    @staticmethod
    def _set_in_list(arr: list) -> Callable[[any], None]:
        """ Возвращает функцию установки значения в список. """
        return lambda x: arr.append(x)

    def _set_it(self, field: str) -> Callable[[any], None]:
        """ Возвращает функцию установки значения в аргумент класса. """
        return lambda x: setattr(self, field, x)


class UsurtBaseTable:

    def __init__(self, doc: HintDocument, rows: int = 1, cols: int = 4,
                 columns_width: tuple[Inches, ...] = None):
        self.doc = doc
        self.table = self.doc.add_table(rows, cols, style='Table Grid')
        self.table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        self.last_row = self.table.rows[0]
        self.rows = rows - 1
        if columns_width:
            for i, w in enumerate(columns_width):
                for c in self.table.columns[i].cells:
                    c.width = w

    def add_row(self, *parts: str, p_style: str = None, char_style: str = None) -> int:
        """
        Добавляет строку в таблицу, вставляя parts в колонки, слева направо.

        :param parts: текст для вставки в колонку.
        :param p_style: стиль параграфа вставки.
        :param char_style: стиль текста вставки.
        :return: индекс добаленного ряда
        """
        if self.last_row is None:
            self.last_row = self.table.add_row()
            self.rows += 1
        for c, part in zip(self.last_row.cells, parts):
            if part:
                c: _Cell = c
                p = c.paragraphs[0]
                p.style = p_style
                p.add_run(part, style=char_style)
        self.last_row = None
        return self.rows

    def merge(self, row_ids: tuple[int, ...], *, cells: tuple[int, int]):
        """
        Соединяет cells в рядах с индексами row_ids.

        :param row_ids: индексы рядов.
        :param cells: диапазон колонок.
        :return:
        """
        if len(row_ids) == 1:
            # горизонтальное соединение
            _cells = self.table.row_cells(row_ids[0])
            c1 = _cells[cells[0]]
            c2 = _cells[cells[1]]
            c1.merge(c2)
        else:
            for y, (c1, c2) in enumerate(zip(*(self.table.row_cells(i) for i in row_ids))):
                if y in cells:
                    c1.merge(c2)

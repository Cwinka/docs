from typing import Callable
from docx.document import Document as HintDocument
from docx.text.paragraph import Paragraph
from docx.table import _Cell, Table
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Length
from docparser import DocxEnumTag
from docx.oxml import CT_Tbl


class Field:
    columns: int = 1
    rows: int = 1
    owner: DocxEnumTag = None
    value = None

    def __call__(self, value: str):
        raise NotImplementedError


class LineField(Field):
    def __init__(self, columns: int, owner: DocxEnumTag):
        """
        :param columns: Колличество колонок занимаемое значением.
        :param owner: Тэг владелец значения.
        """
        self.columns = columns
        self.owner = owner

    def __call__(self, value: str):
        self.value = value


class MultiField:
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

    @staticmethod
    def _set_in_list(arr: list) -> Callable[[any], None]:
        """ Возвращает функцию установки значения в список. """
        return lambda x: arr.append(x)

    def _set_it(self, field: str) -> Callable[[any], None]:
        """ Возвращает функцию установки значения в аргумент класса. """
        return lambda x: setattr(self, field, x)


class UsurtBaseTable:

    def __init__(self, p: Paragraph, rows: int = 1, cols: int = 4, *, width: Length,
                 columns_width: tuple[Inches, ...] = None):
        """ Создаёт таблицу внутри """
        self.table = self._make_table(rows, cols, width, p)
        self.table.style = 'Table Grid'
        self.table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        self.last_row = self.table.rows[0]
        self.rows = rows - 1
        self.cols = cols
        if columns_width:
            for i, w in enumerate(columns_width):
                for c in self.table.columns[i].cells:
                    c.width = w

    @staticmethod
    def _make_table(rows: int, cols: int, width: Length, p: Paragraph):
        tbl = CT_Tbl.new_tbl(rows, cols, width)
        return Table(tbl, p)

    def make_base_headings(self):
        self.add_row('Фамилия Имя Отчество обучающегося', 'Группа,форма обучения (ц, б, к)',
                  'Руководитель практики от УрГУПС')
        self.merge((0,), cells=(2, 3))
        self.add_row('', '', 'Должность', 'Ф.И.О.')
        self.merge((0, 1), cells=(0, 1))

    def apply(self):
        self.table._parent._p.addnext(self.table._tbl)

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

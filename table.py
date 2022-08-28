from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import CT_Tbl
from docx.shared import Length, Inches
from docx.table import Table
from docx.text.paragraph import Paragraph


class DocxTable:
    """ Таблица для docx документа. """

    def __init__(self, p: Paragraph, rows: int = 1, cols: int = 4, *, width: Length,
                 columns_width: tuple[Inches, ...] = None):
        self.table = self._make_table(rows, cols, width, p)  # создание экземпляра таблицы.
        # добавление стилей в таблицу.
        self.table.style = 'Table Grid'
        self.table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        self.last_row = self.table.rows[0]
        self.rows = rows - 1
        self.cols = cols
        if columns_width:  # если задана ширина колонок, установить для каждой ячейки.
            for i, w in enumerate(columns_width):
                for c in self.table.columns[i].cells:
                    c.width = w

    @staticmethod
    def _make_table(rows: int, cols: int, width: Length, p: Paragraph):
        tbl = CT_Tbl.new_tbl(rows, cols, width)
        return Table(tbl, p)

    def make_base_headings(self):
        """ Создание базовой шапки таблицы студентов. """
        self.add_row('Фамилия Имя Отчество обучающегося', 'Группа,форма обучения (ц, б, к)',
                  'Руководитель практики от УрГУПС', align=True)
        self.merge((0,), cells=(2, 3))
        self.add_row('', '', 'Должность', 'Ф.И.О.')
        self.merge((0, 1), cells=(0, 1))

    def apply(self):
        """ Добавляет таблцу в документ. """
        self.table._parent._p.addnext(self.table._tbl)

    def add_row(self, *parts: str, p_style: str = None, char_style: str = None, align: bool = False) -> int:
        """
        Добавляет строку в таблицу, вставляя parts в колонки, слева направо.

        :param parts: текст для вставки в колонку.
        :param p_style: стиль параграфа вставки.
        :param char_style: стиль текста вставки.
        :param align: выравнивать по центру.
        :return: индекс добаленного ряда
        """
        if self.last_row is None:  # если таблица пустая, добавить один ряд.
            self.last_row = self.table.add_row()
            self.rows += 1
        for cell, content in zip(self.last_row.cells, parts):  # в каждую ячейку разместить своё значение.
            if content:
                p = cell.paragraphs[0]

                p.style = p_style  # применение стилей.
                if align:
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

                p.add_run(content, style=char_style)  # добавление текста в параграф.
        self.last_row = None
        return self.rows

    def merge(self, row_ids: tuple[int, ...], *, cells: tuple[int, int]):
        """
        Соединяет индесы ячеек cells в рядах с индексами row_ids.

        :param row_ids: индексы рядов.
        :param cells: диапазон колонок.
        :return:
        """
        if len(row_ids) == 1:
            # горизонтальное соединение.
            _cells = self.table.row_cells(row_ids[0])  # все ячейки в ряде.
            c1 = _cells[cells[0]]  # левая ячейка.
            c2 = _cells[cells[1]]  # правая ячейка.
            c1.merge(c2)
        else:
            # пронумеровав, соединить левую ячейку с правой, только если номер указан в cells.
            for y, (c1, c2) in enumerate(zip(*(self.table.row_cells(i) for i in row_ids))):
                if y in cells:
                    c1.merge(c2)

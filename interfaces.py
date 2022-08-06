from pathlib import Path
from typing import Generator, Callable, TypeVar, Type
import pymorphy2
from loguru import logger
import re
import openpyxl
from docx.document import Document as HintDocument
from docx.table import _Cell
from docx.text.paragraph import Run, Paragraph
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches


class UsurtDoc:

    _morf = pymorphy2.MorphAnalyzer(lang='ru')

    def intro(self) -> Generator[str, any, None]:
        """ Содержит итератор по строкам титутульного листа документа. """
        raise NotImplementedError

    def order(self) -> Generator[str, any, None]:
        """ Содержит итератор по строкам приказа документа. """
        raise NotImplementedError

    def ending(self) -> Generator[str, any, None]:
        """ Содержит итератор по строкам которые заканчивают документ. """
        raise NotImplementedError

    def make_table(self, doc: 'HintDocument') -> 'UsurtBaseTable':
        """ Создаёт таблицу с заголовками в документе. """
        raise NotImplementedError

    def make_all_tables(self, doc: 'HintDocument'):
        """ Создает все необходимые таблицы в документе. """

    def _morf_to(self, text: str, target: str) -> str:
        """
        Приводит текст в нужный падеж.

        :param text: текст для постановки в нужный падеж.
        :param target: граммема.
        :return:
        """
        proper = ''
        for cleared, orig in self._splitter(text):
            inf = self._inflect(cleared, target)
            if inf.lower() == cleared.lower():
                inf = cleared
            proper += re.sub(cleared, inf, orig) + ' '
        return proper

    @staticmethod
    def _splitter(text: str) -> tuple[str, str]:
        """
        Разделяет текст на слова.

        :param text: текст для разделения на слова.
        :return: очищенное слово без знаков, оригинальное слово.
        """
        for part in text.split(' '):
            if not part:
                continue
            yield part.strip('()'), part

    def _inflect(self, word: str, target: str) -> str:
        """
        Приводит слово в нужный падеж.
        Метод не обрабатыет слова, написанные в верхнем регистре.
        Метод не обрабатыет слова, длиной меньше 3 символов.

        :param word: слово.
        :param target: граммема.
        :return: слово поставленное в указаннай падеж.
        """
        if word.isupper() or len(word) < 3:
            return word
        m = self._morf.parse(word)[0]
        inf = m.inflect({target})
        if inf is None:
            logger.warning(f'Не удалось привести слово "{word}" к таргету "{target}".')
            return word
        return inf.word


class XlsxData:

    def get_setter(self, xlsx_field: str) -> tuple[[Callable[[str], None]], int, int] | None:
        """
        Возвращает функцию установки значения поля xlsx_field в аргумент класса.

        :param xlsx_field: имя xlsx поля.
        :return: функция установки, колличество строк которое занимают значения (если None, тогда значения идут до
                 следующего ряда с установленным полем A), колличество стольбцов для одного значения.
        """
        raise NotImplementedError

    def get_unset_fields(self) -> tuple[str]:
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


class UnsetFieldsError(Exception):
    def __init__(self, m: str):
        super().__init__(m)
        self.err = m


class XlsxDataParser:

    _XlsxData = TypeVar('_XlsxData')

    @classmethod
    def parse(cls, path: Path, data_cls: Type[_XlsxData]) -> _XlsxData:
        """
        Парсит данные из xlsx файла path.

        :param path: путь до файла xlsx.
        :param data_cls: тип класса XlsxData.
        :return:
        """
        wb_obj = openpyxl.load_workbook(path)
        sheet = wb_obj.active
        data = data_cls()
        rows = sheet.iter_rows(values_only=Type)
        for row in rows:
            cls._set_xlsx_value_to_data(row, rows, data)

        unset = data.get_unset_fields()
        if unset:
            err = "\n".join(unset)
            raise UnsetFieldsError(f'Недостаточно данных для формирования docx документа, '
                                   f'следующие поля xlsx документа {path.as_posix()} должны быть устанвлены: \n{err}')

        return data

    @classmethod
    def _set_xlsx_value_to_data(cls, row, rows, data: '_XlsxData'):
        """
        Устанавливает значение поля из xlsx в соответсвующее поле структуры data.

        :param row: текущий ряд.
        :param rows: генератор всех рядов.
        :param data: структура данных.
        :return:
        """
        s = cls._get_setter(row, data)
        if s:
            setter, _rows, _cols = s
            if _rows and _rows == 1:
                value = cls._extract_from_row(row, 1)
            else:
                value, row = cls._multi_line_values(rows, row, _rows, _cols)
                if row:
                    cls._set_xlsx_value_to_data(row, rows, data)
            setter(value)

    @staticmethod
    def _get_setter(row, data: '_XlsxData') -> tuple[Callable, int, int] | None:
        """
        Возвращает результат метода XlsxData.get_setter.

        :param row: текущий ряд.
        :param data: структура данных.
        :return:
        """
        xlsx_field: str = row[0]
        if xlsx_field:
            line_field = xlsx_field.strip().replace('\n', ' ')
            if not line_field.startswith('#'):
                # # обозначает коментарий
                s = data.get_setter(line_field)
                if s:
                    return s

    @classmethod
    def _multi_line_values(cls, rows_gen, current_row, rows: int | None, cols: int):
        """
        Извлекает многостроковое значение из xlsx.

        :param rows_gen: генератор рядов
        :param current_row: текущий ряд
        :param rows: колличество рядов соответсвующих значений
        :param cols: колличество колонок соответсвующих значений
        :return:
        """
        values = [cls._extract_from_row(current_row, cols)]
        last_row = None
        if rows:
            rows -= 1  # так как current_row уже добавлен
            for row in rows_gen:
                if rows == 0:
                    last_row = row
                    break
                values.append(cls._extract_from_row(row, cols))
                rows -= 1
        else:
            for row in rows_gen:
                if any(row) and row[0] is None:  # первое поле (A) для ключей
                    values.append(cls._extract_from_row(row, cols))
                else:
                    last_row = row
                    break
        return values, last_row

    @staticmethod
    def _extract_from_row(row, cols: int):
        return row[1] if cols == 1 else row[1:cols+1]

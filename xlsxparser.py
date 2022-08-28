from pathlib import Path
from typing import Type, Iterator
from interfaces import Field, XlsxData
import openpyxl


class XlsxDataParserError(Exception):
    pass


class XlsxDataParser:
    """
    Парсер данных xlsx документа.
    """

    def __init__(self, path: Path):
        if path.suffix not in ('.xlsx', 'xls'):
            raise XlsxDataParserError(f'Неподходящий формат документа {path}. Необходим документ в формате xls/xlsx.')
        wb_obj = openpyxl.load_workbook(path)
        self.sheet = wb_obj.active

    def parse(self, keeper: Type[XlsxData]) -> XlsxData:
        """
        Парсит данные из xlsx файла path в хранилище данных keeper.

        :param keeper: хранилище данных.
        :return:
        """
        keep = keeper()
        rows = self.sheet.iter_rows(values_only=True)
        for row in rows:
            self._set_xlsx_value_in_keeper(row, rows, keep)
        return keep

    @classmethod
    def _set_xlsx_value_in_keeper(cls, row: list[str], rows: Iterator[list[str]], keeper: XlsxData):
        """
        Устанавливает значение поля из xlsx в соответсвующее поле структуры data.

        :param row: текущий ряд.
        :param rows: генератор всех рядов.
        :param keeper: хранилище данных.
        :return:
        """
        row_key = row[0]
        if field := cls._get_field(row_key, keeper):
            if field.rows and field.rows == 1:
                value = cls._extract_from_row(row, 1)
            else:
                value, row = cls._extract_multi_row_value(rows, row, field.rows, field.columns)
                if row:
                    cls._set_xlsx_value_in_keeper(row, rows, keeper)
            field(value)

    @staticmethod
    def _get_field(key: str, keeper: XlsxData) -> Field | None:
        """
        Возвращает поле из keeper для хранения значения.

        :param key: ключ, по которому необходимо найти поле в хранилище keeper.
        :param keeper: хранилище данных.
        :return:
        """
        if key:
            line_field = key.strip().replace('\n', ' ')
            if not line_field.startswith('#'):  # # обозначает коментарий
                return keeper.get_field(line_field)

    @classmethod
    def _extract_multi_row_value(cls, rows_iter: Iterator[list[str]], current_row: list[str], rows: int | None,
                                 cols: int):
        """
        Извлекает многостроковое значение из xlsx.

        :param rows_iter: итератор по рядам.
        :param current_row: текущий ряд.
        :param rows: колличество рядов значений.
        :param cols: колличество колонок значений.
        :return: Последний просмотренный ряд.
        """
        values = [cls._extract_from_row(current_row, cols)]
        last_row = None
        if rows:
            rows -= 1  # так как current_row уже добавлен
            while rows:
                row = next(rows_iter)
                values.append(cls._extract_from_row(row, cols))
                rows -= 1
            last_row = next(rows_iter)
        else:
            for row in rows_iter:
                if any(row) and row[0] is None:  # первое поле (A) для ключей
                    values.append(cls._extract_from_row(row, cols))
                else:
                    last_row = row
                    break
        return values, last_row

    @staticmethod
    def _extract_from_row(row, cols: int):
        return row[1] if cols == 1 else row[1:cols+1]

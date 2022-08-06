from pathlib import Path
from typing import TypeVar, Type
from interfaces import Field
import openpyxl


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
        if field := cls._get_data_field(row, data):
            if field.rows and field.rows == 1:
                value = cls._extract_from_row(row, 1)
            else:
                value, row = cls._multi_line_values(rows, row, field.rows, field.columns)
                if row:
                    cls._set_xlsx_value_to_data(row, rows, data)
            field(value)

    @staticmethod
    def _get_data_field(row, data: '_XlsxData') -> Field | None:
        """
        Возвращает поле из data для установки значения.

        :param row: текущий ряд.
        :param data: структура данных.
        :return:
        """
        xlsx_field: str = row[0]
        if xlsx_field:
            line_field = xlsx_field.strip().replace('\n', ' ')
            if not line_field.startswith('#'):  # # обозначает коментарий
                return data.get_field(line_field)

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

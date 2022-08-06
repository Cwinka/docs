import argparse
import sys
from pathlib import Path

from docx.document import Document as HintDocument
from loguru import logger

from docparser import DocxWithTags, DocxEnumTag, UnknownDueDate
from interfaces import XlsxData, Field, UnsetFieldError
from xlsxparser import XlsxDataParser


class UsurtXlsxTrainingData(XlsxData):

    def __init__(self):
        self.columns: dict[str, Field] = {
            "Вид практики": Field(1, DocxEnumTag.KIND, rows=1),
            "Тип практики": Field(1, DocxEnumTag.AIM, rows=1),
            "Курс": Field(1, DocxEnumTag.GRADE, rows=1),
            "Факультет": Field(1, DocxEnumTag.FACULTY, rows=1),
            "Группа": Field(1, DocxEnumTag.GROUP, rows=1),
            "Форма обучения": Field(1, DocxEnumTag.STUDY_TYPE, rows=1),
            "Специализация": Field(1, DocxEnumTag.SPECIALIZATION, rows=1),
            "Период практики (годы)": Field(1, DocxEnumTag.PERIOD_YEARS, rows=1),
            "Период практики (дни)": Field(1, DocxEnumTag.PERIOD_DATE, rows=1),
            "Кафедра": Field(1, DocxEnumTag.PULPIT, rows=1),
            "Должность руководителя практики": Field(1, DocxEnumTag.DIRECTOR, rows=1),
            "ФИО руководителя практики": Field(1, DocxEnumTag.DIRECTOR_NAME, rows=1),
            "ФИО студентов. Группа, форма обучения": Field(1, DocxEnumTag.TRAINING_TABLE),
        }

    def __iter__(self):
        return iter(self.columns.values())

    def get_field(self, xlsx_field: str) -> Field | None:
        """ Возвращает поле, соответствующее значению xlsx_field. """
        return self.columns.get(xlsx_field)

    def get_unset_fields(self) -> tuple[tuple[str, Field]]:
        return tuple((s, field) for s, field in self.columns.items() if field.value is None)


# class UsurtTrainingDoc(UsurtDoc):
#
#     def __init__(self, data: UsurtXlsxTrainingData):
#         self.data = data
#
#     def make_table(self, doc: HintDocument) -> 'UsurtBaseTable':
#         t = UsurtBaseTable(doc, columns_width=(Inches(5),))
#         t.add_row('Фамилия Имя Отчество обучающегося', 'Группа,форма обучения (ц, б, к)',
#                   'Руководитель практики от УрГУПС')
#         t.merge((0,), cells=(2, 3))
#         t.add_row('', '', 'Должность', 'Ф.И.О.')
#         t.merge((0, 1), cells=(0, 1))
#         return t
#
#     def make_all_tables(self, doc: 'HintDocument'):
#         students = self.make_table(doc)
#         director = self.data.director
#         director_name = self.data.director_name
#         for student in self.data.students:
#             name, group = student
#             students.add_row(name, group, director, director_name)


def check_filled(data: XlsxData, doc: DocxWithTags):
    if unset := data.get_unset_fields():
        unset_tags = tuple(field[1].owner for field in unset)
        used_tags = doc.get_used_tags()
        for tag in used_tags:
            if tag in unset_tags:
                err = "\n".join(field[0] for field in unset)
                raise UnsetFieldError(f'Недостаточно данных для формирования docx документа, '
                                      f'следующие поля должны быть установлены: \n{err}')


def main():
    logger.remove()
    logger.add(sys.stdout, colorize=True, format="<level>{level}</level> | <level>{message}</level>")
    parser = argparse.ArgumentParser()
    parser.add_argument('docx', type=str, help='путь до шаблона docx документа.')
    parser.add_argument('xlsx', type=str, help='путь до xlsx документа с данными.')
    parser.add_argument('-o', '-out', type=str, help='путь до нового doc документа.')
    args = parser.parse_args()
    xlsx_path = Path(args.xlsx)
    docx_path = Path(args.docx)
    out = Path(args.o) if args.o else Path(f'{docx_path.stem}-prepared.docx')

    data = XlsxDataParser.parse(xlsx_path, UsurtXlsxTrainingData)
    doc = DocxWithTags(docx_path, init=True)

    try:
        check_filled(data, doc)
    except UnsetFieldError as e:
        logger.error(e)
        exit(1)

    for field in data:
        try:
            doc.replace_tag(field.owner, field.value)
        except UnknownDueDate as e:
            logger.error(e)
            exit(1)

    logger.info(f'Документ {out.as_posix()} успешно создан по шаблону {docx_path.as_posix()} на '
                f'основе данных из {xlsx_path.as_posix()}.')
    doc.save(out)


if __name__ == '__main__':
    main()

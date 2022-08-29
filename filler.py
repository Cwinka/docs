import argparse
import sys
from pathlib import Path

from docx.shared import Inches
from docx.text.paragraph import Paragraph
from loguru import logger
from typing import Iterator
from docparser import TaggedDoc, DocxEnumTag, UnknownDueDate, TaggedDocError
from interfaces import XlsxData, LineField, UnsetFieldError, MultiField, Field
from table import DocxTable
from xlsxparser import XlsxDataParser, XlsxDataParserError


class UsurtData(XlsxData):
    """
    Данные xlsx документа, которые можно ввести.
    """
    def __init__(self):
        self.columns: dict[str, Field] = {
            "Вид практики": LineField(1, DocxEnumTag.KIND),
            "Тип практики": LineField(1, DocxEnumTag.AIM),
            "Курс": LineField(1, DocxEnumTag.GRADE),
            "Факультет": LineField(1, DocxEnumTag.FACULTY),
            "Группа": LineField(1, DocxEnumTag.GROUP),
            "Форма обучения": LineField(1, DocxEnumTag.STUDY_TYPE),
            "Специализация": LineField(1, DocxEnumTag.SPECIALIZATION),
            "Период практики (годы)": LineField(1, DocxEnumTag.PERIOD_YEARS),
            "Период практики (дни)": LineField(1, DocxEnumTag.PERIOD_DATE),
            "Кафедра": LineField(1, DocxEnumTag.PULPIT),
            "Должность руководителя практики": LineField(1, DocxEnumTag.DIRECTOR),
            "ФИО руководителя практики": LineField(1, DocxEnumTag.DIRECTOR_NAME),
            "Группа организаций. Имя организации. ФИО студентов, форма обучения":
                MultiField(2, DocxEnumTag.TABLES),
        }

    def __iter__(self):
        return iter(self.columns.values())

    def help_iter(self) -> Iterator[tuple[str, Field]]:
        return iter(self.columns.items())

    def get_field(self, xlsx_field: str) -> Field | None:
        return self.columns.get(xlsx_field)

    def get(self, tag: DocxEnumTag) -> Field | None:
        for f in self.columns.values():
            if f.owner == tag:
                return f

    def get_unset_fields(self) -> tuple[tuple[str, Field]]:
        """ Возвращает кортэж из каноничного имени поля и самого поля. """
        return tuple((s, field) for s, field in self.columns.items() if field.value is None)


def check_filled(data: XlsxData, doc: TaggedDoc):
    """ Проверяет, все ли необходимые данные заполнены в xlsx. """
    if unset := data.get_unset_fields():
        unset_tags = tuple(field[1].owner for field in unset)  # распаковка только enum тэгов.
        used_tags = doc.get_used_tags()  # получение использованных enum тегов в шаблоне docx.
        for tag in used_tags:
            if tag in unset_tags:
                err = "\n".join(field[0] for field in unset)
                raise UnsetFieldError(f'Недостаточно данных для формирования docx документа, '
                                      f'следующие поля должны быть установлены: \n{err}')


def fill_tables(doc: TaggedDoc, tag: DocxEnumTag, xl_data: XlsxData):
    """
    Заполняет все таблицы данными и вставляет в документ.

    :param doc: документ.
    :param tag: имя тэга с которого начать вставлять таблицы.
    :param xl_data: хранилище данных.
    :return:
    """
    def _new_p(after: Paragraph):
        """  Создание нового параграфа сразу поле предыдущего. """
        new_p = doc._d.add_paragraph()
        after._p.addnext(new_p._p)
        return new_p

    def _add_student(_name: str, _budget: str):
        nonlocal n_students
        table.add_row(f'{n_students}. {_name}', f'{group}, {_budget}', director, director_name)
        n_students += 1

    table_paragraph = doc._hit_paragraphs[tag].pop()  # параграф, в котором найден тэг таблицы.
    paragraph = _new_p(table_paragraph)  # вставка нового неформатированного параграфа.
    table_paragraph._element.getparent().remove(table_paragraph._p)  # удаление параграфа с тэгом.

    director = xl_data.get(DocxEnumTag.DIRECTOR).value  # должность преподователя
    director_name = xl_data.get(DocxEnumTag.DIRECTOR_NAME).value  # фио преподователя
    group = xl_data.get(DocxEnumTag.GROUP).value  # номер группы студентов.

    n_students = 1  # нумерация студентов.
    n_org = 1  # нумерация органицаций.
    for table_data in xl_data.get(DocxEnumTag.TABLES).value:  # для каждой группы данных таблицы.
        table = DocxTable(paragraph, width=doc.width, columns_width=(Inches(5),))  # создание экземпляра таблицы
        table.make_base_headings()  # создание шапки.

        name, _ = table_data[0]
        if _ is None:
            paragraph.add_run(f'{n_org}. {name}\n')
            paragraph.style = 'Heading 4'
        else:
            _add_student(table_data[0][0], table_data[0][1])  # имя организации опущено, первый кортеж - студент

        n_sub_org = 1  # нумерация филиалов органицаций.
        for line in table_data[1:]:  # для каждой строки
            name = line[0]
            budget = line[1]
            if budget and all(filter(lambda x: x[0].isupper(), name.strip().split(' '))):  # встретилось имя студента
                _add_student(name, budget)
            else:
                r = table.add_row(f'{n_org}.{n_sub_org}. {name}', p_style='Heading 6')  # встретилось имя филиала
                n_sub_org += 1
                table.merge((r,), cells=(0, table.cols-1))
        n_org += 1
        paragraph = _new_p(paragraph)
        table.apply()


class ListTagsAction(argparse.Action):

    def __init__(self, option_strings, dest, nargs=None, **kwargs):
        super().__init__(option_strings, dest, nargs=0, **kwargs)

    def __call__(self, parser, namespace, values, option_string=None):
        for canon, field in UsurtData().help_iter():
            print(f'Xlsx поле: "{canon}", тэг docx: "<{field.owner.value}>"')
        exit(0)


def main():
    logger.remove()
    logger.add(sys.stdout, colorize=True, format="<level>{level}</level> | <level>{message}</level>")
    parser = argparse.ArgumentParser()
    parser.add_argument('docx', type=str, help='путь до шаблона docx документа.')
    parser.add_argument('xlsx', type=str, help='путь до xlsx документа с данными.')
    parser.add_argument('-o', '-out', type=str, help='путь до нового docx документа.')
    parser.add_argument('-lt', '-list-tags', action=ListTagsAction, help='отобразить список доступных тэгов.')

    args = parser.parse_args()
    xlsx_path = Path(args.xlsx)
    docx_path = Path(args.docx)
    out = Path(args.o) if args.o else Path(f'{docx_path.stem}-prepared.docx')

    try:
        xl_data = XlsxDataParser(xlsx_path).parse(UsurtData)
        doc = TaggedDoc(docx_path, init=True)
        check_filled(xl_data, doc)
    except (XlsxDataParserError, TaggedDocError, UnsetFieldError) as e:
        logger.error(e)
        exit(1)

    for field in xl_data:
        match field.owner:
            case DocxEnumTag.TABLES:
                fill_tables(doc, field.owner, xl_data)
            case _:
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

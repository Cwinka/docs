import argparse
import sys
from pathlib import Path

from docx.document import Document as HintDocument
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from loguru import logger

from docparser import TaggedDoc, DocxEnumTag, UnknownDueDate
from interfaces import XlsxData, LineField, UnsetFieldError, MultiField, UsurtBaseTable, Field
from xlsxparser import XlsxDataParser


class UsurtData(XlsxData):

    def __init__(self):
        self.columns: dict[str, LineField] = {
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

    def get_field(self, xlsx_field: str) -> LineField | None:
        return self.columns.get(xlsx_field)

    def get(self, tag: DocxEnumTag) -> Field | None:
        for f in self.columns.values():
            if f.owner == tag:
                return f

    def get_unset_fields(self) -> tuple[tuple[str, LineField]]:
        """ Возвращает кортэж из каноничного имени поля и самого поля. """
        return tuple((s, field) for s, field in self.columns.items() if field.value is None)


def check_filled(data: XlsxData, doc: TaggedDoc):
    if unset := data.get_unset_fields():
        unset_tags = tuple(field[1].owner for field in unset)
        used_tags = doc.get_used_tags()
        for tag in used_tags:
            if tag in unset_tags:
                err = "\n".join(field[0] for field in unset)
                raise UnsetFieldError(f'Недостаточно данных для формирования docx документа, '
                                      f'следующие поля должны быть установлены: \n{err}')


def fill_tables(doc: TaggedDoc, tag: DocxEnumTag, xl_data: XlsxData):
    def _new_p(after: Paragraph):
        # new_p = OxmlElement("w:p")
        new_p = doc._d.add_paragraph()
        after._p.addnext(new_p._p)
        return new_p

    paragraphs = doc._found_p[tag]
    director = xl_data.get(DocxEnumTag.DIRECTOR).value
    director_name = xl_data.get(DocxEnumTag.DIRECTOR_NAME).value
    group = xl_data.get(DocxEnumTag.GROUP).value

    found = paragraphs.pop()
    p = _new_p(found)
    found._element.getparent().remove(found._p)

    # print()
    # paragraph_styles = [s for s in doc._d.styles if s.type == WD_STYLE_TYPE.RUNS]
    # print(paragraph_styles)

    students = 1
    for part in xl_data.get(DocxEnumTag.TABLES).value:
        t = UsurtBaseTable(p, width=doc.width, columns_width=(Inches(5),))
        t.make_base_headings()

        org_common_name, _ = part[0]
        if _ is None:
            p.add_run(f'{org_common_name}\n')
            p.style = 'Heading 4'

        for line in part[1:]:
            name, budget = line
            student_name = name.strip()
            if budget and all(filter(lambda x: x[0].isupper(), student_name.split(' '))):
                t.add_row(f'{students}. {name}', f'{group}, {budget}', director, director_name)
                students += 1
            else:
                r = t.add_row(name, p_style='Heading 6')
                t.merge((r,), cells=(0, t.cols-1))
        p = _new_p(p)
        t.apply()


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

    xl_data = XlsxDataParser(xlsx_path).parse(UsurtData)
    doc = TaggedDoc(docx_path, init=True)

    try:
        check_filled(xl_data, doc)
    except UnsetFieldError as e:
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

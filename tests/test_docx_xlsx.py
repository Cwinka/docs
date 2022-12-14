import pytest

from docparser import TaggedDoc, UnknownDueDate
from interfaces import DocxEnumTag
from xlsxparser import XlsxDataParser, TagData
from tests.test_docx import DOCX_RESOURCE, DOCX_RESOURCE_BAD, flat_docx, save_path
from tests.test_xlsx import XLSX_RESOURCE, XLSX_RESOURCE_BAD


def test_replace(save_path):
    data = XlsxDataParser(XLSX_RESOURCE).parse(TagData)
    doc = TaggedDoc(DOCX_RESOURCE, init=True)

    for field in data:
        match field.owner:
            case DocxEnumTag.TABLES:
                pass
            case _:
                doc.replace_tag(field.owner, field.value)

    doc.save(save_path)
    assert save_path.exists()
    text = flat_docx(save_path)
    for field in data:
        if field.owner == DocxEnumTag.TABLES:
            continue
        assert text.find(field.owner.value) == -1, f'Найден незаменный тэг "{field.owner}" в тексте.'


def test_bad_replace(save_path):
    data = XlsxDataParser(XLSX_RESOURCE_BAD).parse(TagData)
    doc = TaggedDoc(DOCX_RESOURCE_BAD, init=True)

    for field in data:
        match field.owner:
            case DocxEnumTag.TABLES:
                pass
            case DocxEnumTag.PERIOD_DAYS:
                with pytest.raises(UnknownDueDate):
                    doc.replace_tag(field.owner, field.value)
            case _:
                doc.replace_tag(field.owner, field.value)

    doc.save(save_path)
    assert save_path.exists()
    text = flat_docx(save_path)
    for field in data:
        if field.owner in (DocxEnumTag.TABLES, DocxEnumTag.PERIOD_DAYS):
            continue
        assert text.find(field.owner.value) == -1, f'Найден незаменный тэг "{field.owner}" в тексте.'


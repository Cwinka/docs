import pytest
from pathlib import Path
from xlsxparser import XlsxDataParser, TagData, XlsxData, Field, XlsxDataParserError
from interfaces import DocxEnumTag

XLSX_RESOURCE = Path('tests/samples/s1.xlsx')
XLSX_RESOURCE_BAD = Path('tests/samples/s2.xlsx')
XLSX_RESOURCE_CORRUPT = Path('tests/samples/s3.xlsx')


def test_read():
    data = XlsxDataParser(XLSX_RESOURCE).parse(TagData)
    assert isinstance(data, TagData)
    assert isinstance(data, XlsxData)


def test_tags():
    data = XlsxDataParser(XLSX_RESOURCE).parse(TagData)
    v = data.get(DocxEnumTag.GROUP)
    assert DocxEnumTag.GROUP == v.owner
    assert isinstance(v.value, str)
    for field in data:
        assert isinstance(field, Field)
        assert isinstance(field.owner, DocxEnumTag)

    assert len(data.get_unset_fields()) == 0


def test_tag_help():
    t = TagData()

    for (canon, field) in t.help_iter():
        assert isinstance(canon, str)
        enc = canon.encode('utf8')
        assert isinstance(enc, bytes)
        assert canon == enc.decode('utf8')

        assert isinstance(field.rows, int | None)
        assert isinstance(field.columns, int)


def test_bad_tags():
    data = XlsxDataParser(XLSX_RESOURCE_BAD).parse(TagData)
    assert len(data.get_unset_fields()) > 0
    d = data.get(DocxEnumTag.PERIOD_DAYS)
    assert d.owner == DocxEnumTag.PERIOD_DAYS
    assert isinstance(d.value, str)

    for field in data:
        v = str(field.value)
        enc = v.encode('utf8')
        assert isinstance(enc, bytes)
        assert v == enc.decode('utf8')


def test_corrupt_xlsx():
    with pytest.raises(XlsxDataParserError):
        XlsxDataParser(Path('does not exist')).parse(TagData)
        XlsxDataParser(XLSX_RESOURCE_CORRUPT).parse(TagData)


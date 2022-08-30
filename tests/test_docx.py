import os

import pytest
from pathlib import Path
from docparser import TaggedDoc, TaggedDocError, DocxEnumTag, UnknownDueDate
import docx
from functools import reduce

DOCX_RESOURCE = Path('tests/samples/s1.docx')
DOCX_RESOURCE_BAD = Path('tests/samples/s2.docx')
DOCX_RESOURCE_CORRUPT = Path('tests/samples/s3.docx')


@pytest.fixture()
def save_path(request):
    def teardown():
        os.remove(path.as_posix())
    path = Path('.test_tmp_file')
    request.addfinalizer(teardown)
    return path


def flat_docx(path: Path) -> str:
    new_doc = docx.Document(path)
    return reduce(lambda x1, x2: f'{x1}\n{x2}', (p.text for p in new_doc.paragraphs))


def test_read():
    doc = TaggedDoc(DOCX_RESOURCE, init=True)
    assert isinstance(doc, TaggedDoc)
    assert doc.width > 0


def test_tags(save_path):
    doc = TaggedDoc(DOCX_RESOURCE, init=True)
    assert len(doc.get_used_tags()) != 0

    content = "SoMeStRangeString"
    doc.replace_tag(DocxEnumTag.GRADE, content)

    doc.save(save_path)
    assert save_path.exists()
    text = flat_docx(save_path)
    assert text.find(content) != -1


def test_tags_bad():
    doc = TaggedDoc(DOCX_RESOURCE_BAD, init=True)
    for t in doc.get_used_tags():
        match t:
            case DocxEnumTag.PERIOD_DAYS:
                with pytest.raises(UnknownDueDate):
                    doc.replace_tag(t, "Foo")
            case _:
                doc.replace_tag(t, "Foo")


def test_corrupt_file():
    with pytest.raises(TaggedDocError):
        TaggedDoc(Path('does not exist'), init=True)
        TaggedDoc(DOCX_RESOURCE_CORRUPT, init=True)

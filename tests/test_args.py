import pytest
import subprocess
from pathlib import Path
from docparser import TaggedDoc
from tests.test_docx import DOCX_RESOURCE
from tests.test_xlsx import XLSX_RESOURCE

MAIN = Path('filler.py')
PYTHON = Path('venv/Scripts/python.exe')


def run_main(*args: str):
    subprocess.check_call([PYTHON.as_posix(), MAIN.as_posix(), *args])


def test_proper_args(save_path):
    run_main(DOCX_RESOURCE.as_posix(), XLSX_RESOURCE.as_posix(), '-o', save_path.as_posix())

    assert save_path.exists()
    doc = TaggedDoc(save_path, init=True)
    assert len(doc.get_used_tags()) == 0


def test_invalid_args():
    args = ((DOCX_RESOURCE,), ('', DOCX_RESOURCE), (DOCX_RESOURCE, DOCX_RESOURCE),
            (XLSX_RESOURCE, DOCX_RESOURCE), (XLSX_RESOURCE, ""), ("", ""),
            ("sob", 'gob', "-o", 'bob'), ("", "", "-o", 'bob'), tuple())
    for arg in args:
        with pytest.raises(subprocess.CalledProcessError):
            run_main(*arg)

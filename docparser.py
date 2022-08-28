import re
from pathlib import Path
from enum import Enum
from morfeus import morf
from docx.document import Document as HintDocument
from docx.text.paragraph import Paragraph
from docx import Document
from collections import defaultdict


class DocxEnumTag(Enum):
    KIND = 'KIND'
    AIM = 'AIM'
    GRADE = 'GRADE'
    FACULTY = 'FACULTY'
    GROUP = 'GROUP'
    STUDY_TYPE = 'STUDY_TYPE'
    SPECIALIZATION = 'SPECIALIZATION'
    PERIOD_YEARS = 'PERIOD_YEARS'
    PERIOD_DATE = 'PERIOD_DAYS'
    PULPIT = 'PULPIT'
    DIRECTOR = 'DIRECTOR'
    DIRECTOR_NAME = 'DIRECTOR_NAME'
    STUDENTS = 'STUDENTS'
    TRAINING_TABLE = 'TRAINING_TABLE'


class _DocxTag:

    def __init__(self, enum: DocxEnumTag, due: str = None):
        self.enum = enum
        self.due = due
        self.value = enum.value
        self.name = enum.name

    @staticmethod
    def global_re() -> str:
        """ Возвращает регулярное выражение для поиска любого тэга в тексте. """
        return '<(?P<tag>[A-Z_]+)(:(?P<due>[a-z]+))?>'

    def local_re(self) -> str:
        """ Возвращает регулярное выражение для поиска тэга в тексте. """
        return f'<(?P<tag>{self.value})(:(?P<due>[a-z]+))?>'

    def replace_re(self) -> str:
        """ Возвращает регулярное выражение для замены тэга в тексте. """
        if due := self.due:
            return f'<{self.value}:{due}>'
        return f'<{self.value}>'

    @classmethod
    def from_re(cls, tag: re.Match) -> '_DocxTag':
        e = DocxEnumTag(tag.group('tag'))
        return cls(e, due=tag.group('due'))


class DocsReplaceContent:

    def replace(self, docx: HintDocument):
        """
        Метод вызывается когда тэг заменяется на содержимое данного класса.
        Метод должен расположить необходимое содержимое в документе.
        """
        pass


class UnknownDueDate(Exception):
    pass


class TaggedDoc:
    def __init__(self, path: Path, init: bool = False):
        self._path = path
        self._d: HintDocument = Document(path)
        self._found_tags: dict[DocxEnumTag, list[_DocxTag]] = defaultdict(list)
        self._found_p: dict[DocxEnumTag: set[Paragraph]] = defaultdict(set)
        if init:
            self._parse()

    def save(self, path: Path):
        self._d.save(path)

    def get_used_tags(self) -> list[DocxEnumTag]:
        """
        Возвращает тэги, которые используются в документе.
        :return:
        """
        return list(self._found_tags)

    def replace_tag(self, tag: DocxEnumTag, content: str):
        """ Заменяет все tag внутри документа на content (в соответсвующем падеже) """
        par = self._found_p[tag]
        tags = self._found_tags[tag]
        for p in par:
            for t in tags:
                try:
                    due_text = morf(str(content), t.due).strip()
                except ValueError:
                    raise UnknownDueDate(f'Неизвестный падеж в тэге "{t.name}": "{t.due}".')
                self._replace_text(p, t.replace_re(), due_text)

    @staticmethod
    def _replace_text(paragraph: Paragraph, regex: str, replace_str: str):
        """ Заменяет все вхождения regex на replace_str внутри paragraph """
        while True:
            match = re.search(regex, paragraph.text)
            if not match:
                break
            runs = iter(paragraph.runs)
            start, end = match.start(), match.end()
            for run in runs:
                run_len = len(run.text)
                if start < run_len:
                    break
                start, end = start - run_len, end - run_len
            run_text = run.text
            run.text = f'{run_text[:start]}{replace_str}{run_text[end:]}'
            end -= len(run_text)
            for run in runs:
                if end <= 0:
                    break
                run_text = run.text
                run.text = run_text[end:]
                end -= len(run_text)

            # for run in paragraph.runs:
            #     if run.text == "":
            #         r = run._r
            #         r.getparent().remove(r)

    def _parse(self):
        search_pattern = _DocxTag.global_re()
        for p in self._d.paragraphs:
            if found := re.finditer(search_pattern, p.text):
                for tag in found:
                    t = _DocxTag.from_re(tag)
                    self._found_p[t.enum].add(p)
                    self._found_tags[t.enum].append(t)

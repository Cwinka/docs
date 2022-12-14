import re
from collections import defaultdict
from pathlib import Path
from loguru import logger
from docx import Document
from docx.document import Document as HintDocument
from docx.shared import Length
from docx.text.paragraph import Paragraph
from interfaces import raise_invalid_path, DocxEnumTag
from morfeus import morf


class _DocxTag:
    """
    Сложный тэг в шаблоне docx, хранит информацию о enum и падеже, в который необходимо
    поставить предложение перед заменой.
    """

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


class UnknownDueDate(Exception):
    """ Неизвестный падеж. """
    pass


class TaggedDocError(Exception):
    pass


class TaggedDoc:
    def __init__(self, path: Path, init: bool = False):
        raise_invalid_path(path, TaggedDocError, exts=('.docx',))
        self._path = path  # Путь до шаблона docx.
        try:
            self._d: HintDocument = Document(path)  # Объект библиотеки python-docx.
        except ValueError:
            raise TaggedDocError(f'Неподходящий формат документа {path}. Необходим документ в формате docx.')
        except OSError:
            raise TaggedDocError(f'Документ {path} не является docx документом.')
        self.width: Length = self._d._block_width  # Ширина документа в относительных единицах.
        # Маппинг найденных enum тэгов на структуры тэгов, хранящие
        # дополнительную информацию об использовании тэга.
        self._found_tags: dict[DocxEnumTag, list[_DocxTag]] = defaultdict(list)
        #  Маппинг найденных enum тегов на список параграфов, в которых встречаются найденные тэги.
        self._hit_paragraphs: dict[DocxEnumTag: set[Paragraph]] = defaultdict(set)
        if init:
            self.parse()

    def parse(self):
        """
        Анализирует документ на наличие тэгов.
        :return:
        """
        self._clear()
        search_pattern = _DocxTag.global_re()  # Регулярное выражения для поиска тэгов.
        for p in self._d.paragraphs:
            if found := re.finditer(search_pattern, p.text):
                for tag in found:
                    try:
                        t = _DocxTag.from_re(tag)  # Создание экземпляра сложного тэга из строки.
                    except ValueError:
                        logger.warning(f'Найден несуществующий тэг "{tag.group(0)}" в параграфе "{p.text}".')
                    else:
                        self._hit_paragraphs[t.enum].add(p)  # Сопоставление enum и параграфа где найден тэг.
                        self._found_tags[t.enum].append(t)  # Сопоставление enum и со сложным тэгом.

    def _clear(self):
        self._found_tags.clear()
        self._hit_paragraphs.clear()

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
        paragraphs = self._hit_paragraphs[tag]
        for p in paragraphs:
            for t in self._found_tags[tag]:
                try:
                    due_content = morf(str(content), t.due).strip()  # приведение content в нужный падеж
                except ValueError:
                    raise UnknownDueDate(f'Неизвестный падеж в тэге "{t.name}": "{t.due}".')
                self._replace_text(p, t.replace_re(), due_content)  # замена тэга на content

    @staticmethod
    def _replace_text(paragraph: Paragraph, regex: str, replace_str: str):
        """ Заменяет все вхождения regex на replace_str внутри paragraph """
        while True:
            match = re.search(regex, paragraph.text)
            if not match:
                break
            runs = iter(paragraph.runs)  # итератор по блокам параграфа
            start, end = match.start(), match.end()  # начало и конец совпадения регулярного выражения
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

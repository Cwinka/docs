import re

import pymorphy2
from loguru import logger


_morf = pymorphy2.MorphAnalyzer(lang='ru')


def morf(text: str, due_date: str | None) -> str:
    """
    Изменяет текст в нужный падеж.

    :param text: текст для изменения.
    :param due_date: граммема (падеж).
    :return:
    """
    if due_date is None:
        return text
    proper = ''
    for cleared, orig in _splitter(text):
        inf = _inflect(cleared, due_date)
        if inf.lower() == cleared.lower():
            inf = cleared
        proper += re.sub(cleared, inf, orig) + ' '
    return proper


def _splitter(text: str) -> tuple[str, str]:
    """
    Разделяет текст на слова.

    :param text: текст для разделения на слова.
    :return: очищенное слово без знаков, оригинальное слово.
    """
    for part in text.split(' '):
        if not part:
            continue
        yield part.strip('()'), part


def _inflect(word: str, target: str) -> str:
    """
    Приводит слово в нужный падеж.
    Метод не обрабатыет слова, написанные в верхнем регистре.
    Метод не обрабатыет слова, длиной меньше 3 символов.

    :param word: слово.
    :param target: граммема.
    :return: слово поставленное в указаннай падеж.
    """
    if word.isupper() or len(word) < 3:
        return word
    m = _morf.parse(word)[0]
    inf = m.inflect({target})
    if inf is None:
        logger.warning(f'Не удалось привести слово "{word}" к таргету "{target}".')
        return word
    return inf.word

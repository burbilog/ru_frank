# -*- coding: utf-8 -*-
# furigana_convert.py — «…（かな）」 → Writer-ruby (устойчивая)

import re
import uno

KANJI = r'\u4E00-\u9FFF'
KANA  = r'\u3040-\u309F\u30A0-\u30FFー'       # хира + ката + «ー»

PATTERN = re.compile(fr'([{KANJI}]+[{KANA}]*)[（(]([{KANA}]+)[）)]')

def is_kanji(ch: str) -> bool:
    return '\u4E00' <= ch <= '\u9FFF'


def furigana_convert(*args):
    doc = XSCRIPTCONTEXT.getDocument()  # noqa: F821
    if not doc.supportsService("com.sun.star.text.TextDocument"):
        raise RuntimeError("Откройте документ Writer")

    enum = doc.Text.createEnumeration()
    while enum.hasMoreElements():
        para = enum.nextElement()
        if not para.supportsService("com.sun.star.text.Paragraph"):
            continue

        # Обрабатываем все совпадения в обратном порядке, чтобы не сбить позиции
        matches = list(PATTERN.finditer(para.String))
        for m in reversed(matches):
            word, kana = m.groups()
            abs_start = m.start()

            # Применяем руби ко всему слову
            ruby_cur = doc.Text.createTextCursorByRange(para.getStart())
            ruby_cur.goRight(abs_start, False)
            ruby_cur.goRight(len(word), True)
            ruby_cur.RubyText = kana

            # Убираем «（kana）» или «(kana)»
            del_cur = doc.Text.createTextCursorByRange(para.getStart())
            del_cur.goRight(abs_start + len(word), False)  # Позиция перед скобкой
            del_cur.goRight(len(kana) + 2, True)  # Выделяем скобки с kana
            del_cur.String = ""

    doc.getCurrentController().getFrame().getContainerWindow().setFocus()

g_exportedScripts = (furigana_convert,)

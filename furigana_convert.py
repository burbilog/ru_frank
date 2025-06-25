# -*- coding: utf-8 -*-
# furigana_convert.py — «…（かな）」 → Writer-ruby (устойчивая)

import re
import uno

KANJI = r'\u4E00-\u9FFF'
KANA  = r'\u3040-\u309F\u30A0-\u30FFー'       # хира + ката + «ー»

PATTERN = re.compile(fr'([{KANJI}{KANA}]+)（([{KANA}]+)）')

def is_kanji(ch: str) -> bool:
    return '\u4E00' <= ch <= '\u9FFF'

def core_of(word: str) -> tuple[int, int]:
    """
    Возвращает (offset, length) той части word, которая начинается
    с ПОСЛЕДНЕЙ группы кандзи и захватывает хвостовую хирагану/катакану.
    """
    last_k = max(i for i, c in enumerate(word) if is_kanji(c))
    # идём влево, пока подряд идут кандзи
    start = last_k
    while start > 0 and is_kanji(word[start - 1]):
        start -= 1
    return start, len(word) - start

def furigana_convert(*args):
    doc = XSCRIPTCONTEXT.getDocument()
    if not doc.supportsService("com.sun.star.text.TextDocument"):
        raise RuntimeError("Откройте документ Writer")

    enum = doc.Text.createEnumeration()
    while enum.hasMoreElements():
        para = enum.nextElement()
        if not para.supportsService("com.sun.star.text.Paragraph"):
            continue

        while True:
            m = PATTERN.search(para.String)
            if not m:
                break

            word, kana  = m.groups()
            abs_start   = m.start()

            rel_off, core_len = core_of(word)

            # 1. ставим руби на «ядро»
            ruby_cur = doc.Text.createTextCursorByRange(para.getStart())
            ruby_cur.goRight(abs_start + rel_off, False)
            ruby_cur.goRight(core_len, True)
            ruby_cur.RubyText = kana

            # 2. убираем «（kana）»
            del_cur = doc.Text.createTextCursorByRange(ruby_cur.getEnd())
            del_cur.goRight(len(kana) + 2, True)
            del_cur.String = ""

    doc.getCurrentController().getFrame().getContainerWindow().setFocus()

g_exportedScripts = (furigana_convert,)

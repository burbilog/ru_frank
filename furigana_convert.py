# -*- coding: utf-8 -*-
# furigana_convert.py — «…（かな）」 → Writer-ruby (устойчивая)

import re
import uno

KANJI = r'\u4E00-\u9FFF'
KANA  = r'\u3040-\u309F\u30A0-\u30FFー'       # хира + ката + «ー»

PATTERN = re.compile(fr'([{KANJI}{KANA}]+)（([{KANA}]+)）')

def is_kanji(ch: str) -> bool:
    return '\u4E00' <= ch <= '\u9FFF'

def split_word_readings(word: str, kana: str) -> list[tuple[int, int, str]]:
    """
    Разбивает слово и чтение на отдельные части для каждой группы кандзи.
    Возвращает список (offset, length, reading) для каждой части.
    """
    result = []
    word_pos = 0
    kana_pos = 0
    
    while word_pos < len(word):
        if is_kanji(word[word_pos]):
            # Найти группу подряд идущих кандзи
            kanji_start = word_pos
            while word_pos < len(word) and is_kanji(word[word_pos]):
                word_pos += 1
            
            # Найти следующую хирагану/катакану после кандзи
            kana_start = word_pos
            while word_pos < len(word) and not is_kanji(word[word_pos]):
                word_pos += 1
            
            # Определить длину чтения для этой группы
            # Простая эвристика: распределяем чтение пропорционально
            kanji_count = word_pos - kanji_start
            if kanji_count == 1:
                # Для одного кандзи берем чтение до следующего кандзи или до конца
                reading_end = kana_pos
                while reading_end < len(kana):
                    # Ищем границу чтения (обычно 2-4 символа на кандзи)
                    if reading_end - kana_pos >= 4:
                        break
                    # Проверяем, не начинается ли следующее чтение
                    remaining_kanji = sum(1 for i in range(word_pos, len(word)) if is_kanji(word[i]))
                    remaining_kana = len(kana) - reading_end - 1
                    if remaining_kanji > 0 and remaining_kana / remaining_kanji < 1.5:
                        break
                    reading_end += 1
                
                if reading_end > len(kana):
                    reading_end = len(kana)
                    
                reading = kana[kana_pos:reading_end]
                kana_pos = reading_end
            else:
                # Для группы кандзи берем пропорциональную часть оставшегося чтения
                remaining_kanji = sum(1 for i in range(kanji_start, len(word)) if is_kanji(word[i]))
                remaining_kana = len(kana) - kana_pos
                reading_len = min(remaining_kana, max(2, remaining_kana * kanji_count // remaining_kanji))
                reading = kana[kana_pos:kana_pos + reading_len]
                kana_pos += reading_len
            
            result.append((kanji_start, word_pos - kanji_start, reading))
        else:
            word_pos += 1
    
    return result

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

            # Получаем все части слова с их чтениями
            parts = split_word_readings(word, kana)

            # Применяем руби для каждой части (в обратном порядке)
            for rel_off, part_len, reading in reversed(parts):
                if reading:  # Только если есть чтение
                    ruby_cur = doc.Text.createTextCursorByRange(para.getStart())
                    ruby_cur.goRight(abs_start + rel_off, False)
                    ruby_cur.goRight(part_len, True)
                    ruby_cur.RubyText = reading

            # Убираем «（kana）»
            del_cur = doc.Text.createTextCursorByRange(para.getStart())
            del_cur.goRight(m.end() - 1, False)  # Позиция перед '）'
            del_cur.goRight(1, True)  # Выделяем '）'
            del_cur.String = ""
            
            del_cur = doc.Text.createTextCursorByRange(para.getStart())
            del_cur.goRight(abs_start + len(word), False)  # Позиция перед '（'
            del_cur.goRight(len(kana) + 1, True)  # Выделяем '（kana'
            del_cur.String = ""

    doc.getCurrentController().getFrame().getContainerWindow().setFocus()

g_exportedScripts = (furigana_convert,)

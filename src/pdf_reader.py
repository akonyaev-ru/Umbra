"""Извлечение содержимого PDF для анонимизации.

Здесь используются pdfplumber/pdfminer.six (MIT) и pypdf (BSD-3) — они дают
ТЕКСТ И ТАБЛИЦЫ, по которым строится карта замен. Физическую закраску делает
PyMuPDF (fitz) в doc_processor._write_redacted_pdf; проект из-за этого
распространяется под AGPL v3 (см. README). Прежний вариант этого комментария
запрещал PyMuPDF ради проприетарной поставки — он устарел.

Выход PDF-контура — `[ANON] <имя>.pdf` с физической закраской, а НЕ .docx и не
.txt, как было в ранних версиях.

PDF — разметка вывода, а не структурированный текст: абзацы и таблицы здесь не
читаются готовыми, их приходится ВОССТАНАВЛИВАТЬ по координатам. Контур:
PDF → блоки (абзацы + таблицы) → карта замен → физическая закраска исходного
PDF → '[ANON] <имя>.pdf'.

Склейка строк в абзацы — не косметика, а безопасность: NER, глядя на одну
строку, не видит ФИО, разорванное переносом ('Иванов Иван\nИванович' или
'Ива-\nнов'), и такие ПДн уходили бы в ИИ необезличенными.
"""

import collections
import re

MAX_PDF_PAGES = 1_000

# Разрыв абзаца: шаг между строками заметно больше обычного межстрочного.
# 1.6 подобрано эмпирически — ниже рвёт нумерованные списки, выше склеивает
# соседние абзацы в один.
_PARAGRAPH_GAP_RATIO = 1.6

# Перенос слова в конце строки: буква + дефис.
_HYPHEN_END_RE = re.compile(r'\w-$')


def _join_lines(parts):
    """Склеивает строки одного абзаца, разбирая переносы слов."""
    text = ''
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if not text:
            text = part
        elif _HYPHEN_END_RE.search(text) and part[:1].islower():
            # 'Ива-' + 'нов' → 'Иванов'. Только перед строчной буквой: в
            # 'ООО «Ромашка» -' + 'Исполнитель' дефис настоящий, не перенос.
            text = text[:-1] + part
        else:
            text += ' ' + part
    return text


def _leading(steps):
    """Оценка обычного межстрочного шага — САМЫЙ ЧАСТЫЙ шаг, а не медиана.

    Медиана здесь обманывает: если абзацы короткие, разрывов почти столько же,
    сколько обычных строк, и медиана садится ровно между ними — тогда ни один
    разрыв не превышает порога и весь текст слипается в один абзац. Самый
    частый шаг — это и есть межстрочный интервал документа.

    При равенстве частот берём МЕНЬШИЙ: перепутать межстрочный интервал с
    разрывом абзаца — значит не разорвать там, где надо; обратная ошибка
    (лишний разрыв) хуже, она может рассечь ФИО пополам, и NER его не увидит.
    """
    if not steps:
        return 0
    counts = collections.Counter(round(step) for step in steps)
    top_count = max(counts.values())
    return min(step for step, count in counts.items() if count == top_count)


def _paragraphs(page):
    """Абзацы страницы: [(top, text)]. Строки группируются по зазорам."""
    lines = page.extract_text_lines(layout=False, strip=True, return_chars=False)
    lines = [line for line in lines if line.get('text', '').strip()]
    if not lines:
        return []

    steps = [after['top'] - before['top'] for before, after in zip(lines, lines[1:])]
    leading = _leading(steps)

    out = []
    buffer = []
    top = lines[0]['top']
    for index, line in enumerate(lines):
        if buffer and leading > 0:
            step = line['top'] - lines[index - 1]['top']
            if step > leading * _PARAGRAPH_GAP_RATIO:
                out.append((top, _join_lines(buffer)))
                buffer = []
                top = line['top']
        buffer.append(line['text'])
    if buffer:
        out.append((top, _join_lines(buffer)))
    return [(t, text) for t, text in out if text.strip()]


def _page_blocks(page):
    """Блоки одной страницы в порядке чтения (сверху вниз)."""
    items = []
    tables = page.find_tables()
    boxes = [table.bbox for table in tables]

    for table in tables:
        rows = [[(cell or '').strip() for cell in row] for row in table.extract()]
        if any(any(cell for cell in row) for row in rows):
            items.append((table.bbox[1], {'type': 'table', 'rows': rows}))

    def outside_tables(obj):
        # Слова внутри рамки таблицы уже попали в её ячейки — иначе текст
        # задвоился бы: один раз в таблице, второй раз абзацем.
        x = (obj['x0'] + obj['x1']) / 2
        y = (obj['top'] + obj['bottom']) / 2
        return not any(x0 <= x <= x1 and top <= y <= bottom
                       for x0, top, x1, bottom in boxes)

    body = page.filter(outside_tables) if boxes else page
    for top, text in _paragraphs(body):
        items.append((top, {'type': 'paragraph', 'text': text}))

    items.sort(key=lambda item: item[0])
    return [block for _top, block in items]


def block_texts(blocks):
    """Плоский список всех текстов блоков в порядке обхода.

    Единая точка: и анонимизация, и сборка DOCX, и построение карты при
    восстановлении ходят по блокам ЧЕРЕЗ неё, поэтому порядок меток
    ([ФИО_1], [ФИО_2], ...) совпадает по построению, а не по совпадению.
    """
    texts = []
    for block in blocks:
        if block['type'] == 'table':
            for row in block['rows']:
                texts.extend(row)
        else:
            texts.append(block['text'])
    return texts


def extract_blocks(path):
    """PDF → список блоков: {'type':'paragraph','text':str} и
    {'type':'table','rows':[[str,...],...]} в порядке чтения.
    Пустой список, если текста нет (скан без OCR — об этом предупреждает
    вызывающий)."""
    import pdfplumber
    from pypdf import PdfReader

    # Проверки — на pypdf: он парсит лениво, а сообщения об ошибках уже
    # выверены. Тяжёлый разбор раскладки делает только pdfplumber ниже.
    reader = PdfReader(path)
    if len(reader.pages) > MAX_PDF_PAGES:
        raise ValueError(f'В PDF слишком много страниц (максимум {MAX_PDF_PAGES}).')
    if reader.is_encrypted:
        # Пробуем пустой пароль (часто PDF «зашифрован» без пароля владельца).
        # decrypt возвращает PasswordType (0/NOT_DECRYPTED = неудача) — проверяем
        # результат, иначе дальше молча извлеклась бы пустота вместо явной ошибки.
        try:
            ok = reader.decrypt('')
        except Exception:
            import logging
            logging.warning("Failed to decrypt PDF with empty password", exc_info=True)
            ok = False
        if not ok:
            raise ValueError('PDF защищён паролем — снимите защиту и повторите.')

    blocks = []
    failed_pages = []
    with pdfplumber.open(path, password='') as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            try:
                blocks.extend(_page_blocks(page))
            except Exception:
                import logging
                logging.error(f"Failed to extract blocks from page {page_number}", exc_info=True)
                failed_pages.append(page_number)
            finally:
                # Иначе pdfplumber держит разобранную раскладку всех страниц
                # в памяти — на многостраничном деле это сотни мегабайт.
                try:
                    page.flush_cache()
                except Exception:
                    import logging
                    logging.warning("Failed to flush page cache", exc_info=True)
    if failed_pages:
        preview = ', '.join(map(str, failed_pages[:10]))
        more = '…' if len(failed_pages) > 10 else ''
        raise ValueError(f'Не удалось извлечь текст со страниц: {preview}{more}.')
    return blocks


def has_extractable_text(path):
    """True, если из PDF извлекается непустой текст (иначе это скан без OCR)."""
    return any(text.strip() for text in block_texts(extract_blocks(path)))

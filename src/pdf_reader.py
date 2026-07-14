"""Извлечение текста из PDF для анонимизации.

Библиотека — pypdf (BSD-3, чистый Python): безопасна для проприетарной
поставки, в отличие от PyMuPDF/fitz (AGPL — требует раскрытия исходников при
распространении, несовместимо с обфускацией/лицензированием Umbra).

PDF — это разметка вывода, а не структурированный текст: абзацы/таблицы точно
не восстановить. Поэтому контур такой: PDF → извлечённый текст → анонимизация →
`[ANON].txt` (не обратно в PDF). Деанонимизация ответа ИИ идёт в .txt.
"""

import re


def extract_lines(path):
    """Возвращает список строк текста PDF (по одной на визуальную строку,
    с сохранением абзацных разрывов). Пустой список, если текста нет
    (например, скан без OCR — об этом предупреждает вызывающий)."""
    from pypdf import PdfReader

    reader = PdfReader(path)
    if reader.is_encrypted:
        # Пробуем пустой пароль (часто PDF «зашифрован» без пароля владельца).
        # decrypt возвращает PasswordType (0/NOT_DECRYPTED = неудача) — проверяем
        # результат, иначе дальше молча извлеклась бы пустота вместо явной ошибки.
        try:
            ok = reader.decrypt('')
        except Exception:
            ok = False
        if not ok:
            raise ValueError('PDF защищён паролем — снимите защиту и повторите.')

    out = []
    for page in reader.pages:
        try:
            text = page.extract_text() or ''
        except Exception:
            text = ''
        for line in text.split('\n'):
            out.append(line.rstrip())
        out.append('')   # разрыв между страницами
    # Схлопываем хвостовые пустые строки.
    while out and not out[-1].strip():
        out.pop()
    return out


def has_extractable_text(path):
    """True, если из PDF извлекается непустой текст (иначе это скан без OCR)."""
    lines = extract_lines(path)
    return any(l.strip() for l in lines)

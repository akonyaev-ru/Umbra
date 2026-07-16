"""Копирование результата анонимизации в буфер обмена.

Windows: сразу в ДВУХ форматах — файл (CF_HDROP) и его текст (CF_UNICODETEXT).
Целевое приложение само берёт нужный формат: область загрузки файлов / письмо /
папка — получит файл, поле чата ИИ — вставит текст. Один клик покрывает оба
сценария отправки в ИИ.

macOS/Linux: копируется только ТЕКСТ (через буфер tkinter) — положить файл в
буфер средствами Tk нельзя. Поэтому pywin32 импортируется строго внутри
windows-ветки: на других ОС его нет, и импорт на уровне модуля ронял бы запуск.
"""

import os
import sys
import struct


def _read_text_file(path):
    """Читает текст терпимо к кодировке (свои файлы пишем в UTF-8, но ответ мог
    быть сохранён и в cp1251)."""
    raw = open(path, 'rb').read()
    try:
        return raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        return raw.decode('cp1251', errors='replace')


def extract_text(path):
    """Текст результата для вставки в поле чата ИИ. .md/.txt/.csv/.html — как есть;
    .docx — тело + таблицы; .xlsx — значения ячеек (форматирование в буфере всё
    равно не нужно). Возвращает '' только если текста действительно нет."""
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.md', '.txt', '.csv', '.html'):
        return _read_text_file(path)
    if ext == '.docx':
        from docx import Document
        doc = Document(path)
        parts = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    parts.append(cell.text)
        return '\n'.join(p for p in parts if p.strip())
    if ext == '.xlsx':
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        lines = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                cells = [str(v) for v in row if v is not None and str(v).strip()]
                if cells:
                    lines.append('\t'.join(cells))
        return '\n'.join(lines)
    return ''


def _hdrop_blob(paths):
    """DROPFILES-структура для CF_HDROP: заголовок 20 байт + список путей в
    UTF-16-LE с двойным нулём на конце."""
    # DROPFILES { DWORD pFiles=20; POINT pt(0,0); BOOL fNC=0; BOOL fWide=1 }
    header = struct.pack('<IiiII', 20, 0, 0, 0, 1)
    files = '\0'.join(os.path.abspath(p) for p in paths) + '\0\0'
    return header + files.encode('utf-16-le')


def copy_result(path, tk_master=None):
    """На Windows кладёт файл (CF_HDROP) и его текст (CF_UNICODETEXT) в буфер обмена.
    На Mac/Linux копирует только текст через tkinter.

    Возвращает True, только если в буфер РЕАЛЬНО что-то положено. Врать об успехе
    нельзя: иначе пользователь жмёт «Вставить» в чате ИИ и молча получает пустоту.
    Исключения пробрасываются вызывающему."""
    text = extract_text(path)

    if sys.platform == 'win32':
        import win32clipboard
        blob = _hdrop_blob([path])
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_HDROP, blob)
            if text:
                win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
        finally:
            win32clipboard.CloseClipboard()
        # Файл в буфере есть всегда — успех не зависит от наличия текста.
        return True

    # macOS/Linux: класть в буфер нечего, если нет окна Tk или текст пуст.
    if tk_master is None or not text:
        return False
    tk_master.clipboard_clear()
    tk_master.clipboard_append(text)
    tk_master.update()  # Required for tkinter clipboard to register
    return True

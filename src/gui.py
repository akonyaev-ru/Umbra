import os
import sys
import threading
import subprocess
import webbrowser
import json
from pathlib import Path
import customtkinter as ctk
from tkinter import messagebox, filedialog
from PIL import Image
import naming
import ctypes
import pypdf
from docx import Document
from nlp_engine import NLPProcessor
from md_merge import merge_ai_md
from tkinterdnd2 import TkinterDnD, DND_FILES

class TkinterDnD_CTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

try:
    # Per-Monitor V2 DPI awareness (-4 = DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2).
    ctypes.windll.user32.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
    ctypes.windll.user32.SetProcessDpiAwarenessContext.restype = ctypes.c_bool
    ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))

    # Fix for blurry Taskbar icons in Windows (groups app uniquely)
    myappid = 'com.antigravity.umbra.v1'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    import logging
    logging.warning("Failed to set AppID", exc_info=True)

# Следуем системной теме Windows (светлая/тёмная).
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("green")


def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# Шрифт заголовка «U M B R A» — Audiowide (решение владельца 2026-07-10;
# заменил FIFA Welcome). Файл Audiowide-Regular.ttf (Google Fonts, лицензия OFL)
# вшит в поставку (--add-data в build.py) и грузится ПРИВАТНО через GDI
# AddFontResourceExW(FR_PRIVATE): шрифт видит только этот процесс, установка в
# систему не нужна → одинаково работает на любом ПК, куда раздали .exe.
# Грузить надо ДО создания CTkFont. При любой ошибке молча остаёмся на
# системном шрифте (заголовок — не функциональность).
TITLE_FONT_FAMILY = "Audiowide"
try:
    import ctypes as _ct
    _gdi32 = _ct.windll.gdi32
    # Явные argtypes (см. §10.3 AGENTS.md): без них указатели на 64-битной
    # Windows могут усекаться и вызов молча не срабатывает.
    _gdi32.AddFontResourceExW.argtypes = [_ct.c_wchar_p, _ct.c_uint32, _ct.c_void_p]
    _gdi32.AddFontResourceExW.restype = _ct.c_int
    _FR_PRIVATE = 0x10
    _gdi32.AddFontResourceExW(get_resource_path('Audiowide-Regular.ttf'), _FR_PRIVATE, None)
except Exception:
    import logging
    logging.warning("Failed to load custom font", exc_info=True)


# --- Палитра (кортежи (светлая, тёмная) — CustomTkinter сам выбирает по теме) ---
BG_COLOR = ("#F8FAFC", "#1B1B1D")        # фон окна
FRAME_COLOR = ("#FFFFFF", "#242427")      # карточки/панели
BORDER_COLOR = ("#E2E8F0", "#3A3A3E")     # границы
ACCENT_COLOR = "#6D529F"                   # заливка кнопок (фиолетовый читается в обеих темах)
ACCENT_HOVER = "#563E81"
ACCENT_TEXT = ("#513A7E", "#9A81D2")       # фиолетовый ТЕКСТ с достаточным контрастом
HOVER_COLOR = ("#F2EFF9", "#292434")       # ховер карточки
SELECTED_BG = ("#E6DDF2", "#31254A")       # выбранная карточка
SUCCESS_COLOR = ("#513A7E", "#9A81D2")
ERROR_COLOR = ("#DC2626", "#F87171")
WARNING_COLOR = ("#B45309", "#FBBF24")
SUPPORTED_EXTENSIONS = ('.docx', '.xlsx', '.txt', '.pdf', '.csv', '.html')
TEXT_COLOR = ("#0F172A", "#F1F5F9")        # основной текст
MUTED_COLOR = ("#64748B", "#94A3B8")       # приглушённый текст
PROCESSING_BTN = ("#C4B6DF", "#453466")    # видимая кнопка во время обработки


class DocumentCard(ctk.CTkFrame):
    def __init__(self, master, file_info, command, is_selected=False, **kwargs):
        bg_color = SELECTED_BG if is_selected else ("#F1F5F9", "#242427")
        border_col = ACCENT_COLOR if is_selected else ("#CBD5E1", "#3F3F46")
        super().__init__(master, height=48, fg_color=bg_color,
                         corner_radius=8, border_width=2,
                         border_color=border_col, **kwargs)
        self.grid_propagate(False)
        self.file_info = file_info
        self.command = command
        self.is_selected = is_selected

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Фрейм для текста (только имя файла)
        self.text_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.text_frame.grid(row=0, column=0, sticky="w", padx=(16, 12))

        # Имя файла
        self.name_lbl = ctk.CTkLabel(
            self.text_frame,
            text=file_info['name'],
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold" if is_selected else "normal"),
            text_color=TEXT_COLOR,
            justify="left",
            anchor="w",
            wraplength=440,
        )
        self.name_lbl.pack(anchor="w")

        bind_targets = [self, self.text_frame, self.name_lbl]

        for w in bind_targets:
            w.bind("<Button-1>", self.on_click)
            w.bind("<Enter>", self.on_enter)
            w.bind("<Leave>", self.on_leave)

    def on_enter(self, event):
        if not self.is_selected:
            self.configure(fg_color=HOVER_COLOR)

    def on_leave(self, event):
        if not self.is_selected:
            self.configure(fg_color=("#F1F5F9", "#242427"))

    def on_click(self, event):
        self.command(self.file_info['path'])


class UmbraApp(TkinterDnD_CTk):
    def __init__(self, process_callback, deanon_callback=None, cleanup_callback=None):
        super().__init__(fg_color=BG_COLOR)

        self.process_callback = process_callback
        self.deanon_callback = deanon_callback
        self.cleanup_callback = cleanup_callback
        self.files = []             # документы, которые дал пользователь (дроп/кнопка)
        self.selected_files = set()
        # Ставится в True только после успешной привязки windnd (конец __init__),
        # но читается из _render_cards — инициализируем здесь, до первой отрисовки.
        self.dnd_enabled = False
        self.last_out_paths = []
        self.last_out_path = None

        self.title("Umbra")
        self.geometry("720x540")
        self.minsize(680, 540)
        self.resizable(True, True)

        # Иконка окна/панели задач. iconbitmap ставим сразу как фолбэк, но
        # CustomTkinter после инициализации донастраивает окно и может сбросить
        # её, а сам Tk к тому же GDI-масштабирует один кадр .ico под HiDPI
        # (размыто). Поэтому после реализации окна повторно применяем чёткую
        # иконку Windows-нативно через WM_SETICON (см. _apply_crisp_icon).
        try:
            self.iconbitmap(get_resource_path('icon_purple.ico'))
        except Exception:
            import logging
            logging.warning("Failed to load window icon", exc_info=True)
        self.after(200, self._apply_crisp_icon)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)  # список тянется

        # --- ШАПКА (лого + название) ---
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, pady=(28, 20), padx=32, sticky="ew")
        self.header_frame.grid_columnconfigure(1, weight=1)

        try:
            logo_img = Image.open(get_resource_path('logo.png'))
            img_w, img_h = logo_img.size
            target_h = 64
            target_w = int((img_w / img_h) * target_h)
            self.logo_ctk = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(target_w, target_h))
            self.logo_label = ctk.CTkLabel(self.header_frame, image=self.logo_ctk, text="")
            self.logo_label.grid(row=0, column=0, rowspan=2, padx=(0, 16))
        except Exception:
            import logging
            logging.warning("Failed to load logo", exc_info=True)

        # Разрядка заголовка — ТОНКИМИ пробелами U+2009 (уже обычного): буквы
        # ближе друг к другу, но вордмарк остаётся разреженным. Записаны
        # escape-последовательностями — в исходнике нет невидимых символов.
        self.title_label = ctk.CTkLabel(
            self.header_frame,
            text="U\u2009M\u2009B\u2009R\u2009A",
            font=ctk.CTkFont(family=TITLE_FONT_FAMILY, size=24, weight="bold"),
            text_color=TEXT_COLOR
        )
        self.title_label.grid(row=0, column=1, sticky="w", pady=(8, 0))

        self.subtitle_label = ctk.CTkLabel(
            self.header_frame,
            text="Анонимизация юридических документов",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=ACCENT_TEXT
        )
        self.subtitle_label.grid(row=1, column=1, sticky="w")

        self.header_switch_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.header_switch_frame.grid(row=0, column=2, rowspan=2, sticky="e")

        self.md_mode = ctk.BooleanVar(value=False)
        self.md_switch = ctk.CTkSwitch(
            self.header_switch_frame,
            text="",
            width=36,
            variable=self.md_mode,
            progress_color=ACCENT_COLOR,
            button_color="#513A7E",
            button_hover_color="#3B2A5C",
            switch_height=18, switch_width=36,
        )
        self.md_switch.grid(row=0, column=0, sticky="e", padx=(0, 8), pady=(4, 0))

        self.md_label = ctk.CTkLabel(
            self.header_switch_frame,
            text="Умный конвертер",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=MUTED_COLOR
        )
        self.md_label.grid(row=0, column=1, sticky="e")
        self.md_label.bind("<Button-1>", lambda e: self.md_switch.toggle())

        # --- СПИСОК ОТКРЫТЫХ ДОКУМЕНТОВ ---
        # border_width=2: однопиксельная граница на HiDPI почти не читалась.
        self.content_wrapper = ctk.CTkFrame(self, fg_color=FRAME_COLOR, corner_radius=16,
                                            border_width=2, border_color=BORDER_COLOR)
        self.content_wrapper.grid(row=1, column=0, padx=32, sticky="nsew")
        self.content_wrapper.grid_columnconfigure(0, weight=1)
        self.content_wrapper.grid_rowconfigure(1, weight=1)

        self.list_header_frame = ctk.CTkFrame(self.content_wrapper, fg_color="transparent")
        self.list_header_frame.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 12))
        self.list_header_frame.grid_columnconfigure(0, weight=1)

        self.list_title = ctk.CTkLabel(
            self.list_header_frame,
            text="Документы",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=MUTED_COLOR
        )
        self.list_title.grid(row=0, column=0, sticky="w")

        # ВНИМАНИЕ: self.md_mode уже создан выше (строка ~196) и привязан к
        # переключателю «Умный конвертер». Повторное создание здесь порождало
        # ВТОРОЙ BooleanVar, который читал export_md, тогда как переключатель
        # менял ПЕРВЫЙ — тумблер был мёртв. Не воссоздавать переменную здесь.

        # Равноправная альтернатива перетаскиванию: на macOS/Linux windnd
        # недоступен, и эта кнопка там — единственный способ дать документ.
        self.btn_pick = ctk.CTkButton(
            self.list_header_frame,
            text="Выбрать файл",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="transparent", hover_color=HOVER_COLOR, text_color=ACCENT_TEXT,
            border_width=2,   # 1 px на HiDPI почти не читался
            border_color=BORDER_COLOR,
            width=120, height=32, corner_radius=8,
            command=self.pick_file,
        )
        self.btn_pick.grid(row=0, column=1, sticky="e")

        self.files_scroll = ctk.CTkScrollableFrame(
            self.content_wrapper,
            fg_color="transparent"
        )
        self.files_scroll.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))
        
        # Полностью отключаем ползунок
        if hasattr(self.files_scroll, '_scrollbar'):
            self.files_scroll._scrollbar.grid_forget()
            self.files_scroll._scrollbar.grid = lambda *args, **kwargs: None
            
        self.document_cards = []

        # --- НИЖНЯЯ ПАНЕЛЬ (статус, прогресс, кнопки) ---
        self.action_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_frame.grid(row=2, column=0, padx=32, pady=(20, 24), sticky="ew")
        self.action_frame.grid_columnconfigure(0, weight=1)

        self.status_container = ctk.CTkFrame(self.action_frame, fg_color="transparent")
        self.status_container.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        self.status_container.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(
            self.status_container,
            text="Выберите документ для обработки.\nОбязательно проверяйте результат перед отправкой.",
            text_color=MUTED_COLOR,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            wraplength=640,
        )
        self.status_label.grid(row=0, column=0)

        self.progress_bar = ctk.CTkProgressBar(
            self.status_container,
            mode="indeterminate",
            progress_color=ACCENT_COLOR,
            fg_color=BORDER_COLOR,
            height=4
        )
        self.progress_bar.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self.progress_bar.set(0)
        self.progress_bar.grid_remove()

        # Две главные кнопки в ряд: слева «Анонимизировать», справа «Деанонимизировать».
        self.buttons_row = ctk.CTkFrame(self.action_frame, fg_color="transparent")
        self.buttons_row.grid(row=1, column=0, sticky="ew")
        self.buttons_row.grid_columnconfigure((0, 1), weight=1)

        self.btn_anonymize = ctk.CTkButton(
            self.buttons_row,
            text="Анонимизировать",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color=ACCENT_COLOR, hover_color=ACCENT_HOVER, text_color="#FFFFFF",
            height=48, corner_radius=8,
            command=self.start_processing,
        )
        self.btn_anonymize.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.btn_deanon = ctk.CTkButton(
            self.buttons_row,
            text="Деанонимизировать",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color=ACCENT_COLOR, hover_color=ACCENT_HOVER, text_color="#FFFFFF",
            height=48, corner_radius=8,
            command=self.start_deanon,
        )
        self.btn_deanon.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        # Кнопка "Скопировать для ИИ" — только после АНОНИМИЗАЦИИ (кладёт в буфер
        # обезличенный файл + текст). После деанонимизации НЕ показываем: там
        # результат с реальными ПДн, копировать его «для ИИ» опасно.
        self._copy_ai_default = "Скопировать для ИИ"
        self.btn_copy_ai = ctk.CTkButton(
            self.action_frame,
            text=self._copy_ai_default,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color="transparent", hover_color=HOVER_COLOR, text_color=ACCENT_TEXT,
            border_width=1, border_color=BORDER_COLOR, height=40, corner_radius=8,
            command=self.copy_for_ai,
        )
        self.btn_copy_ai.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.btn_copy_ai.grid_remove()

        # Кнопка "Открыть папку" появляется после любой успешной обработки.
        self.btn_open_folder = ctk.CTkButton(
            self.action_frame,
            text="Открыть папку с результатом",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color="transparent", hover_color=HOVER_COLOR, text_color=ACCENT_TEXT,
            border_width=1, border_color=BORDER_COLOR, height=40, corner_radius=8,
            command=self.open_result_folder,
        )
        self.btn_open_folder.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        self.btn_open_folder.grid_remove()

        # --- ПОДВАЛ (автор) ---
        self.credit_label = ctk.CTkLabel(
            self.action_frame,
            text="Создатель: Алексей Коняев · inbox@akonyaev.ru",
            font=ctk.CTkFont(family="Segoe UI", size=11, underline=True),
            text_color=MUTED_COLOR,
            cursor="hand2",
        )
        self.credit_label.grid(row=4, column=0, pady=(16, 0))
        self.credit_label.bind(
            "<Button-1>",
            lambda e: webbrowser.open("mailto:inbox@akonyaev.ru"),
        )

        self.is_processing = False
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Drag-and-drop (Windows). Мы используем безопасную 64-битную реализацию
        # на ctypes (safe_hook_dropfiles) вместо windnd, который ломает приложение
        # на 64-битных системах из-за усечения указателей окна.
        if sys.platform == 'win32':
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind('<<Drop>>', self._on_drop_event)
                self.dnd_enabled = True
            except Exception as exc:
                print(f"Drag-and-drop недоступен: {exc}", file=sys.stderr)

        # Кнопки «Обновить» больше нет, а список стартует пустым: без этого
        # вызова пользователь увидел бы пустую панель без единой подсказки.
        self.refresh_files()

    def _on_drop_event(self, event):
        files = self.tk.splitlist(event.data)
        self._on_drop(files)

    def _on_drop(self, filenames):
        """Приём перетащенных файлов и папок — основной способ дать документ."""
        if self.is_processing:
            return
        existing = {f["path"] for f in self.files}
        added = 0
        rejected = []
        truncated = False

        def accept(path):
            nonlocal added
            if path not in existing:
                self.files.append({
                    "name": os.path.basename(path),
                    "path": path,
                    "type": os.path.splitext(path)[1].lower(),
                })
                existing.add(path)
                added += 1
            self.selected_files.add(path)

        for raw in filenames:
            path = os.path.abspath(raw)
            if os.path.isdir(path):
                for root, dirs, files in os.walk(path, followlinks=False):
                    dirs[:] = [name for name in dirs
                               if not os.path.islink(os.path.join(root, name))]
                    for file in files:
                        if file.lower().endswith(SUPPORTED_EXTENSIONS):
                            accept(os.path.join(root, file))
                            if added >= 1_000:
                                truncated = True
                                break
                    if truncated:
                        break
            elif path.lower().endswith(SUPPORTED_EXTENSIONS):
                accept(path)
            else:
                rejected.append(path)

        self.refresh_files(reset_status=False)
        # Молча проглоченный дроп читается как «программа сломана» — для
        # DnD-first это главный источник недоверия. Объясняем отказ.
        if truncated:
            self.set_status("Добавлены первые 1000 файлов из папки.", color=WARNING_COLOR)
        elif rejected and not added:
            self.set_status(self._reject_reason(rejected), color=ERROR_COLOR)
        elif added == 1 and not rejected and len(filenames) == 1:
            # Умный Drag-and-Drop: проверяем, может ли это быть ответом ИИ
            path = list(self.selected_files)[0]
            anon_path, orig_path = self._resolve_deanon_pair(path)
            if anon_path == path and orig_path:
                # Нашли пару (смарт-матч с единственным паспортом)!
                # Автозапуск деанонимизации
                self.start_deanon()

    def _empty_state_text(self):
        """Пустой список — главный экран программы, он же инструкция.
        windnd работает только на Windows, поэтому на macOS/Linux
        единственный честный совет — кнопка выбора файла."""
        if self.dnd_enabled:
            return ("Перетащите документ сюда\n"
                    "или нажмите «Выбрать файл».")
        return "Нажмите «Выбрать файл»."


    @staticmethod
    def _reject_reason(rejected):
        exts = {os.path.splitext(path)[1].lower() for path in rejected}
        if '.doc' in exts:
            return ("Старый формат .doc не поддерживается: его открытие может "
                    "запускать макросы Word.\n"
                    "Сохраните документ в Word как .docx и перетащите снова.")
        listed = ', '.join(sorted(ext for ext in exts if ext)) or 'без расширения'
        return (f"Не подходит: {listed}.\n"
                "Пожалуйста, используйте поддерживаемый формат документа.")

    def _apply_crisp_icon(self, window=None):
        """Ставит чёткую иконку окна и панели задач Windows-нативно.

        Tk `iconbitmap` берёт один кадр из .ico и GDI-масштабирует его под
        текущий DPI → размыто. Здесь мы просим саму ОС загрузить кадр точного
        пиксельного размера из многоразрешённого .ico (LoadImageW) и ставим его
        через WM_SETICON для маленькой (заголовок) и большой (панель задач,
        Alt-Tab) иконок. Также повторно применяем iconbitmap, т.к. CustomTkinter
        мог сбросить нашу иконку при своей донастройке масштаба.

        window — целевое окно (главное или окно активации). По умолчанию self.
        """
        window = window if window is not None else self
        icon_path = get_resource_path('icon_purple.ico')
        try:
            window.iconbitmap(icon_path)
        except Exception:
            import logging
            logging.warning("Failed to set window iconbitmap", exc_info=True)

        if not (sys.platform == 'win32' and os.path.exists(icon_path)):
            return

        try:
            import ctypes
            from ctypes import wintypes

            user32 = ctypes.windll.user32
            # Явные сигнатуры обязательны: HICON/HWND — 64-битные хендлы,
            # иначе ctypes усечёт их до int и вызов сломается без ошибки.
            user32.LoadImageW.restype = wintypes.HANDLE
            user32.LoadImageW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR,
                                          wintypes.UINT, ctypes.c_int, ctypes.c_int,
                                          wintypes.UINT]
            user32.SendMessageW.restype = ctypes.c_ssize_t
            user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT,
                                            ctypes.c_size_t, ctypes.c_ssize_t]
            user32.GetAncestor.restype = wintypes.HWND
            user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]

            IMAGE_ICON = 1
            LR_LOADFROMFILE = 0x00000010
            WM_SETICON = 0x0080
            ICON_SMALL, ICON_BIG = 0, 1
            SM_CXICON, SM_CYICON = 11, 12
            SM_CXSMICON, SM_CYSMICON = 49, 50
            GA_ROOT = 2

            # Процесс Per-Monitor-V2-aware → системные метрики уже с учётом DPI,
            # поэтому запрашиваем ровно тот размер, который нужен ОС.
            cx_big = user32.GetSystemMetrics(SM_CXICON)
            cx_small = user32.GetSystemMetrics(SM_CXSMICON)
            cy_small = user32.GetSystemMetrics(SM_CYSMICON)
            
            cx_icon = user32.GetSystemMetrics(11)
            cy_icon = user32.GetSystemMetrics(12)
            hicon_big = user32.LoadImageW(0, icon_path, IMAGE_ICON, cx_icon, cy_icon, LR_LOADFROMFILE)
            hicon_small = user32.LoadImageW(None, icon_path, IMAGE_ICON,
                                            cx_small, cy_small, LR_LOADFROMFILE)

            # Кнопку в панели задач держит КОРНЕВОЕ окно (GA_ROOT) — оно, а не
            # сам Tk-HWND, определяет иконку в таскбаре. Ставим на оба на всякий
            # случай (заголовок берёт иконку с Tk-окна).
            win_id = window.winfo_id()
            root_hwnd = user32.GetAncestor(win_id, GA_ROOT) or win_id
            targets = {win_id, root_hwnd}

            for hwnd in targets:
                if hicon_big:
                    user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon_big)
                if hicon_small:
                    user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon_small)
        except Exception:
            import logging
            logging.warning("Failed to send WM_SETICON message", exc_info=True)

    def select_file(self, path):
        if self.is_processing:
            return
        if path in self.selected_files:
            self.selected_files.remove(path)
        else:
            self.selected_files.add(path)
        self._render_cards()

    def _render_cards(self):
        for card in self.document_cards:
            card.destroy()
        self.document_cards.clear()

        if hasattr(self.files_scroll, '_scrollbar'):
            if len(self.files) <= 3:
                self.files_scroll._scrollbar.grid_remove()
            else:
                self.files_scroll._scrollbar.grid()

        if not self.files:
            lbl = ctk.CTkLabel(
                self.files_scroll,
                text=self._empty_state_text(),
                text_color=MUTED_COLOR,
                font=ctk.CTkFont(family="Segoe UI", size=14),
                justify="center",
            )
            lbl.pack(pady=60)
            self.document_cards.append(lbl)
            self.selected_files.clear()
            self.btn_anonymize.configure(state="disabled")
            return

        self.btn_anonymize.configure(state="normal")

        for file_info in self.files:
            is_selected = (file_info['path'] in self.selected_files)
            card = DocumentCard(self.files_scroll, file_info, self.select_file, is_selected=is_selected)
            card.pack(fill="x", pady=(0, 8), padx=(8, 8))
            self.document_cards.append(card)

    def pick_file(self):
        """Выбор файла через диалог — альтернатива перетаскиванию.
        Добавляет его в список и выбирает."""
        if self.is_processing:
            return
        from tkinter import filedialog
        paths = filedialog.askopenfilenames(
            title="Выберите документ",
            filetypes=[("Документы", "*.pdf *.docx *.xlsx *.txt *.csv *.html"), ("Все файлы", "*.*")])
        if not paths:
            return
        for path in paths:
            path = os.path.abspath(path)
            if path not in [f['path'] for f in self.files]:
                self.files.append({
                    'name': os.path.basename(path), 'path': path,
                    'type': os.path.splitext(path)[1].lower(),
                })
            self.selected_files.add(path)
        self.refresh_files(reset_status=False)

    def refresh_files(self, reset_status=True):
        if self.is_processing:
            return

        # Список — только то, что дал пользователь. Файлы, которых уже нет
        # на диске (перемещены/удалены после дропа), отсеиваем.
        self.files = [f for f in self.files if os.path.exists(f['path'])]

        # Выбор — множественный (self.selected_files, набор путей). Отдельного
        # self.selected_file_path в этой модели нет: обращение к нему роняло
        # refresh_files (AttributeError) и карточки не отрисовывались вовсе.
        if not self.files:
            self.selected_files.clear()
        else:
            # Оставляем в выборе только пути, которые ещё присутствуют в списке.
            available = {f['path'] for f in self.files}
            self.selected_files = {p for p in self.selected_files if p in available}

        # reset_status=False сохраняет сообщение о результате обработки,
        # чтобы оно не затиралось при автоматическом обновлении после успеха.
        if reset_status:
            self.btn_open_folder.grid_remove()
            self.btn_copy_ai.grid_remove()
            if self.files:
                self.set_status("Выберите документ и нажмите «Анонимизировать».\n"
                                "Будут скрыты ФИО, организации, адреса, телефоны, e-mail и счета.\n"
                                "Обязательно проверяйте результат перед отправкой.")
            elif self.dnd_enabled:
                self.set_status("Перетащите документ в окно или нажмите «Выбрать файл».\n"
                                "Обязательно проверяйте результат перед отправкой.")
            else:
                self.set_status("Нажмите «Выбрать файл».\n"
                                "Обязательно проверяйте результат перед отправкой.")

        self._render_cards()

    def set_status(self, text, color=MUTED_COLOR):
        self.status_label.configure(text=text, text_color=color)

    def _set_controls_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        self.btn_pick.configure(state=state)

    def open_result_folder(self):
        path = self.last_out_path
        if not path or not os.path.exists(path):
            self.set_status("Файл результата не найден.", color=ERROR_COLOR)
            return
        try:
            if sys.platform == 'win32':
                subprocess.Popen(['explorer', '/select,', os.path.normpath(path)])
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', '-R', path])
            else:
                subprocess.Popen(['xdg-open', os.path.dirname(path)])
        except Exception:
            import logging
            logging.error("Failed to open file path", exc_info=True)
            if sys.platform == 'win32':
                os.startfile(os.path.dirname(path))

    def copy_for_ai(self):
        """Кладёт обезличенный файл И его текст в буфер обмена: вставить в поле
        чата ИИ (текст) или в область загрузки файлов/письмо (файл)."""
        path = self.last_out_path
        if not path or not os.path.exists(path):
            self.set_status("Файл результата не найден.", color=ERROR_COLOR)
            return
        try:
            import clipboard_util
            # copy_result возвращает False, если в буфер положить было нечего:
            # сообщаем честно, а не показываем «Скопировано» над пустым буфером.
            if not clipboard_util.copy_result(path, self):
                self.set_status("Не удалось скопировать текст этого формата.\n"
                                "Нажмите «Открыть папку» и приложите файл к чату вручную.",
                                color=ERROR_COLOR)
                return
            # Короткое подтверждение текстом (не анимация — простая смена подписи).
            if sys.platform != 'win32':
                self.btn_copy_ai.configure(text="Текст скопирован")
            else:
                self.btn_copy_ai.configure(text="Скопировано")
            self.after(1600, lambda: self.btn_copy_ai.configure(text=self._copy_ai_default))
        except Exception as e:
            self.set_status(f"Не удалось скопировать: {e}", color=ERROR_COLOR)

    def _bind_clipboard_keycodes(self, widget):
        """Ctrl+V/C/X/A по ФИЗИЧЕСКИМ кодам клавиш: в русской раскладке keysym
        кириллический и обычные бинды <Control-v> не срабатывают (§10.6
        AGENTS.md). ВАЖНО: виртуальное событие генерируем на event.widget —
        внутреннем tkinter.Entry, у которого есть классовые бинды <<Paste>> и
        которому реально пришло нажатие. Генерация на CTk-обёртке (фрейме)
        уходит в пустоту. 'break' гасит клавишу, чтобы в английской раскладке
        вставка не сработала дважды (наш бинд + классовый)."""
        actions = {86: '<<Paste>>', 67: '<<Copy>>', 88: '<<Cut>>', 65: '<<SelectAll>>'}
        def handler(event):
            seq = actions.get(event.keycode)
            if seq:
                event.widget.event_generate(seq)
                return 'break'
        widget.bind('<Control-KeyPress>', handler)
        widget._clip_handler = handler   # прямой доступ для автотестов

    def _attach_entry_context_menu(self, widget):
        """Контекстное меню правой кнопки для поля ввода: Вставить/Копировать/
        Вырезать/Выделить всё. Привычный для пользователя способ вставки —
        работает всегда, независимо от раскладки и горячих клавиш."""
        import tkinter as tk
        entry = widget._entry   # внутренний tkinter.Entry (цель команд меню)
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="Вставить",
                         command=lambda: entry.event_generate('<<Paste>>'))
        menu.add_command(label="Копировать",
                         command=lambda: entry.event_generate('<<Copy>>'))
        menu.add_command(label="Вырезать",
                         command=lambda: entry.event_generate('<<Cut>>'))
        menu.add_separator()
        menu.add_command(label="Выделить всё",
                         command=lambda: entry.event_generate('<<SelectAll>>'))

        def popup(event):
            entry.focus_set()
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
            return 'break'
        widget.bind('<Button-3>', popup)

    def on_closing(self):
        if self.is_processing:
            messagebox.showinfo(
                "Обработка не завершена",
                "Дождитесь завершения обработки. Готовые файлы публикуются только целиком."
            )
            return
        self.destroy()

    # ---------------- Деанонимизация (восстановление данных) ---------------- #
    def _resolve_deanon_pair(self, selected):
        """По выбранному документу определяет пару (ответ ИИ, оригинал).

        Ответ ИИ ищется в порядке предпочтения: '[ОТВЕТ] <имя>.md' (юрист
        сохранил ответ отдельно), '[ANON] <имя>.md' (перезаписал наш экспорт),
        '[ANON] <имя>.<ext>' (классический docx/txt-контур). Оригинал — рядом на
        диске и среди данных нам файлов. Возвращает (anon_path, orig_path) или (None, None)."""
        folder, fname = os.path.split(selected)
        base, _ext = os.path.splitext(fname)
        # Пара может лежать не в одной папке: юрист мог перетащить оригинал и
        # ответ ИИ из разных мест (типично — ответ из Загрузок). Поэтому ищем
        # и рядом на диске, и среди всех файлов, которые пользователь нам дал.
        known_paths = [f['path'] for f in self.files]

        def find_by_name(target_name):
            p = os.path.join(folder, target_name)
            if os.path.exists(p):
                return p
            for known in known_paths:
                if os.path.basename(known) == target_name and os.path.exists(known):
                    return known
            return None

        def manifests():
            result = []
            for path in Path(folder).glob('*.umbra.json'):
                if not naming.has_marker(path.name, naming.ANON):
                    continue
                try:
                    if path.stat().st_size > 1024 * 1024:
                        continue
                    with open(path, encoding='utf-8') as stream:
                        data = json.load(stream)
                    if data.get('schema') == 1 and isinstance(data.get('original_name'), str):
                        result.append(data)
                except (OSError, UnicodeError, json.JSONDecodeError, AttributeError):
                    continue
            return result

        records = manifests()
        if naming.has_marker(base, naming.ANSWER) or naming.has_marker(base, naming.ANON):
            clean = naming.strip_markers(base)
            names = {record['original_name'] for record in records
                     if naming.strip_markers(Path(record.get('output_name', '')).stem) == clean}
            if len(names) == 1:
                orig = find_by_name(names.pop())
                if orig:
                    return selected, orig
            return None, None

        # Выбран оригинал: паспорт даёт точное имя результата и исключает
        # неоднозначный выбор «первого попавшегося расширения».
        matching = [record for record in records if record['original_name'] == fname]
        candidates = []
        for record in matching:
            output_name = record.get('output_name', '')
            output_stem = naming.strip_markers(Path(output_name).stem)
            for answer_ext in ('.md', '.txt', '.docx', '.xlsx'):
                answer = find_by_name(naming.build_name(output_stem, naming.ANSWER, answer_ext))
                if answer and answer not in candidates:
                    candidates.append(answer)
            anon = find_by_name(output_name)
            if anon and anon not in candidates:
                candidates.append(anon)
        explicit_answers = [path for path in candidates
                            if naming.has_marker(Path(path).stem, naming.ANSWER)]
        if len(explicit_answers) == 1:
            return explicit_answers[0], selected
        if not explicit_answers and len(candidates) == 1:
            return candidates[0], selected

        if not matching:
            # Умный Drag-and-Drop: если выбран неизвестный файл, а паспорт всего один,
            # мы предполагаем, что этот файл — ответ от ИИ для этого единственного паспорта.
            if len(records) == 1:
                orig = find_by_name(records[0]['original_name'])
                if orig:
                    return selected, orig

        return None, None

    def start_deanon(self):
        if self.is_processing: return
        targets = list(self.selected_files)
        if not targets:
            messagebox.showerror("Ошибка", "Файлы не выбраны.")
            self.refresh_files()
            return
            
        pairs = []
        for target in targets:
            if not os.path.exists(target): continue
            anon_path, orig_path = self._resolve_deanon_pair(target)
            # Метки ищем в ИМЕНИ файла, а не в полном пути: папка вида
            # 'C:\\Мои [ANON] дела\\' иначе выдавала бы себя за помеченный файл
            # и спрашивать ответ ИИ мы бы перестали.
            if (not anon_path and orig_path is None and len(targets) == 1
                    and not naming.has_any_marker(Path(target).stem)):
                from tkinter import filedialog
                picked = filedialog.askopenfilename(title="Выберите файл с ответом ИИ", initialdir=os.path.dirname(target), filetypes=[("Безопасный ответ ИИ", "*.md *.docx *.xlsx *.txt"), ("Все файлы", "*.*")])
                if picked:
                    anon_path, orig_path = picked, target
            if anon_path and orig_path:
                pairs.append((anon_path, orig_path))
                
        if not pairs:
            self.set_status("Пары для восстановления (Оригинал + Ответ ИИ) не найдены.", color=ERROR_COLOR)
            return

        self.is_processing = True
        self.last_out_paths = []
        self.last_out_path = None
        self.btn_open_folder.grid_remove()
        self.btn_copy_ai.grid_remove()
        self.btn_deanon.configure(state="disabled", fg_color=PROCESSING_BTN, text_color="#FFFFFF", text="Восстановление…")
        self.btn_anonymize.configure(state="disabled")
        self._set_controls_enabled(False)
        self.set_status("Восстановление данных…", color=ACCENT_TEXT)
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(0)
        self.progress_bar.grid()
        self._render_cards()

        threading.Thread(target=self._deanon_thread, args=(pairs,), daemon=True).start()

    def _deanon_thread(self, pairs):
        total = len(pairs)
        success_count = 0
        errors = []
        notices = []
        restored = []
        for i, (anon_path, orig_path) in enumerate(pairs):
            self._ui(lambda idx=i: self.progress_bar.set(idx / total))
            self._ui(lambda fname=os.path.basename(anon_path): self.set_status(f"Восстановление {fname}...", color=ACCENT_TEXT))
            try:
                result = self.deanon_callback(anon_path, orig_path)
                if getattr(result, 'needs_confirmation', False):
                    manual = [index for index, item in enumerate(result.deletions)
                              if not item.get('auto')]
                    confirmed = None
                    if manual:
                        decision = {'value': False}
                        ready = threading.Event()

                        def ask():
                            decision['value'] = messagebox.askyesno(
                                "Удаления из ответа ИИ",
                                f"ИИ предложил удалить блоков: {len(manual)}. "
                                "Применить эти удаления?"
                            )
                            ready.set()

                        self._ui(ask)
                        ready.wait()
                        confirmed = set(manual) if decision['value'] else set()
                    out_path = result.save(confirmed)
                    stats = getattr(result, 'stats', {})
                    result = (out_path, stats)
                out_path, stats = result if isinstance(result, tuple) else (result, {})
                if out_path:
                    self.last_out_path = out_path
                    success_count += 1
                    # Оба контракта возврата (кортеж и MdDeanonSession.save)
                    # сходятся здесь, и здесь результат ТОЧНО уже на диске —
                    # единственная безопасная точка, чтобы наметить уборку.
                    restored.append((orig_path, out_path, anon_path))
                if stats.get('unresolved'):
                    notices.append(
                        f"{os.path.basename(anon_path)}: не восстановлено меток — "
                        f"{stats['unresolved']}"
                    )
                notices.extend(
                    f"{os.path.basename(anon_path)}: {notice}"
                    for notice in stats.get('notices', [])
                )
            except Exception as exc:
                errors.append((os.path.basename(anon_path), str(exc)))

        # Уборка — только ПОСЛЕ всего цикла и только по успешным парам. Внутри
        # цикла было бы нельзя: два ответа ИИ на один оригинал — обычное дело, а
        # удаление паспорта после первой пары обрубило бы вторую на полпути.
        removed = 0
        if self.cleanup_callback:
            for orig_path, out_path, answer_path in restored:
                try:
                    removed += len(self.cleanup_callback(orig_path, out_path, answer_path))
                except Exception:
                    import logging
                    logging.warning("Cleanup failed during deanon", exc_info=True)

        try:
            if success_count > 0:
                self._ui(lambda: self.progress_bar.set(1.0))
                msg = f"Успешно восстановлено файлов: {success_count} из {total}."
                if removed:
                    msg += f" Убрано промежуточных файлов: {removed}."
                if errors:
                    msg += "\nОшибки: " + "; ".join(f"{name} — {error}" for name, error in errors[:3])
                if notices:
                    msg += "\nПроверьте: " + "; ".join(notices[:3])
                color = WARNING_COLOR if errors or notices else SUCCESS_COLOR
                self._ui(lambda message=msg, chosen=color: self.set_status(message, color=chosen))
                self._ui(self.btn_open_folder.grid)
            else:
                detail = errors[0][1] if errors else "Не удалось восстановить данные."
                self._ui(lambda message=detail: self.set_status(f"Ошибка: {message}", color=ERROR_COLOR))
        finally:
            self._ui(self._deanon_done)

    def _deanon_done(self):
        """Разблокирует UI после восстановления (или отмены/сохранения сессии)."""
        self.progress_bar.stop()
        self.progress_bar.grid_remove()
        self.progress_bar.configure(mode="indeterminate")
        self.is_processing = False
        self.btn_deanon.configure(state="normal", fg_color=ACCENT_COLOR,
                                  text_color="#FFFFFF", text="Деанонимизировать")
        self.btn_anonymize.configure(state="normal")
        self._set_controls_enabled(True)
        self.refresh_files(reset_status=False)

    def start_processing(self):
        if self.is_processing: return
        targets = [t for t in self.selected_files if os.path.exists(t)]
        if not targets:
            messagebox.showerror("Ошибка", "Файлы не выбраны.")
            self.refresh_files()
            return

        self.is_processing = True
        self.last_out_paths = []
        self.last_out_path = None
        self.btn_open_folder.grid_remove()
        self.btn_copy_ai.grid_remove()
        self.btn_anonymize.configure(state="disabled", fg_color=PROCESSING_BTN, text_color="#FFFFFF", text="Обработка…")
        self.btn_deanon.configure(state="disabled")
        self._set_controls_enabled(False)

        self.set_status("Обработка документов…", color=ACCENT_TEXT)
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(0)
        self.progress_bar.grid()
        self._render_cards()

        opts = {
            'hide_names': True,
            'hide_locations': True,
            'hide_orgs': True,
            'smart_contract_mode': False,
            'export_md': bool(self.md_mode.get()),
        }

        thread = threading.Thread(target=self._process_thread, args=(targets, opts))
        thread.daemon = True
        thread.start()

    def _ui(self, func):
        """Безопасно маршалит вызов в главный поток Tk. Если окно уже закрыто,
        self.after бросает TclError — глушим, чтобы не уронить фоновый поток."""
        try:
            self.after(0, func)
        except Exception:
            import logging
            logging.warning("Failed to dispatch UI call to main thread", exc_info=True)

    def _process_thread(self, file_paths, opts):
        total = len(file_paths)
        success_count = 0
        errors = []                       # [(имя_файла, текст_ошибки)]
        try:
            for i, file_path in enumerate(file_paths):
                self._ui(lambda idx=i: self.progress_bar.set(idx / total))
                self._ui(lambda fname=os.path.basename(file_path): self.set_status(f"Обработка {fname}...", color=ACCENT_TEXT))
                # Ошибка на ОДНОМ файле не должна прерывать весь пакет и терять
                # отчёт об уже обработанных: ловим пофайлово, копим и продолжаем.
                try:
                    out_path = self.process_callback(file_path, **opts)
                except Exception as e:
                    errors.append((os.path.basename(file_path), str(e)))
                    continue
                if out_path:
                    self.last_out_paths.append(out_path)
                    self.last_out_path = out_path
                    success_count += 1

            def on_result():
                self.progress_bar.set(1.0)
                if success_count > 0:
                    if success_count == 1 and not errors:
                        name = os.path.basename(self.last_out_path)
                        self.set_status(f"Готово: {name}\nНажмите «Скопировать для ИИ»...", color=SUCCESS_COLOR)
                    else:
                        msg = f"Успешно обработано файлов: {success_count} из {total}."
                        if errors:
                            msg += "\nНе удалось: " + "; ".join(f"{n} — {e}" for n, e in errors[:3])
                            if len(errors) > 3:
                                msg += f" и ещё {len(errors) - 3}"
                        self.set_status(msg, color=SUCCESS_COLOR)
                    self.btn_copy_ai.grid()
                    self.btn_open_folder.grid()
                elif errors:
                    self.set_status(f"Ошибка: {errors[0][1]}", color=ERROR_COLOR)
                else:
                    self.set_status("Не удалось обработать файлы.", color=ERROR_COLOR)
            self._ui(on_result)
        except Exception as e:
            msg = str(e)
            self._ui(lambda: self.set_status(f"Ошибка: {msg}", color=ERROR_COLOR))
        finally:
            def on_done():
                self.progress_bar.stop()
                self.progress_bar.grid_remove()
                self.progress_bar.configure(mode="indeterminate")
                self.is_processing = False
                self.btn_anonymize.configure(state="normal", fg_color=ACCENT_COLOR, text_color="#FFFFFF", text="Анонимизировать")
                self.btn_deanon.configure(state="normal")
                self._set_controls_enabled(True)
                self.refresh_files(reset_status=False)
            self._ui(on_done)

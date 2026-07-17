import logging
from gui import UmbraApp
from doc_processor import DocumentProcessor
from nlp_engine import NLPProcessor

import tempfile
import os

log_path = os.path.join(tempfile.gettempdir(), 'umbra.log')

# Настройка глобального логгера (файл создастся только при реальной записи ошибки)
logging.basicConfig(
    level=logging.WARNING,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[logging.FileHandler(log_path, delay=True)]
)

doc_proc_instance = None

def get_doc_proc():
    global doc_proc_instance
    if doc_proc_instance is None:
        nlp = NLPProcessor()
        doc_proc_instance = DocumentProcessor(nlp)
    return doc_proc_instance

def process_file_wrapper(file_path, **opts):
    # Вызывается из фонового потока GUI. НЕ показываем messagebox отсюда
    # (Tkinter не потокобезопасен) — даём исключению всплыть в _process_thread,
    # который выводит ошибку в UI через self.after(0, ...).
    proc = get_doc_proc()
    return proc.process_file(file_path, **opts)

def deanonymize_wrapper(new_path, source_path):
    # Восстановление данных: документ от ИИ + ключ или оригинал.
    proc = get_doc_proc()
    return proc.deanonymize_file(new_path, source_path)

def cleanup_wrapper(source_path, out_path, answer_path=None):
    # Уборка промежуточных файлов. Вызывается GUI только после того, как
    # результат сохранён на диск (см. _deanon_thread).
    proc = get_doc_proc()
    return proc.cleanup_intermediates(source_path, out_path, answer_path)

def main():
    app = UmbraApp(process_callback=process_file_wrapper,
                   deanon_callback=deanonymize_wrapper,
                   cleanup_callback=cleanup_wrapper)
    app.mainloop()

if __name__ == "__main__":
    main()

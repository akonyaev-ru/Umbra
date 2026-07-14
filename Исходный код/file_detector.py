import os
import psutil
import pythoncom
import win32com.client

def get_open_files():
    """
    Scans for open .docx files in Microsoft Word and open .txt files in Notepad/Notepad++.
    Returns a list of dicts: {'name': filename, 'path': full_path, 'type': ext}
    """
    open_files = []
    seen_paths = set()
    
    # 1. Detect open Word documents
    com_initialized = False
    try:
        # We must initialize COM in the current thread since this might be called
        # from a GUI background thread. CoInitialize бросает com_error с кодом
        # RPC_E_CHANGED_MODE, если поток уже инициализирован в другой модели —
        # в этом случае COM всё равно годен к использованию, просто не деинициализируем.
        try:
            pythoncom.CoInitialize()
            com_initialized = True
        except Exception:
            com_initialized = False

        try:
            word = win32com.client.GetActiveObject("Word.Application")
            for doc in word.Documents:
                # Отдельный try на документ: один файл в модальном/битом состоянии
                # не должен прерывать перечисление остальных.
                try:
                    path = doc.FullName
                except Exception:
                    continue
                if os.path.exists(path) and path.lower().endswith('.docx'):
                    if path not in seen_paths:
                        open_files.append({
                            'name': os.path.basename(path),
                            'path': path,
                            'type': '.docx'
                        })
                        seen_paths.add(path)
        except Exception:
            pass # Word not running or no permissions
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # 2. Detect open Text files in Notepad and Notepad++
    try:
        for p in psutil.process_iter(['name', 'cmdline']):
            try:
                name = p.info['name']
                if name and ('notepad' in name.lower()):
                    cmd = p.info['cmdline']
                    if cmd and len(cmd) > 1:
                        for arg in cmd[1:]:
                            # Skip standard arguments
                            if arg.startswith('-') or arg.startswith('/'):
                                continue
                            if os.path.exists(arg) and arg.lower().endswith('.txt'):
                                path = os.path.abspath(arg)
                                if path not in seen_paths:
                                    open_files.append({
                                        'name': os.path.basename(path),
                                        'path': path,
                                        'type': '.txt'
                                    })
                                    seen_paths.add(path)
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    except Exception:
        pass

    return open_files

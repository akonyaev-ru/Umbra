"""Сборка БЕЗ обфускации (обычный PyInstaller --onefile).

Резервный путь, пока не куплена лицензия PyArmor: build.py (PyArmor-триал) не
может обфусцировать весь проект («out of license»). Этот скрипт даёт рабочий exe
с актуальным кодом, но без защиты исходников. Лицензионный шлюз (licensing.py)
работает независимо от обфускации.
"""

import os
import sys
import subprocess


def build():
    import customtkinter
    ctk_path = os.path.dirname(customtkinter.__file__)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconsole",
        "--onefile",
        "--icon=icon_purple.ico",
        "--version-file=version_info.txt",
        "--add-data=icon_purple.ico;.",
        "--add-data=logo.png;.",
        "--add-data=Audiowide-Regular.ttf;.",
        f"--add-data={ctk_path};customtkinter/",
        "--add-data=models;models/",
        "--copy-metadata=natasha",
        "--copy-metadata=slovnet",
        "--copy-metadata=navec",
        "--copy-metadata=razdel",
        # Ленивые импорты внутри методов PyInstaller обычно находит, но третьи
        # стороны подстрахуем явно.
        "--hidden-import=markdown_it",
        "--hidden-import=pypdf",
        "--hidden-import=win32clipboard",
        "--distpath=..",
        "--name=Umbra",
        "main.py",
    ]
    print("Running PyInstaller (no obfuscation):")
    print(" ".join(cmd))
    subprocess.run(cmd, check=True)


if __name__ == "__main__":
    build()

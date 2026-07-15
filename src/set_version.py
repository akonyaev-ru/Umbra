"""Обновляет version_info.txt номером версии из тега релиза.

Версия берётся из переменной окружения UMBRA_VERSION (напр. «v2.1.1»), которую
CI подставляет из github.ref_name. Без тега скрипт ничего не меняет — остаётся
значение, записанное в файле. Нужно, чтобы бинарник, собранный из тега v2.1.1,
не сообщал в свойствах «2.1.0» (иначе не отличить пропатченную сборку от старой).

Запуск: UMBRA_VERSION=v2.1.1 python set_version.py
"""
import os
import re


def main():
    tag = os.environ.get("UMBRA_VERSION", "").lstrip("vV").strip()
    if not tag:
        print("UMBRA_VERSION not set - version_info.txt unchanged")
        return

    parts = re.findall(r"\d+", tag)[:4]
    while len(parts) < 4:
        parts.append("0")
    nums = ", ".join(parts)
    dotted = ".".join(parts[:3])

    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "version_info.txt")
    text = open(path, encoding="utf-8").read()
    text = re.sub(r"filevers=\(\d+,\s*\d+,\s*\d+,\s*\d+\)", f"filevers=({nums})", text)
    text = re.sub(r"prodvers=\(\d+,\s*\d+,\s*\d+,\s*\d+\)", f"prodvers=({nums})", text)
    text = re.sub(r"StringStruct\('FileVersion', '[^']*'\)",
                  f"StringStruct('FileVersion', '{dotted}')", text)
    text = re.sub(r"StringStruct\('ProductVersion', '[^']*'\)",
                  f"StringStruct('ProductVersion', '{dotted}')", text)
    open(path, "w", encoding="utf-8").write(text)
    print(f"version_info.txt updated to version {dotted}")


if __name__ == "__main__":
    main()

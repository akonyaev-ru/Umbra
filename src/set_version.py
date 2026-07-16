"""Обновляет version_info.txt номером версии из тега релиза.

Версия берётся из переменной окружения UMBRA_VERSION (напр. «v2.1.1»), которую
CI подставляет из github.ref_name. Без тега скрипт ничего не меняет — остаётся
значение, записанное в файле. Нужно, чтобы бинарник, собранный из тега v2.1.1,
не сообщал в свойствах «2.1.0» (иначе не отличить пропатченную сборку от старой).

Запуск: UMBRA_VERSION=v2.1.1 python set_version.py
"""
import os
import re
import tempfile


def main():
    raw_tag = os.environ.get("UMBRA_VERSION", "").strip()
    if not raw_tag:
        print("UMBRA_VERSION not set - version_info.txt unchanged")
        return
    match = re.fullmatch(r"[vV]?(\d+)\.(\d+)(?:\.(\d+))?(?:\.(\d+))?", raw_tag)
    if not match:
        raise ValueError(
            "UMBRA_VERSION must be an exact numeric version such as v2026.2 or v2.1.1"
        )
    parts = list(match.groups(default="0"))
    nums = ", ".join(parts)
    dotted = ".".join(parts[:3])

    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "version_info.txt")
    with open(path, encoding="utf-8") as stream:
        text = stream.read()
    text = re.sub(r"filevers=\(\d+,\s*\d+,\s*\d+,\s*\d+\)", f"filevers=({nums})", text)
    text = re.sub(r"prodvers=\(\d+,\s*\d+,\s*\d+,\s*\d+\)", f"prodvers=({nums})", text)
    text = re.sub(r"StringStruct\('FileVersion', '[^']*'\)",
                  f"StringStruct('FileVersion', '{dotted}')", text)
    text = re.sub(r"StringStruct\('ProductVersion', '[^']*'\)",
                  f"StringStruct('ProductVersion', '{dotted}')", text)
    fd, temporary = tempfile.mkstemp(prefix=".version-", suffix=".txt",
                                     dir=os.path.dirname(path))
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(text)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        try:
            os.remove(temporary)
        except FileNotFoundError:
            pass
    print(f"version_info.txt updated to version {dotted}")


if __name__ == "__main__":
    main()

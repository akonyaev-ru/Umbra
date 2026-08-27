"""Security boundaries shared by document import and export code.

The application handles documents that may be both private and hostile.  This
module keeps the fail-closed checks small, deterministic and easy to test.
"""

from __future__ import annotations

import contextlib
import hashlib
import json
import os
from pathlib import Path
import stat
import tempfile
import zipfile
from xml.etree import ElementTree
from defusedxml import ElementTree as SafeElementTree


MAX_INPUT_BYTES = 100 * 1024 * 1024
MAX_TEXT_BYTES = 20 * 1024 * 1024
MAX_ZIP_MEMBERS = 5_000
MAX_ZIP_UNCOMPRESSED_BYTES = 250 * 1024 * 1024
MAX_ZIP_RATIO = 200
MAX_XLSX_CELLS = 1_000_000
MANIFEST_SCHEMA = 2
# Схема 1 — паспорта прежних версий: без поля algo_version. Читать их можно, но
# восстановление по ним отклоняется (см. load_matching_manifest): правила
# обезличивания с тех пор менялись, и метки указывали бы на другие значения.
SUPPORTED_MANIFEST_SCHEMAS = (1, 2)

# Byte-identical safe parts from python-docx 1.2.0's empty default template.
# The dependency is pinned; updating it requires reviewing these fingerprints.
_SAFE_DOCX_CUSTOM_PARTS = {
    "customxml/_rels/item1.xml.rels",
    "customxml/item1.xml",
    "customxml/itemprops1.xml",
}
_SAFE_DOCX_CUSTOM_HASHES = {
    "customxml/item1.xml": "a86086ffc5d8e83ebd6c71a55d1d2efaa31b137977f5f3a752366e1023612144",
    "customxml/itemprops1.xml": "c542307b13ec29a8b546217bb37936ab4822e044b265d2952985ec3d6afed24e",
}
_SAFE_DOCX_THUMBNAIL_HASH = (
    "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead"
)


def sha256_file(path: str | os.PathLike[str]) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def check_input_file(path: str | os.PathLike[str], *, text: bool = False) -> None:
    path = os.fspath(path)
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Файл не найден: {path}")
    limit = MAX_TEXT_BYTES if text else MAX_INPUT_BYTES
    size = os.path.getsize(path)
    if size > limit:
        raise ValueError(
            f"Файл слишком большой ({size // (1024 * 1024)} МБ). "
            f"Допустимо не более {limit // (1024 * 1024)} МБ."
        )


def _is_reparse_point(path: str) -> bool:
    try:
        info = os.lstat(path)
    except FileNotFoundError:
        return False
    attrs = getattr(info, "st_file_attributes", 0)
    marker = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0)
    return os.path.islink(path) or bool(marker and attrs & marker)


def unique_output_path(path: str | os.PathLike[str]) -> str:
    """Never overwrite an existing output (including links/reparse points)."""
    candidate = os.path.abspath(os.fspath(path))
    parent = os.path.dirname(candidate)
    if _is_reparse_point(parent):
        raise ValueError("Папка назначения является ссылкой/reparse point.")
    if not os.path.lexists(candidate):
        return candidate
    stem, suffix = os.path.splitext(candidate)
    for number in range(2, 10_000):
        numbered = f"{stem} ({number}){suffix}"
        if not os.path.lexists(numbered):
            return numbered
    raise FileExistsError("Не удалось подобрать свободное имя выходного файла.")


@contextlib.contextmanager
def atomic_output(final_path: str | os.PathLike[str]):
    """Yield a same-directory temporary path and atomically publish it."""
    final_path = os.path.abspath(os.fspath(final_path))
    parent = os.path.dirname(final_path)
    os.makedirs(parent, exist_ok=True)
    if os.path.lexists(final_path) or _is_reparse_point(parent):
        raise FileExistsError(f"Выходной путь уже занят или небезопасен: {final_path}")
    suffix = os.path.splitext(final_path)[1]
    fd, temporary = tempfile.mkstemp(prefix=".umbra-", suffix=suffix, dir=parent)
    os.close(fd)
    try:
        yield temporary
        if not os.path.isfile(temporary):
            raise RuntimeError("Обработчик не создал выходной файл.")
        if os.path.lexists(final_path):
            raise FileExistsError(f"Выходной путь был занят во время обработки: {final_path}")
        os.replace(temporary, final_path)
    finally:
        try:
            os.remove(temporary)
        except FileNotFoundError:
            pass


def write_json_atomic(path: str | os.PathLike[str], payload: dict) -> None:
    path = os.path.abspath(os.fspath(path))
    with atomic_output(path) as temporary:
        with open(temporary, "w", encoding="utf-8", newline="\n") as stream:
            json.dump(payload, stream, ensure_ascii=False, indent=2, sort_keys=True)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
    try:
        os.chmod(path, stat.S_IRUSR | stat.S_IWUSR)
    except OSError:
        pass


def write_manifest(output_path: str, original_path: str, options: dict,
                   algo_version: str) -> str:
    manifest_path = output_path + ".umbra.json"
    payload = {
        "schema": MANIFEST_SCHEMA,
        "original_name": os.path.basename(original_path),
        "original_sha256": sha256_file(original_path),
        "output_name": os.path.basename(output_path),
        "algo_version": algo_version,
        "options": {key: bool(value) for key, value in options.items()},
    }
    write_json_atomic(manifest_path, payload)
    return manifest_path


def find_matching_manifests(original_path: str) -> list:
    """Все паспорта, привязанные к этому НЕИЗМЕНЁННОМУ оригиналу: [(path, data)],
    новые первыми. Пустой список — привязанных паспортов нет."""
    original_path = os.path.abspath(original_path)
    expected_hash = sha256_file(original_path)
    # Имя паспорта НЕ фильтруем: привязка к оригиналу проверяется ниже по
    # содержимому (original_name + original_sha256), а имя как ключ — хрупко.
    # Фильтр по "<stem> [ANON]" ломался от любой смены схемы имён и выдавал
    # «Не найден паспорт», что читается как подмена оригинала, а не как баг имён.
    candidates = sorted(
        Path(original_path).parent.glob("*.umbra.json"),
        key=lambda p: p.stat().st_mtime_ns,
        reverse=True,
    )
    valid = []
    for candidate in candidates:
        check_input_file(candidate, text=True)
        try:
            with open(candidate, "r", encoding="utf-8") as stream:
                data = json.load(stream)
        except (OSError, UnicodeError, json.JSONDecodeError):
            continue
        if (
            isinstance(data, dict)
            and data.get("schema") in SUPPORTED_MANIFEST_SCHEMAS
            and data.get("original_name") == os.path.basename(original_path)
            and data.get("original_sha256") == expected_hash
            and isinstance(data.get("options"), dict)
        ):
            valid.append((candidate, data))
    return valid


def load_matching_manifest(original_path: str, algo_version: str) -> dict:
    """Find the newest valid sidecar bound to this unchanged original."""
    valid = find_matching_manifests(original_path)
    if not valid:
        raise ValueError(
            "Не найден паспорт обезличивания для неизменённого оригинала. "
            "Обезличьте этот файл заново и используйте созданный ответ."
        )
    option_sets = {json.dumps(item[1]["options"], sort_keys=True) for item in valid}
    if len(option_sets) > 1:
        raise ValueError(
            "Найдено несколько паспортов с разными настройками. "
            "Удалите устаревшие файлы *.umbra.json или обезличьте оригинал заново."
        )
    # Карта замен нигде не хранится: она строится повторной анонимизацией
    # оригинала. Если правила обезличивания с тех пор изменились (обновление
    # Umbra), состав и нумерация меток будут другими — [ФИО_3] из старого ответа
    # ИИ указал бы на другого человека. Тихо подставить чужие данные хуже, чем
    # отказать, поэтому проверка жёсткая.
    stored_version = valid[0][1].get("algo_version")
    if stored_version != algo_version:
        raise ValueError(
            "Паспорт обезличивания создан другой версией Umbra "
            f"(в паспорте: {stored_version or 'не указана'}, сейчас: {algo_version}). "
            "Правила обезличивания с тех пор изменились, и метки указывали бы "
            "не на те данные. Обезличьте оригинал заново и отправьте ИИ новый файл."
        )
    return valid[0][1]


def validate_ooxml(path: str, kind: str, *, reject_formulas: bool = False) -> None:
    """Reject active/external/opaque OOXML content and decompression bombs."""
    check_input_file(path)
    if kind not in {"docx", "xlsx"}:
        raise ValueError(f"Неизвестный тип OOXML: {kind}")
    forbidden_prefixes = (
        ("word/media/", "word/embeddings/", "word/activex/", "word/charts/",
         "word/diagrams/", "word/glossary/")
        if kind == "docx"
        else ("xl/media/", "xl/embeddings/", "xl/activex/", "xl/externallinks/",
              "xl/charts/", "xl/pivotcache/", "xl/querytables/",
              "xl/slicercaches/", "xl/model/", "customxml/")
    )
    forbidden_files = {
        "docprops/custom.xml",
        "word/vbaproject.bin",
        "xl/vbaproject.bin",
        "xl/connections.xml",
    }
    try:
        archive = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as exc:
        raise ValueError("Файл повреждён или не является корректным OOXML-документом.") from exc
    with archive:
        infos = archive.infolist()
        if len(infos) > MAX_ZIP_MEMBERS:
            raise ValueError("В документе слишком много внутренних частей.")
        total = 0
        rel_files = []
        sheet_files = []
        seen_names = set()
        names_lower = {info.filename.replace("\\", "/").lower(): info.filename for info in infos}
        custom_parts = {name: original for name, original in names_lower.items()
                        if name.startswith("customxml/")}
        if custom_parts:
            if kind != "docx":
                raise ValueError("Пользовательские XML-данные в Excel не поддерживаются.")
            # python-docx's clean template contains a fixed, empty bibliography.
            # The relationship id may be re-serialized, while the data parts are
            # required to match the pinned template exactly.
            if set(custom_parts) != _SAFE_DOCX_CUSTOM_PARTS or any(
                hashlib.sha256(archive.read(original)).hexdigest()
                != _SAFE_DOCX_CUSTOM_HASHES[name]
                for name, original in custom_parts.items()
                if name in _SAFE_DOCX_CUSTOM_HASHES
            ):
                raise ValueError("Документ содержит пользовательские XML-данные.")

        thumbnail_name = names_lower.get("docprops/thumbnail.jpeg")
        if kind == "docx" and thumbnail_name:
            if (hashlib.sha256(archive.read(thumbnail_name)).hexdigest()
                    != _SAFE_DOCX_THUMBNAIL_HASH):
                raise ValueError(
                    "Документ содержит пользовательскую миниатюру предпросмотра. "
                    "Отключите сохранение миниатюры в свойствах Word."
                )

        for info in infos:
            name = info.filename.replace("\\", "/")
            name_lower = name.lower()
            if name_lower in seen_names:
                raise ValueError("В документе обнаружены дублирующиеся ZIP-части.")
            seen_names.add(name_lower)
            if name.startswith("/") or ".." in Path(name).parts:
                raise ValueError("В документе обнаружен небезопасный путь ZIP.")
            if name_lower.startswith("_xmlsignatures/"):
                raise ValueError(
                    "Подписанные OOXML-документы не изменяются: сохраните отдельную "
                    "неподписанную копию для обезличивания."
                )
            total += info.file_size
            if total > MAX_ZIP_UNCOMPRESSED_BYTES:
                raise ValueError("Распакованный документ превышает безопасный лимит.")
            if info.file_size and info.compress_size == 0:
                raise ValueError("В документе обнаружена подозрительная ZIP-часть.")
            if info.compress_size and info.file_size / info.compress_size > MAX_ZIP_RATIO:
                raise ValueError("Документ похож на ZIP-бомбу.")
            if name_lower in forbidden_files or name_lower.startswith(forbidden_prefixes):
                raise ValueError(
                    "Документ содержит вложения, медиа, макросы или пользовательские данные, "
                    "которые Umbra не может безопасно обезличить."
                )
            if name_lower.endswith(".rels"):
                rel_files.append(name)
            if kind == "xlsx" and name_lower.startswith("xl/worksheets/") and name_lower.endswith(".xml"):
                sheet_files.append(name)
        for name in rel_files:
            try:
                root = SafeElementTree.fromstring(archive.read(name))
            except ElementTree.ParseError as exc:
                raise ValueError("Повреждён файл связей OOXML.") from exc
            for rel in root:
                target = rel.attrib.get("Target", "").strip()
                scheme = target.split(':', 1)[0].lower() if ':' in target else ''
                looks_external = (
                    rel.attrib.get("TargetMode", "").lower() == "external"
                    or scheme in {"http", "https", "ftp", "file", "mailto", "tel"}
                    or target.startswith(("//", "\\\\"))
                    or (len(target) > 2 and target[1] == ':' and target[0].isalpha())
                )
                if looks_external:
                    raise ValueError(
                        "Документ содержит внешнюю ссылку. Удалите гиперссылки и "
                        "внешние подключения перед обработкой."
                    )
        if reject_formulas and kind == "xlsx":
            for name in sheet_files:
                try:
                    root = SafeElementTree.fromstring(archive.read(name))
                except ElementTree.ParseError as exc:
                    raise ValueError("Повреждён XML листа Excel.") from exc
                if any(element.tag.rsplit("}", 1)[-1] == "f" for element in root.iter()):
                    raise ValueError(
                        "Формулы Excel запрещены: ответ ИИ должен содержать только значения."
                    )


def scrub_extended_properties(path: str) -> None:
    """Remove personal/hidden values from docProps/app.xml in-place."""
    sensitive = {
        "Manager", "Company", "HyperlinkBase", "Template",
        "HeadingPairs", "TitlesOfParts",
    }
    folder = os.path.dirname(os.path.abspath(path))
    suffix = os.path.splitext(path)[1]
    fd, temporary = tempfile.mkstemp(prefix=".umbra-props-", suffix=suffix, dir=folder)
    os.close(fd)
    try:
        with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temporary, "w") as target:
            target.comment = source.comment
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename.lower() == "docprops/app.xml":
                    try:
                        root = SafeElementTree.fromstring(data)
                    except ElementTree.ParseError as exc:
                        raise ValueError("Повреждены расширенные свойства OOXML.") from exc
                    for element in list(root):
                        if element.tag.rsplit("}", 1)[-1] in sensitive:
                            element.clear()
                    data = ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)
                target.writestr(info, data)
        os.replace(temporary, path)
    finally:
        try:
            os.remove(temporary)
        except FileNotFoundError:
            pass

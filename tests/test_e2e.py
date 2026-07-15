import os
import sys
import shutil
from pathlib import Path

# Добавляем src в PYTHONPATH
src_dir = Path(__file__).parent.parent / 'src'
sys.path.insert(0, str(src_dir))

from docx import Document
from doc_processor import DocumentProcessor
from nlp_engine import NLPProcessor

def create_test_docx(path):
    doc = Document()
    doc.add_heading('Договор оказания услуг', 0)
    p = doc.add_paragraph('г. Москва, ')
    p.add_run('14 июля 2026 г.').bold = True
    
    doc.add_paragraph('ООО "Рога и Копыта" в лице генерального директора Иванова Ивана Ивановича, с одной стороны, и Microsoft Corporation, с другой стороны, заключили договор.')
    
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = 'Услуга'
    table.cell(0, 1).text = 'Стоимость'
    table.cell(1, 0).text = 'Разработка ПО'
    table.cell(1, 1).text = '150 000,00'
    
    doc.save(path)

def test_e2e_anonymize_deanonymize(tmp_path):
    # Подготавливаем тестовый файл
    test_file = tmp_path / 'test.docx'
    create_test_docx(str(test_file))
    
    # Меняем текущую директорию на src, чтобы nlp_engine нашел модели
    os.chdir(str(src_dir))
    
    # Анонимизация
    nlp = NLPProcessor()
    processor = DocumentProcessor(nlp)
    # process_file сохраняет anon-файл рядом.
    result = processor.process_file(str(test_file), export_md=True)
    anon_path = result[0] if isinstance(result, tuple) else result
    
    assert os.path.exists(anon_path)
    assert anon_path.endswith('.md')
    
    # Читаем MD, симулируем ответ ИИ (ничего не меняли)
    with open(anon_path, 'r', encoding='utf-8') as f:
        md_text = f.read()
    
    assert '[ФИО_1]' in md_text
    assert '[ОРГАНИЗАЦИЯ_1]' in md_text

    # ГЛАВНАЯ проверка продукта: в файле, уходящем ИИ, НЕ должно остаться НИ ОДНОЙ
    # исходной единицы ПДн. Раньше её не было — регрессия, при которой имя утекает
    # в «обезличенный» md, проходила тест зелёной (см. аудит).
    for leaked in ('Иванова Ивана Ивановича', 'Иванов Иван Иванович',
                   'Microsoft Corporation', 'Рога и Копыта', '150 000,00'):
        assert leaked not in md_text, f'ПДн утекли в [ANON].md: {leaked!r}'

    # Английское название организации должно быть замаскировано (прежнее «ИЛИ»
    # всегда удовлетворялось русской компанией и Microsoft не проверяло вовсе).
    assert '[ОРГАНИЗАЦИЯ_2]' in md_text or '[ОРГАНИЗАЦИЯ_3]' in md_text
    assert '[СУММА_1]' in md_text or '[СУММА_2]' in md_text
    
    # Деанонимизация (md_text + оригинал)
    session = processor.deanonymize_file(anon_path, str(test_file))
    session.save()
    deanon_path = session.out_path
    
    assert os.path.exists(deanon_path)
    
    # Проверяем, что реальные данные вернулись
    doc = Document(deanon_path)
    full_text = '\n'.join([p.text for p in doc.paragraphs])
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                full_text += '\n' + cell.text
                
    assert 'Иванова Ивана Ивановича' in full_text
    assert 'Microsoft Corporation' in full_text
    assert '150 000,00' in full_text
    assert '[ФИО_1]' not in full_text

def test_detectors_hide_pii():
    """Регрессия на детекторы, которые раньше пропускали ПДн (см. аудит)."""
    os.chdir(str(src_dir))
    nlp = NLPProcessor()
    expected = {
        'e-mail: ivanov@почта.рф': '[EMAIL]',            # кириллический домен (IDN)
        'тел. (495) 123-45-67': '[ТЕЛЕФОН]',             # городской формат
        'звоните 123-45-67': '[ТЕЛЕФОН]',                # без кода страны
        'паспорт 4509 123456': '[ПАСПОРТ]',              # без слова «серия»
        'Страховое свидетельство: 112-233-445 95': '[СНИЛС]',  # синоним СНИЛС
        'Дата рождения: 15.06.1985': '[ДР]',             # обратный порядок
    }
    for text, tag in expected.items():
        out = nlp.anonymize_text(text)
        assert tag in out, f'{text!r} -> {out!r} (ждали {tag})'

    # Даты НЕ должны дробиться в [СУММА].YYYY (over-anon).
    dt = nlp.anonymize_text('Договор от 27.07.2006, срок до 31.12.2026.')
    assert '27.07.2006' in dt and '31.12.2026' in dt and '[СУММА' not in dt

    # «ГК РФ» — указание права, не адрес: должно сохраниться.
    assert 'ГК РФ' in nlp.anonymize_text('Согласно ст. 431 ГК РФ.')


def test_xlsx_numeric_and_comment(tmp_path):
    """ИНН/телефон, сохранённые как ЧИСЛО, и текст примечания к ячейке — тоже ПДн."""
    os.chdir(str(src_dir))
    import openpyxl
    from openpyxl.comments import Comment
    nlp = NLPProcessor()
    proc = DocumentProcessor(nlp)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['A1'] = 7707083893                 # ИНН как число
    ws['A2'] = 5                          # короткое число — не ПДн, не трогаем
    ws['B1'].comment = Comment('Клиент Петров, ИНН 7707083893', 'Сидоров')
    src = tmp_path / 'reestr.xlsx'
    wb.save(str(src))

    out = proc.process_file(str(src))
    w = openpyxl.load_workbook(out).active
    assert '7707083893' not in str(w['A1'].value)
    assert w['A2'].value == 5
    assert '7707083893' not in (w['B1'].comment.text or '')


def test_underscore_label_merge_invariant(tmp_path):
    """Метки с подчёркиванием ([ИНН_ЮЛ_1]) должны быть видны инварианту слияния:
    если ИИ выкидывает клаузу с ИНН, значение берётся из оригинала, а не теряется."""
    os.chdir(str(src_dir))
    nlp = NLPProcessor()
    proc = DocumentProcessor(nlp)
    d = Document()
    d.add_paragraph('Реквизиты: ИНН 7701234567, КПП 770101001.')
    src = tmp_path / 'inn.docx'
    d.save(str(src))

    md_path = proc.process_file(str(src), export_md=True)
    md_text = open(md_path, encoding='utf-8').read()
    assert '[ИНН_ЮЛ_1]' in md_text

    # Симулируем ответ ИИ, который ВЫКИНУЛ клаузу с ИНН.
    ai = md_text.replace('ИНН [ИНН_ЮЛ_1], ', '')
    ans = tmp_path / 'answer.md'
    ans.write_text(ai, encoding='utf-8')

    session = proc.deanonymize_file(str(ans), str(src))
    session.save()
    full = '\n'.join(p.text for p in Document(session.out_path).paragraphs)
    assert '7701234567' in full, 'реальный ИНН потерян при деанонимизации'


def test_docx_core_properties_scrubbed(tmp_path):
    """Метаданные документа (автор/заголовок) не должны утекать в [ANON].docx."""
    os.chdir(str(src_dir))
    nlp = NLPProcessor()
    proc = DocumentProcessor(nlp)
    d = Document()
    d.core_properties.author = 'Кознова Мария Петровна'
    d.core_properties.title = 'Договор с ООО Ромашка'
    d.add_paragraph('Стороны заключили договор.')
    src = tmp_path / 'meta.docx'
    d.save(str(src))

    out = proc.process_file(str(src))
    cp = Document(out).core_properties
    assert not cp.author
    assert not cp.title


if __name__ == '__main__':
    # Запуск тестов
    import pytest
    pytest.main([__file__])

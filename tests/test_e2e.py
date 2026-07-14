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
    # Проверяем, что английское название и сумма без рублей поймались
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

if __name__ == '__main__':
    # Запуск тестов
    import pytest
    pytest.main([__file__])

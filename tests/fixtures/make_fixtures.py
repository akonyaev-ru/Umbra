# -*- coding: utf-8 -*-
"""Генератор PDF-фикстур для тестов. В CI НЕ запускается.

Почему PDF лежит в репозитории готовым, а не собирается тестом на лету:
кириллица в PDF требует встроенного TTF, а путей к шрифтам с кириллицей нет
ни на macOS-, ни на Linux-раннерах (встроенные шрифты reportlab — латиница).
Готовая фикстура делает тесты одинаковыми на всех трёх ОС и не тянет reportlab
в зависимости CI.

Данные в фикстуре ВЫМЫШЛЕННЫЕ. Именно поэтому для неё сделано исключение из
правила .gitignore, запрещающего класть в репозиторий документы (*.pdf).

Перегенерация (нужен reportlab и Windows-шрифт Times):
    python tests/fixtures/make_fixtures.py
"""

import os

from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

HERE = os.path.dirname(os.path.abspath(__file__))
FONT = r'C:\Windows\Fonts\times.ttf'

# ФИО намеренно разорвано переносом строки ('Иванова Ивана' / 'Ивановича') —
# так проверяется, что склейка строк в абзацы возвращает NER целое имя.
LINES = [
    'ДОГОВОР оказания услуг',
    '',
    'ООО «Ромашка» в лице генерального директора Иванова Ивана',
    'Ивановича, действующего на основании устава, с одной стороны,',
    'и Петров Пётр Петрович, с другой стороны, заключили договор.',
    '',
    'Контактный телефон: +7 999 123-45-67, e-mail: ivanov@romashka.ru',
]

TABLE = [
    ['Услуга', 'Стоимость', 'Исполнитель'],
    ['Разработка ПО', '150 000,00', 'Сидоров С.С.'],
    ['Поддержка', '50 000,00', 'Кузнецов К.К.'],
]


def build(path):
    pdfmetrics.registerFont(TTFont('Times', FONT))
    pdf = canvas.Canvas(path, pagesize=A4)
    pdf.setFont('Times', 12)

    y = 800
    for line in LINES:
        pdf.drawString(60, y, line)
        y -= 18

    y -= 20
    left, width, height = 60, 150, 20
    for row_index, row in enumerate(TABLE):
        for cell_index, value in enumerate(row):
            x = left + cell_index * width
            top = y - row_index * height - height
            pdf.rect(x, top, width, height)
            pdf.drawString(x + 4, top + 6, value)
    pdf.save()


if __name__ == '__main__':
    target = os.path.join(HERE, 'contract_ru.pdf')
    build(target)
    print('готово:', target)

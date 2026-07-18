import os
import sys
from razdel import tokenize
from navec import Navec
from slovnet import NER

def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)

import re

# Основы родовых слов, которые NER часто ошибочно принимает за ORG/PER, но
# которые НЕ являются персональными данными: роли сторон, типы организаций и
# документов, должности. Сверка идёт по началу слова (startswith), чтобы
# покрыть падежи: «Поставщик/Поставщика/Поставщику» и т.п. Спан скрывается
# только если ВСЕ его слова — из этого списка (реальное имя рядом сохранит спан).
STOP_STEMS = (
    # === РОЛИ СТОРОН ПО ВИДАМ ДОГОВОРОВ (ГК РФ) ===
    # Купля-продажа, поставка, контрактация, дарение, рента
    'продав', 'покупател', 'поставщик', 'заготовител', 'производител',
    'дарител', 'одаряем', 'получател', 'плательщик', 'рентополучател',
    # Аренда, прокат, лизинг, наём, ссуда
    'арендодател', 'арендатор', 'наймодател', 'нанимател',
    'лизингодател', 'лизингополучател', 'ссудодател', 'ссудополучател',
    # Подряд, услуги, НИОКР, проектирование
    'подрядчик', 'субподрядчик', 'генподрядчик', 'заказчик', 'исполнител',
    'проектировщик', 'изыскател', 'оценщик', 'аудитор',
    # Перевозка, транспортная экспедиция
    'перевозчик', 'отправител', 'грузоотправител', 'грузополучател',
    'пассажир', 'экспедитор', 'клиент', 'фрахтовщик', 'фрахтовател',
    # Заём, кредит, факторинг, банковские
    'займодав', 'заимодав', 'заёмщик', 'заемщик', 'кредитор', 'должник',
    'фактор', 'вкладчик', 'банк', 'брокер',
    # Хранение
    'хранител', 'поклажедател',
    # Страхование
    'страховщик', 'страховател', 'выгодоприобретател', 'застрахован',
    # Поручение, комиссия, агентирование, доверительное управление
    'доверител', 'поверенн', 'комитент', 'комиссионер', 'принципал',
    'агент', 'учредител',
    # Интеллектуальная собственность, концессия, дистрибуция
    'правооблад', 'пользовател', 'приобретател', 'лицензиар', 'лицензиат',
    'франчайз', 'дистрибьютор', 'дилер', 'автор',
    # Обеспечение обязательств (залог, поручительство, гарантия)
    'залогодател', 'залогодержател', 'поручител', 'гарант', 'бенефициар',
    # Уступка требования, перевод долга (цессия)
    'цедент', 'цессионари',
    # Инвестиции, строительство, ДДУ, ГЧП, концессии, корпоративные
    'инвестор', 'застройщик', 'дольщик', 'концедент', 'концессионер',
    'эмитент', 'акционер',
    # Энергоснабжение, связь, ЖКХ, потребители
    'энергоснабжающ', 'ресурсоснабжающ', 'абонент', 'оператор', 'потребител',
    # Трудовые, семейные и услуги физлицам
    'работодател', 'работник', 'пациент', 'обучающ', 'супруг',
    # Общие обозначения сторон
    'сторон', 'контрагент', 'участник', 'товарищ', 'депонент',
    'раскрывающ', 'принимающ', 'субагент',
    # === ОБОБЩЁННЫЕ ТИПЫ ЛИЦ / ОРГАНИЗАЦИЙ ===
    'обществ', 'компани', 'организаци', 'предприяти', 'учреждени', 'фирм',
    'корпораци', 'товариществ', 'кооператив', 'ассоциаци', 'холдинг',
    'юридическ', 'физическ', 'предпринимател',
    # === ТИПЫ ДОКУМЕНТОВ ===
    'договор', 'соглашени', 'контракт', 'приложени', 'дополнени', 'устав',
    'положени', 'регламент', 'накладн', 'спецификаци', 'протокол',
    'упд', 'счет-фактур', 'эдо', 'торг-', 'предписани', 'доверенност',
    'рекомендаци', 'руководств',
    # === ДОЛЖНОСТИ И ПРОФЕССИИ ===
    'генеральн', 'директор', 'руководител', 'президент', 'председател',
    'бухгалтер', 'управляющ', 'представител', 'сварщик', 'стропальщик',
    'электромеханик',
    # === РАСПРОСТРАНЁННЫЕ НАРИЦАТЕЛЬНЫЕ (часто пишутся с заглавной как
    #     определяемые термины и НЕ являются ПДн: «Товар», «Заказ», «Решение») ===
    'товар', 'решени', 'заказ', 'платформ', 'маркетплейс', 'склад', 'кластер',
    'кабинет', 'поставк', 'доставк', 'логистик', 'упаковк', 'сборк', 'маркировк',
    'квитанц', 'выписк', 'скриншот', 'реестр', 'оферт', 'услуг', 'продукц',
    'претензи', 'убытк', 'неустойк', 'штраф', 'госпошлин', 'пошлин', 'компенсац',
    'ущерб', 'вознагражд', 'платёж', 'платеж', 'сумм', 'стоимост', 'оборуд',
    'флипчарт', 'цветопроб', 'протирк', 'ресивер', 'убираем', 'гостиничн',
    'инженерн', 'корреспонденци', 'приёмк', 'приемк', 'транспортирован',
    # Заголовки, юридические и общие термины, часто пишущиеся капсом
    'студи', 'помещени', 'квартир', 'подписан', 'услови', 'подпис', 
    'обязанност', 'ответственност', 'проч', 'реквизит', 'закон', 'кодекс',
    # Аббревиатуры кодексов и законов
    'гк', 'гпк', 'нк', 'ук', 'коап', 'фз', 'апк',
)

_LETTER_RE = re.compile(r'[a-zA-Zа-яёА-ЯЁ]')
_WORD_RE = re.compile(r'[a-zA-Zа-яёА-ЯЁ]+')

# Метки-заменители по категориям.
TAG_PER = '[ФИО]'
TAG_ORG = '[ОРГАНИЗАЦИЯ]'
TAG_PROJECT = '[ПРОЕКТ]'
TAG_ADDR = '[АДРЕС]'
TAG_EMAIL = '[EMAIL]'
TAG_PHONE = '[ТЕЛЕФОН]'
TAG_MONEY = '[СУММА]'
TAG_ACCOUNT = '[СЧЁТ]'
TAG_INN_UL = '[ИНН_ЮЛ]'
TAG_INN_FL = '[ИНН_ФЛ]'
TAG_OGRN = '[ОГРН]'
TAG_KPP = '[КПП]'
TAG_INDEX = '[ИНДЕКС]'
TAG_DOCNO = '[НОМЕР]'
TAG_BIK = '[БИК]'
TAG_SWIFT = '[SWIFT]'
TAG_PASSPORT = '[ПАСПОРТ]'
TAG_PASSPORT_ISSUED = '[ВЫДАН]'
TAG_PODRAZD = '[КОД_ПОДР]'
TAG_SNILS = '[СНИЛС]'
TAG_DATE_BIRTH = '[ДР]'
TAG_PLACE_BIRTH = '[МЕСТО_РОЖДЕНИЯ]'
TAG_DATE_ISSUE = '[ДАТА_ВЫДАЧИ]'
TAG_KADASTR = '[КАДАСТР]'
TAG_EGRN = '[ЕГРН]'
TAG_DATE = '[ДАТА]'

# Регулярное выражение для всех дат (вида DD.MM.YYYY, YYYY-MM-DD и т.д.)
_DATE_RE = re.compile(r'\b(?:[0-3]?\d[\./\-][01]?\d[\./\-](?:19|20)?\d{2}|(?:19|20)\d{2}[\./\-][01]?\d[\./\-][0-3]?\d)\b')

# Границы предложений для адресных паттернов
_SENT_END = (
    r'(?<!\bг)(?<!\bул)(?<!\bпр)(?<!\bпл)(?<!\bш)(?<!\bнаб)(?<!\bпер)(?<!\bбульв)'
    r'(?<!\bд)(?<!\bдом)(?<!\bстр)(?<!\bкорп)(?<!\bк)(?<!\bкв)(?<!\bофис)(?<!\bпом)'
    r'(?<!\bэт)(?<!\bобл)(?<!\bпос)(?<!\bп)(?<!\bс)(?<!\bдер)(?<!\bмкр)(?<!\bим)'
    r'(?<!\b[А-ЯЁA-Z])'
    r'\.(?=\s+[А-ЯЁA-Z\[]|\s*$)'
)
_ADDR_CHUNK = r'(?:(?!' + _SENT_END + r')[^\n,])+'


def _is_stop_word(w):
    """True, если слово — родовой термин (роль/тип/должность) из STOP_STEMS."""
    wl = w.lower()
    return any(wl.startswith(stem) for stem in STOP_STEMS)


# Обозначения страны/государства-в-целом: не ПДн, а указание применимого права
# («ГК РФ», «право Российской Федерации»). Проверяется ТОЛЬКО для гео-спанов (LOC),
# поэтому фамилии (спаны PER) эти корни не затрагивают — риска утечки нет.
_COUNTRY_STEMS = ('рф', 'росс', 'федерац')


def _is_country_ref(s):
    """True, если гео-спан целиком состоит из слов-обозначений страны/государства."""
    words = _WORD_RE.findall(s)
    return bool(words) and all(
        any(w.lower().startswith(stem) for stem in _COUNTRY_STEMS) for w in words)


def _appears_lowercase(token, text):
    """True, если слово встречается в тексте со строчной буквы — значит это
    нарицательное слово (напр. «решения»), а не имя собственное. Такие слова
    НЕ протягиваем по документу, чтобы не скрыть, например, «Решение суда»."""
    for m in re.finditer(r'\b' + re.escape(token) + r'\b', text, re.IGNORECASE):
        if text[m.start()].islower():
            return True
    return False


class PlaceholderMapper:
    """Присваивает уникальные нумерованные метки ([ФИО_1], [ОРГАНИЗАЦИЯ_2]…) и
    хранит карту {метка: исходное значение}. Одинаковое значение в одной
    категории всегда получает одну и ту же метку — для консистентности анонимного
    документа и точного восстановления."""

    def __init__(self):
        self.mapping = {}          # метка -> исходное значение
        self._value_to_ph = {}     # (категория, значение) -> метка
        self._counters = {}        # категория -> счётчик

    def placeholder(self, category, value):
        key = (category, value)
        ph = self._value_to_ph.get(key)
        if ph is not None:
            return ph
        n = self._counters.get(category, 0) + 1
        self._counters[category] = n
        ph = category[:-1] + '_' + str(n) + ']'   # '[ФИО]' -> '[ФИО_1]'
        self._value_to_ph[key] = ph
        self.mapping[ph] = value
        return ph


class TransformerSpan:
    def __init__(self, start, stop, type_):
        self.start = start
        self.stop = stop
        self.type = type_

class TransformerMarkup:
    def __init__(self, spans):
        self.spans = spans

class NLPProcessor:
    def __init__(self):
        self._compile_regexes()
        self._ner_cache = {}
        
        try:
            from transformers import pipeline
            self.hf_ner = pipeline("ner", model="Babelscape/wikineural-multilingual-ner", aggregation_strategy="simple")
            self.use_hf = True
        except ImportError:
            self.use_hf = False
            navec_path = get_resource_path(os.path.join("models", "navec_news_v1_1B_250K_300d_100q.tar"))
            slovnet_path = get_resource_path(os.path.join("models", "slovnet_ner_news_v1.tar"))
            
            if not os.path.exists(navec_path) or not os.path.exists(slovnet_path):
                raise FileNotFoundError(f"Models not found. Checked: {navec_path}, {slovnet_path}")

            self.navec = Navec.load(navec_path)
            self.ner = NER.load(slovnet_path)
            self.ner.navec(self.navec)

    def _ner_markup(self, text):
        """NER-разметка текста с мемоизацией (см. self._ner_cache)."""
        markup = self._ner_cache.get(text)
        if markup is None:
            if getattr(self, 'use_hf', False):
                hf_result = self.hf_ner(text)
                spans = []
                for ent in hf_result:
                    group = ent['entity_group']
                    if group in ('PER', 'ORG', 'LOC'):
                        spans.append(TransformerSpan(ent['start'], ent['end'], group))
                markup = TransformerMarkup(spans)
            else:
                markup = self.ner(text)
            self._ner_cache[text] = markup
        return markup

    def clear_ner_cache(self):
        """Сбрасывает кеш NER. Вызывается перед обработкой нового документа,
        чтобы разметка предыдущих документов не копилась в памяти в сессии GUI."""
        self._ner_cache.clear()

    def _compile_regexes(self):
        # Разделители тысяч в суммах: обычный, неразрывный (U+00A0) и узкий
        # неразрывный (U+202F) пробелы — строим через chr() (без невидимых
        # символов в исходнике).
        sep = ' ' + chr(0xA0) + chr(0x202F)
        # Число суммы: цифра, затем «(необяз. разделитель)+цифра», но НЕ более 20
        # повторов. Прежний r'\d(?:[\d.,sep]*\d)?' был неоднозначен и давал
        # квадратичный бэктрекинг — строка из сплошных сумм (31 КБ) считалась ~90 с
        # и подвешивала интерфейс. Ограничение длины (реальная сумма короче) не даёт
        # «числу» растянуться на всю строку и держит время линейным.
        num = r'\d(?:[.,' + sep + r']?\d){0,20}'
        cur = (r'рублей|рубля|рубль|руб\.|руб|копеек|копейки|копейка|коп\.|коп|'
               r'долларов|доллара|доллар|долл\.|USD|у\.\s?е\.|евро|EUR|\$|€')
               
        self.regexes = [
            # 1. Проекты и Организации в кавычках (специфичные паттерны)
            (re.compile(r'(?:проект|объект|программа|инициатива)\s+([«"„][\wА-Яа-яёЁ\s\-\.]+[»""»])'), TAG_PROJECT, 1),
            (re.compile(r'(?:ООО|ОАО|ЗАО|АО|ПАО|НАО|ГУП|МУП|ФГУП|НКО|АНО|КФХ|ПК|ТСЖ|НП|Банк|Фонд)\s+([«"„][\wА-Яа-яёЁ\s\-\.]+[»""»])'), TAG_ORG, 1),
            (re.compile(r'(ИП\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]\.)'), TAG_PER, 1),
            (re.compile(r'(ИП\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)'), TAG_PER, 1),

            # 2. Недвижимость
            (re.compile(r'\b(\d{2}:\d{2}:\d{6,7}:\d{1,5}\s*-\s*\d{2}/\d{3}/\d{4}-\d+)\b'), TAG_EGRN, 1),
            (re.compile(r'\b(\d{2}:\d{2}:\d{6,7}:\d{1,5})\b'), TAG_KADASTR, 1),
            
            # 3. Паспорт и личные документы
            (re.compile(r'(?:выдан\s*:?\s*)((?:ГУ|УМВД|МВД|ОВД|УФМС|ФМС|ОУФМС|ТП|ОТДЕЛЕНИЕ|ОТДЕЛ)[^,;\n]{5,120}?)(?=,\s*(?:[Дд]ата|код\s+подразделения)|\s*$)'), TAG_PASSPORT_ISSUED, 1),
            (re.compile(r'(?:серия\s+)(\d{4}\s*(?:№|No|N|no|Но)?\s*\d{6})'), TAG_PASSPORT, 1),
            # Паспорт без слова «серия»: «паспорт 4509 123456», «паспорт РФ 45 09 123456».
            (re.compile(r'(?:(?i:паспорт(?:\s+гражданина)?(?:\s+рф)?)\s*[:№]?\s*)(\d{2}\s?\d{2}\s*(?:№|No|N|no|Но)?\s*\d{6})'), TAG_PASSPORT, 1),
            # СНИЛС по официальным синонимам, а не только по слову «СНИЛС» (регистр триггера не важен).
            (re.compile(r'(?:(?i:СНИЛС|страхов\w*\s+свидетельств\w*|страхов\w*\s+номер\w*)\s*[№:]?\s*)(\d{3}[-\s]?\d{3}[-\s]?\d{3}[-\s]?\d{2}|\d{11})'), TAG_SNILS, 1),
            (re.compile(r'(?:код\s+подразделения\s*[:№]?\s*)(\d{3}[-\s]?\d{3})'), TAG_PODRAZD, 1),
            (re.compile(r'(\d{2}\.\d{2}\.\d{4})(?=\s*(?:года\s+рождения|г\.\s*р\.))'), TAG_DATE_BIRTH, 1),
            # Обратный порядок: «Дата рождения: 15.06.1985» (регистр триггера не важен).
            (re.compile(r'(?:(?i:дата\s+рождения)\s*:?\s*)(\d{2}\.\d{2}\.\d{4})'), TAG_DATE_BIRTH, 1),
            (re.compile(r'(?:[Дд]ата\s+выдачи\s*:?\s*)(\d{2}\.\d{2}\.\d{4})'), TAG_DATE_ISSUE, 1),
            (re.compile(r'(?:место\s+рождения\s*:\s*)([^,\n]{3,200}?)(?=,\s*(?:паспорт|серия)|\s*$)'), TAG_PLACE_BIRTH, 1),
            
            # 4. Адреса (расширенные с границами предложений)
            (re.compile(r'(?:зарегистрирован[а-яё]*\s+по\s+адресу\s*:\s*)([^\n]+?)(?=,\s*(?:СНИЛС|паспорт|серия|телефон)|\s*\(далее|' + _SENT_END + r'|\s*$)'), TAG_ADDR, 1),
            (re.compile(r'(?:адрес\s+регистрации\s*:\s*)([^\n]+?)(?=,\s*(?:СНИЛС|паспорт|серия|телефон)|\s*\(далее|' + _SENT_END + r'|\s*$)'), TAG_ADDR, 1),
            (re.compile(r'(?:(?:юридический|почтовый|фактический|адрес проживания|адрес)\s*(?:адрес)?\s*:\s*)((?:г(?:ород|\.)?\s+)?[А-ЯЁа-яёA-Za-z][^\n]{10,200}?)(?=,\s*(?:ИНН|ОГРН|СНИЛС|паспорт|телефон|расчётный|р/с|БИК)|' + _SENT_END + r'|\s*$)'), TAG_ADDR, 1),
            (re.compile(r'(\bг\.?\s+[А-ЯЁ][а-яё]+\s*,\s*(?:ул|пр|наб|пер|бульв|ш|пл)\.?\s+[А-ЯЁа-яё]' + _ADDR_CHUNK + r'(?:\s*,\s*(?:д|дом|стр|корп|к|кв|офис|пом|эт)\.?\s*' + _ADDR_CHUNK + r')+)'), TAG_ADDR, 1),
            
            # 5. Финансовые/налоговые идентификаторы (в порядке убывания длины)
            (re.compile(r'(?:р/с|к/с|расч[её]тный\s+счёт|расч[её]тный\s+счет|счёт|счет)\s*[№:\s]*(\d{20})'), TAG_ACCOUNT, 1),
            (re.compile(r'\b(\d{20})\b'), TAG_ACCOUNT, 1),
            (re.compile(r'(?:ОГРН[ИП]?\s*[:/№]?\s*)(\d{13,15})'), TAG_OGRN, 1),
            (re.compile(r'\b(\d{15})\b'), TAG_OGRN, 1),
            (re.compile(r'\b(\d{13})\b'), TAG_OGRN, 1),
            (re.compile(r'(?:ИНН(?:[\s\w()]{0,80}?)?\s*[:/№]?\s*)(\d{12})(?!\d)'), TAG_INN_FL, 1),
            (re.compile(r'\b(\d{12})\b'), TAG_INN_FL, 1),
            (re.compile(r'(?:ИНН(?:[\s\w()]{0,80}?)?\s*[:/№]?\s*)(\d{10})(?!\d)'), TAG_INN_UL, 1),
            (re.compile(r'\b(\d{10})\b'), TAG_INN_UL, 1),
            (re.compile(r'(?:КПП\s*[:/№]?\s*)(\d{9})'), TAG_KPP, 1),
            (re.compile(r'(?:БИК(?:[\s\w]{0,80}?)?\s*[:/№]?\s*)(\d{9})'), TAG_BIK, 1),
            (re.compile(r'\b(04\d{7})\b'), TAG_BIK, 1), # БИК всегда с 04
            (re.compile(r'(?:SWIFT|свифт)\s*[:/]?\s*([A-Z]{6}[A-Z0-9]{2,5})'), TAG_SWIFT, 1),

            # 6. Контакты
            # E-mail, включая кириллические домены (IDN, напр. «почта.рф») и локальную часть.
            (re.compile(r'(?<![a-zA-Zа-яёА-ЯЁ0-9_.+-])([a-zA-Zа-яёА-ЯЁ0-9_.+-]{1,64}@[a-zA-Zа-яёА-ЯЁ0-9-]{1,63}(?:\.[a-zA-Zа-яёА-ЯЁ0-9-]{1,63}){1,8})(?![a-zA-Zа-яёА-ЯЁ0-9_.+-])'), TAG_EMAIL, 1),
            # Городской формат со скобками, в т.ч. без кода страны: «(495) 123-45-67», «(49232) 1-23-45».
            (re.compile(r'(?<!\d)(\(\d{3,5}\)[\s\-]?\d{1,3}[\s\-]?\d{2}[\s\-]?\d{2})(?!\d)'), TAG_PHONE, 1),
            # Городской номер 3-2-2 через дефис без кода страны: «123-45-67» (не часть более длинного кода).
            (re.compile(r'(?<![\d-])(\d{3}-\d{2}-\d{2})(?![\d-])'), TAG_PHONE, 1),
            (re.compile(r'(?<!\d)((?:\+?7|8)[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2})(?!\d)'), TAG_PHONE, 1),
            
            # 7. Денежные суммы и номера документов (Оставлено из Umbra)
            # Номер документа после «№». Хвостовые «.» / «-» / «/» НЕ захватываем,
            # чтобы не проглотить точку конца предложения («Спецификации № 1.»).
            (re.compile(r'(№\s*)([\w./\-]*\d[\w/\-]*)'), TAG_DOCNO, 2),
            # Число обёрнуто в АТОМАРНУЮ группу (?>...): без неё, когда за суммой
            # нет слова валюты, движок откатывал «число» посимвольно на каждой
            # стартовой позиции — квадратичное время на длинных рядах сумм.
            (re.compile(r'(?i)((?>' + num + r')(?:\s*\([^)]*\))?)\s*(' + cur + r')(?![а-яёА-ЯЁa-z])'), TAG_MONEY, 1),
            (re.compile(r'(?i)(\$|€|USD|EUR)\s*(' + num + r')'), TAG_MONEY, 2),
            (re.compile(r'(?<!\d)(\d{1,3}(?:[ \xA0\u202F]\d{3})+(?:[,.]\d{2})?)(?!\d)'), TAG_MONEY, 1),
            # Десятичная сумма. Хвост «(?!\.\d)» не даёт спутать её с префиксом даты:
            # в «27.07.2006» фрагмент «27.07» больше НЕ считается суммой.
            (re.compile(r'(?<!\d)(\d+[,.]\d{2})(?!\d)(?!\.\d)'), TAG_MONEY, 1),
            
            # 8. Английские названия и fallback номера
            (re.compile(r'\b(\d{11})\b'), TAG_DOCNO, 1),
            (re.compile(r'\b(\d{14})\b'), TAG_DOCNO, 1),
            (re.compile(r'\b(\d{6})\b'), TAG_INDEX, 1),
            (re.compile(r'(?<![a-zA-Zа-яёА-ЯЁ])([A-Z][a-zA-Z0-9-]{2,}(?:\s+[A-Z][a-zA-Z0-9-]{2,})*)(?![a-zA-Zа-яёА-ЯЁ])'), TAG_ORG, 1),
        ]

    def _keep_ner_span(self, span_text):
        """Решает, оставлять ли NER-спан (True — скрывать сущность).

        Отсекает два вида ложных срабатываний:
        1. Спан без единой буквы (чистое число/пунктуация, напр. «29») — это не
           имя; числовые ПДн ловятся отдельными regex.
        2. Спан, все слова которого — родовые термины из STOP_STEMS (роль,
           тип организации/документа, должность), напр. «Поставщик», «Общество».
        """
        if not _LETTER_RE.search(span_text):
            return False
        words = _WORD_RE.findall(span_text)
        if words and all(_is_stop_word(w) for w in words):
            return False
        return True

    def collect_terms(self, text):
        """Возвращает (surfaces, tokens) — словари {текст: метка} имён собственных
        (PER→[ФИО], ORG→[ОРГАНИЗАЦИЯ]) для сквозной протяжки. surfaces —
        многословные фразы, tokens — отдельные значимые слова (фамилии, названия)."""
        surfaces, tokens = {}, {}
        if not text or not text.strip():
            return surfaces, tokens
        markup = self._ner_markup(text)
        for span in markup.spans:
            if span.type not in ('PER', 'ORG'):
                continue
            s = text[span.start:span.stop]
            if not self._keep_ner_span(s):
                continue
            tag = TAG_PER if span.type == 'PER' else TAG_ORG
            s2 = s.strip(' \t\r\n«»"\'()')
            if ' ' in s2 and len(s2) >= 5:
                surfaces.setdefault(s2, tag)
            # Отдельные значимые слова (фамилии, названия). Нарицательные
            # отсекаются двумя фильтрами: STOP_STEMS здесь (напр. «Решения»,
            # «Товар») и _appears_lowercase при сборке паттернов (напр. «Интернет»).
            for w in _WORD_RE.findall(s2):
                if len(w) >= 4 and not _is_stop_word(w):
                    # Приоритет [ФИО] над [ОРГАНИЗАЦИЯ], если слово встретилось в обоих.
                    if w not in tokens or tag == TAG_PER:
                        tokens[w] = tag
        return surfaces, tokens

    def build_extra_patterns(self, surfaces, tokens, full_text):
        """Строит [(regex, метка)] для протяжки известных имён/названий по всему
        документу (ловит пропуски NER и падежные формы). Нарицательные слова
        (встречаются со строчной буквы) исключаются."""
        patterns = []
        for s, tag in sorted(surfaces.items(), key=lambda kv: len(kv[0]), reverse=True):
            patterns.append((re.compile(re.escape(s), re.IGNORECASE), tag))
        for t, tag in tokens.items():
            if _appears_lowercase(t, full_text):
                continue
            # Корень (без 1–2 последних букв) + любое окончание → падежные формы.
            root = t if len(t) <= 5 else t[:max(5, len(t) - 2)]
            # re.IGNORECASE обязателен: без него фамилия, набранная ЗАГЛАВНЫМИ
            # (шапки, блоки подписей «ДИРЕКТОР ИВАНОВ И.И.» — сплошь в юрдокументах),
            # не совпадёт с корнем из title-case NER-спана и утечёт в открытом виде.
            # Нарицательные уже отсеяны _appears_lowercase выше, поэтому расширение
            # на все регистры безопасно.
            patterns.append((re.compile(r'\b' + re.escape(root) + r'[а-яёА-ЯЁ]*', re.IGNORECASE), tag))
        return patterns

    def extract_entities(self, text, hide_names=True, hide_locations=True,
                         hide_orgs=True, hide_dates=False, extra_patterns=None):
        """
        Runs NER on the text and returns a list of dictionaries with spans and replacement tags.
        extra_patterns — доп. regex (протяжка имён по документу), их совпадения тоже скрываются.
        """
        if not text or not text.strip():
            return []

        # A failed NER pass is a privacy boundary failure.  Continuing with only
        # regexes would silently leak names and organisations.
        markup = self._ner_markup(text)

        replacements = []
        for span in markup.spans:
            category = None
            if span.type == 'PER' and hide_names:
                category = TAG_PER
            elif span.type == 'LOC' and hide_locations:
                # «РФ», «Российская Федерация» — указание юрисдикции, не адрес.
                if not _is_country_ref(text[span.start:span.stop]):
                    category = TAG_ADDR
            elif span.type == 'ORG' and hide_orgs:
                category = TAG_ORG

            if category and self._keep_ner_span(text[span.start:span.stop]):
                replacements.append({'start': span.start, 'stop': span.stop, 'category': category})

        # Regex: скрываем только ЧУВСТВИТЕЛЬНУЮ группу (остальное — «руб.», «№»,
        # «д.» — остаётся в тексте).
        for pattern, category, group in self.regexes:
            for match in pattern.finditer(text):
                s, e = match.span(group)
                if s >= 0 and text[s:e].strip():
                    replacements.append({'start': s, 'stop': e, 'category': category})
                    
        if hide_dates:
            for match in _DATE_RE.finditer(text):
                s, e = match.start(), match.end()
                if s >= 0 and text[s:e].strip():
                    replacements.append({'start': s, 'stop': e, 'category': TAG_DATE})

        # Сквозная протяжка: имена/названия, пропущенные NER в этом абзаце.
        if extra_patterns:
            for pattern, category in extra_patterns:
                for match in pattern.finditer(text):
                    if match.group(0).strip():
                        replacements.append({'start': match.start(), 'stop': match.end(), 'category': category})

        # Разрешение перекрытий ОБЪЕДИНЕНИЕМ (union), а не отбрасыванием: если
        # два спана перекрываются «ступенькой», просто выкинуть один означало бы
        # оставить его непокрытый край в открытом виде (утечка ПДн). Поэтому
        # перекрывающиеся спаны сливаются в один максимальный интервал; категорию
        # берём у спана с меньшим start (детерминизм — важно для точного round-trip).
        # Для непересекающихся спанов (подавляющее большинство) результат прежний.
        replacements.sort(key=lambda x: (x['start'], -x['stop']))
        merged = []
        for rep in replacements:
            if merged and rep['start'] < merged[-1]['stop']:
                if rep['stop'] > merged[-1]['stop']:
                    merged[-1]['stop'] = rep['stop']
            else:
                merged.append(rep)
        # К каждой замене добавляем 'text' — исходную скрываемую подстроку (для
        # карты замен). Наружу отдаём справа-налево: вызывающий применяет замены
        # с конца, чтобы не сдвигать индексы.
        for rep in merged:
            rep['text'] = text[rep['start']:rep['stop']]
        merged.reverse()
        return merged

    def anonymize_text(self, text, hide_names=True, hide_locations=True, hide_orgs=True,
                       hide_dates=False, extra_patterns=None, mapper=None):
        """Анонимизирует строку. С mapper — нумерованные метки ([ФИО_1]…) +
        наполнение карты замен; без mapper — категорийные метки ([ФИО])."""
        replacements = self.extract_entities(text, hide_names, hide_locations, hide_orgs,
                                             hide_dates=hide_dates, extra_patterns=extra_patterns)
        # Метки-номера присваиваем в порядке документа (слева направо), а
        # применяем справа-налево, чтобы не сдвигать индексы.
        if mapper is not None:
            for rep in reversed(replacements):
                rep['ins'] = mapper.placeholder(rep['category'], rep['text'])
        else:
            for rep in replacements:
                rep['ins'] = rep['category']

        res_text = text
        for rep in replacements:
            res_text = res_text[:rep['start']] + rep['ins'] + res_text[rep['stop']:]
        return res_text

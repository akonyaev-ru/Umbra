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
    except Exception:
        base_path = os.path.abspath(".")
    
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
    'даритель', 'одаряем', 'получател', 'плательщик', 'рентополучател',
    # Аренда, прокат, лизинг, наём, ссуда
    'арендодател', 'арендатор', 'наймодател', 'наниматель',
    'лизингодател', 'лизингополучател', 'ссудодател', 'ссудополучател',
    # Подряд, услуги, НИОКР, проектирование
    'подрядчик', 'субподрядчик', 'генподрядчик', 'заказчик', 'исполнител',
    'проектировщик', 'изыскател',
    # Перевозка, транспортная экспедиция
    'перевозчик', 'отправител', 'грузоотправител', 'грузополучател',
    'пассажир', 'экспедитор', 'клиент',
    # Заём, кредит, факторинг, банковские
    'займодав', 'заимодав', 'заёмщик', 'заемщик', 'кредитор', 'должник',
    'фактор', 'вкладчик',
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
    # Инвестиции, строительство, ДДУ, ГЧП, концессии
    'инвестор', 'застройщик', 'дольщик', 'концедент', 'концессионер',
    # Энергоснабжение, связь, ЖКХ, потребители
    'энергоснабжающ', 'ресурсоснабжающ', 'абонент', 'оператор', 'потребител',
    # Трудовые и услуги физлицам
    'работодател', 'работник', 'пациент', 'обучающ',
    # Общие обозначения сторон
    'сторон', 'контрагент', 'участник', 'товарищ', 'депонент',
    # === ОБОБЩЁННЫЕ ТИПЫ ЛИЦ / ОРГАНИЗАЦИЙ ===
    'обществ', 'компани', 'организаци', 'предприяти', 'учреждени', 'фирм',
    'корпораци', 'товариществ', 'кооператив', 'ассоциаци', 'холдинг',
    'юридическ', 'физическ', 'предпринимател',
    # === ТИПЫ ДОКУМЕНТОВ ===
    'договор', 'соглашени', 'контракт', 'приложени', 'дополнени', 'устав',
    'положени', 'регламент', 'накладн', 'спецификаци', 'протокол',
    # === ДОЛЖНОСТИ ===
    'генеральн', 'директор', 'руководител', 'президент', 'председател',
    'бухгалтер', 'управляющ', 'представител',
    # === РАСПРОСТРАНЁННЫЕ НАРИЦАТЕЛЬНЫЕ (часто пишутся с заглавной как
    #     определяемые термины и НЕ являются ПДн: «Товар», «Заказ», «Решение») ===
    'товар', 'решени', 'заказ', 'платформ', 'маркетплейс', 'склад', 'кластер',
    'кабинет', 'поставк', 'доставк', 'логистик', 'упаковк', 'сборк', 'маркировк',
    'квитанц', 'выписк', 'скриншот', 'реестр', 'оферт', 'услуг', 'продукц',
    'претензи', 'убытк', 'неустойк', 'штраф', 'госпошлин', 'пошлин', 'компенсац',
    'ущерб', 'вознагражд', 'платёж', 'платеж', 'сумм',
)

_LETTER_RE = re.compile(r'[a-zA-Zа-яёА-ЯЁ]')
_WORD_RE = re.compile(r'[a-zA-Zа-яёА-ЯЁ]+')

# Метки-заменители по категориям (вместо единого [СКРЫТО]).
TAG_PER = '[ФИО]'
TAG_ORG = '[ОРГАНИЗАЦИЯ]'
TAG_ADDR = '[АДРЕС]'         # география (LOC) + компоненты адреса
TAG_EMAIL = '[EMAIL]'
TAG_PHONE = '[ТЕЛЕФОН]'
TAG_MONEY = '[СУММА]'
TAG_ACCOUNT = '[СЧЁТ]'
TAG_INN = '[ИНН]'
TAG_OGRN = '[ОГРН]'
TAG_KPP = '[КПП]'
TAG_INDEX = '[ИНДЕКС]'
TAG_DOCNO = '[НОМЕР]'


def _is_stop_word(w):
    """True, если слово — родовой термин (роль/тип/должность) из STOP_STEMS."""
    wl = w.lower()
    return any(wl.startswith(stem) for stem in STOP_STEMS)


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


class NLPProcessor:
    def __init__(self):
        navec_path = get_resource_path(os.path.join("models", "navec_news_v1_1B_250K_300d_100q.tar"))
        slovnet_path = get_resource_path(os.path.join("models", "slovnet_ner_news_v1.tar"))
        
        if not os.path.exists(navec_path) or not os.path.exists(slovnet_path):
            raise FileNotFoundError(f"Models not found. Checked: {navec_path}, {slovnet_path}")

        self.navec = Navec.load(navec_path)
        self.ner = NER.load(slovnet_path)
        self.ner.navec(self.navec)

        # Кеш NER-разметки по тексту абзаца. Документ обходится дважды (сбор имён
        # + анонимизация), и NER — самая дорогая операция. NER детерминирован
        # (один текст → одна разметка), поэтому мемоизация НЕ меняет результат, а
        # лишь убирает повторный прогон одного и того же абзаца во втором проходе.
        # Живёт в пределах обработки одного документа (clear_ner_cache).
        self._ner_cache = {}

        self._compile_regexes()

    def _ner_markup(self, text):
        """NER-разметка текста с мемоизацией (см. self._ner_cache)."""
        markup = self._ner_cache.get(text)
        if markup is None:
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
        num = r'\d(?:[\d.,' + sep + r']*\d)?'
        cur = (r'рублей|рубля|рубль|руб\.|руб|копеек|копейки|копейка|коп\.|коп|'
               r'долларов|доллара|доллар|долл\.|USD|у\.\s?е\.|евро|EUR|\$|€')
        # Кортежи (pattern, категория, группа). «группа» — номер capture-группы
        # с ЧУВСТВИТЕЛЬНОЙ частью (её скрываем/нумеруем); остальное (руб., №,
        # «д.») остаётся в тексте. 0 = вся совпавшая строка.
        self.regexes = [
            (re.compile(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'), TAG_EMAIL, 0),
            (re.compile(r'(?<!\d)(?:\+?7|8)[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}(?!\d)'), TAG_PHONE, 0),
            # Номер документа после № — код (группа 2), «№ » остаётся.
            (re.compile(r'(№\s*)([\w./\-]*\d[\w./\-]*)'), TAG_DOCNO, 2),
            # Денежная сумма перед валютой: число + необязательная сумма прописью
            # в скобках (группа 1 — скрываем целиком, иначе пропись выдала бы сумму),
            # валюта остаётся. Пример: «110 500 (Сто десять тысяч пятьсот) рублей».
            (re.compile(r'(?i)(' + num + r'(?:\s*\([^)]*\))?)\s*(' + cur + r')(?![а-яёА-ЯЁa-z])'), TAG_MONEY, 1),
            # Валюта перед суммой: число (группа 2).
            (re.compile(r'(?i)(\$|€|USD|EUR)\s*(' + num + r')'), TAG_MONEY, 2),
            # Идентификаторы по числу цифр.
            (re.compile(r'\b\d{20}\b'), TAG_ACCOUNT, 0),
            (re.compile(r'\b\d{15}\b'), TAG_OGRN, 0),
            (re.compile(r'\b\d{13}\b'), TAG_OGRN, 0),
            (re.compile(r'\b\d{12}\b'), TAG_INN, 0),
            (re.compile(r'\b\d{10}\b'), TAG_INN, 0),
            (re.compile(r'\b\d{9}\b'), TAG_KPP, 0),
            (re.compile(r'\b\d{11}\b'), TAG_DOCNO, 0),
            (re.compile(r'\b\d{14}\b'), TAG_DOCNO, 0),
            (re.compile(r'\b\d{6}\b'), TAG_INDEX, 0),
            # Адрес: обозначение (группа 2 — число/буква/римская цифра), ключевое слово остаётся.
            (re.compile(r'(?i)(дом|д\.|корп\.|корпус|пом\.|помещение|помещ\.|комната|комн\.|ком\.|офис|оф\.|этаж|эт\.|кв\.|квартира|строение|стр\.|литера|лит\.|подъезд)\s*([IVXLCDM]{1,7}(?![A-Za-zА-Яа-яёЁ0-9])|\d+[а-яА-Яa-zA-Z]?(?:\s*[/-]\s*\d+[а-яА-Я]?)?)'), TAG_ADDR, 2),
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
        try:
            markup = self._ner_markup(text)
        except Exception:
            return surfaces, tokens
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
                         hide_orgs=True, extra_patterns=None):
        """
        Runs NER on the text and returns a list of dictionaries with spans and replacement tags.
        extra_patterns — доп. regex (протяжка имён по документу), их совпадения тоже скрываются.
        """
        if not text or not text.strip():
            return []

        # NER — самая хрупкая часть конвейера. Если она упадёт на конкретном
        # фрагменте, НЕ отбрасываем весь абзац: regex-слой (телефоны, e-mail,
        # счета, суммы, идентификаторы) и сквозная протяжка имён работают без NER
        # и обязаны отработать всё равно — иначе НЕанонимизированный текст молча
        # утёк бы. (Под --noconsole sys.stdout is None, поэтому не логируем.)
        try:
            markup = self._ner_markup(text)
        except Exception:
            markup = None

        replacements = []
        if markup is not None:
            for span in markup.spans:
                category = None
                if span.type == 'PER' and hide_names:
                    category = TAG_PER
                elif span.type == 'LOC' and hide_locations:
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
                       extra_patterns=None, mapper=None):
        """Анонимизирует строку. С mapper — нумерованные метки ([ФИО_1]…) +
        наполнение карты замен; без mapper — категорийные метки ([ФИО])."""
        replacements = self.extract_entities(text, hide_names, hide_locations, hide_orgs,
                                             extra_patterns=extra_patterns)
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

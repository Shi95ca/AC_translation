#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import sys
import copy
import json
import time
import urllib.request
import urllib.parse
from datetime import datetime
import xlrd
import openpyxl
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(SCRIPT_DIR, 'custom_courses_cache.tsv')
INSPECTORS_FILE = os.path.join(SCRIPT_DIR, 'inspectors.txt')
COURSES_FILE = os.path.join(SCRIPT_DIR, 'courses1.txt')

DEFAULT_INSPECTORS_CONTENT = """# База инспекторов учебной части Физического факультета МГУ
# Формат: Курс = Инициалы и Фамилия (на английском или на русском)
1 = N.O. Ivanova        # Иванова Наталья Олеговна (1 курс)
2 = N.A. Barsukova      # Барсукова Наталья Алексеевна (2 курс)
3 = V.V. Karminskaya    # Карминская Виктория Владимировна (3 курс)
4 = N.A. Nesterova      # Нестерова Наталья Александровна (4 курс)
5 = I.V. Khomulo        # Хомуло Ирина Вадимовна (5 курс)
6 = T.S. Nesterenko     # Нестеренко Татьяна Сергеевна (6 курс)
mag1 = T.S. Nesterenko  # 1 курс магистратуры
mag2 = T.S. Nesterenko  # 2 курс магистратуры
"""

TRANSLIT_DICT = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e',
    'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'i', 'к': 'k', 'л': 'l', 'м': 'm',
    'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
    'ф': 'f', 'х': 'kh', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'shch',
    'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'iu', 'я': 'ia'
}
for _k, _v in list(TRANSLIT_DICT.items()):
    TRANSLIT_DICT[_k.upper()] = _v.capitalize()
TRANSLIT_TABLE = str.maketrans(TRANSLIT_DICT)


def translit_ru(text):
    return str(text).translate(TRANSLIT_TABLE)


# Исправление неточностей в исходном courses1.txt и базовые общефакультетские курсы
OVERRIDE_COURSES = {
    'научно-исследовательский семинар по современным проблемам физики': 'Research seminar on modern problems of physics',
}

BASE_COURSES = {
    'программирование и информатика': 'Programming and informatics',
    'введение в компьютерные технологии': 'Introduction to computer technologies',
    'введение в технику физического эксперимента': 'Introduction to experimental methods of Physics',
    'общий физический практикум': 'General Physics practicum',
    'математический анализ': 'Mathematical analysis',
    'английский язык': 'English',
    'общая астрономия': 'General astronomy',
    'математическая обработка наблюдений': 'Mathematical processing of observation data',
    'безопасность жизнедеятельности': 'Life safety',
    'физическая культура': 'Physical education',
    'элективные курсы по физической культуре': 'Elective courses in physical education',
    'механика': 'Mechanics',
    'аналитическая геометрия': 'Analytical geometry',
    'молекулярная физика': 'Molecular physics',
    'русский язык и культура речи': 'Russian language and the culture of speech',
    'линейная алгебра': 'Linear algebra',
    'сферическая астрономия': 'Spherical astronomy',
    'теория функций комплексной переменной': 'Function theory of complex variable',
    'электромагнетизм': 'Electricity and magnetism',
    'гравиметрия': 'Gravimetry',
    'галактическая астрономия': 'Galactic astronomy',
    'ознакомительная практика': 'Observation practice',
    'теоретическая механика': 'Theoretical mechanics',
    'атомная физика': 'Atomic physics',
    'оптика': 'Optics',
    'дифференциальные уравнения': 'Differential equations',
    'практикум по радиоэлектронике': 'Radioelectronics practicum',
    'радиофизика': 'Radiophysics',
    'астрометрия': 'Astrometry',
    'интегральные уравнения и вариационное исчисление': 'Integral equations and calculus of variations',
    'геофизика и физика планет': 'Geophysics and planet physics',
    'общая астрофизика': 'General astrophysics',
    'небесная механика': 'Celestial mechanics',
    'методы математической физики': 'Methods of mathematical Physics',
    'электродинамика': 'Electrodynamics',
    'практическая астрофизика': 'Applied astrophysics',
    'физика атомного ядра и частиц': 'Nuclear and particle Physics',
    'специальный астрономический практикум': 'Special astronomy practicum',
    'научно-исследовательская (получение первичных навыков научно-исследовательской работы)': 'Research practice',
    'производственная практика': 'Industrial practice',
    'педагогическая практика': 'Teaching practice',
    'преддипломная практика': 'Pre-graduation practice',
    'научно-исследовательская работа': 'Research work',
    'квантовая теория': 'Quantum theory',
    'теория вероятностей и математическая статистика': 'Probability theory and mathematical statistics',
    'лаборатория специализации': 'Specialization laboratory',
    'термодинамика и статистическая физика': 'Thermodynamics and statistical Physics',
    'основы физики нейтрино': 'Foundations of neutrino physics',
    'основы математического моделирования': 'Foundations of mathematical modeling',
    'численные методы': 'Numerical methods',
    'общая физическая химия': 'General physical chemistry',
    'история россии': 'History of Russia',
    'история и методология физики': 'History and methodology of physics',
    'всеобщая история': 'World history',
    'правоведение': 'Jurisprudence',
    'научно-исследовательский семинар': 'Research seminar',
    'общая теория относительности для астрономов': 'General relativity theory for astronomers',
    'философия': 'Philosophy',
    'экономика': 'Economics',
    'современные проблемы астрономии': 'Modern problems of astronomy',
    'теоретические основы аналитики больших данных': 'Theoretical foundations of big data analysis',
    'работа с базами данных и язык запросов sql': 'Working with databases and the SQL query language',
    'основы программирования и анализа данных на python': 'Foundations of programming and data analysis in Python',
    'введение в космологию. теория однородной вселенной': 'Introduction to cosmology. Theory of homogeneous universe',
    'физические основы циклических ускорителей': 'Physical foundations of cyclic accelerators',
    'теоретическая поддержка экспериментов в физике высоких энергий': 'Theoretical support of modern experiments in high-energy physics',
    'методика современных экспериментов по физике элементарных частиц': 'Methodology of modern experiments in elementary particle physics',
    'избранные главы квантовой теории столкновений': 'Selected chapters of quantum collision theory',
    'основы современной физики элементарных частиц': 'Foundations of modern elementary particle physics',
    'техника и физика ядерного эксперимента': 'Technology and physics of nuclear experiments',
}

MODULES_MAP = {
    'информатика': 'Informatics',
    'информатика и вычислительная физика': 'Informatics and computational Physics',
    'общая физика': 'General Physics',
    'математика': 'Mathematics',
    'фундаментальная и прикладная математика': 'Fundamental and applied Mathematics',
    'астрономия': 'Astronomy',
    'теоретическая физика': 'Theoretical Physics',
    'современное естествознание': 'Modern natural science',
}

MARKS_MAP = {
    'зачет': 'Test',
    'экзамен': 'Exam',
    'диф. зачет': 'Graded test',
    'зачтено': 'passed',
    'отлично': 'excellent',
    'хорошо': 'good',
    'удовлетворительно': 'satisfactory',
    'неудовлетворительно': 'unsatisfactory',
    'не зачтено': 'not passed',
}

SESSION_ATTITUDE_MAP = {
    'сдал сессию': 'Passed session',
    'сдал сессию на "хорошо" и "отлично"': 'Passed session with “good” and “excellent” marks',
    'сдал сессию на "отлично"': 'Passed session with “excellent” marks',
}

REGIONS_MAP = {
    'москва': 'Moscow', 'санкт-петербург': 'Saint Petersburg', 'севастополь': 'Sevastopol',
    'московская': 'Moscow oblast', 'ленинградская': 'Leningrad oblast', 'свердловская': 'Sverdlovsk oblast',
    'тверская': 'Tver oblast', 'нижегородская': 'Nizhny Novgorod oblast', 'новосибирская': 'Novosibirsk oblast',
    'челябинская': 'Chelyabinsk oblast', 'самарская': 'Samara oblast', 'ростовская': 'Rostov oblast',
    'омская': 'Omsk oblast', 'воронежская': 'Voronezh oblast', 'волгоградская': 'Volgograd oblast',
    'саратовская': 'Saratov oblast', 'тюменская': 'Tyumen oblast', 'иркутская': 'Irkutsk oblast',
    'кемеровская': 'Kemerovo oblast', 'оренбургская': 'Orenburg oblast', 'тульская': 'Tula oblast',
    'калужская': 'Kaluga oblast', 'ярославская': 'Yaroslavl oblast', 'владимирская': 'Vladimir oblast',
    'рязанская': 'Ryazan oblast', 'брянская': 'Bryansk oblast', 'курская': 'Kursk oblast',
    'белгородская': 'Belgorod oblast', 'липецкая': 'Lipetsk oblast', 'тамбовская': 'Tambov oblast',
    'смоленская': 'Smolensk oblast', 'ивановская': 'Ivanovo oblast', 'костромская': 'Kostroma oblast',
    'орловская': 'Orel oblast', 'пензенская': 'Penza oblast', 'ульяновская': 'Ulyanovsk oblast',
    'кировская': 'Kirov oblast', 'вологодская': 'Vologda oblast', 'архангельская': 'Arkhangelsk oblast',
    'мурманская': 'Murmansk oblast', 'калининградская': 'Kaliningrad oblast', 'новгородская': 'Novgorod oblast',
    'псковская': 'Pskov oblast', 'астраханская': 'Astrakhan oblast', 'курганская': 'Kurgan oblast',
    'томская': 'Tomsk oblast', 'амурская': 'Amur oblast', 'сахалинская': 'Sakhalin oblast',
    'магаданская': 'Magadan oblast', 'татарстан': 'Republic of Tatarstan',
    'башкортостан': 'Republic of Bashkortostan', 'удмуртская': 'Udmurt Republic',
    'чувашская': 'Chuvash Republic', 'мордовия': 'Republic of Mordovia', 'марий эл': 'Mari El Republic',
    'карелия': 'Republic of Karelia', 'коми': 'Komi Republic', 'крым': 'Republic of Crimea',
    'дагестан': 'Republic of Dagestan', 'бурятия': 'Republic of Buryatia', 'калмыкия': 'Republic of Kalmykia',
    'саха': 'Republic of Sakha (Yakutia)', 'якутия': 'Republic of Sakha (Yakutia)',
    'хакасия': 'Republic of Khakassia', 'адыгея': 'Republic of Adygea',
    'краснодарский': 'Krasnodar krai', 'красноярский': 'Krasnoyarsk krai', 'пермский': 'Perm krai',
    'приморский': 'Primorsky krai', 'ставропольский': 'Stavropol krai', 'хабаровский': 'Khabarovsk krai',
    'алтайский': 'Altai krai', 'забайкальский': 'Zabaykalsky krai', 'камчатский': 'Kamchatka krai',
    'ханты-мансийский': 'Khanty-Mansi autonomous okrug', 'ямало-ненецкий': 'Yamalo-Nenets autonomous okrug',
}


def format_inspector_entry(raw_val):
    val = raw_val.strip()
    if not re.search(r'[А-Яа-яЁё]', val):
        return val
    words = val.split()
    if len(words) == 3 and '.' not in val:
        fam, im, ot = words[0], words[1], words[2]
        i1 = translit_ru(im[0].upper())
        i2 = translit_ru(ot[0].upper())
        fam_en = re.sub(r'ская$', 'skaya', fam, flags=re.IGNORECASE)
        fam_en = translit_ru(fam_en.capitalize())
        return f"{i1}.{i2}. {fam_en}"
    return translit_ru(val)


def load_inspectors_db(filepath=INSPECTORS_FILE):
    if not os.path.exists(filepath):
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(DEFAULT_INSPECTORS_CONTENT)
    db = {}
    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            line_clean = line.split('#', 1)[0].strip()
            if not line_clean or '=' not in line_clean:
                continue
            k, v = line_clean.split('=', 1)
            db[k.strip().lower()] = format_inspector_entry(v)
    return db


def detect_course_and_inspector(parsed_data, inspectors_db):
    info = parsed_data['info']
    is_mag = 'магистр' in (info.get('faculty_line', '') + ' ' + info.get('specialty', '')).lower()
    course = None
    for ord_line in reversed(parsed_data['orders']):
        m = re.search(r'\s([1-6])\s+\d{2}\.\d{2}\.\d{4}\s+приказ', re.sub(r'\s+', ' ', ord_line))
        if m:
            course = int(m.group(1))
            break
    if course is None and parsed_data['semesters']:
        max_sem = max(s['num'] for s in parsed_data['semesters'])
        course = min(6, (max_sem + 1) // 2)
    if course is None:
        grp = info.get('groups', ('', '', ''))[0]
        if re.match(r'^[1-6]\d{2}', grp):
            course = int(grp[0])
    if course is None:
        course = 6
    key = f"mag{course}" if is_mag and f"mag{course}" in inspectors_db else str(course)
    inspector = inspectors_db.get(key) or inspectors_db.get('6', 'T.S. Nesterenko')
    return course, inspector


def load_course_dictionary(courses_file=COURSES_FILE):
    course_dict = {}
    for filepath in [courses_file, CACHE_FILE]:
        if os.path.exists(filepath):
            with open(filepath, 'r', encoding='utf-8') as f:
                for line in f:
                    parts = line.rstrip('\r\n').split('\t')
                    if len(parts) >= 2:
                        ru = re.sub(r'\s+', ' ', parts[0].strip()).lower()
                        en = parts[1].strip()
                        if ru and en:
                            course_dict[ru] = en
    for k, v in BASE_COURSES.items():
        if k not in course_dict:
            course_dict[k] = v
    course_dict.update(OVERRIDE_COURSES)
    return course_dict


def translate_via_api(text_ru):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36'}
    q_enc = urllib.parse.quote(text_ru)
    try:
        url1 = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl=ru&tl=en&dt=t&q={q_enc}"
        req1 = urllib.request.Request(url1, headers=headers)
        with urllib.request.urlopen(req1, timeout=6) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            res = ''.join(seg[0] for seg in data[0] if seg[0])
            if res:
                time.sleep(0.2)
                return res, True
    except Exception:
        pass
    try:
        url2 = f"https://clients5.google.com/translate_a/t?client=dict-chrome-ex&sl=ru&tl=en&q={q_enc}"
        req2 = urllib.request.Request(url2, headers=headers)
        with urllib.request.urlopen(req2, timeout=6) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            if isinstance(data, list) and data:
                res = data[0][0] if isinstance(data[0], list) else str(data[0])
                if res:
                    time.sleep(0.2)
                    return res, True
    except Exception:
        pass
    try:
        url3 = f"https://api.mymemory.translated.net/get?q={q_enc}&langpair=ru|en"
        req3 = urllib.request.Request(url3, headers=headers)
        with urllib.request.urlopen(req3, timeout=6) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            res = data.get('responseData', {}).get('translatedText', '')
            if res and 'MYMEMORY WARNING' not in res.upper():
                return res, True
    except Exception as e:
        print(f"  [!] Онлайн-переводчик недоступен для '{text_ru}' ({e}), использован транслит.")
    return translit_ru(text_ru), False


def google_translate_cached(text_ru, course_dict, save_to_cache=True):
    clean_ru = re.sub(r'\s+', ' ', str(text_ru).strip())
    if not clean_ru:
        return ''
    key = clean_ru.lower()
    if key in course_dict:
        return course_dict[key]
    translated, ok = translate_via_api(clean_ru)
    course_dict[key] = translated
    if save_to_cache and ok:
        with open(CACHE_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{clean_ru}\t{translated}\n")
        print(f"  [Google -> Кэш] {clean_ru}  ==>  {translated}")
    return translated


def translate_course_name(raw_name, course_dict, in_class_str='0', is_extra=False):
    s = re.sub(r'\s+', ' ', raw_name.strip())
    m_empty_kr = re.match(r'^КР\s+(.+?):\s*$', s, flags=re.IGNORECASE)
    if m_empty_kr:
        base_subj = m_empty_kr.group(1).strip()
        subj_en = translate_course_name(base_subj, course_dict, in_class_str, is_extra)
        if is_extra or (in_class_str.isdigit() and int(in_class_str) > 0):
            return subj_en
        return f"Coursework: {subj_en}"
    m_kr_lab = re.match(r'^КР\s+Лаборатория специализации:\s*(.+)$', s, flags=re.IGNORECASE)
    if m_kr_lab:
        return f"Coursework: specialization laboratory: {google_translate_cached(m_kr_lab.group(1), course_dict)}"
    m_kr_nir = re.match(r'^КР\s+Научно-исследовательская работа:\s*(.+)$', s, flags=re.IGNORECASE)
    if m_kr_nir:
        return f"Coursework research work: {google_translate_cached(m_kr_nir.group(1), course_dict)}"
    m_kr_any = re.match(r'^(?:КР|Курсовая работа)\s+([^:]+):\s*(.+)$', s, flags=re.IGNORECASE)
    if m_kr_any:
        sub_en = translate_course_name(m_kr_any.group(1), course_dict)
        top_en = google_translate_cached(m_kr_any.group(2), course_dict)
        return f"Coursework: {sub_en}: {top_en}"
    m_kr_simple = re.match(r'^(?:КР|Курсовая работа)[:\s]+(.+)$', s, flags=re.IGNORECASE)
    if m_kr_simple:
        return f"Coursework: {google_translate_cached(m_kr_simple.group(1), course_dict)}"
    m_mod = re.match(r'^Модуль\s+"([^"]+)"\s+(.+)$', s, flags=re.IGNORECASE)
    if m_mod:
        mod_ru = m_mod.group(1).strip().lower()
        subj_ru = m_mod.group(2).strip()
        mod_en = MODULES_MAP.get(mod_ru) or google_translate_cached(m_mod.group(1), course_dict, False)
        subj_en = translate_course_name(subj_ru, course_dict, in_class_str, is_extra)
        return f'Module "{mod_en}" {subj_en}'
    if s.lower() not in course_dict:
        m_part = re.match(r'^(.+?)\s*\((?:часть|ч\.)\s*(\d+)\)\s*$', s, flags=re.IGNORECASE)
        if m_part:
            base_en = translate_course_name(m_part.group(1), course_dict, in_class_str, is_extra)
            return f"{base_en} (part {m_part.group(2)})"
        m_eng = re.match(r'^(.+?)\s*\(на английском языке\)\s*$', s, flags=re.IGNORECASE)
        if m_eng:
            base_en = translate_course_name(m_eng.group(1), course_dict, in_class_str, is_extra)
            return f"{base_en} (in English)"
    return google_translate_cached(s, course_dict)


def translate_order_line(line_ru, course_dict):
    s = re.sub(r'\s+', ' ', line_ru.strip())
    m = re.match(r'^(.+?)\s+(\d+)\s+(\d{2}\.\d{2}\.\d{4})\s+приказ\s+N\s+([^\s]+)\s+(\d{2}\.\d{2}\.\d{4})\s+с\s+(\d{2}\.\d{2}\.\d{4})$', s)
    if m:
        action_ru, c_num, d1, ord_num, d2, d3 = m.groups()
        if 'зачислен' in action_ru.lower():
            suf = {'1': 'st', '2': 'nd', '3': 'rd'}.get(c_num, 'th')
            action_en = f'Accepted on the {c_num}{suf} course'
        elif 'переведен' in action_ru.lower():
            action_en = 'Transferred to next course'
        else:
            action_en = google_translate_cached(action_ru, course_dict, False)
        ord_num_en = re.sub(r'-?ас$', '-ac', ord_num, flags=re.IGNORECASE)
        return f"{action_en}\t{c_num}\t{d1} order N {ord_num_en}\t{d2} from {d3}"
    return google_translate_cached(s, course_dict, False)


def translate_specialty(spec_ru, course_dict):
    s = re.sub(r'\s+', ' ', spec_ru.strip())
    low = s.lower()
    if 'фундаментальная и прикладная физика' in low:
        return 'Fundamental and applied Physics'
    if 'астрономия' in low:
        return 'Specialization “Astronomy”'
    if low in ('физика', 'направление "физика"', 'специальность "физика"'):
        return 'Physics'
    return google_translate_cached(s, course_dict, False)


def translate_region(reg_ru, course_dict):
    parts = [p.strip() for p in reg_ru.split(',') if p.strip()]
    out = []
    for p in parts:
        low = p.lower()
        if low in ('российская федерация', 'рф', 'россия'):
            out.append('Russian Federation')
            continue
        matched = False
        for rk, rv in REGIONS_MAP.items():
            if rk in low:
                out.append(rv)
                matched = True
                break
        if not matched:
            out.append(google_translate_cached(p, course_dict, False))
    return ', '.join(out)


def translate_address(addr_ru):
    parts = [p.strip() for p in addr_ru.split(',') if p.strip()]
    clean_parts = [p for p in parts if not re.match(r'^(\+7|8)\d{9,10}$', re.sub(r'[\s\-()]', '', p))]
    city_map = {'москва': 'Moscow', 'санкт-петербург': 'Saint Petersburg', 'екатеринбург': 'Ekaterinburg'}
    formatted = []
    for p in clean_parts:
        if p.lower() in city_map:
            formatted.append(city_map[p.lower()])
            continue
        p_t = p.title()
        p_t = re.sub(r'\b(?:Б-Р|Бульвар)\b', 'blvd.', p_t, flags=re.IGNORECASE)
        p_t = re.sub(r'\b(?:Ул\.?|Улица)\b', 'st.', p_t, flags=re.IGNORECASE)
        p_t = re.sub(r'\b(?:Пр-Т|Проспект)\b', 'ave.', p_t, flags=re.IGNORECASE)
        p_t = re.sub(r'\b(?:Пер\.?|Переулок)\b', 'ln.', p_t, flags=re.IGNORECASE)
        p_t = re.sub(r'\b(?:Ш\.?|Шоссе)\b', 'hwy.', p_t, flags=re.IGNORECASE)
        p_t = re.sub(r'\b(?:Наб\.?|Набережная)\b', 'emb.', p_t, flags=re.IGNORECASE)
        p_t = re.sub(r'\bД\.?\s*', '', p_t, flags=re.IGNORECASE).strip()
        p_t = re.sub(r'\bКв\.?\s*', '', p_t, flags=re.IGNORECASE).strip()
        formatted.append(translit_ru(p_t))
    return ', '.join(formatted)


def extract_visible_rows_from_excel(excel_path):
    rows = []
    if excel_path.lower().endswith('.xls'):
        wb = xlrd.open_workbook(excel_path, formatting_info=True)
        sheet = wb.sheet_by_name('CartSt') if 'CartSt' in wb.sheet_names() else wb.sheet_by_index(0)
        for r in range(sheet.nrows):
            rinfo = sheet.rowinfo_map.get(r)
            if rinfo and rinfo.hidden:
                continue
            vals = []
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_DATE:
                    dt = xlrd.xldate_as_tuple(cell.value, wb.datemode)
                    vals.append(f"{dt[2]:02d}.{dt[1]:02d}.{dt[0]:04d}")
                elif cell.ctype == xlrd.XL_CELL_NUMBER and cell.value == int(cell.value):
                    vals.append(str(int(cell.value)))
                else:
                    vals.append(str(cell.value).strip())
            rows.append((r, vals))
    else:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb['CartSt'] if 'CartSt' in wb.sheetnames else wb.active
        for r_idx, row in enumerate(ws.iter_rows(values_only=False), start=1):
            if ws.row_dimensions[r_idx].hidden:
                continue
            vals = []
            for cell in row:
                v = cell.value
                if v is None:
                    vals.append('')
                elif isinstance(v, datetime):
                    vals.append(v.strftime('%d.%m.%Y'))
                elif isinstance(v, float) and v == int(v):
                    vals.append(str(int(v)))
                else:
                    vals.append(str(v).strip())
            rows.append((r_idx - 1, vals))
    return rows


def parse_cartst_excel(excel_path):
    raw_rows = extract_visible_rows_from_excel(excel_path)
    data = {'info': {}, 'orders': [], 'semesters': [], 'extracurricular': []}
    mode = None
    current_sem = None
    for r_idx, row_vals in raw_rows:
        while len(row_vals) < 8:
            row_vals.append('')
        if not any(row_vals):
            continue
        c0 = row_vals[0]
        if r_idx <= 4 and ('факультет' in c0.lower() or 'отделение' in c0.lower()):
            data['info']['faculty_line'] = c0
        elif c0.startswith('Фамилия, имя, отчество'):
            data['info']['fio_full'] = f"{row_vals[1]} {row_vals[2]} {row_vals[3]}".strip()
            data['info']['fio'] = f"{row_vals[1]} {row_vals[2]}".strip()
        elif c0.startswith('Дата рождения'):
            data['info']['dob'] = row_vals[1].replace('г.', '').strip()
        elif c0.startswith('Договор'):
            data['info']['contract'] = row_vals[1]
        elif c0.startswith('Финансирование'):
            data['info']['funding'] = row_vals[1]
        elif c0.startswith('Специальность'):
            data['info']['specialty'] = row_vals[1]
        elif c0.startswith('Специализация'):
            data['info']['specialization'] = row_vals[1]
        elif c0.startswith('Группа'):
            data['info']['groups'] = (row_vals[1], row_vals[4].split('.')[0], row_vals[5].split('.')[0])
        elif c0.startswith('Постоянная прописка'):
            data['info']['region'] = row_vals[1]
        elif c0.startswith('Адрес'):
            data['info']['address'] = row_vals[1]
        elif c0.startswith('Общежитие'):
            data['info']['campus'] = row_vals[1]
        elif c0 == 'Приказы':
            mode = 'orders'
        elif c0.startswith('Выполнение учебного плана в'):
            mode = 'sem'
            m_sem = re.search(r'в\s+(\d+)\s+семестре', c0)
            sem_num = int(m_sem.group(1)) if m_sem else len(data['semesters']) + 1
            current_sem = {'num': sem_num, 'courses': [], 'attitude': ''}
            data['semesters'].append(current_sem)
        elif re.match(r'^(Дополнительно\s+сдано|Факультатив)', c0, flags=re.IGNORECASE):
            mode = 'extra'
            current_sem = None
        elif mode == 'orders':
            data['orders'].append(c0)
        elif mode in ('sem', 'extra'):
            if c0.startswith('Название') and row_vals[1].startswith('Всего'):
                continue
            if c0.startswith('Отношение к сессии'):
                if current_sem is not None:
                    current_sem['attitude'] = row_vals[1]
                continue
            if row_vals[1] or row_vals[3]:
                item = {
                    'name': c0,
                    'hours': row_vals[1],
                    'in_class': row_vals[2].split('.')[0],
                    'rep_form': row_vals[3],
                    'mark': row_vals[5],
                    'date': row_vals[6]
                }
                if mode == 'sem' and current_sem is not None:
                    current_sem['courses'].append(item)
                elif mode == 'extra':
                    data['extracurricular'].append(item)
    data['semesters'] = [s for s in data['semesters'] if s['courses']]
    return data


def fix_document_defaults(doc):
    """Убирает стандартные 10 пт отступа Word (w:after=200), из-за которых раздувались строки таблиц."""
    styles_elm = doc.styles.element
    doc_defs = styles_elm.find(qn('w:docDefaults'))
    if doc_defs is not None:
        ppr_def = doc_defs.find(qn('w:pPrDefault'))
        if ppr_def is not None:
            ppr = ppr_def.find(qn('w:pPr'))
            if ppr is not None:
                sp = ppr.find(qn('w:spacing'))
                if sp is None:
                    sp = OxmlElement('w:spacing')
                    ppr.append(sp)
                sp.set(qn('w:before'), '0')
                sp.set(qn('w:after'), '0')
                sp.set(qn('w:line'), '240')
                sp.set(qn('w:lineRule'), 'auto')
        rpr_def = doc_defs.find(qn('w:rPrDefault'))
        if rpr_def is not None:
            rpr = rpr_def.find(qn('w:rPr'))
            if rpr is not None:
                rfonts = rpr.find(qn('w:rFonts'))
                if rfonts is not None:
                    for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
                        rfonts.set(qn(attr), 'Times New Roman')
                    for t_attr in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
                        if qn(t_attr) in rfonts.attrib:
                            del rfonts.attrib[qn(t_attr)]
    for st_name in ('Normal', 'normal'):
        if st_name in doc.styles:
            st = doc.styles[st_name]
            st.paragraph_format.space_before = Pt(0)
            st.paragraph_format.space_after = Pt(0)
            st.paragraph_format.line_spacing = 1.0
            st.font.name = 'Times New Roman'
            st.font.size = Pt(11)


def format_paragraph_compact(p, line_spacing=1.0, keep_with_next=False, align=None):
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = line_spacing
    pf.keep_with_next = keep_with_next
    if align is not None:
        p.alignment = align
    if p._p.pPr is not None:
        rpr = p._p.pPr.find(qn('w:rPr'))
        if rpr is not None:
            for tag in ('w:highlight', 'w:sz', 'w:szCs'):
                el = rpr.find(qn(tag))
                if el is not None:
                    rpr.remove(el)


def normalize_existing_runs_font(p, size_pt=11):
    """Унифицирует шрифт в клонированных из шаблона абзацах (например, если ФИО подписанта правили в Word)."""
    for run in p.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size_pt)
        run.font.color.rgb = RGBColor(0, 0, 0)
        rpr = run._r.get_or_add_rPr()
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = OxmlElement('w:rFonts')
            rpr.insert(0, rfonts)
        for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
            rfonts.set(qn(attr), 'Times New Roman')
        for t_attr in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
            if qn(t_attr) in rfonts.attrib:
                del rfonts.attrib[qn(t_attr)]


def add_styled_run(p, text, bold=False, italic=False, size_pt=11):
    run = p.add_run(str(text))
    run.bold = bold
    run.italic = italic
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(0, 0, 0)
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rfonts.set(qn(attr), 'Times New Roman')
    return run


def set_cell_compact(cell, text, bold=False, italic=False, keep_with_next=False):
    for extra_p in cell.paragraphs[1:]:
        cell._tc.remove(extra_p._p)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        p._p.remove(r._r)
    format_paragraph_compact(p, line_spacing=1.0, keep_with_next=keep_with_next)
    add_styled_run(p, text, bold=bold, italic=italic)


def ordinal_en(n):
    return f"{n}{'tsnrhtdd'[(n//10%10!=1)*(n%10<4)*n%10::4]}"


def fill_courses_table(doc, body, sem_table_proto, sem_row_proto, courses_list, course_dict, is_extra=False):
    new_tbl_elm = copy.deepcopy(sem_table_proto)
    body.insert(len(body) - 1, new_tbl_elm)
    tbl = docx.table.Table(new_tbl_elm, doc)
    while len(tbl.rows) > 1:
        tbl._tbl.remove(tbl.rows[-1]._tr)
    for c_h in tbl.rows[0].cells:
        set_cell_compact(c_h, c_h.text.strip(), bold=True, keep_with_next=True)
    for c_item in courses_list:
        new_tr = copy.deepcopy(sem_row_proto)
        tbl._tbl.append(new_tr)
        rcells = [docx.table._Cell(tc, tbl) for tc in new_tr.findall(qn('w:tc'))]
        c_name_en = translate_course_name(c_item['name'], course_dict, c_item['in_class'], is_extra=is_extra)
        set_cell_compact(rcells[0], c_name_en)
        hours_en = re.sub(r'кр\s*-\s*', 'cr – ', c_item['hours'])
        set_cell_compact(rcells[1], hours_en)
        set_cell_compact(rcells[2], c_item['in_class'])
        set_cell_compact(rcells[3], MARKS_MAP.get(c_item['rep_form'].lower(), c_item['rep_form']))
        set_cell_compact(rcells[4], MARKS_MAP.get(c_item['mark'].lower(), c_item['mark']))
        set_cell_compact(rcells[5], c_item['date'])


def update_inspector_table(insp_tbl_elm, doc, inspector_name, stamp_date):
    tbl = docx.table.Table(insp_tbl_elm, doc)
    for row in list(tbl.rows):
        unique_tcs = []
        for c in row.cells:
            if c._tc not in unique_tcs:
                unique_tcs.append(c._tc)
        row_txt = ' '.join(docx.table._Cell(tc, tbl).text for tc in unique_tcs).lower()
        if 'inspector' in row_txt:
            right_cell = docx.table._Cell(unique_tcs[-1], tbl)
            p0 = right_cell.paragraphs[0]
            for r in list(p0.runs):
                p0._p.remove(r._r)
            format_paragraph_compact(p0, line_spacing=1.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            add_styled_run(p0, inspector_name, bold=False)
        elif 'seal' in row_txt:
            right_cell = docx.table._Cell(unique_tcs[-1], tbl)
            p0 = right_cell.paragraphs[0]
            for r in list(p0.runs):
                p0._p.remove(r._r)
            format_paragraph_compact(p0, line_spacing=1.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            add_styled_run(p0, stamp_date, bold=True)
    if len(tbl.rows) > 2 and not tbl.rows[-1].cells[0].text.strip() and not tbl.rows[-1].cells[-1].text.strip():
        tbl._tbl.remove(tbl.rows[-1]._tr)


def build_translated_docx(parsed_data, template_path, output_path, course_dict, inspector_name, stamp_date):
    doc = docx.Document(template_path)
    fix_document_defaults(doc)
    info = parsed_data['info']

    m_year = re.search(r'(\d{4})', info.get('faculty_line', ''))
    if len(doc.paragraphs) > 0:
        format_paragraph_compact(doc.paragraphs[0])
    if m_year and len(doc.paragraphs) > 1:
        p1 = doc.paragraphs[1]
        old_p1_txt = p1.text.strip()
        new_p1_txt = re.sub(r'\b\d{4}\b', m_year.group(1), old_p1_txt) if re.search(r'\b\d{4}\b', old_p1_txt) else f"Faculty of Physics main department Full-time education {m_year.group(1)}"
        for r in list(p1.runs):
            p1._p.remove(r._r)
        format_paragraph_compact(p1)
        add_styled_run(p1, new_p1_txt, bold=True)

    if doc.sections[0].footer.paragraphs:
        fp = doc.sections[0].footer.paragraphs[0]
        fp.text = ''
        add_styled_run(fp, f"\t\t\t\t{datetime.now().strftime('%d.%m.%Y')}")

    t0 = doc.tables[0]
    for r_row in t0.rows:
        for c_cell in r_row.cells:
            if c_cell.paragraphs:
                format_paragraph_compact(c_cell.paragraphs[0])

    set_cell_compact(t0.cell(0, 1), translit_ru(info.get('fio', '')))
    set_cell_compact(t0.cell(1, 1), info.get('dob', ''))
    contract_ru = info.get('contract', '').lower()
    set_cell_compact(t0.cell(2, 1), 'No contract' if 'без договора' in contract_ru else google_translate_cached(info.get('contract', ''), course_dict, False))
    funding_ru = info.get('funding', '').lower()
    set_cell_compact(t0.cell(3, 1), 'Federal budget of Russian Federation' if 'федеральный бюджет' in funding_ru else google_translate_cached(info.get('funding', ''), course_dict, False))
    set_cell_compact(t0.cell(4, 1), translate_specialty(info.get('specialty', ''), course_dict))
    set_cell_compact(t0.cell(5, 1), google_translate_cached(info.get('specialization', ''), course_dict, False) if info.get('specialization') else '')

    grp_main, grp_ex, grp_st = info.get('groups', ('', '', ''))
    set_cell_compact(t0.cell(6, 1), translit_ru(grp_main))
    set_cell_compact(t0.cell(6, 2), grp_ex)
    set_cell_compact(t0.cell(6, 3), grp_st)
    set_cell_compact(t0.cell(7, 1), translate_region(info.get('region', ''), course_dict))
    set_cell_compact(t0.cell(8, 1), translate_address(info.get('address', '')))
    campus_ru = info.get('campus', '').lower()
    set_cell_compact(t0.cell(9, 1), 'Resides' if campus_ru == 'проживает' else 'Does not reside')

    # Сохраняем эталонную таблицу семестра, таблицу инспектора и блок подписи начальника отдела из template.docx
    body = doc.element.body
    sem_table_proto = copy.deepcopy(doc.tables[1]._tbl)
    sem_row_proto = copy.deepcopy(doc.tables[1].rows[1]._tr)
    insp_tbl_orig = doc.tables[-1]._tbl
    inspector_table_proto = copy.deepcopy(insp_tbl_orig)

    idx_after_insp = body.index(insp_tbl_orig) + 1
    cert_footer_protos = [
        copy.deepcopy(el) for el in body[idx_after_insp:]
        if not el.tag.endswith('sectPr')
    ]

    idx_after_t0 = body.index(t0._tbl) + 1
    for child in list(body[idx_after_t0:]):
        if not child.tag.endswith('sectPr'):
            body.remove(child)

    p_ord_title = doc.add_paragraph()
    format_paragraph_compact(p_ord_title, keep_with_next=True)
    add_styled_run(p_ord_title, 'Orders', bold=True)
    for ord_ru in parsed_data['orders']:
        p_o = doc.add_paragraph()
        format_paragraph_compact(p_o)
        add_styled_run(p_o, translate_order_line(ord_ru, course_dict))

    for sem in parsed_data['semesters']:
        p_sp = doc.add_paragraph()
        format_paragraph_compact(p_sp)
        p_st = doc.add_paragraph()
        format_paragraph_compact(p_st, keep_with_next=True)
        add_styled_run(p_st, f"Performance of the curriculum in the {ordinal_en(sem['num'])} semester", bold=True)
        fill_courses_table(doc, body, sem_table_proto, sem_row_proto, sem['courses'], course_dict, is_extra=False)
        att_ru = sem['attitude'].strip().lower()
        att_en = SESSION_ATTITUDE_MAP.get(att_ru) or (google_translate_cached(sem['attitude'], course_dict, False) if att_ru else '')
        p_att = doc.add_paragraph()
        format_paragraph_compact(p_att)
        add_styled_run(p_att, 'Attitude to the session:\t\t\t\t', bold=True, italic=True)
        add_styled_run(p_att, att_en, bold=False, italic=True)

    if parsed_data['extracurricular']:
        p_sp_ex = doc.add_paragraph()
        format_paragraph_compact(p_sp_ex)
        p_ex = doc.add_paragraph()
        format_paragraph_compact(p_ex, keep_with_next=True)
        add_styled_run(p_ex, 'Extracurricular courses', bold=True)
        fill_courses_table(doc, body, sem_table_proto, sem_row_proto, parsed_data['extracurricular'], course_dict, is_extra=True)

    update_inspector_table(inspector_table_proto, doc, inspector_name, stamp_date)
    body.insert(len(body) - 1, inspector_table_proto)

    # Вставляем блок заверения начальника иностранного отдела, считанный напрямую из template.docx
    has_non_empty_cert = any(
        el.tag.endswith('p') and ''.join(el.itertext()).strip()
        for el in cert_footer_protos
    )
    if has_non_empty_cert:
        for el in cert_footer_protos:
            body.insert(len(body) - 1, el)
            if el.tag.endswith('p'):
                p_cloned = docx.text.paragraph.Paragraph(el, doc)
                format_paragraph_compact(p_cloned)
                normalize_existing_runs_font(p_cloned)
    else:
        for line_text in [
            'I certify this is an accurate translation of the document\t\t',
            'Evgeny Kurov\t\t\t',
            'International Relations Office\t\t\t\t',
            'Faculty of Physics, MSU'
        ]:
            p_end = doc.add_paragraph()
            format_paragraph_compact(p_end)
            add_styled_run(p_end, line_text)

    doc.save(output_path)
    print(f"\n[OK] Готов файл: {output_path}")


def find_template_file():
    p = os.path.join(SCRIPT_DIR, 'template.docx')
    if os.path.exists(p):
        return p
    raise FileNotFoundError("Не найден файл шаблона 'template.docx' в папке со скриптом.")


def pick_files_via_dialog():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        paths = filedialog.askopenfilenames(
            title='Выберите Excel-файл(ы) учебной карточки',
            filetypes=[('Файлы Excel', '*.xls *.xlsx *.xlsm'), ('Все файлы', '*.*')]
        )
        root.destroy()
        return list(paths)
    except Exception:
        raw = input("Введите путь к файлу Excel (.xls / .xlsx): ").strip().strip('"')
        return [raw] if raw else []


def select_excel_files_interactive():
    local_xls = sorted([
        os.path.join(SCRIPT_DIR, f) for f in os.listdir(SCRIPT_DIR)
        if f.lower().endswith(('.xls', '.xlsx')) and not f.startswith('~$')
    ])
    print("=" * 68)
    print("  АВТОПЕРЕВОД УЧЕБНЫХ КАРТОЧЕК (Иностранный отдел)")
    print("=" * 68)
    if not local_xls:
        print("В папке скрипта нет файлов .xls/.xlsx — открываю окно выбора файла...")
        return pick_files_via_dialog()
    print("Найденные Excel-файлы в папке:")
    for i, fp in enumerate(local_xls, start=1):
        print(f"  [{i}] {os.path.basename(fp)}")
    print("  [A] Перевести ВСЕ файлы из списка")
    print("  [O] Выбрать другой файл через Проводник")
    choice = input("\nВаш выбор (Enter = [1], номер, A или O): ").strip().lower()
    if choice == '' or choice == '1':
        return [local_xls[0]]
    if choice in ('a', 'ф', 'all', 'все'):
        return local_xls
    if choice in ('o', 'о', 'open'):
        return pick_files_via_dialog()
    if choice.isdigit() and 1 <= int(choice) <= len(local_xls):
        return [local_xls[int(choice) - 1]]
    print("[!] Некорректный выбор.")
    return []


def normalize_date_input(raw_str, default_date):
    s = raw_str.strip()
    if not s:
        return default_date
    s = re.sub(r'[/\-,\s]+', '.', s)
    parts = s.split('.')
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
        if y < 100:
            y += 2000
        return f"{d:02d}.{m:02d}.{y:04d}"
    return s


def main():
    tpl_path = find_template_file()
    inspectors_db = load_inspectors_db()
    c_dict = load_course_dictionary()
    cli_files = [a for a in sys.argv[1:] if os.path.exists(a)]
    excel_files = cli_files if cli_files else select_excel_files_interactive()
    if not excel_files:
        print("Файлы не выбраны. Завершение работы.")
        return
    today_str = datetime.now().strftime('%d.%m.%Y')
    for xls_path in excel_files:
        parsed = parse_cartst_excel(xls_path)
        course_num, auto_inspector = detect_course_and_inspector(parsed, inspectors_db)
        fio_ru = parsed['info'].get('fio_full') or parsed['info'].get('fio', '')
        print("\n" + "-" * 68)
        print(f"Файл:      {os.path.basename(xls_path)}")
        print(f"Студент:   {fio_ru}  ->  {translit_ru(parsed['info'].get('fio', ''))}")
        print(f"Курс:      {course_num} (автоопределен инспектор: {auto_inspector})")
        if parsed['extracurricular']:
            print(f"Доп. курс: найдено дисциплин сверх плана — {len(parsed['extracurricular'])}")
        print("-" * 68)
        raw_date = input(f"Какую дату с оригинала справки поставить в конце? [Enter = {today_str}]: ")
        stamp_date = normalize_date_input(raw_date, today_str)
        raw_insp = input(f"Инспектор [Enter = {auto_inspector}, либо введите курс 1-6 или ФИО]: ").strip()
        if not raw_insp:
            final_inspector = auto_inspector
        elif raw_insp.lower() in inspectors_db:
            final_inspector = inspectors_db[raw_insp.lower()]
        else:
            final_inspector = format_inspector_entry(raw_insp)
        base_no_ext = os.path.splitext(xls_path)[0]
        out_docx = f"{base_no_ext}_EN.docx"
        build_translated_docx(parsed, tpl_path, out_docx, c_dict, final_inspector, stamp_date)


if __name__ == '__main__':
    main()
"""Створює порожні шаблони журналів у папці logs/ (виконується один раз під час підготовки пакета).

logs/run_log.xlsx    — аркуші «Інструкція», «Основні запуски», «Відхилення»;
logs/blind_key.xlsx  — ключ кодів для ручної перевірки наосліп (доступний лише тому, хто готує коди);
logs/blind_check.xlsx — форма ручної перевірки наосліп.

Порядок основних запусків: 10 раундів, у кожному всі 9 поєднань «модель × обсяг» у випадковому порядку
(зерно генератора 2026). Непарні раунди виконує одна дослідниця, парні — інша; два раунди на день.
Наявні файли не перезаписуються без параметра --force.
"""
import os
import random
import sys

from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

import xlsx_style

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGS = os.path.join(ROOT, "logs")
MODELS = [("ChatGPT", "chatgpt", "gpt"), ("Gemini", "gemini", "gemini"), ("DeepSeek", "deepseek", "ds")]
LEVELS = ["S", "M", "L"]
EXECUTORS = ("Вероніка", "Аліна")
YES_NO = '"так,ні"'
STATES = '"повна,обірвана,відмова,помилка сервісу,ліміт вичерпано"'
PEOPLE = '"Вероніка,Аліна"'

COLUMNS = [  # назва, ширина, тип введення (None — заповнено заздалегідь або формула)
    ("№", 5, None), ("Раунд", 8, None), ("Модель", 10, None), ("Обсяг", 7, None), ("Повтор", 8, None),
    ("Файл відповіді", 34, None), ("Виконавиця", 11, PEOPLE), ("Дата запуску", 12, "date"),
    ("Тривалість відповіді, хв", 11, "num"), ("Назва моделі в інтерфейсі", 22, "text"),
    ("Тимчасовий чат", 10, YES_NO), ("Натискань «Продовжити»", 11, "int"), ("Стан відповіді", 13, STATES),
    ("Відповідь збережено", 11, YES_NO), ("Налаштування перевірено", 12, YES_NO),
    ("Входить до аналізу", 10, YES_NO), ("Причина виключення", 24, "text"), ("Примітки", 40, "text"),
]

INSTRUCTION = [
    "Журнал запусків: правила заповнення",
    "Аркуші: «Основні запуски» — 90 запусків експерименту; «Відхилення» — зміни протоколу, внесені після початку "
    "пілоту, з датою й причиною. Пілотні запуски записано в окремому журналі logs/pilot_log.xlsx.",
    "Налаштування для основних запусків. ChatGPT: модель GPT-6 Instant, режим міркувань вимкнено, тимчасовий чат. "
    "Gemini: модель 3.6 Flash, тимчасовий чат (якщо недоступний — новий чат). DeepSeek: кнопки DeepThink і Search "
    "вимкнено, новий чат. Режими Auto (модель обирає сервіс) і пошук в інтернеті не використовуються. Поле власних "
    "інструкцій в обліковому записі порожнє. Перед першим запуском обидві виконавиці перевіряють, що в їхніх "
    "облікових записах доступні ті самі моделі.",
    "1. Рядки виконуються згори вниз: порядок у межах раунду випадковий і зафіксований заздалегідь. "
    "Один раунд (9 рядків) — один сеанс роботи; виконавиця раунду вказана в колонці «Виконавиця».",
    "2. Перед запуском перевіряються налаштування сервісу; у колонці «Налаштування перевірено» ставиться «так».",
    "3. Відкривається тимчасовий чат (у DeepSeek — новий чат). Увесь текст файлу requests/request_<обсяг>.txt "
    "вставляється одним повідомленням і надсилається; до запиту нічого не дописується.",
    "4. Якщо відповідь зупинилася й з’явилася кнопка «Продовжити» (Continue), вона натискається не більше трьох "
    "разів; кількість натискань записується в журнал.",
    "5. Відповідь копіюється кнопкою копіювання блоку коду (або всієї відповіді) і зберігається без змін "
    "як текстовий файл UTF-8 за шляхом із колонки «Файл відповіді».",
    "6. Записуються дата (ДД.ММ.РРРР), тривалість генерації в хвилинах, назва моделі, яку показує інтерфейс "
    "(якщо інтерфейс її не показує — назва, яку на запитання повідомляє сама модель), стан відповіді та позначка "
    "«Відповідь збережено».",
    "7. «Входить до аналізу» = «ні» ставиться лише для запуску з помилкою виконання (не той запит чи рівень, "
    "інші налаштування) або коли інтерфейс показав іншу модель; причина записується. Такий запуск повторюється: "
    "унизу додається рядок із тими самими моделлю, обсягом і повтором та новим файлом (наприклад, gpt_M_03_r.txt).",
    "8. Відповіді не редагуються, навіть якщо містять помилки: помилки моделей і є предметом вимірювання.",
    "9. Перевірка збережених відповідей: python3 tools/run_all.py main. "
    "Результати записуються в results/main.xlsx і results/main.txt.",
    "Стан відповіді: повна — відповідь завершена (з натисканнями «Продовжити» або без них); обірвана — після трьох "
    "натискань відповідь не завершена; відмова — модель не згенерувала дані; помилка сервісу — збій інтерфейсу після "
    "початку відповіді; ліміт вичерпано — сервіс не прийняв запит через безкоштовний ліміт (запуск переноситься на "
    "наступну добу). Для кожного стану, крім «повна», у колонці «Примітки» коротко вказується причина.",
]


def add_validation(ws, col_letter, kind, last_row):
    rng = f"{col_letter}2:{col_letter}{last_row}"
    if kind in (YES_NO, STATES, PEOPLE):
        dv = DataValidation(type="list", formula1=kind, allow_blank=True)
    elif kind == "date":
        dv = DataValidation(type="date", operator="between", formula1="DATE(2026,1,1)", formula2="DATE(2027,12,31)",
                            allow_blank=True, error="Дата у форматі ДД.ММ.РРРР")
    elif kind == "int":
        dv = DataValidation(type="whole", operator="between", formula1="0", formula2="3", allow_blank=True,
                            error="Від 0 до 3")
    else:
        return
    ws.add_data_validation(dv)
    dv.add(rng)


def journal_sheet(wb, title, rows, columns, spare=20):
    ws = wb.create_sheet(title)
    xlsx_style.header(ws, [c[0] for c in columns], [c[1] for c in columns])
    last = len(rows) + spare + 1
    for r in range(2, last + 1):
        values = rows[r - 2] if r - 2 < len(rows) else []
        for ci, (_, _, kind) in enumerate(columns, start=1):
            cell = ws.cell(row=r, column=ci, value=values[ci - 1] if ci - 1 < len(values) else None)
            cell.border = xlsx_style.BORDER
            if kind is not None:
                cell.fill = xlsx_style.FILL_INPUT
            if kind == "date":
                cell.number_format = "DD.MM.YYYY"
    for ci, (_, _, kind) in enumerate(columns, start=1):
        add_validation(ws, ws.cell(row=1, column=ci).column_letter, kind, last)
    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=len(columns)).column_letter}{last}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws, last


def main_rows():
    rnd = random.Random(2026)
    combos = [(m, lv) for m in MODELS for lv in LEVELS]
    rows, rep = [], {c: 0 for c in combos}
    for rnd_no in range(1, 11):
        order = combos[:]
        rnd.shuffle(order)
        for m, lv in order:
            rep[(m, lv)] += 1
            k = rep[(m, lv)]
            rows.append([len(rows) + 1, rnd_no, m[0], lv, k, f"answers/main/{m[1]}/{m[2]}_{lv}_{k:02d}.txt",
                         EXECUTORS[(rnd_no - 1) % 2]])
    return rows


def build_run_log(path):
    wb = Workbook()
    xlsx_style.text_sheet(wb, "Інструкція", INSTRUCTION, first=True)
    journal_sheet(wb, "Основні запуски", main_rows(), COLUMNS)
    dev = wb.create_sheet("Відхилення")
    heads = ["Дата (ДД.ММ.РРРР)", "Етап протоколу", "Що змінено", "Причина", "Вплив на висновок", "Хто вніс"]
    xlsx_style.header(dev, heads, [13, 14, 48, 48, 36, 12])
    xlsx_style.body(dev, [[None] * 6 for _ in range(30)])
    wb.active = 1
    wb.save(path)


def build_blind(folder):
    wb = Workbook()
    ws = xlsx_style.sheet(wb, "Ключ", ["Код відповіді", "Модель", "Обсяг", "Файл відповіді"],
                          [[f"K-{k:02d}", None, None, None] for k in range(1, 10)], [12, 12, 8, 40], first=True)
    for r in range(2, 11):
        for c in "BCD":
            ws[f"{c}{r}"].fill = xlsx_style.FILL_INPUT
    wb.save(os.path.join(folder, "blind_key.xlsx"))

    wb = Workbook()
    heads = ["Код відповіді", "id вакансії", "Таблиця", "id запису", "Правило", "Ручний висновок", "Коментар",
             "Висновок валідатора", "Збіг"]
    ws = xlsx_style.sheet(wb, "Перевірка", heads, [], [12, 11, 22, 14, 9, 14, 36, 14, 8], first=True)
    dv = DataValidation(type="list", formula1='"виконано,порушено,не оцінюється"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("F2:F400")
    for r in range(2, 401):
        for c in range(1, 10):
            cell = ws.cell(row=r, column=c)
            cell.border = xlsx_style.BORDER
            if c <= 7:
                cell.fill = xlsx_style.FILL_INPUT
            if c in (2, 4, 5):
                cell.number_format = "@"  # текст: номер правила 3.10 не перетворюється на 3.1
        ws[f"I{r}"] = f'=IF(OR(F{r}="",H{r}=""),"",IF(F{r}=H{r},"так","ні"))'
    ws["K1"], ws["K2"] = "Частка збігів", '=IFERROR(COUNTIF(I2:I400,"так")/(COUNTIF(I2:I400,"так")+COUNTIF(I2:I400,"ні")),"")'
    ws["K2"].number_format = "0.0%"
    ws["K1"].font = xlsx_style.BOLD
    ws.column_dimensions["K"].width = 14
    xlsx_style.text_sheet(wb, "Інструкція", [
        "Ручна перевірка наосліп",
        "Відповіді для перевірки лежать у папці blind/ під кодами K-01…K-09; модель і обсяг за кодом не визначаються.",
        "У кожній відповіді обирається одна вакансія з усіма пов’язаними записами: компанія, профіль HR, навички "
        "вакансії, відгуки на неї та історія їхніх статусів.",
        "Для кожного такого запису й кожного застосовного правила (таблиця 6.2 звіту) у формі заповнюється рядок: "
        "код відповіді, id вакансії, таблиця, id запису (для VacancySkill — vacancyId/skillId), правило, ручний висновок "
        "і коментар.",
        "Колонку «Висновок валідатора» заповнює команда python3 tools/blind_check.py compare після завершення "
        "ручної перевірки; колонка «Збіг» і частка збігів обчислюються формулами.",
    ])
    wb.save(os.path.join(folder, "blind_check.xlsx"))


def main(force=False):
    os.makedirs(LOGS, exist_ok=True)
    targets = [os.path.join(LOGS, n) for n in ("run_log.xlsx", "blind_key.xlsx", "blind_check.xlsx")]
    existing = [t for t in targets if os.path.exists(t)]
    if existing and not force:
        sys.exit("Журнали вже існують: " + ", ".join(os.path.relpath(t, ROOT) for t in existing) +
                 ". Для перезапису потрібен параметр --force.")
    build_run_log(targets[0])
    build_blind(LOGS)


if __name__ == "__main__":
    main(force="--force" in sys.argv)
    print("шаблони журналів створено")

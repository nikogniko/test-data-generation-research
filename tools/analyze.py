"""Підсумкові таблиці й рішення за H1 і H2 (таблиця 5.2 звіту ЛР6).

Отримує список запусків (дані рядка журналу, усі перевірки валідатора й підсумок) і записує результати
у двох форматах з однаковим змістом: книгу Excel (write) і текстовий файл (write_txt).
Аркуші книги для пілоту й основних запусків:
  «Запуски»      — рядок на запуск: дані журналу й показники валідатора;
  «Правила»      — для кожного запуску й правила кількість перевірок «виконано / порушено / не оцінюється»;
  «Порушення»    — кожна перевірка з результатом «порушено» або «не оцінюється» з деталями.
Лише для основних запусків:
  «Модель і обсяг», «Рішення», «Повторюваність».
У текстовому файлі ті самі дані подано розділами; перевірки за правилами для основних запусків
підсумовано за поєднаннями «модель × обсяг».

Середня половина запусків — проміжок між першим і третім квартилем (Q1–Q3): Q1 — медіана меншої
половини впорядкованого ряду, Q3 — медіана більшої половини; для 10 запусків це 3-тє і 8-ме значення.
Межі, що збігаються, вважаються перетином.
"""
from datetime import date, datetime
from statistics import median as _median

from openpyxl import Workbook

import xlsx_style
from validate import BAD, NA, OK, number, pct

LEVELS = ["S", "M", "L"]
RULES = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "2.1", "2.2", "2.3", "2.4", "2.5", "2.6"] + \
        [f"3.{k}" for k in range(1, 11)]
SHARES = [("group1", "Група 1, %"), ("group2", "Група 2, %"), ("group2_main", "Група 2 без 2.5, %"),
          ("group3", "Група 3, %"), ("group3_core", "Група 3 без 3.1–3.3, %"), ("overall", "Загальна, %"),
          ("overall_main", "Загальна без 2.5, %")]


def median(values):
    values = [v for v in values if v is not None]
    return round(_median(values), 2) if values else None


def middle_half(values):
    v = sorted(x for x in values if x is not None)
    if not v:
        return None, None
    if len(v) == 1:
        return v[0], v[0]
    return median(v[: len(v) // 2]), median(v[(len(v) + 1) // 2:])


def readable(r):
    return r["summary"] is not None and r["summary"]["state"] == "прочитано"


def rule_2_5(s):
    d = (s or {}).get("rule_2_5")
    return f"{d['порушено']} / {d['оцінено']}" if d else None


def run_rows(runs, pilot):
    rows = []
    for r in runs:
        s = r["summary"]
        lead = [r["no"], r["model"]] + ([r["interface"]] if pilot else [r["round"], r["rep"]]) + \
               [r["level"], r["file"], r["date"]] + ([r["duration"]] if pilot else []) + \
               [r["presses"], r["state"]] + ([r["assessment"]] if pilot else [])
        tail = [s["state"] if s else "файл не знайдено", (s or {}).get("chars"), (s or {}).get("records_total"),
                (s or {}).get("volume_deviation")] + [pct(s, k) for k, _ in SHARES] + \
               [rule_2_5(s), pct(s, "usable_records"), pct(s, "p1")]
        rows.append(lead + tail)
    return rows


def run_headers(pilot):
    lead = ["№", "Модель"] + (["Назва моделі в інтерфейсі"] if pilot else ["Раунд", "Повтор"]) + \
           ["Обсяг", "Файл відповіді", "Дата запуску"] + (["Тривалість, хв"] if pilot else []) + \
           ["Натискань «Продовжити»", "Стан (журнал)"] + (["Оцінка обсягу"] if pilot else [])
    tail = ["Стан (валідатор)", "Символів", "Записів", "Відхилення від заданого обсягу"] + \
           [name for _, name in SHARES] + ["Правило 2.5: порушено / оцінено", "Придатні записи, %", "Записи для П1, %"]
    lw = [5, 11] + ([24] if pilot else [8, 8]) + [7, 34, 12] + ([10] if pilot else []) + [12, 12] + \
         ([12] if pilot else [])
    return lead + tail, lw + [14, 10, 9, 13] + [10] * len(SHARES) + [14, 11, 11]


def rule_rows(runs):
    rows = []
    for r in runs:
        if not readable(r):
            continue
        for rule in RULES:
            d = r["summary"]["per_rule"].get(rule)
            if d:
                rows.append([r["file"], r["model"], r["level"], rule, d[OK], d[BAD], d[NA]])
    return rows


def violation_rows(runs):
    return [[r["file"], r["model"], r["level"], c["stage"], c["rule"], c["table"], number(c["record"]), c["record_id"],
             c["result"], c["detail"]]
            for r in runs if r["run"] is not None for c in r["run"].rows if c["result"] != OK]


def models_table(runs):
    rows = []
    for m in sorted({r["model"] for r in runs}):
        for lv in [x for x in LEVELS if any(r["level"] == x and r["model"] == m for r in runs)]:
            rr = [r for r in runs if r["model"] == m and r["level"] == lv]
            ok = [r for r in rr if readable(r)]
            ys = [pct(r["summary"], "overall_main") for r in ok]
            lo, hi = middle_half(ys)
            complete = [r for r in ok if r["summary"].get("format_1_2") == "так"]
            ys_clean = [y for y in ys if y is not None]
            rows.append({"model": m, "level": lv, "runs": len(rr), "readable": len(ok),
                         "complete_%": round(100 * len(complete) / len(rr), 1) if rr else None,
                         "median": median(ys), "q1": lo, "q3": hi,
                         "min": min(ys_clean) if ys_clean else None, "max": max(ys_clean) if ys_clean else None,
                         "g2": median([pct(r["summary"], "group2_main") for r in ok]),
                         "g3": median([pct(r["summary"], "group3") for r in ok]),
                         "g3x": median([pct(r["summary"], "group3_core") for r in ok]),
                         "r25": sum(1 for r in ok if (r["summary"].get("rule_2_5") or {}).get("порушено")),
                         "usable": median([pct(r["summary"], "usable_records") for r in ok]),
                         "p1": median([pct(r["summary"], "p1") for r in ok]),
                         "presses": median([r["presses"] for r in rr if isinstance(r["presses"], (int, float))])})
    return rows


def decisions(runs, table):
    """H1: частка групи 3 > частки групи 2 (без 2.5) більш ніж у половині запусків кожної моделі.
    H2: середні половини загальної частки (без 2.5) однієї пари моделей не перетинаються щонайменше на
    двох рівнях обсягу, і меншу частку має та сама модель."""
    h1 = []
    for m in sorted({r["model"] for r in runs}):
        rr = [r for r in runs if r["model"] == m and r["level"] in LEVELS]
        wins = 0
        for r in rr:
            g2, g3 = pct(r["summary"], "group2_main"), pct(r["summary"], "group3")
            if readable(r) and g2 is not None and g3 is not None and g3 > g2:
                wins += 1
        h1.append([m, wins, len(rr), "так" if rr and wins > len(rr) / 2 else "ні"])
    h1_ok = bool(h1) and all(x[3] == "так" for x in h1)

    h2, apart = [], {}
    for lv in LEVELS:
        rows = [t for t in table if t["level"] == lv and t["q1"] is not None]
        for a in range(len(rows)):
            for b in range(a + 1, len(rows)):
                x, y = rows[a], rows[b]
                better = x["model"] if x["q3"] < y["q1"] else (y["model"] if y["q3"] < x["q1"] else None)
                pair = f"{x['model']} / {y['model']}"
                h2.append([lv, pair, f"{x['q1']}–{x['q3']}", f"{y['q1']}–{y['q3']}",
                           "ні" if better else "так", better or "—"])
                if better:
                    apart.setdefault((pair, better), []).append(lv)
    stable = [f"{p}: менша частка в {b} (рівні: {', '.join(lv)})" for (p, b), lv in apart.items() if len(lv) >= 2]
    return h1, h1_ok, h2, stable


def repeat_rows(runs):
    rows = []
    for m in sorted({r["model"] for r in runs}):
        for lv in LEVELS:
            rr = [r for r in runs if r["model"] == m and r["level"] == lv]
            if not rr:
                continue
            for rule in RULES:
                k = sum(1 for r in rr if r["summary"] and r["summary"]["per_rule"].get(rule, {}).get(BAD))
                cat = "не траплялося" if k == 0 else ("постійне" if k > len(rr) / 2 else "епізодичне")
                rows.append([m, lv, rule, k, len(rr), cat])
    return rows


def write(path, runs, pilot):
    wb = Workbook()
    heads, widths = run_headers(pilot)
    xlsx_style.sheet(wb, "Запуски", heads, run_rows(runs, pilot), widths, first=True)
    xlsx_style.sheet(wb, "Правила", ["Файл відповіді", "Модель", "Обсяг", "Правило", "Виконано", "Порушено",
                                     "Не оцінюється"], rule_rows(runs), [34, 11, 7, 9, 10, 10, 12])
    xlsx_style.sheet(wb, "Порушення", ["Файл відповіді", "Модель", "Обсяг", "Етап", "Правило", "Таблиця",
                                       "№ запису", "id запису", "Результат", "Деталі"],
                     violation_rows(runs), [34, 11, 7, 6, 8, 24, 9, 11, 14, 60])
    result = None
    if not pilot:
        table = models_table(runs)
        cols = ["model", "level", "runs", "readable", "complete_%", "median", "q1", "q3", "min", "max",
                "g2", "g3", "g3x", "r25", "usable", "p1", "presses"]
        xlsx_style.sheet(wb, "Модель і обсяг",
                         ["Модель", "Обсяг", "Запусків", "Прочитано", "Повних відповідей, %",
                          "Медіана загальної частки без 2.5, %", "Q1, %", "Q3, %", "Мінімум, %", "Максимум, %",
                          "Медіана групи 2 без 2.5, %", "Медіана групи 3, %", "Медіана групи 3 без 3.1–3.3, %",
                          "Запусків із порушенням 2.5", "Медіана придатних записів, %", "Медіана П1, %",
                          "Медіана натискань «Продовжити»"],
                         [[t[c] for c in cols] for t in table], [11, 7, 9, 10, 11, 14, 8, 8, 10, 10, 12, 11, 14,
                                                                   12, 13, 10, 13])
        h1, h1_ok, h2, stable = decisions(runs, table)
        ws = xlsx_style.sheet(wb, "Рішення", ["H1: модель", "Запусків, де група 3 > група 2 без 2.5",
                                              "Усього запусків", "Більш ніж у половині"], h1, [26, 18, 14, 16])
        ws.auto_filter.ref = None
        r = len(h1) + 3
        ws.cell(row=r, column=1, value="Рішення H1").font = xlsx_style.BOLD
        ws.cell(row=r, column=2, value="H1 підтверджено" if h1_ok else "H1 не підтверджено")
        r += 2
        xlsx_style.header(ws, ["H2: обсяг", "Пара моделей", "Q1–Q3 першої, %", "Q1–Q3 другої, %",
                               "Проміжки перетинаються", "Менша частка в"], row=r, freeze=False)
        xlsx_style.body(ws, h2, start=r + 1)
        r += len(h2) + 2
        ws.cell(row=r, column=1, value="Рішення H2").font = xlsx_style.BOLD
        ws.cell(row=r, column=2, value="H2 підтверджено" if stable else "H2 не підтверджено")
        for k, line in enumerate(stable, start=1):
            ws.cell(row=r + k, column=2, value=line)
        xlsx_style.sheet(wb, "Повторюваність", ["Модель", "Обсяг", "Правило", "Запусків із порушенням",
                                                "Усього запусків", "Категорія"], repeat_rows(runs),
                         [11, 7, 9, 14, 12, 15])
        result = table, h1, h1_ok, h2, stable
    wb.save(path)
    return result


SHORT = [("group1", "Г1"), ("group2", "Г2"), ("group2_main", "Г2*"), ("group3", "Г3"), ("group3_core", "Г3°"),
         ("overall", "Заг"), ("overall_main", "Заг*")]
LEGEND = [
    "Позначення",
    "  Г1, Г2, Г3 — частка порушених перевірок, %: Г1 — правила 1.3–1.6, Г2 — 2.1–2.6, Г3 — 3.1–3.10",
    "  Г2* — група 2 без правила 2.5; Г3° — група 3 без правил 3.1–3.3",
    "  Заг — усі правила 1.3–3.10; Заг* — усі правила без 2.5 (за Г2*, Г3 і Заг* перевіряються H1 і H2)",
    "  2.5 — правило 2.5 «дата завершення активної вакансії пізніша за дату запуску»: порушено / оцінено",
    "  «порушено / оцінено» — кількість перевірок; у дужках — перевірки з результатом «не оцінюється»",
    "  «—» — значення немає: відповідь не прочитано або поле журналу не заповнено",
]


def _cell(v):
    """Текст клітинки й ознака числа (числа вирівнюються праворуч)."""
    if v is None or v == "":
        return "—", False
    if isinstance(v, (datetime, date)):
        return v.strftime("%d.%m.%Y"), False
    if isinstance(v, float):
        return f"{v:.2f}".rstrip("0").rstrip(".").replace(".", ","), True
    if isinstance(v, int) and not isinstance(v, bool):
        return str(v), True
    return str(v), False


def text_table(headers, rows):
    cells = [[(h, False) for h in headers]] + [[_cell(v) for v in r] for r in rows]
    w = [max(len(r[i][0]) for r in cells) for i in range(len(headers))]

    def line(r):
        return "  ".join(t.rjust(w[i]) if num else t.ljust(w[i]) for i, (t, num) in enumerate(r)).rstrip()
    return [line(cells[0]), "  ".join("-" * x for x in w)] + [line(r) for r in cells[1:]]


def _groups(runs, pilot):
    """Стовпці зведень за правилами: для пілоту — кожен запуск, для основних — «модель × обсяг»."""
    ok = [r for r in runs if readable(r)]
    if pilot:
        return [(f"№{r['no']}", [r]) for r in ok]
    groups = [(f"{m} {lv}", [r for r in ok if r["model"] == m and r["level"] == lv])
              for m in sorted({r["model"] for r in ok}) for lv in LEVELS]
    return [g for g in groups if g[1]]


def rule_matrix(runs, pilot):
    groups = _groups(runs, pilot)
    rows = []
    for rule in RULES:
        row = [rule]
        for _, rr in groups:
            d = [r["summary"]["per_rule"].get(rule, {}) for r in rr]
            bad, ok, na = (sum(x.get(k, 0) for x in d) for k in (BAD, OK, NA))
            row.append(f"{bad}/{bad + ok}" + (f" (+{na})" if na else ""))
        rows.append(row)
    return ["Правило"] + [name for name, _ in groups], rows


def write_txt(path, runs, pilot):
    source = "logs/pilot_log.xlsx, аркуш «Журнал»" if pilot else "logs/run_log.xlsx, аркуш «Основні запуски»"
    out = ["Результати пілоту" if pilot else "Результати основних запусків",
           f"Джерело: журнал {source}; файл створює tools/run_all.py.", ""] + LEGEND

    def section(title, lines):
        out.extend(["", "", title, ""] + lines)

    if pilot:
        section("1. Запуски", text_table(
            ["№", "Модель", "Обсяг", "Назва моделі в інтерфейсі", "Дата", "Хв", "«Продовжити»",
             "Стан", "Оцінка обсягу", "Файл відповіді"],
            [[r["no"], r["model"], r["level"], r["interface"], r["date"], r["duration"],
              r["presses"], r["state"], r["assessment"], r["file"]] for r in runs]))
    else:
        section("1. Запуски", text_table(
            ["№", "Раунд", "Модель", "Обсяг", "Повтор", "Дата", "«Продовжити»", "Стан", "Файл відповіді"],
            [[r["no"], r["round"], r["model"], r["level"], r["rep"], r["date"], r["presses"], r["state"],
              r["file"]] for r in runs]))
    section("2. Обсяг відповідей", text_table(
        ["№", "Модель", "Обсяг", "Стан (валідатор)", "Символів", "Записів", "Відхилення від обсягу",
         "Придатні записи, %", "Записи для П1, %"],
        [[r["no"], r["model"], r["level"], r["summary"]["state"] if r["summary"] else "файл не знайдено",
          (r["summary"] or {}).get("chars"), (r["summary"] or {}).get("records_total"),
          (r["summary"] or {}).get("volume_deviation"), pct(r["summary"], "usable_records"),
          pct(r["summary"], "p1")] for r in runs]))
    section("3. Частки порушених перевірок, %", text_table(
        ["№", "Модель", "Обсяг"] + [name for _, name in SHORT] + ["2.5"],
        [[r["no"], r["model"], r["level"]] + [pct(r["summary"], k) for k, _ in SHORT] + [rule_2_5(r["summary"])]
         for r in runs]))
    heads, rows = rule_matrix(runs, pilot)
    section("4. Перевірки за правилами: порушено / оцінено" + ("" if pilot else " (сума за 10 повторами)"),
            text_table(heads, rows))
    bad = violation_rows(runs)
    section("5. Порушення й неоцінені перевірки", text_table(
        ["Файл відповіді", "Етап", "Правило", "Таблиця", "№ запису", "id запису", "Результат", "Деталі"],
        [[v[0]] + v[3:] for v in bad]) if bad else ["Немає."])
    if not pilot:
        table = models_table(runs)
        section("6. Модель і обсяг", text_table(
            ["Модель", "Обсяг", "Запусків", "Прочитано", "Повних, %", "Медіана Заг*", "Q1", "Q3", "Мінімум",
             "Максимум"],
            [[t[c] for c in ("model", "level", "runs", "readable", "complete_%", "median", "q1", "q3", "min",
                             "max")] for t in table]) + [""] + text_table(
            ["Модель", "Обсяг", "Медіана Г2*", "Медіана Г3", "Медіана Г3°", "Запусків із порушенням 2.5",
             "Медіана придатних, %", "Медіана П1, %", "Медіана «Продовжити»"],
            [[t[c] for c in ("model", "level", "g2", "g3", "g3x", "r25", "usable", "p1", "presses")]
             for t in table]))
        h1, h1_ok, h2, stable = decisions(runs, table)
        section("7. Рішення", text_table(["H1: модель", "Запусків, де Г3 > Г2*", "Усього запусків",
                                         "Більш ніж у половині"], h1) +
                ["", "Рішення H1: " + ("H1 підтверджено" if h1_ok else "H1 не підтверджено"), ""] +
                text_table(["H2: обсяг", "Пара моделей", "Q1–Q3 першої, %", "Q1–Q3 другої, %",
                            "Проміжки перетинаються", "Менша частка в"],
                           [[x[0], x[1], x[2].replace(".", ","), x[3].replace(".", ","), x[4], x[5]] for x in h2]) +
                ["", "Рішення H2: " + ("H2 підтверджено" if stable else "H2 не підтверджено")] +
                [f"  {line}" for line in stable])
        rep = {(m, lv, rule): (k, n) for m, lv, rule, k, n, _ in repeat_rows(runs)}
        keys = [(m, lv) for m in sorted({r["model"] for r in runs}) for lv in LEVELS
                if (m, lv, RULES[0]) in rep]
        section("8. Повторюваність: у скількох повторах порушено правило",
                ["0 — не траплялося; не більше половини повторів — епізодичне; більше половини — постійне.", ""] +
                text_table(["Правило"] + [f"{m} {lv}" for m, lv in keys],
                           [[rule] + ["{}/{}".format(*rep[(m, lv, rule)]) for m, lv in keys] for rule in RULES]))
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")

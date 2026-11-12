"""Перевірка інструментів на даних, для яких правильний результат відомий заздалегідь.

1. Контрольні набори рівнів S, M, L, XL (tools/reference.py): валідатор не повинен знаходити порушень.
2. Пробні випадки: у набір рівня M вноситься по одному порушенню (25 випадків) і додаються 2 контрольні
   випадки; валідатор має знайти саме внесене порушення й нічого іншого.
3. Імітація: три умовні «моделі» × 3 рівні × 10 запусків із випадково внесеними порушеннями з пробних
   випадків. Це не відповіді моделей; імітація перевіряє, що tools/analyze.py правильно рахує медіани,
   квартилі й застосовує правила рішення H1 і H2.

Запуск: python3 tools/check_tools.py
Результат: tool_check/tool_check.xlsx (аркуші «Контрольні набори», «Пробні випадки», «Імітація — модель і
обсяг», «Імітація — рішення»). Проміжні файли не зберігаються: усі набори відтворюються з коду.
"""
import json
import os
import random
import sys

from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import analyze  # noqa: E402
import reference  # noqa: E402
import validate  # noqa: E402
import xlsx_style  # noqa: E402

RUN_DATE = "12.11.2026"
wrap = reference.wrap


def check_text(spec, text, level, run_id):
    run, info = validate.validate(spec, text, run_id, validate.parse_run_date(RUN_DATE))
    info["level"], info["chars"] = level, len(text)
    return run, validate.summarize(spec, run, info, level)


def first(d, table, cond):
    return next(r for r in d[table] if cond(r))


def setv(table, cond, **kw):
    def f(d):
        first(d, table, cond).update(kw)
    return f


def ANY(r):
    return True


def ACTIVE(r):
    return r["status"] == "ACTIVE"


def BOTH(r):
    return r["minSalary"] is not None and r["maxSalary"] is not None and r["salaryPeriod"] == "PER_MONTH"


def MIN_ONLY(r):
    return r["minSalary"] is not None and r["maxSalary"] is None and r["salaryPeriod"] == "PER_MONTH"


def other_company_hr(d):
    v = d["Vacancy"][0]
    v["hrProfileId"] = first(d, "HrProfile", lambda h: h["companyId"] != v["companyId"])["id"]


def pending_company(d):
    v = first(d, "Vacancy", ACTIVE)
    first(d, "Company", lambda c: c["id"] == v["companyId"])["verificationStatus"] = "PENDING"


def duplicate_application(d):
    plain = [a for a in d["Application"] if a["status"] not in ("HIRED", "REJECTED", "WITHDRAWN")]
    a, b = plain[0], plain[1]
    b["vacancyId"], b["studentProfileId"] = a["vacancyId"], a["studentProfileId"]


def history_back(d):
    a = first(d, "Application", lambda r: r["status"] == "INTERVIEW_INVITED")
    a["status"] = "VIEWED"
    last = max((e for e in d["ApplicationStatusHistory"] if e["applicationId"] == a["id"]), key=lambda e: e["createdAt"])
    d["ApplicationStatusHistory"].append({"id": "e999", "applicationId": a["id"], "fromStatus": "INTERVIEW_INVITED",
                                          "toStatus": "VIEWED", "changedByUserId": "u2",
                                          "createdAt": last["createdAt"][:11] + "20:00:00Z"})


def two_hired(d):
    hired = first(d, "Application", lambda r: r["status"] == "HIRED")
    other = first(d, "Application", lambda r: r["vacancyId"] == hired["vacancyId"] and r["status"] == "REJECTED")
    other["status"] = "HIRED"
    for e in d["ApplicationStatusHistory"]:
        if e["applicationId"] == other["id"] and e["toStatus"] == "REJECTED":
            e["toStatus"] = "HIRED"


def cases(spec):
    full = wrap(reference.build(spec, "M"))
    c = lambda rule, what, mutate=None, raw=None: {"rule": rule, "what": what, "mutate": mutate, "raw": raw}  # noqa
    return [
        c("1.1", "відповідь обірвано посередині", raw=full[: len(full) // 2]),
        c("1.1", "замість даних — пояснення", raw="Ось як можна згенерувати такі дані: спочатку створіть компанії..."),
        c("1.2", "немає таблиці історії статусів", lambda d: d.pop("ApplicationStatusHistory")),
        c("1.3", "у вакансії немає closingDate", lambda d: d["Vacancy"][0].pop("closingDate")),
        c("1.4", "рік заснування компанії записано текстом", lambda d: d["Company"][0].update(foundationYear="2015")),
        c("1.4", "дата народження без часу (2004-05-12)",
          lambda d: d["StudentProfile"][0].update(birthDate="2004-05-12")),
        c("1.5", "статус відгуку APPROVED", lambda d: d["Application"][5].update(status="APPROVED")),
        c("1.6", "назва вакансії довша за 200 символів", lambda d: d["Vacancy"][3].update(title="Д" * 201)),
        c("2.1", "мінімальна зарплата −5000", setv("Vacancy", MIN_ONLY, minSalary=-5000)),
        c("2.2", "мінімум 40 000, максимум 30 000", setv("Vacancy", BOTH, minSalary=40000, maxSalary=30000)),
        c("2.3", "максимальна зарплата без мінімальної", setv("Vacancy", BOTH, minSalary=None)),
        c("2.4", "2 000 000 за місяць", setv("Vacancy", MIN_ONLY, minSalary=2000000)),
        c("2.5", "активна вакансія з минулою датою завершення", setv("Vacancy", ACTIVE, closingDate="2026-11-01T00:00:00Z")),
        c("2.6", "активна вакансія з причиною закриття", setv("Vacancy", ACTIVE, closeReason="MANUAL")),
        c("3.1", "два користувачі з однаковим e-mail", lambda d: d["User"][5].update(email=d["User"][4]["email"])),
        c("3.2", "відгук на вакансію, якої немає", lambda d: d["Application"][5].update(vacancyId="v99")),
        c("3.3", "навичка, якої немає в довіднику", lambda d: d["VacancySkill"][0].update(skillId=999)),
        c("3.4", "профіль студента належить HR", lambda d: d["StudentProfile"][0].update(userId="u1")),
        c("3.5", "HR однієї компанії публікує вакансію іншої", other_company_hr),
        c("3.6", "активна вакансія в неперевіреної компанії", pending_company),
        c("3.7", "два відгуки одного студента на ту саму вакансію", duplicate_application),
        c("3.8", "відгук датовано раніше за публікацію вакансії",
          setv("Application", lambda r: r["status"] != "HIRED", createdAt="2026-09-20T10:00:00Z")),
        c("3.9", "в історії статус повернуто з INTERVIEW_INVITED на VIEWED", history_back),
        c("3.10", "є найнятий, а вакансія ще активна",
          setv("Vacancy", lambda r: r["status"] == "CLOSED", status="ACTIVE", closedAt=None, closeReason=None)),
        c("3.10", "двоє найнятих на одну вакансію", two_hired),
        c("—", "контроль: пояснення до блоку коду й після нього", raw="Ось дані:\n" + full + "\nГотово."),
        c("—", "контроль: без змін", raw=full),
    ]


def reference_rows(spec):
    rows = []
    for level in spec["levels"]:
        run, s = check_text(spec, wrap(reference.build(spec, level)), level, f"reference_{level}")
        bad = sum(1 for r in run.rows if r["result"] != validate.OK)
        rows.append([level, s["records_total"], len(run.rows), bad, s["volume_deviation"],
                     "як очікувано" if bad == 0 and s["volume_deviation"] == 0 else "НЕ ЗБІГАЄТЬСЯ"])
    return rows


def case_rows(spec):
    rows = []
    for n, c in enumerate(cases(spec), start=1):
        if c["raw"] is not None:
            text = c["raw"]
        else:
            data = reference.build(spec, "M")
            c["mutate"](data)
            text = wrap(data)
        _, s = check_text(spec, text, "M", f"case_{n:02d}")
        bad = {r: v[validate.BAD] for r, v in s["per_rule"].items() if v[validate.BAD]}
        target = bad.pop(c["rule"], 0)
        na = sum(v[validate.NA] for v in s["per_rule"].values())
        ok = (target >= 1 and not bad) if c["rule"] != "—" else not bad
        rows.append([n, c["rule"], c["what"], target, sum(bad.values()), ", ".join(bad), na,
                     "так" if s.get("wrapper_text") else "ні", "як очікувано" if ok else "НЕ ЗБІГАЄТЬСЯ"])
    return rows


PROFILES = {"імітація-А": (0, 2, None), "імітація-Б": (2, 4, None), "імітація-В": (3, 6, "3.")}


def simulation(spec):
    """«Модель» А: 0–2 порушення на запуск, Б: 2–4, В: 3–6 лише з групи 3 (без правил 1.1–1.3)."""
    rnd = random.Random(20261116)
    pool = [c for c in cases(spec) if c["mutate"] is not None and c["rule"] not in ("1.2", "1.3")]
    runs = []
    for model, (lo, hi, prefer) in PROFILES.items():
        for level in analyze.LEVELS:
            for n in range(1, 11):
                data = reference.build(spec, level)
                options = [c for c in pool if prefer is None or c["rule"].startswith(prefer)]
                for c in rnd.sample(options, rnd.randint(lo, hi)):
                    try:
                        c["mutate"](data)
                    except (StopIteration, IndexError):
                        pass  # на малому обсязі для деяких випадків немає потрібного запису
                run, s = check_text(spec, wrap(data), level, f"{model}_{level}_{n:02d}")
                runs.append({"model": model, "level": level, "rep": n, "presses": 0, "run": run, "summary": s})
    table = analyze.models_table(runs)
    return table, analyze.decisions(runs, table)


def main():
    spec = json.load(open(validate.SPEC, encoding="utf-8"))
    wb = Workbook()
    refs = reference_rows(spec)
    xlsx_style.sheet(wb, "Контрольні набори", ["Рівень", "Записів", "Перевірок", "Порушено або не оцінюється",
                                               "Відхилення від обсягу", "Результат"], refs,
                     [8, 10, 11, 14, 12, 16], first=True)
    case_list = case_rows(spec)
    xlsx_style.sheet(wb, "Пробні випадки", ["№", "Правило", "Внесене порушення", "Знайдено порушень цього правила",
                                            "Інших порушень", "Інші правила", "Не оцінюється", "Текст поза блоком",
                                            "Результат"], case_list, [5, 9, 46, 14, 10, 12, 12, 10, 16])
    table, (h1, h1_ok, h2, stable) = simulation(spec)
    cols = ["model", "level", "runs", "median", "q1", "q3", "min", "max", "g2", "g3"]
    xlsx_style.sheet(wb, "Імітація — модель і обсяг",
                     ["«Модель»", "Обсяг", "Запусків", "Медіана загальної частки без 2.5, %", "Q1, %", "Q3, %",
                      "Мінімум, %", "Максимум, %", "Медіана групи 2 без 2.5, %", "Медіана групи 3, %"],
                     [[t[c] for c in cols] for t in table], [13, 7, 9, 14, 8, 8, 10, 10, 12, 11])
    rows = [["H1", x[0], f"{x[1]} із {x[2]}", x[3]] for x in h1]
    rows.append(["H1", "рішення", "", "підтверджено" if h1_ok else "не підтверджено"])
    rows += [["H2", f"{x[0]}: {x[1]}", f"{x[2]} і {x[3]}", f"перетинаються: {x[4]}; менша частка: {x[5]}"] for x in h2]
    rows.append(["H2", "рішення", "; ".join(stable), "підтверджено" if stable else "не підтверджено"])
    xlsx_style.sheet(wb, "Імітація — рішення", ["Гіпотеза", "Об’єкт", "Значення", "Висновок"], rows, [10, 34, 40, 40])
    out = os.path.join(validate.ROOT, "tool_check", "tool_check.xlsx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    return refs, case_list, table, h1, h1_ok, stable


if __name__ == "__main__":
    refs, case_list, table, h1, h1_ok, stable = main()
    for r in refs:
        print(r)
    print(sum(1 for c in case_list if c[-1] == "як очікувано"), "із", len(case_list), "пробних випадків як очікувано")
    for c in case_list:
        if c[-1] != "як очікувано":
            print("  ", c)
    print("H1:", h1, h1_ok)
    print("H2:", stable)

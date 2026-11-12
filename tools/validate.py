"""Валідатор однієї відповіді моделі: перевірка за 22 правилами ЛР5 (1.1–3.10) у чотири етапи.

Етап 1. Формат: відповідь читається як JSON-об'єкт, у ньому є вісім таблиць-масивів (правила 1.1–1.2).
Етап 2. Схема: обов'язкові поля, типи, списки значень, довжина тексту (1.3–1.6).
Етап 3. Правила предметної області для одного запису (2.1–2.6).
Етап 4. Зв'язки й правила між записами (3.1–3.10).

Відповідь не виправляється: читається вміст першого блоку коду, а якщо блоку немає — увесь текст.
Результат кожної перевірки: «виконано», «порушено» або «не оцінюється». «Не оцінюється» ставиться,
коли для перевірки бракує даних, помилку в яких уже враховано іншим правилом (наприклад, відгук
посилається на вакансію, якої немає, — це порушення 3.2, і правило 3.8 для нього не перевіряється).

Правило 2.5 («дата завершення активної вакансії пізніша за сьогоднішню») перевіряється відносно
дати запуску з журналу. Оскільки в запиті дати немає, порушення 2.5 залежать від того, яку дату модель
вважає поточною; тому 2.5 подається окремим показником і не входить до часток, за якими
перевіряються H1 і H2 (основні частки, суфікс _main).

Перевірка однієї відповіді з виведенням усіх перевірок у файл Excel:
  python3 tools/validate.py <файл відповіді> <рівень S|M|L|XL> <дата ДД.ММ.РРРР> <результат.xlsx>
Усі відповіді з журналу перевіряє tools/run_all.py.
"""
import json
import os
import re
import sys
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SPEC = os.path.join(ROOT, "spec.json")

OK, BAD, NA = "виконано", "порушено", "не оцінюється"
STAGE = {"1.1": 1, "1.2": 1, "1.3": 2, "1.4": 2, "1.5": 2, "1.6": 2}
DT_RE = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})$")
DAY_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
PROCESS = ["SENT", "VIEWED", "SHORTLISTED", "INTERVIEW_INVITED", "OFFERED", "HIRED"]
TERMINAL = {"HIRED", "REJECTED", "WITHDRAWN"}
INT32 = 2 ** 31


def stage_of(rule):
    return STAGE.get(rule, 3 if rule.startswith("2.") else 4)


def extract(text):
    """Повертає (рядок для розбору, спосіб, чи є текст поза блоком коду)."""
    m = re.search(r"```[a-zA-Z]*\s*\n", text)
    if not m:
        return text.strip(), "увесь текст", False
    rest = text[m.end():]
    end = rest.find("```")
    body = rest if end < 0 else rest[:end]
    outside = text[:m.start()] + ("" if end < 0 else rest[end + 3:])
    return body.strip(), "перший блок коду", bool(outside.strip())


def parse_dt(v):
    """Дата й час у форматі ISO 8601, який Prisma приймає для полів DateTime."""
    if not isinstance(v, str) or not DT_RE.match(v):
        return None
    try:
        return datetime.fromisoformat(v.replace("Z", "+00:00"))
    except ValueError:
        return None


def to_dt(v):
    """Дата для правил 2.5, 3.8 і 3.9: повна дата й час або дата без часу. Дату без часу Prisma не приймає
    (порушення 1.4), але сама дата відома, тож правила, що від неї залежать, перевіряються."""
    d = parse_dt(v)
    if d is None and isinstance(v, str) and DAY_RE.match(v):
        d = parse_dt(v + "T00:00:00Z")
    return d


def type_ok(f, v):
    t = f["type"]
    if t in ("string", "enum"):
        return isinstance(v, str)
    if t == "int":
        return isinstance(v, int) and not isinstance(v, bool) and -INT32 <= v < INT32
    if t == "datetime":
        return parse_dt(v) is not None
    return False


class Run:
    def __init__(self, spec, run_id):
        self.spec, self.run_id = spec, run_id
        self.rows = []
        self.bad_records = {}  # (таблиця, номер запису) -> групи правил, порушені в записі

    def check(self, rule, result, table="", idx="", rid="", detail=""):
        self.rows.append({"run_id": self.run_id, "stage": stage_of(rule), "rule": rule, "table": table,
                          "record": idx, "record_id": rid, "result": result, "detail": detail})
        if result == BAD and table != "" and idx != "":
            self.bad_records.setdefault((table, idx), set()).add(rule.split(".")[0])


def validate(spec, text, run_id, run_date):
    run = Run(spec, run_id)
    body, mode, wrapper = extract(text)
    info = {"run_id": run_id, "extraction": mode, "wrapper_text": wrapper, "readable": False}

    # Етап 1. Формат
    try:
        data = json.loads(body)
        ok = isinstance(data, dict)
    except ValueError as e:
        data, ok = None, False
        info["parse_error"] = str(e)[:120]
    run.check("1.1", OK if ok else BAD, detail="" if ok else "не JSON-об'єкт")
    if not ok:
        return run, info
    info["readable"] = True
    tables = spec["tables"]
    missing = [t for t in tables if not isinstance(data.get(t), list)]
    run.check("1.2", BAD if missing else OK, detail=", ".join(missing))
    recs = {t: (data[t] if isinstance(data.get(t), list) else []) for t in tables}
    info["counts"] = {t: len(recs[t]) for t in tables}

    # Етап 2. Схема: valid[(table, i)] = множина полів, які пройшли 1.3–1.6
    valid = {}
    for t, tdef in tables.items():
        for i, r in enumerate(recs[t]):
            good = set()
            if not isinstance(r, dict):
                run.check("1.3", BAD, t, i, "", "запис не є об'єктом")
                valid[(t, i)] = good
                continue
            rid = r.get("id", "")
            absent = [n for n, f in tdef["fields"].items() if f["required"] and r.get(n) is None]
            run.check("1.3", BAD if absent else OK, t, i, rid, ", ".join(absent))
            for n, f in tdef["fields"].items():
                v = r.get(n)
                if v is None:
                    if not f["required"]:
                        good.add(n)  # порожнє необов'язкове поле коректне, але до знаменників не входить
                    continue
                if not type_ok(f, v):
                    run.check("1.4", BAD, t, i, rid, f"{n}={str(v)[:40]}")
                    if f["type"] == "datetime" and to_dt(v) is not None:
                        good.add(n)  # дата без часу: формат порушено, але дата придатна для правил 2.5, 3.8, 3.9
                    continue
                run.check("1.4", OK, t, i, rid, n)
                if f["type"] == "enum":
                    if v not in f["values"]:
                        run.check("1.5", BAD, t, i, rid, f"{n}={v}")
                        continue
                    run.check("1.5", OK, t, i, rid, n)
                if "max" in f:
                    if len(v) > f["max"]:
                        run.check("1.6", BAD, t, i, rid, f"{n}: {len(v)} > {f['max']}")
                        continue
                    run.check("1.6", OK, t, i, rid, n)
                good.add(n)
            valid[(t, i)] = good

    def has(t, i, *names):
        return all(n in valid[(t, i)] for n in names)

    def val(t, i, n):
        r = recs[t][i]
        return r.get(n) if isinstance(r, dict) and n in valid[(t, i)] else None

    def rid(t, i):
        r = recs[t][i]
        return r.get("id", "") if isinstance(r, dict) else ""

    # Етап 3. Правила одного запису (усі стосуються Vacancy)
    today = run_date
    limits = spec["salary_limits"]
    for i, r in enumerate(recs["Vacancy"]):
        if not isinstance(r, dict):
            continue
        v = lambda n: val("Vacancy", i, n)
        mn, mx = r.get("minSalary"), r.get("maxSalary")
        if mn is not None:
            if has("Vacancy", i, "minSalary"):
                run.check("2.1", OK if mn >= 0 else BAD, "Vacancy", i, rid("Vacancy", i), f"minSalary={mn}")
            else:
                run.check("2.1", NA, "Vacancy", i, rid("Vacancy", i))
        if mn is not None and mx is not None:
            if has("Vacancy", i, "minSalary", "maxSalary"):
                run.check("2.2", OK if mx >= mn else BAD, "Vacancy", i, rid("Vacancy", i), f"{mn}..{mx}")
            else:
                run.check("2.2", NA, "Vacancy", i, rid("Vacancy", i))
        if mx is not None:
            run.check("2.3", OK if mn is not None else BAD, "Vacancy", i, rid("Vacancy", i))
        if mn is not None or mx is not None:
            if has("Vacancy", i, "minSalary", "maxSalary", "salaryPeriod"):
                lim = limits.get(v("salaryPeriod") or "PER_MONTH")
                over = [x for x in (mn, mx) if x is not None and x > lim]
                run.check("2.4", BAD if over else OK, "Vacancy", i, rid("Vacancy", i), f"межа {lim}")
            else:
                run.check("2.4", NA, "Vacancy", i, rid("Vacancy", i))
        if r.get("status") == "ACTIVE":
            if has("Vacancy", i, "status", "closingDate"):
                cd = to_dt(v("closingDate")).date()
                run.check("2.5", OK if cd > today else BAD, "Vacancy", i, rid("Vacancy", i), f"closingDate={cd}")
            else:
                run.check("2.5", NA, "Vacancy", i, rid("Vacancy", i))
        if has("Vacancy", i, "status", "closedAt", "closeReason"):
            closed_fields = r.get("closedAt") is not None or r.get("closeReason") is not None
            bad = closed_fields and v("status") != "CLOSED"
            run.check("2.6", BAD if bad else OK, "Vacancy", i, rid("Vacancy", i), f"status={v('status')}")
        else:
            run.check("2.6", NA, "Vacancy", i, rid("Vacancy", i))

    # Етап 4. Зв'язки й правила між записами
    index = {}  # таблиця -> {id: номер першого запису з цим id}
    for t, tdef in tables.items():
        for key in tdef["unique"]:
            seen = {}
            for i in range(len(recs[t])):
                if not isinstance(recs[t][i], dict):
                    continue
                if not has(t, i, *key):
                    if all(recs[t][i].get(k) is not None for k in key):
                        run.check("3.1", NA, t, i, rid(t, i), "+".join(key))
                    continue
                k = tuple(val(t, i, n) for n in key)
                if k in seen:
                    run.check("3.1", BAD, t, i, rid(t, i), f"{'+'.join(key)} повторює запис {seen[k]}")
                else:
                    seen[k] = i
                    run.check("3.1", OK, t, i, rid(t, i), "+".join(key))
            if key == ["id"]:
                index[t] = {k[0]: i for k, i in seen.items()}

    def target(t, i, field):
        """Номер запису, на який веде посилання, або None (якщо посилання порожнє чи не перевірене)."""
        ref = tables[t]["fields"][field]["ref"]
        x = val(t, i, field)
        return None if x is None else index.get(ref, {}).get(x)

    for t, tdef in tables.items():
        for n, f in tdef["fields"].items():
            for i in range(len(recs[t])):
                if not isinstance(recs[t][i], dict) or recs[t][i].get(n) is None:
                    continue
                if "ref" in f:
                    if not has(t, i, n) or f["ref"] in missing:
                        run.check("3.2", NA, t, i, rid(t, i), n)
                    else:
                        found = target(t, i, n) is not None
                        run.check("3.2", OK if found else BAD, t, i, rid(t, i), f"{n}={val(t, i, n)}")
                if "dict" in f:
                    if not has(t, i, n):
                        run.check("3.3", NA, t, i, rid(t, i), n)
                    else:
                        ok = str(val(t, i, n)) in spec["dictionaries"][f["dict"]]
                        run.check("3.3", OK if ok else BAD, t, i, rid(t, i), f"{n}={val(t, i, n)}")

    # 3.4 роль користувача відповідає профілю
    for t, role in (("HrProfile", "HR"), ("StudentProfile", "STUDENT")):
        for i in range(len(recs[t])):
            if not isinstance(recs[t][i], dict):
                continue
            u = target(t, i, "userId")
            if u is None or not has("User", u, "role"):
                run.check("3.4", NA, t, i, rid(t, i))
            else:
                actual = val("User", u, "role")
                run.check("3.4", OK if actual == role else BAD, t, i, rid(t, i), f"роль {actual}")

    # 3.5 вакансія і HR з однієї компанії; 3.6 активна вакансія лише в перевіреній компанії
    for i in range(len(recs["Vacancy"])):
        if not isinstance(recs["Vacancy"][i], dict):
            continue
        h = target("Vacancy", i, "hrProfileId")
        if h is None or target("Vacancy", i, "companyId") is None or not has("HrProfile", h, "companyId"):
            run.check("3.5", NA, "Vacancy", i, rid("Vacancy", i))
        else:
            same = val("HrProfile", h, "companyId") == val("Vacancy", i, "companyId")
            run.check("3.5", OK if same else BAD, "Vacancy", i, rid("Vacancy", i))
        if val("Vacancy", i, "status") == "ACTIVE":
            c = target("Vacancy", i, "companyId")
            if c is None or not has("Company", c, "verificationStatus"):
                run.check("3.6", NA, "Vacancy", i, rid("Vacancy", i))
            else:
                vs = val("Company", c, "verificationStatus")
                run.check("3.6", OK if vs == "APPROVED" else BAD, "Vacancy", i, rid("Vacancy", i), vs)

    # 3.7 один відгук студента на вакансію; 3.8 відгук у межах відкритості вакансії
    pairs = {}
    for i in range(len(recs["Application"])):
        if not isinstance(recs["Application"][i], dict):
            continue
        a = rid("Application", i)
        if has("Application", i, "vacancyId", "studentProfileId"):
            k = (val("Application", i, "vacancyId"), val("Application", i, "studentProfileId"))
            run.check("3.7", BAD if k in pairs else OK, "Application", i, a,
                      f"повторює відгук {pairs[k]}" if k in pairs else "")
            pairs.setdefault(k, a or f"№ {i + 1}")
        else:
            run.check("3.7", NA, "Application", i, a)
        vi = target("Application", i, "vacancyId")
        if vi is None or not has("Application", i, "createdAt") or not has("Vacancy", vi, "closingDate", "publishedAt"):
            run.check("3.8", NA, "Application", i, a)
            continue
        created = to_dt(val("Application", i, "createdAt"))
        pub = val("Vacancy", vi, "publishedAt")
        closing = to_dt(val("Vacancy", vi, "closingDate")).date()
        bad = pub is None or created < to_dt(pub) or created.date() > closing
        run.check("3.8", BAD if bad else OK, "Application", i, a, f"створено {created.date()}, вакансія до {closing}")

    # 3.9 історія статусів
    hist = {}
    for j in range(len(recs["ApplicationStatusHistory"])):
        if not isinstance(recs["ApplicationStatusHistory"][j], dict):
            continue
        ai = target("ApplicationStatusHistory", j, "applicationId")
        if ai is not None:
            hist.setdefault(ai, []).append(j)
    for i in range(len(recs["Application"])):
        if not isinstance(recs["Application"][i], dict):
            continue
        a = rid("Application", i)
        events = hist.get(i, [])
        h = "ApplicationStatusHistory"
        if h in missing or not has("Application", i, "status") or any(not has(h, j, "toStatus", "fromStatus", "createdAt") for j in events):
            run.check("3.9", NA, "Application", i, a)
            continue
        events.sort(key=lambda j: to_dt(val(h, j, "createdAt")))
        problem = history_problem([(val(h, j, "fromStatus"), val(h, j, "toStatus")) for j in events],
                                  val("Application", i, "status"))
        run.check("3.9", BAD if problem else OK, "Application", i, a, problem or "")

    # 3.10 не більше одного найнятого; вакансія з найнятим закрита з причиною HIRED
    hired = {}
    for i in range(len(recs["Application"])):
        if isinstance(recs["Application"][i], dict) and val("Application", i, "status") == "HIRED":
            vi = target("Application", i, "vacancyId")
            if vi is not None:
                hired.setdefault(vi, []).append(i)
    for vi, apps in sorted(hired.items()):
        if not has("Vacancy", vi, "status", "closeReason"):
            run.check("3.10", NA, "Vacancy", vi, rid("Vacancy", vi))
            continue
        problems = []
        if len(apps) > 1:
            problems.append(f"найнято {len(apps)}")
        if val("Vacancy", vi, "status") != "CLOSED" or val("Vacancy", vi, "closeReason") != "HIRED":
            problems.append(f"вакансія {val('Vacancy', vi, 'status')}/{val('Vacancy', vi, 'closeReason')}")
        run.check("3.10", BAD if problems else OK, "Vacancy", vi, rid("Vacancy", vi), "; ".join(problems))

    info["records_total"] = sum(len(recs[t]) for t in tables)
    return run, info


def history_problem(events, final_status):
    """Повертає опис першої проблеми в історії статусів або None, якщо історія коректна."""
    if not events:
        return "немає історії"
    if events[0] != (None, "SENT"):
        return f"перший запис {events[0][0]} -> {events[0][1]}, а має бути null -> SENT"
    current, before_withdraw = "SENT", None
    for k, (frm, to) in enumerate(events[1:], start=2):
        if frm != current:
            return f"запис {k}: fromStatus {frm}, а попередній toStatus {current}"
        if current == "WITHDRAWN":
            if to != before_withdraw:
                return f"запис {k}: з WITHDRAWN можна лише повернутися до {before_withdraw}"
        elif current in TERMINAL:
            return f"запис {k}: зміна після {current}"
        elif to == "WITHDRAWN":
            before_withdraw = current
        elif to != "REJECTED" and (to not in PROCESS or PROCESS.index(to) <= PROCESS.index(current)):
            return f"запис {k}: перехід {current} -> {to} не дозволено"
        current = to
    if current != final_status:
        return f"останній статус в історії {current}, у відгуку {final_status}"
    return None


GROUPS = {
    "group1": ["1.3", "1.4", "1.5", "1.6"],
    "group2": ["2.1", "2.2", "2.3", "2.4", "2.5", "2.6"],
    "group2_main": ["2.1", "2.2", "2.3", "2.4", "2.6"],
    "group3": [f"3.{k}" for k in range(1, 11)],
    "group3_core": [f"3.{k}" for k in range(4, 11)],
    "rule_2_5": ["2.5"],
}
GROUPS["overall"] = GROUPS["group1"] + GROUPS["group2"] + GROUPS["group3"]
GROUPS["overall_main"] = GROUPS["group1"] + GROUPS["group2_main"] + GROUPS["group3"]


def expected_counts(spec, level):
    """Кількість записів, задана в запиті: таблиця -> (число, чи це мінімум)."""
    lv = spec["levels"][level]
    return {t: (n, t in lv["minimum"]) for t, n in lv["counts"].items()}


def summarize(spec, run, info, level):
    s = dict(info)
    per_rule = {}
    for r in run.rows:
        d = per_rule.setdefault(r["rule"], {OK: 0, BAD: 0, NA: 0})
        d[r["result"]] += 1
    s["per_rule"] = per_rule

    def share(rules):
        bad = sum(per_rule.get(x, {}).get(BAD, 0) for x in rules)
        done = bad + sum(per_rule.get(x, {}).get(OK, 0) for x in rules)
        return {"порушено": bad, "оцінено": done, "частка_%": round(100 * bad / done, 2) if done else None}

    if not info["readable"]:
        s["state"] = "не вдалося прочитати"
        return s
    s["state"] = "прочитано"
    s["format_1_2"] = "так" if per_rule.get("1.2", {}).get(BAD, 0) == 0 else "ні"
    for name, rules in GROUPS.items():
        s[name] = share(rules)
    dev = 0
    for t, (n, is_min) in expected_counts(spec, level).items():
        got = info["counts"][t]
        dev += max(0, n - got) if is_min else abs(n - got)
    s["volume_deviation"] = dev
    total = info["records_total"]
    bad_any = len(run.bad_records)
    s["usable_records"] = {"придатні": total - bad_any, "усього": total,
                           "частка_%": round(100 * (total - bad_any) / total, 2) if total else None}
    fmt_ok = total - sum(1 for g in run.bad_records.values() if "1" in g)
    p1 = sum(1 for g in run.bad_records.values() if "1" not in g and g & {"2", "3"})
    s["p1"] = {"з_правильним_форматом": fmt_ok, "з_них_з_порушеннями_груп_2_3": p1,
               "частка_%": round(100 * p1 / fmt_ok, 2) if fmt_ok else None}
    return s


def parse_run_date(v):
    """Дата запуску: значення дати з Excel або рядок ДД.ММ.РРРР."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return datetime.strptime(str(v).strip(), "%d.%m.%Y").date()


def check(spec, raw_path, level, run_date, run_id=None):
    """Перевіряє файл відповіді; повертає (усі перевірки, підсумок)."""
    text = open(raw_path, encoding="utf-8").read()
    run_id = run_id or os.path.splitext(os.path.basename(raw_path))[0]
    run, info = validate(spec, text, run_id, parse_run_date(run_date))
    info["level"], info["chars"] = level, len(text)
    return run, summarize(spec, run, info, level)


def number(idx):
    """Номер запису в масиві таблиці для виведення: від 1 (у перевірках зберігається індекс від 0)."""
    return idx + 1 if isinstance(idx, int) else idx


def pct(s, key):
    return (s.get(key) or {}).get("частка_%") if s else None


def write_xlsx(path, run, s):
    """Усі перевірки однієї відповіді: аркуші «Підсумок» і «Перевірки»."""
    from openpyxl import Workbook
    import xlsx_style
    wb = Workbook()
    rows = [["Файл", s["run_id"]], ["Рівень обсягу", s["level"]], ["Стан", s["state"]],
            ["Записів", s.get("records_total")], ["Відхилення від заданого обсягу", s.get("volume_deviation")]]
    rows += [[f"Частка порушених перевірок: {name}, %", pct(s, name)] for name in GROUPS]
    rows += [[f"Правило {rule}: виконано / порушено / не оцінюється", f"{d[OK]} / {d[BAD]} / {d[NA]}"]
             for rule, d in sorted(s["per_rule"].items(), key=lambda x: [int(p) for p in x[0].split(".")])]
    xlsx_style.sheet(wb, "Підсумок", ["Показник", "Значення"], rows, [52, 30], first=True)
    cols = ["stage", "rule", "table", "record", "record_id", "result", "detail"]
    xlsx_style.sheet(wb, "Перевірки", ["Етап", "Правило", "Таблиця", "№ запису", "id запису", "Результат", "Деталі"],
                     [[r[c] for c in cols[:3]] + [number(r["record"])] + [r[c] for c in cols[4:]] for r in run.rows],
                     [6, 9, 24, 9, 12, 15, 60])
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    spec = json.load(open(SPEC, encoding="utf-8"))
    raw, level, run_date, out = sys.argv[1:5]
    run, summary = check(spec, raw, level, run_date)
    write_xlsx(out, run, summary)
    print(json.dumps({k: pct(summary, k) for k in GROUPS}, ensure_ascii=False))

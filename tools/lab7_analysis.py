"""ЛР7: набір даних 90 основних запусків, основний аналіз, перевірки стійкості, повторна перевірка знайдених
порушень і знеособлення відповідей.

  python3 tools/lab7_analysis.py

Запускається з кореня репозиторію test-data-generation-research. Валідатор (tools/validate.py) і правила
підсумків не змінюються: скрипт лише застосовує їх до відповідей із журналу logs/run_log.xlsx.
Результати: results/lab7_dataset.xlsx і results/lab7.txt; знеособлені копії — answers_anonymized/.
"""
import json
import os
import random
import re
import sys
from datetime import datetime as _dt
from statistics import median as _median

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import validate  # noqa: E402
import xlsx_style as xs  # noqa: E402
from validate import BAD, NA, OK  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402

SPEC = json.load(open(os.path.join(ROOT, "spec.json"), encoding="utf-8"))
TABLES = list(SPEC["tables"])
LEVELS, MODELS = ["S", "M", "L"], ["ChatGPT", "Gemini", "DeepSeek"]
RULES = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "2.1", "2.2", "2.3", "2.4", "2.5", "2.6"] + \
        [f"3.{k}" for k in range(1, 11)]
GROUPS = {
    "g1": ["1.3", "1.4", "1.5", "1.6"],
    "g2": ["2.1", "2.2", "2.3", "2.4", "2.6"],          # група 2 без 2.5
    "g3": [f"3.{k}" for k in range(1, 11)],
    "g3core": [f"3.{k}" for k in range(4, 11)],          # група 3 без механічних 3.1–3.3
}
GROUPS["main"] = GROUPS["g1"] + GROUPS["g2"] + GROUPS["g3"]   # основний показник: 1.3–3.10 без 2.5
GROUPS["with25"] = GROUPS["main"] + ["2.5"]
GROUPS["r25"] = ["2.5"]


# ---------------------------------------------------------------- журнал і оцінювання відповідей

def read_log(path=os.path.join(ROOT, "logs", "run_log.xlsx")):
    ws = load_workbook(path, data_only=True)["Основні запуски"]
    names = [c.value for c in ws[1]]
    rows = [dict(zip(names, r)) for r in ws.iter_rows(min_row=2, values_only=True)]
    return [r for r in rows if r.get("Файл відповіді")]


def runs_from_log(rows, prefer_original=False):
    """90 запусків. Якщо для номера є два рядки (файл відповіді й копія), береться рядок «Входить до аналізу = так»,
    а з prefer_original — файл відповіді."""
    by_no = {}
    for r in rows:
        by_no.setdefault(int(r["№"]), []).append(r)
    out = []
    for no in sorted(by_no):
        cand = by_no[no]
        if len(cand) > 1:
            pick = [r for r in cand if (r["Входить до аналізу"] == "так") != prefer_original]
            cand = pick or cand
        out.append(cand[0])
    assert len(out) == 90, len(out)
    return out


def share(per, rules):
    bad = sum(per.get(x, {}).get(BAD, 0) for x in rules)
    done = bad + sum(per.get(x, {}).get(OK, 0) for x in rules)
    return bad, done, (round(100 * bad / done, 3) if done else None)


def evaluate(r):
    rel = r["Файл відповіді"]
    text = open(os.path.join(ROOT, rel), encoding="utf-8").read()
    run, s = validate.check(SPEC, os.path.join(ROOT, rel), r["Обсяг"], r["Дата запуску"], run_id=rel)
    per = s["per_rule"]
    rec = {
        "no": int(r["№"]), "round": int(str(r["Раунд"]).lstrip("R")), "rep": int(r["Повтор"]), "model": r["Модель"],
        "level": r["Обсяг"], "executor": r["Виконавиця"], "date": r["Дата запуску"],
        "duration": r["Тривалість відповіді, хв"], "interface": (r["Назва моделі в інтерфейсі"] or "").strip(),
        "presses": r["Натискань «Продовжити»"], "state": r["Стан відповіді"], "included": r["Входить до аналізу"],
        "file": rel, "chars": len(text), "readable": s["readable"], "records": s.get("records_total") or 0,
        "counts": s.get("counts") or {}, "voldev": s.get("volume_deviation"), "per_rule": per,
        "violations": [c for c in run.rows if c["result"] == BAD],
        "not_evaluated": [c for c in run.rows if c["result"] == NA],
        "usable": s.get("usable_records"), "p1": s.get("p1"),
    }
    for key, rules in GROUPS.items():
        bad, done, pct = share(per, rules) if s["readable"] else (None, None, None)
        rec[key] = {"bad": bad, "done": done, "pct": pct if done else None}
    rec["has_share"] = rec["main"]["pct"] is not None
    rec["n_viol"] = sum(1 for c in rec["violations"] if c["rule"] != "2.5")   # з 1.1 і 1.2
    rec["any_viol"] = rec["n_viol"] > 0
    return rec


# ---------------------------------------------------------------- описові підсумки й правила рішення

def median(v):
    v = [x for x in v if x is not None]
    return round(_median(v), 3) if v else None


def middle_half(values):
    v = sorted(x for x in values if x is not None)
    if not v:
        return None, None
    if len(v) == 1:
        return v[0], v[0]
    return median(v[: len(v) // 2]), median(v[(len(v) + 1) // 2:])


def quart(values):
    ys = [v for v in values if v is not None]
    q1, q3 = middle_half(ys)
    return {"n": len(ys), "median": median(ys), "q1": q1, "q3": q3,
            "min": min(ys) if ys else None, "max": max(ys) if ys else None}


def cell(runs, m, lv):
    rr = sorted((r for r in runs if r["model"] == m and r["level"] == lv), key=lambda z: z["rep"])
    q = quart([r["main"]["pct"] for r in rr])
    return dict(q, runs=len(rr), complete=sum(r["state"] == "повна" for r in rr),
                truncated=sum(r["state"] == "обірвана" for r in rr),
                bad=sum(r["main"]["bad"] or 0 for r in rr), done=sum(r["main"]["done"] or 0 for r in rr),
                any=sum(r["any_viol"] and r["has_share"] for r in rr),
                values=[r["main"]["pct"] for r in rr], n_viol=[r["n_viol"] if r["has_share"] else None for r in rr],
                r25_bad=sum(r["r25"]["bad"] or 0 for r in rr), r25_done=sum(r["r25"]["done"] or 0 for r in rr),
                r25_runs=sum((r["r25"]["bad"] or 0) > 0 for r in rr),
                presses=[r["presses"] for r in rr], duration=median([r["duration"] for r in rr]),
                records=[r["records"] for r in rr if r["has_share"]])


def h1(runs, a="g3", b="g2", den="all"):
    """H1: у більш ніж половині запусків моделі частка a більша за частку b. den='all' — знаменник усі запуски
    моделі (запуск без частки рахується «не на користь»; den='share' — лише запуски з часткою."""
    out = {}
    for m in MODELS:
        rr = [r for r in runs if r["model"] == m]
        if den == "share":
            rr = [r for r in rr if r["has_share"]]
        wins = sum(1 for r in rr if r[a]["pct"] is not None and r[b]["pct"] is not None and r[a]["pct"] > r[b]["pct"])
        zeros = sum(1 for r in rr if r[a]["pct"] == 0 and r[b]["pct"] == 0)
        out[m] = {"wins": wins, "n": len(rr), "zeros": zeros, "no_share": sum(not r["has_share"] for r in rr),
                  "ok": wins > len(rr) / 2}
    return out, all(v["ok"] for v in out.values())


def h2(runs, fn=lambda r: r["main"]["pct"]):
    """H2: Q1–Q3 пари моделей не перетинаються на одному рівні щонайменше на двох рівнях обсягу."""
    cells = {(m, lv): quart([fn(r) for r in runs if r["model"] == m and r["level"] == lv])
             for m in MODELS for lv in LEVELS}
    rows, apart = [], {}
    for lv in LEVELS:
        for a in range(3):
            for b in range(a + 1, 3):
                x, y = cells[(MODELS[a], lv)], cells[(MODELS[b], lv)]
                better = None
                if x["q1"] is not None and y["q1"] is not None:
                    better = MODELS[a] if x["q3"] < y["q1"] else (MODELS[b] if y["q3"] < x["q1"] else None)
                pair = f"{MODELS[a]} / {MODELS[b]}"
                rows.append({"level": lv, "pair": pair, "x": (x["q1"], x["q3"]), "y": (y["q1"], y["q3"]),
                             "better": better})
                if better:
                    apart.setdefault((pair, better), []).append(lv)
    stable = {k: v for k, v in apart.items() if len(v) >= 2}
    return {"cells": cells, "rows": rows, "ok": bool(stable), "stable": stable,
            "apart": sum(1 for x in rows if x["better"])}


def totals(runs):
    t = {}
    for m in MODELS:
        rr = [r for r in runs if r["model"] == m]
        t[m] = {"runs": len(rr), "bad": sum(r["main"]["bad"] or 0 for r in rr),
                "done": sum(r["main"]["done"] or 0 for r in rr), "any": sum(r["any_viol"] and r["has_share"] for r in rr),
                "r25_bad": sum(r["r25"]["bad"] or 0 for r in rr), "r25_done": sum(r["r25"]["done"] or 0 for r in rr)}
    return t


def rule_table(runs):
    out = {}
    for rule in RULES:
        for m in MODELS:
            for lv in LEVELS:
                rr = [r for r in runs if r["model"] == m and r["level"] == lv]
                bad = sum(r["per_rule"].get(rule, {}).get(BAD, 0) for r in rr)
                ok = sum(r["per_rule"].get(rule, {}).get(OK, 0) for r in rr)
                out[(rule, m, lv)] = {"bad": bad, "done": bad + ok,
                                      "runs": sum(1 for r in rr if r["per_rule"].get(rule, {}).get(BAD))}
    return out


# ---------------------------------------------------------------- перевірки стійкості

def robustness(P, S1):
    """Кожна перевірка: які запуски, яким показником, і чи змінюються рішення за H1 і H2."""
    def row(code, title, runs, h1_res, h2_res, note=""):
        t = totals(runs)
        return {"code": code, "title": title, "n": len(runs), "totals": t,
                "h1": {m: (v["wins"], v["n"]) for m, v in h1_res[0].items()}, "h1_ok": h1_res[1],
                "h2_apart": h2_res["apart"], "h2_ok": h2_res["ok"], "note": note}
    out = [row("P", "Основний аналіз (90 запусків)", P, h1(P), h2(P))]
    out.append(row("S1", "gpt_L_09 — файл відповіді з тегом <escape> замість копії без тегу", S1, h1(S1), h2(S1)))
    out.append(row("S2", "Знаменник H1 — лише запуски з часткою (без gpt_L_03 і ds_M_10)", P,
                   h1(P, den="share"), h2(P)))
    top = {}
    for m in MODELS:
        worst = max((r for r in P if r["model"] == m and r["has_share"]), key=lambda r: r["n_viol"])
        if worst["n_viol"]:
            top[m] = worst["file"]
    drop = [r for r in P if r["file"] not in top.values()]
    out.append(row("S3", "Без запуску з найбільшою кількістю порушень у кожній моделі, де вони є (" +
                   ", ".join(os.path.basename(f)[:-4] for f in top.values()) + ")", drop, h1(drop), h2(drop)))
    out.append(row("A1", "Показник — абсолютна кількість порушень у запуску", P, h1(P),
                   h2(P, lambda r: r["n_viol"] if r["has_share"] else None)))
    out.append(row("A2", "Загальна частка разом із правилом 2.5", P, h1(P), h2(P, lambda r: r["with25"]["pct"])))
    out.append(row("A3", "Група 3 без правил 3.1–3.3", P, h1(P, "g3core"), h2(P, lambda r: r["g3core"]["pct"])))
    for who in ("Вероніка", "Аліна"):
        rr = [r for r in P if r["executor"] == who]
        out.append(row("В-" + who[0], f"Лише запуски виконавиці {who} (45)", rr, h1(rr), h2(rr)))
    loo = []
    for k in range(1, 11):
        rr = [r for r in P if r["round"] != k]
        a, b = h1(rr), h2(rr)
        loo.append((k, a[1], b["ok"], {m: (v["wins"], v["n"]) for m, v in a[0].items()}))
    out.append({"code": "Р-1", "title": "Без одного раунду (10 варіантів по 81 запуску)", "n": 81, "totals": None,
                "h1": None, "h1_ok": any(x[1] for x in loo), "h2_apart": None, "h2_ok": any(x[2] for x in loo),
                "note": f"H1 підтверджено в {sum(x[1] for x in loo)} із 10, H2 — у {sum(x[2] for x in loo)} із 10",
                "loo": loo})
    return out


# ---------------------------------------------------------------- незалежна повторна перевірка порушень

ORDER = ["SENT", "VIEWED", "SHORTLISTED", "INTERVIEW_INVITED", "OFFERED", "HIRED"]


def _d(v):
    try:
        return _dt.fromisoformat(str(v).replace("Z", "+00:00"))
    except ValueError:
        return None


def _history_ok(events, final):
    if not events or (events[0].get("fromStatus"), events[0].get("toStatus")) != (None, "SENT"):
        return False
    cur, before_w = "SENT", None
    for e in events[1:]:
        f, t = e.get("fromStatus"), e.get("toStatus")
        if f != cur or cur in ("HIRED", "REJECTED"):
            return False
        if cur == "WITHDRAWN":
            if t != before_w:
                return False
        elif t == "WITHDRAWN":
            before_w = cur
        elif t != "REJECTED" and (t not in ORDER or ORDER.index(t) <= ORDER.index(cur)):
            return False
        cur = t
    return cur == final


def verify_violation(text, rule, table, idx, run_date):
    """Окремий код (без validate.check): чи справді запис порушує правило. Повертає (так/ні, пояснення)."""
    if rule == "1.1":
        try:
            json.loads(text.strip())
            return False, "текст читається як JSON"
        except ValueError as e:
            return True, f"json.loads: {str(e)[:60]}"
    data = json.loads(validate.extract(text)[0])
    rec = data[table][idx]
    if rule == "1.4":
        bad = [k for k, v in rec.items() if isinstance(v, str) and k.endswith(("At", "Date")) and _d(v) is None]
        return bool(bad), "неіснуюча дата: " + ", ".join(f"{k}={rec[k]}" for k in bad)
    if rule == "2.5":
        cd = _d(rec["closingDate"])
        return cd.date() <= run_date, f"closingDate {cd.date()} не пізніше дати запуску {run_date}"
    if rule == "3.1":
        same = [i for i, x in enumerate(data[table]) if i != idx and x.get("id") == rec.get("id")]
        return bool(same), f"id {rec.get('id')} повторюється в записах № {', '.join(str(i + 1) for i in same)}"
    if rule == "1.5":
        fields = SPEC["tables"][table]["fields"]
        bad = [f"{k}={v}" for k, v in rec.items() if fields.get(k, {}).get("values") and v is not None
               and v not in fields[k]["values"]]
        return bool(bad), "значення поза списком: " + ", ".join(bad)
    if rule == "3.2":
        fields = SPEC["tables"][table]["fields"]
        bad = [f"{k}={rec[k]}" for k, f in fields.items() if f.get("ref") and rec.get(k) is not None
               and rec[k] not in {x.get("id") for x in data.get(f["ref"], [])}]
        return bool(bad), "немає запису, на який посилається: " + ", ".join(bad)
    if rule == "3.5":
        hr = {h["id"]: h for h in data["HrProfile"]}[rec["hrProfileId"]]
        return hr["companyId"] != rec["companyId"], f"компанія вакансії {rec['companyId']}, HR — {hr['companyId']}"
    if rule == "3.7":
        same = [a["id"] for a in data["Application"][:idx]
                if (a["vacancyId"], a["studentProfileId"]) == (rec["vacancyId"], rec["studentProfileId"])]
        return bool(same), f"пара ({rec['vacancyId']}, {rec['studentProfileId']}) уже є в {', '.join(same)}"
    if rule == "3.8":
        v = {x["id"]: x for x in data["Vacancy"]}[rec["vacancyId"]]
        c, pub, cl = _d(rec["createdAt"]), _d(v.get("publishedAt") or ""), _d(v["closingDate"])
        return pub is None or c < pub or c.date() > cl.date(), \
            f"відгук {rec['createdAt']}, публікація {v.get('publishedAt')}"
    if rule == "3.9":
        ev = sorted([h for h in data["ApplicationStatusHistory"] if h.get("applicationId") == rec["id"]],
                    key=lambda h: _d(h["createdAt"]))
        chain = " → ".join(f"{h.get('fromStatus')}→{h.get('toStatus')}" for h in ev)
        return not _history_ok(ev, rec["status"]), f"статус {rec['status']}; історія: {chain}"[:150]
    if rule == "3.10":
        hired = [a["id"] for a in data["Application"] if a.get("vacancyId") == rec["id"] and a.get("status") == "HIRED"]
        return len(hired) > 1 or rec.get("status") != "CLOSED" or rec.get("closeReason") != "HIRED", \
            f"вакансія {rec.get('status')}/{rec.get('closeReason')}, найнято: {', '.join(hired)}"
    return False, "правило не перевіряється повторно"


def verify_all(runs):
    notes, ok = {}, 0
    for r in runs:
        text = open(os.path.join(ROOT, r["file"]), encoding="utf-8").read()
        rd = validate.parse_run_date(r["date"])
        for c in r["violations"]:
            yes, why = verify_violation(text, c["rule"], c["table"], c["record"], rd)
            notes[(r["file"], c["rule"], c["record"])] = ("підтверджено: " if yes else "НЕ підтверджено: ") + why
            ok += yes
    return notes, ok


# ---------------------------------------------------------------- знеособлення

EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
PHONE = re.compile(r'("primaryPhone"\s*:\s*)"[^"]*"')


def anonymize(text):
    seen = {}

    def mail(m):
        seen.setdefault(m.group(0), len(seen) + 1)
        return f"user{seen[m.group(0)]}@example.invalid"
    out = EMAIL.sub(mail, text)
    out, phones = PHONE.subn(r'\1"[masked]"', out)
    return out, len(seen), phones


def anonymize_all():
    """answers/** → answers_anonymized/**; валідатор запускається знову, підсумки порівнюються з оригіналом."""
    rows = []
    pilot = load_workbook(os.path.join(ROOT, "logs", "pilot_log.xlsx"), data_only=True)["Журнал"]
    names = [c.value for c in pilot[1]]
    meta = {}
    for r in pilot.iter_rows(min_row=2, values_only=True):
        d = dict(zip(names, r))
        if d.get("Файл відповіді"):
            meta[d["Файл відповіді"]] = (d["Обсяг"], d["Дата запуску"])
    for r in read_log():
        meta[r["Файл відповіді"]] = (r["Обсяг"], r["Дата запуску"])
    for base, _, files in sorted(os.walk(os.path.join(ROOT, "answers"))):
        for name in sorted(files):
            if not name.endswith(".txt"):
                continue
            src = os.path.join(base, name)
            rel = os.path.relpath(src, ROOT).replace(os.sep, "/")
            dst = os.path.join(ROOT, "answers_anonymized", os.path.relpath(src, os.path.join(ROOT, "answers")))
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            anon, n_mail, n_phone = anonymize(open(src, encoding="utf-8").read())
            open(dst, "w", encoding="utf-8").write(anon)
            same = "—"
            if rel in meta:
                level, date = meta[rel]
                _, a = validate.check(SPEC, src, level, date, run_id=rel)
                _, b = validate.check(SPEC, dst, level, date, run_id=rel)
                same = "так" if a["per_rule"] == b["per_rule"] else "ні"
            rows.append([rel, "answers_anonymized/" + rel[len("answers/"):], n_mail, n_phone, same])
    return rows


# ---------------------------------------------------------------- ручна перевірка наосліп (крок 9)

def blind_summary():
    """Частка збігів ручних висновків із валідатором (logs/blind_check.xlsx після blind_check.py compare)."""
    path = os.path.join(ROOT, "logs", "blind_check.xlsx")
    if not os.path.exists(path):
        return None
    ws = load_workbook(path, data_only=True)["Перевірка"]
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0] and r[5]]
    match = [r for r in rows if r[5] == r[7]]
    return {"rows": len(rows), "match": len(match), "diff": [r for r in rows if r[5] != r[7]],
            "codes": len({r[0] for r in rows}), "records": len({(r[0], r[2], r[3]) for r in rows})}


# ---------------------------------------------------------------- книга results/lab7_dataset.xlsx

def dash(x):
    return "—" if x is None else x


DICTIONARY = [
    ("№, Раунд, Повтор", "Номер запуску в журналі, раунд (9 умов у випадковому порядку), повтор умови", "ціле", "—",
     "1–90; 1–10; 1–10", "журнал"),
    ("Модель", "Вебсервіс (рівень чинника «модель»)", "називна", "—", "ChatGPT, Gemini, DeepSeek", "журнал"),
    ("Назва в інтерфейсі", "Назва моделі, яку показав інтерфейс (DeepSeek — яку назвала модель); пробіли на краях прибрано",
     "текст", "—", "GPT-6; 4 Hydrogen; DeepSeek-V4.1-Flash", "журнал"),
    ("Обсяг", "Рівень обсягу запиту", "порядкова", "—", "S (≥62 записи), M (≥123), L (≥246)", "журнал"),
    ("Виконавиця, Дата", "Хто й коли виконав запуск", "називна; дата", "ДД.ММ.РРРР",
     "Вероніка (непарні раунди), Аліна (парні); 14.11–18.11.2026", "журнал"),
    ("Тривалість, хв", "Тривалість генерації відповіді", "дійсне", "хвилини", "> 0", "журнал"),
    ("Натискань «Продовжити»", "Скільки разів натиснуто «Продовжити»", "ціле", "разів", "0–3", "журнал"),
    ("Стан", "Стан відповіді за журналом", "називна", "—", "повна, обірвана", "журнал"),
    ("Частку пораховано", "Чи є в запуску оцінені перевірки (JSON прочитано, записи є)", "так/ні", "—", "так, ні",
     "валідатор"),
    ("Файл відповіді", "Шлях до файлу, який перевірено (для запуску 80 — копія без службового тегу)", "текст", "—",
     "answers/main/<сервіс>/<файл>", "журнал"),
    ("Символів", "Довжина файлу відповіді", "ціле", "символи", "≥ 0", "розрахунок"),
    ("Записів; Записів: <таблиця>", "Кількість записів у відповіді й у кожній із 8 таблиць", "ціле", "записи",
     "≥ 0", "валідатор"),
    ("Відхилення від обсягу", "Сума |задано − фактично| за таблицями; для історії статусів — лише нестача", "ціле",
     "записи", "≥ 0", "валідатор"),
    ("1.1; 1.2", "Відповідь читається як JSON; є всі 8 таблиць-масивів", "називна", "—", "виконано, порушено",
     "валідатор"),
    ("Г1, Г2*, Г3, Г3°, Заг*: порушено / оцінено / %", "Перевірки групи правил: Г1 — 1.3–1.6; Г2* — 2.1–2.4, 2.6; "
     "Г3 — 3.1–3.10; Г3° — 3.4–3.10; Заг* — 1.3–3.10 без 2.5 (основний показник)", "ціле; дійсне",
     "перевірки; %", "≥ 0; 0–100", "валідатор"),
    ("2.5: порушено / оцінено", "Дата завершення активної вакансії пізніша за дату запуску (окремий показник)",
     "ціле", "перевірки", "≥ 0", "валідатор"),
    ("Порушень у запуску", "Абсолютна кількість порушених перевірок усіх правил, крім 2.5", "ціле", "перевірки",
     "≥ 0", "валідатор"),
    ("Придатні записи, %; П1, %", "Записи без жодного порушення; записи з правильним форматом, але з порушенням "
     "груп 2–3", "дійсне", "%", "0–100", "валідатор"),
]


def decisions():
    return [
        ["Р1", "Для запуску 80 (gpt_L_09) у журналі два рядки: файл відповіді та його копія без службового тегу "
         "<escape>; в аналіз узято копію (рядок «Входить до аналізу = так»), тож у вибірці 90 запусків.",
         "Тег потрапив у файл під час копіювання з інтерфейсу; JSON після тегу повний; файл відповіді збережено "
         "без змін.",
         "Перевірка S1 повторює аналіз за файлом із тегом."],
        ["Р2", "Запуски 21 (gpt_L_03 — вісім порожніх масивів) і 90 (ds_M_10 — текст обривається посеред JSON) у "
         "журналі позначено «Входить до аналізу = ні» з причиною. Обидва залишено у вибірці 90 запусків як "
         "результат і подано окремо; частку для них не рахують і нулем не замінюють.",
         "Позначка журналу суперечить таблиці 3.2 ЛР6: обірвана відповідь і відповідь, що не читається як JSON, "
         "не вилучаються, а входять до аналізу як результат; діє правило протоколу.",
         "Описові частки — за 88 запусками; у H1 ці запуски рахуються «не на користь» (знаменник 30); перевірка "
         "S2 — знаменник 29."],
        ["Р3", "Правило 2.5 перевірено відносно дати запуску з журналу й подано окремим показником.",
         "Розділ 4 і таблиця 5.2 ЛР6: у запиті немає поточної дати.",
         "До часток H1 і H2 не входить; перевірка A2 показує частку разом із 2.5."],
        ["Р4", "Пропущені значення не заповнювали; для запусків без частки записано «—», а не 0. Інших пропусків "
         "у журналі немає, крім «Причини виключення» для запусків із позначкою «так».",
         "Пропуск і нульова частка мають різний зміст.",
         "Кількість запусків із часткою (88) наведено в кожній таблиці результатів."],
        ["Р5", "Нетипові запуски не вилучали: gpt_M_06 — 18 порушень правила 3.1; gpt_L_10, ds_L_04, ds_L_08 — "
         "по 8.",
         "Кожне з цих порушень підтверджено повторною перевіркою за текстом відповіді.",
         "Перевірка S3 повторює аналіз без запуску з найбільшою кількістю порушень у кожній моделі."],
        ["Р6", "Знеособлення: копії всіх відповідей у answers_anonymized/ (e-mail → userN@example.invalid, "
         "телефони → [masked]).", "Таблиця 10.1 ЛР6.",
         "Підсумки валідатора для копій збігаються з підсумками для вихідних файлів."],
    ]


RUN_HEAD = (["№", "Раунд", "Повтор", "Модель", "Назва в інтерфейсі", "Обсяг", "Виконавиця", "Дата", "Тривалість, хв",
             "Натискань «Продовжити»", "Стан", "Частку пораховано", "Файл відповіді", "Символів", "Записів"] +
            [f"Записів: {t}" for t in TABLES] + ["Відхилення від обсягу", "1.1", "1.2"] +
            sum(([f"{g}: порушено", f"{g}: оцінено", f"{g}: %"] for g in ("Г1", "Г2*", "Г3", "Г3°", "Заг*")), []) +
            ["2.5: порушено", "2.5: оцінено", "Порушень у запуску", "Придатні записи, %", "П1, %"])


def run_row(r):
    flag = lambda rule: (("порушено" if r["per_rule"].get(rule, {}).get(BAD) else "виконано")
                         if rule in r["per_rule"] else "—")
    return ([r["no"], r["round"], r["rep"], r["model"], r["interface"], r["level"], r["executor"],
             r["date"].strftime("%d.%m.%Y"), r["duration"], r["presses"], r["state"],
             "так" if r["has_share"] else "ні", r["file"], r["chars"], r["records"]] +
            [r["counts"].get(t, "—") for t in TABLES] + [dash(r["voldev"]), flag("1.1"), flag("1.2")] +
            sum(([dash(r[k]["bad"]), dash(r[k]["done"]), dash(r[k]["pct"])] for k in ("g1", "g2", "g3", "g3core", "main")), [])
            + [dash(r["r25"]["bad"]), dash(r["r25"]["done"]), r["n_viol"],
               dash((r["usable"] or {}).get("частка_%")), dash((r["p1"] or {}).get("частка_%"))])


def write_book(path, P, rob, notes, anon, blind):
    wb = Workbook()
    xs.text_sheet(wb, "Опис", [
        "ЛР7. Набір даних 90 основних запусків і результати аналізу",
        "Експеримент 3 сервіси × 3 рівні обсягу × 10 повторів (90 запусків, одиниця аналізу — відповідь моделі на "
        "один запит), основні запуски 14–18.11.2026. Сервіси chatgpt.com («GPT-6»), gemini.google.com («4 Hydrogen»), "
        "chat.deepseek.com («DeepSeek-V4.1-Flash», DeepThink і Search вимкнено); безкоштовні облікові записи; одна "
        "схема бази даних із 8 таблиць.",
        "Перевірки (одне правило, застосоване до одного запису) — елементи вимірювання всередині запуску, а не обсяг "
        "вибірки.",
        "Аркуші: «Словник даних», «Рішення щодо даних», «Запуски» (рядок на запуск), «Правила» (запуск × правило), "
        "«Порушення» (кожна порушена й неоцінена перевірка з повторною перевіркою), «Модель і обсяг» (розкид 10 "
        "повторів), «Правила за умовами», «Рішення H1 H2», «Стійкість», «Перевірка наосліп», «Знеособлення».",
        "Відтворення: python3 tools/lab7_analysis.py з кореня репозиторію.",
    ], width=130, first=True)
    xs.sheet(wb, "Словник даних", ["Поле", "Зміст", "Тип", "Одиниця", "Допустимі значення", "Джерело"],
             [list(x) for x in DICTIONARY], [26, 60, 14, 12, 36, 12])
    xs.sheet(wb, "Рішення щодо даних", ["№", "Що зроблено", "Чому", "Вплив"], decisions(), [6, 60, 55, 50])
    xs.sheet(wb, "Запуски", RUN_HEAD, [run_row(r) for r in P],
             [5, 7, 7, 10, 18, 7, 11, 11, 9, 10, 9, 9, 40, 9, 8] + [9] * 8 + [9, 9, 9] + [9] * 15 + [9] * 5)
    rr = []
    for r in P:
        for rule in RULES:
            d = r["per_rule"].get(rule)
            if d:
                rr.append([r["no"], os.path.basename(r["file"]), r["model"], r["level"], rule, d[OK], d[BAD], d[NA]])
    xs.sheet(wb, "Правила", ["№", "Файл", "Модель", "Обсяг", "Правило", "Виконано", "Порушено", "Не оцінюється"], rr,
             [5, 22, 10, 7, 8, 10, 10, 12])
    vr = []
    for r in P:
        for c in r["violations"] + r["not_evaluated"]:
            vr.append([r["no"], os.path.basename(r["file"]), r["model"], r["level"], c["stage"], c["rule"], c["table"],
                       (c["record"] + 1) if isinstance(c["record"], int) else "", c["record_id"], c["result"],
                       c["detail"], notes.get((r["file"], c["rule"], c["record"]),
                                              "не оцінюється: потрібне поле вже порушує інше правило")])
    xs.sheet(wb, "Порушення", ["№", "Файл", "Модель", "Обсяг", "Етап", "Правило", "Таблиця", "№ запису", "id запису",
                               "Результат", "Деталі", "Повторна перевірка за текстом відповіді"], vr,
             [5, 22, 10, 7, 6, 8, 24, 8, 10, 12, 46, 52])
    mr = []
    for m in MODELS:
        for lv in LEVELS:
            c = cell(P, m, lv)
            mr.append([m, lv, c["runs"], c["complete"], c["truncated"], c["n"], c["bad"], c["done"],
                       round(100 * c["bad"] / c["done"], 3) if c["done"] else "—", c["any"],
                       dash(c["min"]), dash(c["q1"]), dash(c["median"]), dash(c["q3"]), dash(c["max"])] +
                      [dash(v) for v in c["values"]] + [", ".join("—" if x is None else str(x) for x in c["n_viol"]),
                                                       f"{c['r25_bad']} / {c['r25_done']}", c["r25_runs"],
                                                       ", ".join(f"{p:g}" for p in c["presses"]), c["duration"]])
    xs.sheet(wb, "Модель і обсяг",
             ["Модель", "Обсяг", "Запусків", "Повних", "Обірваних", "Із часткою", "Порушено (сума)", "Оцінено (сума)",
              "Частка від суми, %", "Запусків із ≥1 порушенням", "Мінімум, %", "Q1, %", "Медіана, %", "Q3, %",
              "Максимум, %"] + [f"Повтор {k}, %" for k in range(1, 11)] +
             ["Порушень у повторах 1–10", "2.5: порушено / оцінено", "Запусків із порушенням 2.5",
              "«Продовжити» в повторах 1–10", "Медіана тривалості, хв"],
             mr, [10, 7, 9, 8, 9, 9, 10, 11, 10, 11, 9, 8, 9, 8, 9] + [8] * 10 + [24, 14, 12, 22, 12])
    rt = rule_table(P)
    xs.sheet(wb, "Правила за умовами", ["Правило"] + [f"{m} {lv}" for m in MODELS for lv in LEVELS],
             [[rule] + [f"{rt[(rule, m, lv)]['bad']} / {rt[(rule, m, lv)]['done']}" +
                        (f"; запусків {rt[(rule, m, lv)]['runs']}" if rt[(rule, m, lv)]["runs"] else "")
                        for m in MODELS for lv in LEVELS] for rule in RULES], [9] + [20] * 9)
    a, a_ok = h1(P)
    b = h2(P)
    hr = [[m, v["wins"], v["n"], v["zeros"], v["no_share"], "так" if v["ok"] else "ні"] for m, v in a.items()]
    hr += [["H1 підтверджено", "так" if a_ok else "ні", "", "", "", ""], ["", "", "", "", "", ""]]
    hr += [[x["level"], x["pair"], f"{x['x'][0]}–{x['x'][1]}", f"{x['y'][0]}–{x['y'][1]}",
            "ні" if x["better"] else "так", x["better"] or "—"] for x in b["rows"]]
    hr += [["H2 підтверджено", "так" if b["ok"] else "ні", "", "", "", ""]]
    xs.sheet(wb, "Рішення H1 H2", ["H1: модель / H2: обсяг", "Г3 > Г2* / пара", "Запусків / Q1–Q3 першої",
                                   "Обидві частки 0 / Q1–Q3 другої", "Без частки / перетин", "Рішення / менша частка"],
             hr, [22, 24, 20, 22, 18, 18])
    sr = []
    for x in rob:
        t = x["totals"]
        sr.append([x["code"], x["title"], x["n"]] +
                  (sum(([f"{t[m]['bad']} / {t[m]['done']}", t[m]["any"]] for m in MODELS), []) if t else [""] * 6) +
                  ([" · ".join(f"{x['h1'][m][0]}/{x['h1'][m][1]}" for m in MODELS)] if x["h1"] else [x["note"]]) +
                  ["так" if x["h1_ok"] else "ні", dash(x["h2_apart"]), "так" if x["h2_ok"] else "ні"])
    xs.sheet(wb, "Стійкість", ["Код", "Що змінено", "Запусків"] +
             sum(([f"{m}: порушено / оцінено", f"{m}: запусків із порушенням"] for m in MODELS), []) +
             ["H1: Г3 > Г2* (ChatGPT · Gemini · DeepSeek)", "H1 підтверджено", "Пар без перетину Q1–Q3",
              "H2 підтверджено"], sr, [6, 50, 9] + [13, 11] * 3 + [24, 11, 12, 11])
    if blind:
        br = [["Відповідей (по одній на умову)", blind["codes"]], ["Записів перевірено", blind["records"]],
              ["Пар «запис × правило»", blind["rows"]], ["Збігів із валідатором", blind["match"]],
              ["Частка збігів, %", round(100 * blind["match"] / blind["rows"], 1)],
              ["Розбіжностей", len(blind["diff"])]]
        xs.sheet(wb, "Перевірка наосліп", ["Показник", "Значення"], br, [36, 14])
    xs.sheet(wb, "Знеособлення", ["Оригінал", "Знеособлена копія", "Замінено e-mail (унікальних)",
                                  "Замінено телефонів", "Підсумки валідатора збігаються"], anon, [44, 50, 14, 12, 14])
    wb.save(path)


def summary_text(P, rob, blind, n_verify, n_viol):
    a, a_ok = h1(P)
    b = h2(P)
    lines = ["ЛР7: основний аналіз 90 запусків (tools/lab7_analysis.py)", ""]
    lines.append(f"Стан: повних {sum(r['state'] == 'повна' for r in P)}, обірваних {sum(r['state'] == 'обірвана' for r in P)}; "
                 f"частку пораховано для {sum(r['has_share'] for r in P)}")
    for m in MODELS:
        t = totals(P)[m]
        lines.append(f"{m}: порушено {t['bad']} з {t['done']} перевірок (без 2.5); запусків із порушенням {t['any']}; "
                     f"2.5: {t['r25_bad']} з {t['r25_done']}")
    lines.append("H1: " + "; ".join(f"{m} {v['wins']}/{v['n']}" for m, v in a.items()) +
                 (" — підтверджено" if a_ok else " — не підтверджено"))
    lines.append(f"H2: пар без перетину Q1–Q3 {b['apart']} — " + ("підтверджено" if b["ok"] else "не підтверджено"))
    lines.append(f"Повторна перевірка знайдених порушень: підтверджено {n_verify} з {n_viol}")
    if blind:
        lines.append(f"Перевірка наосліп: збіг {blind['match']} з {blind['rows']} пар у {blind['records']} записах")
    lines += ["", "Стійкість:"] + [f"  {x['code']}: H1 {'так' if x['h1_ok'] else 'ні'}, H2 {'так' if x['h2_ok'] else 'ні'} — "
                                  f"{x['title']}" + (f" ({x['note']})" if x["note"] else "") for x in rob]
    return "\n".join(lines) + "\n"


def main():
    rows = read_log()
    P = [evaluate(r) for r in runs_from_log(rows)]
    S1 = [evaluate(r) for r in runs_from_log(rows, prefer_original=True)]
    rob = robustness(P, S1)
    notes, ok = verify_all(P)
    anon = anonymize_all()
    blind = blind_summary()
    os.makedirs(os.path.join(ROOT, "results"), exist_ok=True)
    write_book(os.path.join(ROOT, "results", "lab7_dataset.xlsx"), P, rob, notes, anon, blind)
    text = summary_text(P, rob, blind, ok, sum(len(r["violations"]) for r in P))
    open(os.path.join(ROOT, "results", "lab7.txt"), "w", encoding="utf-8").write(text)
    print(text)
    return P, S1, rob, notes, anon, blind


if __name__ == "__main__":
    main()

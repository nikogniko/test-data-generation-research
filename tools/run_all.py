"""Перевіряє всі відповіді, записані в журналі запусків, і записує результати в книгу Excel і текстовий файл.

Запуск із кореневої папки пакета:
  python3 tools/run_all.py pilot   — журнал пілоту logs/pilot_log.xlsx       → results/pilot.xlsx, results/pilot.txt
  python3 tools/run_all.py main    — logs/run_log.xlsx, «Основні запуски»   → results/main.xlsx, results/main.txt

Перевіряються рядки, у яких «Відповідь збережено» = «так» і немає позначки «Входить до аналізу» = «ні».
Склад книги й текстового файлу описано в tools/analyze.py.
Потрібні Python 3.10+ і пакет openpyxl.
"""
import json
import os
import sys

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import analyze  # noqa: E402
import validate  # noqa: E402

ROOT = validate.ROOT
LOGS = {"pilot": ("pilot_log.xlsx", "Журнал"), "main": ("run_log.xlsx", "Основні запуски")}


def read_log(phase):
    name, sheet = LOGS[phase]
    ws = load_workbook(os.path.join(ROOT, "logs", name), data_only=True)[sheet]
    names = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        r = dict(zip(names, row))
        if r.get("Модель") and r.get("Файл відповіді"):
            yield r


def collect(phase):
    spec = json.load(open(validate.SPEC, encoding="utf-8"))
    pilot = phase == "pilot"
    runs, problems = [], []
    for r in read_log(phase):
        if r.get("Відповідь збережено") != "так":
            continue
        if r.get("Входить до аналізу") == "ні":
            continue
        item = {"no": r.get("№"), "model": r["Модель"], "level": r["Обсяг"], "file": r["Файл відповіді"],
                "date": r.get("Дата запуску"), "round": r.get("Раунд"), "rep": r.get("Повтор"),
                "duration": r.get("Тривалість відповіді, хв"),
                "interface": (r.get("Назва моделі в інтерфейсі") or "").strip() or None,
                "presses": r.get("Натискань «Продовжити»"), "state": r.get("Стан відповіді"),
                "assessment": None, "run": None, "summary": None}
        if pilot:
            presses = item["presses"] or 0
            item["assessment"] = ("у межах" if presses == 0 else "на межі") if item["state"] == "повна" else \
                ("за межами" if item["state"] in ("обірвана", "відмова") else None)
        raw = os.path.join(ROOT, r["Файл відповіді"])
        if not os.path.exists(raw):
            problems.append(f"немає файлу {r['Файл відповіді']}")
        elif not r.get("Дата запуску"):
            problems.append(f"не вказано дату запуску для {r['Файл відповіді']}")
        else:
            item["run"], item["summary"] = validate.check(spec, raw, r["Обсяг"], r["Дата запуску"],
                                                          run_id=r["Файл відповіді"])
        runs.append(item)
    return runs, problems


def main(phase):
    runs, problems = collect(phase)
    os.makedirs(os.path.join(ROOT, "results"), exist_ok=True)
    out = os.path.join(ROOT, "results", f"{phase}.xlsx")
    result = None
    if runs:
        result = analyze.write(out, runs, pilot=phase == "pilot")
        analyze.write_txt(out[:-len(".xlsx")] + ".txt", runs, pilot=phase == "pilot")
    return out, runs, problems, result


if __name__ == "__main__":
    phase = sys.argv[1] if len(sys.argv) > 1 else "main"
    if phase not in LOGS:
        sys.exit("Параметр: pilot або main")
    out, runs, problems, _ = main(phase)
    rel = os.path.relpath(out, ROOT)
    print(f"Перевірено відповідей: {sum(1 for r in runs if r['summary'])}. Результати: {rel}, {rel[:-5]}.txt")
    for p in problems:
        print("Увага:", p)

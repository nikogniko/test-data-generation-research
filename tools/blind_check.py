"""Ручна перевірка наосліп: підготовка відповідей під кодами й порівняння ручних висновків із валідатором.

  python3 tools/blind_check.py prepare
      Читає logs/blind_key.xlsx (код → файл відповіді) і створює папку blind/ з файлами K-01.txt … K-09.txt.
      JSON у всіх файлах переформатовано однаково (відступ 2, без тексту поза JSON), тому модель не
      впізнається за виглядом відповіді.
  python3 tools/blind_check.py compare
      Для кожного рядка logs/blind_check.xlsx знаходить висновок валідатора щодо того самого запису й правила
      і записує його в колонку «Висновок валідатора»; колонки «Збіг» і частка збігів рахуються формулами.

Висновок валідатора для пари «запис × правило»: «порушено», якщо хоча б одна перевірка порушена;
«виконано», якщо є виконані перевірки й немає порушених; інакше «не оцінюється».
"""
import json
import os
import sys

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import validate  # noqa: E402

ROOT = validate.ROOT
KEY = os.path.join(ROOT, "logs", "blind_key.xlsx")
FORM = os.path.join(ROOT, "logs", "blind_check.xlsx")
BLIND = os.path.join(ROOT, "blind")


def key_rows():
    ws = load_workbook(KEY)["Ключ"]
    for code, model, level, path in ws.iter_rows(min_row=2, values_only=True):
        if code and path:
            yield code, model, level, path


def prepare():
    os.makedirs(BLIND, exist_ok=True)
    for code, _, _, path in key_rows():
        body, _, _ = validate.extract(open(os.path.join(ROOT, path), encoding="utf-8").read())
        try:
            text = json.dumps(json.loads(body), ensure_ascii=False, indent=2)
        except ValueError:
            text = body  # відповідь, яку не вдалося прочитати, передається як є
        with open(os.path.join(BLIND, f"{code}.txt"), "w", encoding="utf-8") as fh:
            fh.write(text)
        print(code, "готово")


def record_key(table, record, data):
    """Ідентифікатор запису у формі: id, а для VacancySkill — vacancyId/skillId."""
    if table == "VacancySkill" and isinstance(record, int):
        r = data["VacancySkill"][record]
        return f"{r.get('vacancyId')}/{r.get('skillId')}" if isinstance(r, dict) else ""
    return ""


def verdicts(spec, path, level, run_date):
    run, _ = validate.check(spec, os.path.join(ROOT, path), level, run_date)
    body, _, _ = validate.extract(open(os.path.join(ROOT, path), encoding="utf-8").read())
    try:
        data = json.loads(body)
    except ValueError:
        data = {}
    out = {}
    for c in run.rows:
        rid = c["record_id"] or record_key(c["table"], c["record"], data)
        out.setdefault((c["table"], str(rid), c["rule"]), []).append(c["result"])
    return out


def merge(results):
    if validate.BAD in results:
        return validate.BAD
    return validate.OK if validate.OK in results else validate.NA


def run_dates():
    """Дата запуску для кожного файлу відповіді з журналу основних запусків."""
    ws = load_workbook(os.path.join(ROOT, "logs", "run_log.xlsx"), data_only=True)["Основні запуски"]
    names = [c.value for c in ws[1]]
    dates = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        r = dict(zip(names, row))
        if r.get("Файл відповіді") and r.get("Дата запуску"):
            dates[r["Файл відповіді"]] = r["Дата запуску"]
    return dates


def compare():
    spec = json.load(open(validate.SPEC, encoding="utf-8"))
    dates = run_dates()
    found = {code: verdicts(spec, path, level, dates[path]) for code, _, level, path in key_rows()}
    wb = load_workbook(FORM)
    ws = wb["Перевірка"]
    filled = 0
    for r in range(2, ws.max_row + 1):
        code, _, table, rid, rule = (ws.cell(row=r, column=c).value for c in range(1, 6))
        if not code or not table or rule is None:
            continue
        results = found.get(code, {}).get((table, str(rid), str(rule)))
        ws.cell(row=r, column=8, value=merge(results) if results else "немає перевірки")
        filled += 1
    wb.save(FORM)
    print(f"Заповнено рядків: {filled}")


if __name__ == "__main__":
    action = sys.argv[1] if len(sys.argv) > 1 else ""
    if action == "prepare":
        prepare()
    elif action == "compare":
        compare()
    else:
        sys.exit("Параметр: prepare або compare")

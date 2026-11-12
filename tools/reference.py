"""Генератор контрольних наборів: коректні дані заданого рівня обсягу, що виконують усі 22 правила.

Контрольний набір не є відповіддю моделі. Він потрібен для перевірки валідатора: на ньому валідатор
не повинен знаходити порушень, а після внесення одного порушення — має знаходити саме його.
Дати в наборі розраховано на перевірку з датою запуску не пізніше 30.12.2026.

Окремий запуск: python3 tools/reference.py M reference_M.txt
Зазвичай генератор викликає tools/check_tools.py.
"""
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SPEC = os.path.join(ROOT, "spec.json")

FIRST = ["Олена", "Андрій", "Марія", "Тарас", "Ірина", "Богдан", "Софія", "Дмитро", "Юлія", "Максим",
         "Наталія", "Остап", "Катерина", "Роман", "Дарина", "Ігор", "Леся", "Василь"]
LAST = ["Коваль", "Мельник", "Шевчук", "Бойко", "Ткаченко", "Кравець", "Олійник", "Лисенко", "Руденко",
        "Савченко", "Петренко", "Марченко", "Гнатюк", "Поліщук", "Іваненко", "Дорошенко", "Левченко", "Зінченко"]
MIDDLE = ["VIEWED", "SHORTLISTED", "INTERVIEW_INVITED", "OFFERED"]
SKILLS = [1, 2, 3, 7, 9, 13, 14, 17, 25, 26, 27, 35, 37, 40, 50, 51, 53, 65]
PROFESSIONS = [1, 2, 3, 4, 5, 6, 7, 9, 10, 13]


def build(spec, level="M"):
    """Набір даних рівня level; склад береться з тих самих параметрів, з яких зібрано розділ «Склад» запиту."""
    p = spec["levels"][level]["params"]
    n_users = p["hr"] + p["students"]
    users = []
    for k in range(n_users):
        users.append({"id": f"u{k + 1}", "clerkUserId": f"user_test_{k + 1:03d}", "email": f"user{k + 1}@example.com",
                      "role": "HR" if k < p["hr"] else "STUDENT", "status": "ACTIVE",
                      "firstName": FIRST[k % len(FIRST)], "lastName": LAST[(k * 7) % len(LAST)],
                      "middleName": None, "createdAt": f"2026-09-{k % 28 + 1:02d}T09:00:00Z"})
    companies = []
    for k in range(p["approved"] + p["pending"]):
        approved = k < p["approved"]
        companies.append({"id": f"c{k + 1}", "registrationType": "COMPANY" if approved else "FOP",
                          "registrationNumber": f"{40000001 + k}", "legalName": f"ТОВ «Компанія {k + 1}»",
                          "publicName": f"Компанія {k + 1}", "foundationYear": 2010 + k % 15,
                          "employeeCount": "SIZE_21_50" if approved else None, "about": "Розробка програмного забезпечення.",
                          "publicEmail": f"hr{k + 1}@company.example.com",
                          "verificationStatus": "APPROVED" if approved else "PENDING",
                          "createdAt": f"2026-08-{k % 28 + 1:02d}T10:00:00Z"})
    hrs = [{"id": f"h{k + 1}", "userId": f"u{k + 1}", "companyId": f"c{k + 1}", "position": "Рекрутер"}
           for k in range(p["hr"])]
    students = [{"id": f"s{k + 1}", "userId": f"u{p['hr'] + k + 1}", "gender": None,
                 "birthDate": f"2003-{k % 12 + 1:02d}-15T00:00:00Z", "contactEmail": f"student{k + 1}@example.com",
                 "primaryPhone": f"+38050{k + 1:07d}", "desiredPosition": "Junior-розробник", "minSalary": None,
                 "visibility": "APPLIED_ONLY"} for k in range(p["students"])]

    vacancies = []

    def vac(status, hr_index, closing, closed=None, reason=None):
        n = len(vacancies) + 1
        published = None if status == "DRAFT" else "2026-10-02T08:00:00Z"
        salary = [(30000, 45000, "PER_MONTH"), (35000, None, "PER_MONTH"), (250, 400, "PER_HOUR"), (None, None, None)][n % 4]
        vacancies.append({"id": f"v{n}", "hrProfileId": f"h{hr_index + 1}", "companyId": f"c{hr_index + 1}",
                          "title": f"Вакансія {n}", "description": "Опис вакансії.",
                          "professionId": PROFESSIONS[n % len(PROFESSIONS)], "minSalary": salary[0],
                          "maxSalary": salary[1], "salaryPeriod": salary[2], "status": status,
                          "createdAt": "2026-10-01T08:00:00Z", "publishedAt": published, "closingDate": closing,
                          "closedAt": closed, "closeReason": reason})
        return f"v{n}"

    approved_hr = list(range(p["approved"]))
    closed_v = [vac("CLOSED", approved_hr[k % len(approved_hr)], "2026-11-30T00:00:00Z", "2026-11-10T12:00:00Z",
                    "HIRED") for k in range(p["closed"])]
    active_v = [vac("ACTIVE", approved_hr[(k + p["closed"]) % len(approved_hr)], "2026-12-31T00:00:00Z")
                for k in range(p["active"])]
    for k in range(p["draft"]):
        vac("DRAFT", p["approved"] + k % p["pending"], "2027-01-31T00:00:00Z")
    vskills = [{"vacancyId": v["id"], "skillId": SKILLS[(3 * j + k) % len(SKILLS)],
                "weight": ["CRITICAL", "IMPORTANT", "NICE_TO_HAVE"][k]}
               for j, v in enumerate(vacancies) for k in range(3)]

    # шляхи статусів: найм і відмова на кожну закриту вакансію, далі відмови, відкликання, проміжні статуси
    plan = []
    for v in closed_v:
        plan.append((v, ["SENT", "VIEWED", "SHORTLISTED", "INTERVIEW_INVITED", "OFFERED", "HIRED"]))
        plan.append((v, ["SENT", "VIEWED", "REJECTED"]))
    open_v = active_v + closed_v
    rejected_left = p["rejected"] - len(closed_v)
    for k in range(max(0, rejected_left)):
        plan.append((open_v[k % len(open_v)], ["SENT", "REJECTED"]))
    for k in range(p["withdrawn"]):
        plan.append((open_v[(k + 1) % len(open_v)], ["SENT", "VIEWED", "WITHDRAWN"]))
    k = 0
    while len(plan) < p["apps"]:
        steps = k % len(MIDDLE) + 1
        path = ["SENT"] + (MIDDLE[:steps] if k % 3 else MIDDLE[steps - 1:steps])
        plan.append((open_v[k % len(open_v)], path))
        k += 1
    plan = plan[:p["apps"]]

    hr_user = {v["id"]: "u" + v["hrProfileId"][1:] for v in vacancies}
    used, apps, history = set(), [], []
    for n, (v, path) in enumerate(plan, start=1):
        s = n % p["students"]
        tries = 0
        while (v, s) in used:
            s, tries = (s + 1) % p["students"], tries + 1
            if tries > p["students"]:
                raise ValueError(f"на вакансію {v} більше відгуків, ніж студентів")
        used.add((v, s))
        day = 3 + n % 6
        apps.append({"id": f"a{n}", "vacancyId": v, "studentProfileId": f"s{s + 1}", "status": path[-1],
                     "coverLetter": None, "createdAt": f"2026-10-{day:02d}T10:00:00Z"})
        prev = None
        for step, st in enumerate(path):
            student_action = step == 0 or st == "WITHDRAWN" or prev == "WITHDRAWN"
            history.append({"id": f"e{len(history) + 1}", "applicationId": f"a{n}", "fromStatus": prev,
                            "toStatus": st,
                            "changedByUserId": f"u{p['hr'] + s + 1}" if student_action else hr_user[v],
                            "createdAt": f"2026-10-{day:02d}T{10 + step:02d}:00:00Z"})
            prev = st
    return {"User": users, "Company": companies, "HrProfile": hrs, "StudentProfile": students,
            "Vacancy": vacancies, "VacancySkill": vskills, "Application": apps, "ApplicationStatusHistory": history}


def wrap(data):
    return "```json\n" + json.dumps(data, ensure_ascii=False, indent=1) + "\n```\n"


if __name__ == "__main__":
    spec = json.load(open(SPEC, encoding="utf-8"))
    data = build(spec, sys.argv[1])
    with open(sys.argv[2], "w", encoding="utf-8") as fh:
        fh.write(wrap(data))
    print({t: len(v) for t, v in data.items()})

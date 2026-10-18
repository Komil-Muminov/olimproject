"""Окно «Листы и ИНН» без экрана: строка на каждый лист каждого файла (брать ли, колонка ИНН, название,
«без ИНН — как есть») → настройки чтения для scan_files."""
from analysis_list_map import RAW
from column_entry import to_mapping
from column_spec import SpecError

NEED_INN = "впишите колонку ИНН или отметьте «без ИНН»"


def choices(props: list) -> list:
    """Что нашла программа: на каждом листе колонка ИНН и название (found — нашла сама)."""
    return [{"path": p["path"], "sheet": name, "use": True, "inn": sh.get("inn") or "", "name": sh.get("name") or "",
             "raw": False, "found": bool(sh.get("inn")), "repeats": bool(sh.get("repeats"))}
            for p in props for name, sh in p["sheets"].items()]


def to_mappings(props: list, rows: list) -> tuple:
    """→ ({путь: [настройки листов]}, [(номер строки, что не так)]). Файлы без выбранных листов не попадают."""
    sheets = {p["path"]: p["sheets"] for p in props}
    out, problems = {}, []
    for i, r in enumerate(rows):
        if not r["use"]:
            continue
        if r["raw"]:
            m = {"sheet": r["sheet"], "kind": RAW}
        elif not r["inn"].strip():
            problems.append((i, NEED_INN))
            continue
        else:
            try:
                m = to_mapping(r["sheet"], sheets[r["path"]][r["sheet"]],
                               {"inn": r["inn"].strip(), "name": r["name"].strip()}, {"repeats": r["repeats"]})
            except SpecError as e:
                problems.append((i, str(e)))
                continue
        out.setdefault(r["path"], []).append(m)
    return out, problems

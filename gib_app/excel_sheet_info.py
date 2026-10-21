"""Что не видно в значениях ячеек: скрытые листы, формулы без сохранённого результата. Читается прямо
из .xlsx / .xlsm (это zip с xml); для старого .xls таких сведений нет — пусто."""
import re
import zipfile

from file_format import real_path

SHEET_RE = re.compile(r'<sheet\b[^>]*>')
CELL_RE = re.compile(r'<c\b[^>]*?\br="([A-Z]+)\d+"[^>]*?(?:/>|>(.*?)</c>)', re.S)
VALUE_RE = re.compile(r"<v>[^<]")
FORMULA_SCAN_LIMIT = 20_000_000     # большие листы (реестры) — без поиска формул (долго)


def _attr(tag: str, name: str) -> str | None:
    m = re.search(rf'\b{name}="([^"]*)"', tag)
    return m.group(1) if m else None


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def workbook_sheets(path: str) -> dict:
    """{лист: {"hidden": скрыт ли, "xml": путь к листу внутри файла}}; не .xlsx — {}."""
    path = real_path(path)
    if not zipfile.is_zipfile(path):
        return {}
    with zipfile.ZipFile(path) as z:
        if "xl/workbook.xml" not in z.namelist():         # не книга Excel нового формата (например, .xls)
            return {}
        book = z.read("xl/workbook.xml").decode("utf-8", "ignore")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
    targets = {_attr(t, "Id"): _attr(t, "Target") for t in re.findall(r"<Relationship\b[^>]*>", rels)}
    out = {}
    for tag in SHEET_RE.findall(book):
        target = (targets.get(_attr(tag, "r:id")) or "").lstrip("/")
        name = _attr(tag, "name").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
        out[name] = {"hidden": _attr(tag, "state") in ("hidden", "veryHidden"),
                     "xml": target if target.startswith("xl/") else "xl/" + target}
    return out


def empty_formulas(path: str, sheet: str) -> dict:
    """{столбец (с 0): сколько формул без сохранённого значения}; большой лист или не .xlsx — {}."""
    info = workbook_sheets(path).get(sheet)
    if not info:
        return {}
    with zipfile.ZipFile(real_path(path)) as z:
        if z.getinfo(info["xml"]).file_size > FORMULA_SCAN_LIMIT:
            return {}
        xml = z.read(info["xml"]).decode("utf-8", "ignore")
    empty = {}
    for col, body in CELL_RE.findall(xml):
        if body and "<f" in body and not VALUE_RE.search(body):      # формула, а результата нет
            empty[_col_index(col)] = empty.get(_col_index(col), 0) + 1
    return empty

"""Настоящий формат Excel-файла — по содержимому, а не по расширению: «Вписка.xls», внутри которой
современный .xlsx (Excel такой открывает, а читалки определяют формат по расширению и падают)."""
import os
import shutil
import tempfile

ZIP, OLE = b"PK\x03\x04", b"\xd0\xcf\x11\xe0"      # .xlsx / .xlsm — zip; старый .xls — OLE


def real_path(path: str) -> str:
    """Путь, по которому файл прочитается: сам файл или его копия во временной папке с верным расширением."""
    with open(path, "rb") as f:
        sig = f.read(4)
    ext = os.path.splitext(path)[1].lower()
    want = ".xlsx" if sig == ZIP and ext == ".xls" else ".xls" if sig == OLE and ext in (".xlsx", ".xlsm") else None
    if not want:
        return path
    out = os.path.join(tempfile.gettempdir(), f"gib_read_{os.getpid()}_{abs(hash(os.path.abspath(path)))}{want}")
    if not os.path.exists(out) or os.path.getmtime(out) < os.path.getmtime(path):
        shutil.copyfile(path, out)
    return out

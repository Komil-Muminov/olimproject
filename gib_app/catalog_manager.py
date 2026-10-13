"""Менеджер справочника кодов ГИБ (загрузка, сохранение, CRUD операции)."""
import json
import os
from catalog_data import DEFAULT_GIB_CATALOG


def _default_storage_path() -> str:
    appdata = os.environ.get("APPDATA")
    if appdata:
        folder = os.path.join(appdata, "Olim")
        os.makedirs(folder, exist_ok=True)
        return os.path.join(folder, "gib_catalog.json")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "gib_catalog.json")


class GibCatalogManager:
    """Управление статьями классификации ГИБ с сохранением изменений пользователя."""

    def __init__(self, filepath: str | None = None):
        self.filepath = filepath or _default_storage_path()
        self.items: dict[str, str] = {}
        self.load()

    def load(self):
        """Загружает сохраненные пользователем статьи или инициализирует по умолчанию."""
        if os.path.exists(self.filepath):
            try:
                with open(self.filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        self.items = {str(k).strip(): str(v).strip() for k, v in data.items()}
                        return
            except Exception:
                pass
        self.items = dict(DEFAULT_GIB_CATALOG)
        self.save()

    def save(self):
        """Сохраняет текущий справочник в файл JSON."""
        try:
            folder = os.path.dirname(self.filepath)
            if folder:
                os.makedirs(folder, exist_ok=True)
            with open(self.filepath, "w", encoding="utf-8") as f:
                json.dump(self.items, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def get_all(self) -> dict[str, str]:
        """Возвращает все статьи, отсортированные по коду."""
        return dict(sorted(self.items.items(), key=lambda kv: (len(kv[0]), kv[0])))

    def get(self, code: str) -> str | None:
        """Получить наименование по коду ГИБ."""
        return self.items.get(str(code).strip())

    def set(self, code: str, name: str):
        """Добавить или обновить статью расходов."""
        clean_code = str(code).strip()
        clean_name = str(name).strip()
        if not clean_code or not clean_name:
            raise ValueError("Код и наименование не могут быть пустыми.")
        self.items[clean_code] = clean_name
        self.save()

    def delete(self, code: str) -> bool:
        """Удалить статью расходов по коду."""
        clean_code = str(code).strip()
        if clean_code in self.items:
            del self.items[clean_code]
            self.save()
            return True
        return False

    def reset_to_default(self):
        """Сбросить справочник к стандартным значениям."""
        self.items = dict(DEFAULT_GIB_CATALOG)
        self.save()

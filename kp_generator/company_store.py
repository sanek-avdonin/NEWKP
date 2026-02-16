from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List

from kp_generator.models import CompanyProfile  # абсолютный импорт


class CompanyStore:
    def __init__(self, json_path: Path):
        self.json_path = json_path
        self._companies: Dict[str, CompanyProfile] = {}

    def load(self) -> List[CompanyProfile]:
        if not self.json_path.exists():
            raise FileNotFoundError(f"Не найден файл профилей компаний: {self.json_path}")

        try:
            data = json.loads(self.json_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as e:
            # дружелюбное сообщение
            raise ValueError(f"Некорректный JSON в {self.json_path}: {e}") from e

        if not isinstance(data, dict):
            raise ValueError(f"Файл {self.json_path} должен содержать JSON-объект верхнего уровня.")

        companies = data.get("companies", [])
        if not isinstance(companies, list):
            raise ValueError(f"Поле 'companies' в {self.json_path} должно быть списком.")

        self._companies.clear()

        required_keys = ["id", "name", "inn", "address", "phone", "ceo"]

        for idx, row in enumerate(companies):
            if not isinstance(row, dict):
                raise ValueError(f"Элемент companies[{idx}] должен быть объектом (dict).")

            missing = [key for key in required_keys if key not in row]
            if missing:
                # требование: понятное сообщение с указанием компании/индекса
                who = row.get("name", f"№{idx + 1}")
                raise ValueError(f"В компании {who} отсутствуют ключи: {missing}")

            c = CompanyProfile(
                id=str(row["id"]),
                name=str(row["name"]),
                inn=str(row["inn"]),
                address=str(row["address"]),
                phone=str(row["phone"]),
                ceo=str(row["ceo"]),
                logo_path=row.get("logo_path"),
            )
            self._companies[c.id] = c

        return list(self._companies.values())

    def get(self, company_id: str) -> CompanyProfile:
        if company_id not in self._companies:
            raise KeyError(f"Компания с id={company_id} не найдена")
        return self._companies[company_id]

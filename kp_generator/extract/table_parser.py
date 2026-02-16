from __future__ import annotations

from decimal import Decimal
from typing import List
import re

from kp_generator.models import Item  # абсолютный импорт


def _to_decimal_ru(s: str) -> Decimal:
    s = s.replace(" ", "").replace("\u00a0", "").replace(",", ".")
    return Decimal(s)


def parse_items_from_text(raw_text: str) -> List[Item]:
    """
    Эвристика:
    - ищем строки, где есть цена/сумма и количество
    - разделение по 2+ пробелам / табам
    """
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
    items: List[Item] = []

    money_re = re.compile(r"(\d[\d\s\u00a0]*[.,]\d{2})")  # qty_re удалён (не использовался)

    for ln in lines:
        monies = money_re.findall(ln)
        if len(monies) < 1:
            continue

        parts = re.split(r"\s{2,}|\t+", ln)
        if len(parts) < 4:
            continue

        try:
            amount = _to_decimal_ru(monies[-1])
            price = _to_decimal_ru(monies[-2]) if len(monies) >= 2 else amount
        except Exception:
            continue

        qty = None
        unit = ""
        name = ""

        if len(parts) >= 4:
            name = parts[0].strip()
            mid = parts[1:-2]
            if len(mid) >= 1:
                try:
                    qty = _to_decimal_ru(mid[0])
                except Exception:
                    qty = None
            if len(mid) >= 2:
                unit = str(mid[1]).strip()

        if not name or qty is None or qty <= 0:
            continue

        calc_amount = (qty * price).quantize(Decimal("0.01"))
        items.append(Item(
            name=name,
            qty=qty,
            unit=unit,
            price=price.quantize(Decimal("0.01")),
            amount=amount.quantize(Decimal("0.01")) if amount > 0 else calc_amount
        ))

    if not items:
        raise ValueError("Не удалось извлечь таблицу товаров из PDF-текста. Нужен шаблон/другой формат КП.")

    return items

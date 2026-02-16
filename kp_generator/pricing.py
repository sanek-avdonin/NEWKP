from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
import random
from typing import List

from kp_generator.models import Item, VariantSettings  # абсолютный импорт


def _round_to_step(value: Decimal, step: Decimal) -> Decimal:
    if step <= 0:
        return value
    q = (value / step).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return (q * step).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def apply_pricing(items: List[Item], settings: VariantSettings, rng: random.Random) -> List[Item]:
    out: List[Item] = []
    percent_factor = (Decimal("1") + settings.percent_up / Decimal("100"))

    for it in items:
        base = it.price * percent_factor + settings.fixed_add

        if settings.random_spread > 0:
            spread = settings.random_spread
            delta = Decimal(str(rng.uniform(float(-spread), float(spread))))
            base = base + delta

        base = max(base, Decimal("0.01"))
        price_new = _round_to_step(base, settings.rounding_step)
        amount_new = (price_new * it.qty).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        out.append(Item(
            name=it.name,
            qty=it.qty,
            unit=it.unit,
            price=price_new,
            amount=amount_new
        ))

    return out

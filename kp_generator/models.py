from __future__ import annotations
from dataclasses import dataclass
from decimal import Decimal
from typing import Optional


@dataclass
class Item:
    name: str
    qty: Decimal
    unit: str
    price: Decimal
    amount: Decimal  # qty * price


@dataclass
class CompanyProfile:
    id: str
    name: str
    inn: str
    address: str
    phone: str
    ceo: str
    logo_path: Optional[str] = None


@dataclass
class VariantSettings:
    company_id: str
    percent_up: Decimal          # e.g. 1.5 means +1.5%
    fixed_add: Decimal           # rub per item
    random_spread: Decimal       # +/- rub
    rounding_step: Decimal       # 1/10/50/100

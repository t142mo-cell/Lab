# -*- coding: utf-8 -*-
from __future__ import annotations
from dataclasses import dataclass, asdict
from typing import Optional, Dict, Any

@dataclass
class Item:
    seq_id: int
    name: str
    category: str
    quantity: float
    unit: str
    storage_place: str
    packaging: str
    expiry_date: Optional[str] = None
    date_received: Optional[str] = None
    batch_number: str = ""
    responsible: str = ""
    qualification: Optional[str] = None
    reagent_type: Optional[str] = None
    state_register_no: Optional[str] = None
    certified_value: Optional[str] = None
    manufacture_date: Optional[str] = None
    manufacturer: Optional[str] = None
    storage_conditions: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "Item":
        defaults = dict(
            packaging="",
            expiry_date=None, date_received=None, batch_number="", responsible="",
            qualification=None, reagent_type=None,
            state_register_no=None, certified_value=None, manufacture_date=None,
            manufacturer=None, storage_conditions=None
        )
        payload = {**defaults, **d}
        return Item(**payload)
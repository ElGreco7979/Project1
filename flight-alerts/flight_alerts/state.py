"""Remembers which deals were already sent, so each run only alerts about new or cheaper ones."""
import json
from datetime import date
from pathlib import Path

from .models import Deal


class SeenStore:
    def __init__(self, path: Path):
        self.path = path
        self.prices: dict[str, float] = {}
        if path.exists():
            self.prices = json.loads(path.read_text())

    def is_new(self, deal: Deal, realert_drop: float) -> bool:
        old = self.prices.get(deal.key)
        return old is None or deal.price <= old - realert_drop

    def remember(self, deals: list[Deal]) -> None:
        for d in deals:
            self.prices[d.key] = min(d.price, self.prices.get(d.key, d.price))

    def save(self, today: date) -> None:
        # Drop flights that already departed; keys start with "ORIGIN|DEST|YYYY-MM-DD".
        cutoff = today.isoformat()
        self.prices = {k: v for k, v in self.prices.items() if k.split("|")[2][:10] >= cutoff}
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self.prices, indent=0, sort_keys=True))

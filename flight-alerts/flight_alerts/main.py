import argparse
import os
import sys
import tomllib
from datetime import date, timedelta
from pathlib import Path

from . import notify
from .geo import is_europe
from .models import Deal
from .providers import ryanair, travelpayouts
from .state import SeenStore

ROOT = Path(__file__).resolve().parent.parent


def collect(cfg: dict, today: date) -> tuple[list[Deal], list[str]]:
    s = cfg["search"]
    origin, max_price, currency = s["origin"].upper(), float(s["max_price"]), s["currency"].upper()
    start, end = today + timedelta(days=1), today + timedelta(days=int(s["days_ahead"]))
    providers = cfg.get("providers", {})
    deals, errors = [], []

    if providers.get("ryanair", True):
        try:
            deals += ryanair.fetch(origin, max_price, currency, start, end)
        except Exception as e:  # one provider failing shouldn't stop the others
            errors.append(f"Ryanair: {e}")
    if providers.get("travelpayouts", True):
        token = os.environ.get("TRAVELPAYOUTS_TOKEN")
        if not token:
            print("TRAVELPAYOUTS_TOKEN not set; skipping Travelpayouts (only Ryanair fares will be checked).")
        else:
            try:
                deals += travelpayouts.fetch(origin, max_price, currency, start, end, token)
            except Exception as e:
                errors.append(f"Travelpayouts: {e}")

    deals = [d for d in deals if d.price < max_price and d.currency.upper() == currency]
    if s.get("europe_only"):
        deals = [d for d in deals if is_europe(d.country_code)]
    return dedupe(deals), errors


def dedupe(deals: list[Deal]) -> list[Deal]:
    """The same flight can come from several providers; keep the cheapest."""
    best: dict[tuple, Deal] = {}
    for d in deals:
        k = (d.destination, d.departure.strftime("%Y-%m-%dT%H:%M"))
        if k not in best or d.price < best[k].price:
            best[k] = d
    return list(best.values())


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Alert about cheap flights.")
    parser.add_argument("--config", type=Path, default=ROOT / "config.toml")
    parser.add_argument("--state", type=Path, default=ROOT / "state" / "seen.json")
    parser.add_argument("--dry-run", action="store_true", help="print deals, send nothing, don't update state")
    args = parser.parse_args(argv)

    cfg = tomllib.loads(args.config.read_text())
    s, alerts = cfg["search"], cfg.get("alerts", {})
    today = date.today()
    deals, errors = collect(cfg, today)
    for err in errors:
        print(f"warning: {err}", file=sys.stderr)

    store = SeenStore(args.state)
    new = [d for d in deals if store.is_new(d, float(alerts.get("realert_drop", 5)))]
    # Europe first, then cheapest, then cap the message size.
    new.sort(key=lambda d: (not is_europe(d.country_code), d.price))
    new = new[: int(alerts.get("max_per_message", 40))]
    print(f"{len(deals)} fares under {s['max_price']} {s['currency']}, {len(new)} new.")

    if new:
        if args.dry_run:
            print("\n".join(notify.render(new, s["origin"], float(s["max_price"]), s["currency"], as_html=False)))
        else:
            channels = notify.send(new, s["origin"], float(s["max_price"]), s["currency"])
            print(f"Sent via {', '.join(channels)}.")
    if not args.dry_run:
        store.remember(new)
        store.save(today)
    # Fail the run only if every provider failed, so you notice broken setups.
    return 1 if errors and not deals else 0

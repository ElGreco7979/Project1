"""Travelpayouts / Aviasales flight data API: cheapest fares found by travellers' searches, all airlines.

Get a free token at https://www.travelpayouts.com (Tools -> API) and set TRAVELPAYOUTS_TOKEN.
"""
from datetime import date, datetime

from .. import http
from ..models import Deal

PRICES_URL = "https://api.travelpayouts.com/aviasales/v3/prices_for_dates"
CITIES_URL = "https://api.travelpayouts.com/data/en/cities.json"
AIRLINES_URL = "https://api.travelpayouts.com/data/en/airlines.json"


def months_between(start: date, end: date) -> list[str]:
    months, y, m = [], start.year, start.month
    while (y, m) <= (end.year, end.month):
        months.append(f"{y:04d}-{m:02d}")
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return months


def fetch(
    origin: str,
    max_price: float,
    currency: str,
    start: date,
    end: date,
    token: str,
    nights: tuple[int, int] | None = None,
    get_json=http.get_json,
) -> list[Deal]:
    """One-way fares when nights is None, otherwise round trips staying nights[0]..nights[1] nights."""
    cities = {c["code"]: c for c in get_json(CITIES_URL) if c.get("code")}
    airlines = {a["code"]: a.get("name") or a["code"] for a in get_json(AIRLINES_URL) if a.get("code")}
    deals = []
    for month in months_between(start, end):
        data = get_json(
            PRICES_URL,
            {
                "origin": origin,
                "departure_at": month,
                "one_way": "false" if nights else "true",
                "currency": currency.lower(),
                "sorting": "price",
                "limit": 1000,
                "token": token,
            },
        )
        if not data.get("success", True):
            raise RuntimeError(f"Travelpayouts error: {data.get('error')}")
        deals.extend(parse(data, origin, currency, cities, airlines, max_price, start, end, nights))
    return deals


def parse(data, origin, currency, cities, airlines, max_price, start, end, nights=None) -> list[Deal]:
    deals = []
    for row in data.get("data") or []:
        if row.get("price") is None or not row.get("departure_at"):
            continue
        price = float(row["price"])
        departure = parse_time(row["departure_at"])
        if price >= max_price or not (start <= departure.date() <= end):
            continue
        returning = parse_time(row["return_at"]) if row.get("return_at") else None
        if nights and (returning is None or not nights[0] <= (returning.date() - departure.date()).days <= nights[1]):
            continue
        stops = (row.get("transfers") or 0) + (row.get("return_transfers") or 0)
        city = cities.get(row.get("destination", ""), {})
        airline_code = row.get("airline", "")
        name = city.get("name") or row.get("destination", "")
        deals.append(
            Deal(
                provider="Aviasales",
                origin=origin,
                destination=row.get("destination_airport") or row.get("destination", ""),
                destination_name=name + (f" ({stops} stop{'s' if stops > 1 else ''})" if stops else ""),
                country_code=(city.get("country_code") or "").upper(),
                departure=departure.replace(tzinfo=None),
                price=price,
                currency=currency.upper(),
                airline=airlines.get(airline_code, airline_code),
                flight_number=f"{airline_code}{row.get('flight_number', '')}" if airline_code else "",
                link=f"https://www.aviasales.com{row['link']}" if row.get("link") else "",
                return_departure=returning.replace(tzinfo=None) if returning else None,
            )
        )
    return deals


def parse_time(value: str) -> datetime:
    return datetime.fromisoformat(value.replace("Z", "+00:00"))

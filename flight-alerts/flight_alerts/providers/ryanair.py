"""Ryanair's public fare finder (the API behind ryanair.com's "cheap flights" page). No key needed."""
from datetime import date, datetime, timedelta

from .. import http
from ..geo import EUROPE, OTHER
from ..models import Deal

URL = "https://services-api.ryanair.com/farfnd/v4/oneWayFares"
NAME_TO_CODE = {name.lower(): code for code, name in {**EUROPE, **OTHER}.items()}
WINDOW_DAYS = 7  # the API returns only the cheapest fare per destination per window


def fetch(origin: str, max_price: float, currency: str, start: date, end: date, get_json=http.get_json) -> list[Deal]:
    deals = []
    window_start = start
    while window_start <= end:
        window_end = min(window_start + timedelta(days=WINDOW_DAYS - 1), end)
        data = get_json(
            URL,
            {
                "departureAirportIataCode": origin,
                "outboundDepartureDateFrom": window_start.isoformat(),
                "outboundDepartureDateTo": window_end.isoformat(),
                "priceValueTo": max_price,
                "currency": currency,
                "language": "en",
                "market": "en-gb",
            },
        )
        deals.extend(parse(data, origin))
        window_start = window_end + timedelta(days=1)
    return deals


def parse(data: dict, origin: str) -> list[Deal]:
    deals = []
    for fare in data.get("fares") or []:
        out = fare.get("outbound") or {}
        arrival = out.get("arrivalAirport") or {}
        price = out.get("price") or {}
        if price.get("value") is None or not out.get("departureDate"):
            continue
        city = arrival.get("city") or {}
        country = (city.get("countryCode") or "").upper() or NAME_TO_CODE.get(
            (arrival.get("countryName") or "").lower(), ""
        )
        departure = datetime.fromisoformat(out["departureDate"])
        dest = arrival.get("iataCode", "")
        deals.append(
            Deal(
                provider="Ryanair",
                origin=origin,
                destination=dest,
                destination_name=city.get("name") or arrival.get("name") or dest,
                country_code=country,
                departure=departure,
                price=float(price["value"]),
                currency=price.get("currencyCode", ""),
                airline="Ryanair",
                flight_number=out.get("flightNumber", ""),
                link=(
                    "https://www.ryanair.com/gb/en/trip/flights/select?adults=1&isReturn=false"
                    f"&dateOut={departure.date().isoformat()}&originIata={origin}&destinationIata={dest}"
                ),
            )
        )
    return deals

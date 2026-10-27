"""Ryanair's public fare finder (the API behind ryanair.com's "cheap flights" page). No key needed."""
from datetime import date, datetime, timedelta

from .. import http
from ..geo import EUROPE, OTHER
from ..models import Deal

ONE_WAY_URL = "https://services-api.ryanair.com/farfnd/v4/oneWayFares"
ROUND_TRIP_URL = "https://services-api.ryanair.com/farfnd/v4/roundTripFares"
NAME_TO_CODE = {name.lower(): code for code, name in {**EUROPE, **OTHER}.items()}
WINDOW_DAYS = 7  # the API returns only the cheapest fare per destination per window


def fetch(
    origin: str,
    max_price: float,
    currency: str,
    start: date,
    end: date,
    nights: tuple[int, int] | None = None,
    get_json=http.get_json,
) -> list[Deal]:
    """One-way fares when nights is None, otherwise round trips staying nights[0]..nights[1] nights."""
    deals = []
    window_start = start
    while window_start <= end:
        window_end = min(window_start + timedelta(days=WINDOW_DAYS - 1), end)
        params = {
            "departureAirportIataCode": origin,
            "outboundDepartureDateFrom": window_start.isoformat(),
            "outboundDepartureDateTo": window_end.isoformat(),
            "priceValueTo": max_price,
            "currency": currency,
            "language": "en",
            "market": "en-gb",
        }
        if nights:
            params |= {
                "inboundDepartureDateFrom": (window_start + timedelta(days=nights[0])).isoformat(),
                "inboundDepartureDateTo": (window_end + timedelta(days=nights[1])).isoformat(),
                "durationFrom": nights[0],
                "durationTo": nights[1],
            }
        deals.extend(parse(get_json(ROUND_TRIP_URL if nights else ONE_WAY_URL, params), origin))
        window_start = window_end + timedelta(days=1)
    return deals


def parse(data: dict, origin: str) -> list[Deal]:
    deals = []
    for fare in data.get("fares") or []:
        out, back = fare.get("outbound") or {}, fare.get("inbound")
        arrival = out.get("arrivalAirport") or {}
        price = total_price(fare)
        if price is None or not out.get("departureDate") or (back is not None and not back.get("departureDate")):
            continue
        city = arrival.get("city") or {}
        country = (city.get("countryCode") or "").upper() or NAME_TO_CODE.get(
            (arrival.get("countryName") or "").lower(), ""
        )
        departure = datetime.fromisoformat(out["departureDate"])
        returning = datetime.fromisoformat(back["departureDate"]) if back else None
        dest = arrival.get("iataCode", "")
        link = (
            f"https://www.ryanair.com/gb/en/trip/flights/select?adults=1&isReturn={'true' if back else 'false'}"
            f"&dateOut={departure.date().isoformat()}&originIata={origin}&destinationIata={dest}"
        )
        if returning:
            link += f"&dateIn={returning.date().isoformat()}"
        flight_numbers = "/".join(n for n in [out.get("flightNumber"), (back or {}).get("flightNumber")] if n)
        deals.append(
            Deal(
                provider="Ryanair",
                origin=origin,
                destination=dest,
                destination_name=city.get("name") or arrival.get("name") or dest,
                country_code=country,
                departure=departure,
                price=price,
                currency=((fare.get("summary") or {}).get("price") or out.get("price") or {}).get("currencyCode", ""),
                airline="Ryanair",
                flight_number=flight_numbers,
                link=link,
                return_departure=returning,
            )
        )
    return deals


def total_price(fare: dict) -> float | None:
    """The summary holds the total; fall back to adding the legs."""
    summary = ((fare.get("summary") or {}).get("price") or {}).get("value")
    if summary is not None:
        return float(summary)
    legs = [fare.get("outbound"), fare.get("inbound")]
    values = [((leg or {}).get("price") or {}).get("value") for leg in legs if leg is not None]
    return float(sum(values)) if values and None not in values else None

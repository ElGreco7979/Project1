import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from flight_alerts import notify
from flight_alerts.main import dedupe
from flight_alerts.models import Deal
from flight_alerts.providers import ryanair, travelpayouts
from flight_alerts.state import SeenStore

RYANAIR_SAMPLE = {
    "fares": [
        {
            "outbound": {
                "departureAirport": {"iataCode": "LCA", "name": "Larnaca"},
                "arrivalAirport": {
                    "countryName": "Greece",
                    "iataCode": "ATH",
                    "name": "Athens",
                    "city": {"name": "Athens", "code": "ATHENS", "countryCode": "gr"},
                },
                "departureDate": "2026-11-10T06:25:00",
                "price": {"value": 19.99, "currencyCode": "EUR"},
                "flightNumber": "FR1234",
            }
        },
        {"outbound": {"arrivalAirport": {"iataCode": "XXX"}, "price": {}}},  # incomplete row is skipped
    ]
}


ROUND_TRIP_SAMPLE = {
    "fares": [
        {
            "outbound": {**RYANAIR_SAMPLE["fares"][0]["outbound"], "price": {"value": 30, "currencyCode": "EUR"}},
            "inbound": {"departureDate": "2026-11-14T09:10:00", "flightNumber": "FR1235",
                        "price": {"value": 25.5, "currencyCode": "EUR"}},
            "summary": {"price": {"value": 55.5, "currencyCode": "EUR"}},
        }
    ]
}


def deal(dest="ATH", country="GR", price=50.0, day=10, **kw):
    return Deal(
        provider="Test", origin="LCA", destination=dest, destination_name=dest, country_code=country,
        departure=datetime(2026, 11, day, 8, 0), price=price, currency="EUR", **kw,
    )


class RyanairTest(unittest.TestCase):
    def test_parse(self):
        [d] = ryanair.parse(RYANAIR_SAMPLE, "LCA")
        self.assertEqual((d.destination, d.country_code, d.price), ("ATH", "GR", 19.99))
        self.assertIn("destinationIata=ATH", d.link)

    def test_parse_round_trip(self):
        [d] = ryanair.parse(ROUND_TRIP_SAMPLE, "LCA")
        self.assertEqual((d.price, d.nights, d.flight_number), (55.5, 4, "FR1234/FR1235"))
        self.assertIn("isReturn=true", d.link)
        self.assertIn("dateIn=2026-11-14", d.link)

    def test_round_trip_price_falls_back_to_legs(self):
        fare = {**ROUND_TRIP_SAMPLE["fares"][0], "summary": {}}
        self.assertEqual(ryanair.total_price(fare), 55.5)

    def test_round_trip_request(self):
        calls = []
        ryanair.fetch("LCA", 100, "EUR", date(2026, 11, 1), date(2026, 11, 7), nights=(2, 14),
                      get_json=lambda url, params: calls.append((url, params)) or {"fares": []})
        [(url, params)] = calls
        self.assertEqual(url, ryanair.ROUND_TRIP_URL)
        self.assertEqual((params["inboundDepartureDateFrom"], params["inboundDepartureDateTo"]), ("2026-11-03", "2026-11-21"))
        self.assertEqual((params["durationFrom"], params["durationTo"]), (2, 14))

    def test_country_from_name_when_code_missing(self):
        sample = {"fares": [{"outbound": {**RYANAIR_SAMPLE["fares"][0]["outbound"],
                                          "arrivalAirport": {"iataCode": "BGY", "countryName": "Italy"}}}]}
        self.assertEqual(ryanair.parse(sample, "LCA")[0].country_code, "IT")

    def test_fetch_splits_into_weekly_windows(self):
        calls = []
        ryanair.fetch("LCA", 100, "EUR", date(2026, 11, 1), date(2026, 11, 20),
                      get_json=lambda url, params: calls.append(params) or {"fares": []})
        self.assertEqual([c["outboundDepartureDateFrom"] for c in calls], ["2026-11-01", "2026-11-08", "2026-11-15"])
        self.assertEqual(calls[-1]["outboundDepartureDateTo"], "2026-11-20")


class TravelpayoutsTest(unittest.TestCase):
    def fake_get(self, url, params=None):
        if url == travelpayouts.CITIES_URL:
            return [{"code": "BUD", "name": "Budapest", "country_code": "HU"},
                    {"code": "DXB", "name": "Dubai", "country_code": "AE"}]
        if url == travelpayouts.AIRLINES_URL:
            return [{"code": "W6", "name": "Wizz Air"}]
        self.requested.append(params["departure_at"])
        return {"success": True, "data": [
            {"destination": "BUD", "destination_airport": "BUD", "price": 39, "airline": "W6",
             "flight_number": "2312", "departure_at": "2026-11-12T13:05:00+02:00", "transfers": 0,
             "link": "/search/LCA1211BUD1"},
            {"destination": "DXB", "destination_airport": "DXB", "price": 150, "airline": "W6",
             "departure_at": "2026-11-12T13:05:00+02:00"},  # too expensive
            {"destination": "DXB", "destination_airport": "DXB", "price": 60, "airline": "W6",
             "departure_at": "2027-03-01T13:05:00+02:00"},  # outside the window
        ]}

    def test_fetch(self):
        self.requested = []
        deals = travelpayouts.fetch("LCA", 100, "EUR", date(2026, 11, 1), date(2026, 12, 15), "tok",
                                    get_json=self.fake_get)
        self.assertEqual(self.requested, ["2026-11", "2026-12"])
        self.assertEqual(len(deals), 2)  # same BUD row for each month request
        d = deals[0]
        self.assertEqual((d.destination_name, d.country_code, d.airline, d.flight_number), ("Budapest", "HU", "Wizz Air", "W62312"))
        self.assertEqual(d.link, "https://www.aviasales.com/search/LCA1211BUD1")

    def test_round_trip_filters_stay_length(self):
        rows = {"data": [
            {"destination": "BUD", "price": 80, "departure_at": "2026-11-12T13:05:00+02:00",
             "return_at": "2026-11-16T10:00:00+01:00", "transfers": 0, "return_transfers": 1},
            {"destination": "BUD", "price": 70, "departure_at": "2026-11-12T13:05:00+02:00",
             "return_at": "2026-12-20T10:00:00+01:00"},  # 38 nights: too long
            {"destination": "BUD", "price": 60, "departure_at": "2026-11-12T13:05:00+02:00"},  # no return
        ]}
        deals = travelpayouts.parse(rows, "LCA", "EUR", {}, {}, 100, date(2026, 11, 1), date(2026, 12, 31), (2, 14))
        [d] = deals
        self.assertEqual((d.price, d.nights, d.destination_name), (80, 4, "BUD (1 stop)"))

    def test_round_trip_request(self):
        self.requested = []
        params = []
        def get(url, p=None):
            if p:
                params.append(p)
            return self.fake_get(url, p)
        travelpayouts.fetch("LCA", 100, "EUR", date(2026, 11, 1), date(2026, 11, 30), "tok", (2, 14), get_json=get)
        self.assertEqual(params[0]["one_way"], "false")

    def test_months_across_year(self):
        self.assertEqual(travelpayouts.months_between(date(2026, 11, 28), date(2027, 1, 2)), ["2026-11", "2026-12", "2027-01"])


class LogicTest(unittest.TestCase):
    def test_dedupe_keeps_cheapest(self):
        out = dedupe([deal(price=60), deal(price=45), deal(dest="BUD", country="HU")])
        self.assertEqual(sorted(d.price for d in out), [45, 50])

    def test_europe_listed_first(self):
        groups = notify.group([deal(dest="DXB", country="AE", price=20), deal(price=80), deal(dest="BUD", country="HU", price=30)])
        self.assertEqual([t for t, _ in groups], ["🇪🇺 Europe", "🌍 Rest of the world"])
        self.assertEqual([d.destination for d in groups[0][1]], ["BUD", "ATH"])

    def test_render_splits_long_messages(self):
        deals = [deal(dest=f"A{i:02d}", day=1 + i % 28, link="https://example.com/" + "x" * 80) for i in range(80)]
        chunks = notify.render(deals, "LCA", 100, "EUR", as_html=True)
        self.assertGreater(len(chunks), 1)
        self.assertTrue(all(len(c) <= 4096 for c in chunks))

    def test_render_round_trip(self):
        d = deal(return_departure=datetime(2026, 11, 14, 9, 0))
        text = notify.render([d], "LCA", 100, "EUR", as_html=False)[0]
        self.assertIn("new round trip deal from LCA", text)
        self.assertIn("Tue 10 Nov 08:00 → Sat 14 Nov 09:00 (4 nights)", text)

    def test_dedupe_keeps_different_return_dates(self):
        a = deal(return_departure=datetime(2026, 11, 14))
        b = deal(return_departure=datetime(2026, 11, 17))
        self.assertEqual(len(dedupe([a, b])), 2)

    def test_seen_store(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "seen.json"
            store = SeenStore(path)
            self.assertTrue(store.is_new(deal(price=50), 5))
            store.remember([deal(price=50), deal(dest="OLD", day=1)])
            store.save(date(2026, 11, 5))  # the Nov 1 flight is dropped
            store = SeenStore(path)
            self.assertEqual(len(store.prices), 1)
            self.assertFalse(store.is_new(deal(price=48), 5))
            self.assertTrue(store.is_new(deal(price=45), 5))


if __name__ == "__main__":
    unittest.main()

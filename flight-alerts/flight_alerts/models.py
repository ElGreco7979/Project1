from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True)
class Deal:
    provider: str
    origin: str
    destination: str        # IATA airport code
    destination_name: str   # city or airport name, for display
    country_code: str       # ISO 3166-1 alpha-2, upper case ("" if unknown)
    departure: datetime
    price: float            # total for the whole trip
    currency: str
    airline: str = ""
    flight_number: str = ""
    link: str = ""
    return_departure: datetime | None = None  # None for one-way

    @property
    def nights(self) -> int | None:
        return (self.return_departure.date() - self.departure.date()).days if self.return_departure else None

    @property
    def key(self) -> str:
        """Identifies the same flight across runs, independent of its price."""
        return "|".join(
            [
                self.origin,
                self.destination,
                self.departure.strftime("%Y-%m-%dT%H:%M"),
                self.return_departure.strftime("%Y-%m-%dT%H:%M") if self.return_departure else "",
                self.airline,
                self.flight_number,
            ]
        )

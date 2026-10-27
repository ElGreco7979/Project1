"""Country helpers for grouping deals into Europe and the rest of the world."""

# ISO codes of countries treated as Europe (geographic Europe plus transcontinental neighbours).
EUROPE = {
    "AD": "Andorra", "AL": "Albania", "AM": "Armenia", "AT": "Austria", "AZ": "Azerbaijan",
    "BA": "Bosnia and Herzegovina", "BE": "Belgium", "BG": "Bulgaria", "BY": "Belarus",
    "CH": "Switzerland", "CY": "Cyprus", "CZ": "Czechia", "DE": "Germany", "DK": "Denmark",
    "EE": "Estonia", "ES": "Spain", "FI": "Finland", "FO": "Faroe Islands", "FR": "France",
    "GB": "United Kingdom", "GE": "Georgia", "GI": "Gibraltar", "GR": "Greece", "HR": "Croatia",
    "HU": "Hungary", "IE": "Ireland", "IS": "Iceland", "IT": "Italy", "XK": "Kosovo",
    "LI": "Liechtenstein", "LT": "Lithuania", "LU": "Luxembourg", "LV": "Latvia", "MC": "Monaco",
    "MD": "Moldova", "ME": "Montenegro", "MK": "North Macedonia", "MT": "Malta",
    "NL": "Netherlands", "NO": "Norway", "PL": "Poland", "PT": "Portugal", "RO": "Romania",
    "RS": "Serbia", "RU": "Russia", "SE": "Sweden", "SI": "Slovenia", "SK": "Slovakia",
    "SM": "San Marino", "TR": "Turkey", "UA": "Ukraine", "VA": "Vatican City",
}

# Names for common destinations outside Europe; anything else is shown by its code.
OTHER = {
    "AE": "United Arab Emirates", "BH": "Bahrain", "EG": "Egypt", "IL": "Israel", "JO": "Jordan",
    "LB": "Lebanon", "MA": "Morocco", "QA": "Qatar", "SA": "Saudi Arabia", "TN": "Tunisia",
    "KW": "Kuwait", "OM": "Oman", "IN": "India", "TH": "Thailand", "US": "United States",
}


def is_europe(country_code: str) -> bool:
    return country_code.upper() in EUROPE


def country_name(country_code: str) -> str:
    code = country_code.upper()
    return EUROPE.get(code) or OTHER.get(code) or code or "Unknown"

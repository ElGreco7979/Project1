import json
import urllib.parse
import urllib.request

USER_AGENT = "Mozilla/5.0 (compatible; flight-alerts/1.0)"


def get_json(url: str, params: dict | None = None, timeout: int = 30):
    if params:
        url = f"{url}?{urllib.parse.urlencode(params)}"
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.load(resp)


def post_json(url: str, payload: dict, timeout: int = 30):
    data = json.dumps(payload).encode()
    req = urllib.request.Request(
        url, data=data, headers={"User-Agent": USER_AGENT, "Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.load(resp)

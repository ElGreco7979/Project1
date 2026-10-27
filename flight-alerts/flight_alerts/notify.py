"""Formats deals and sends them by Telegram, email, or to the console."""
import html
import os
import smtplib
from email.message import EmailMessage

from . import http
from .geo import country_name, is_europe
from .models import Deal

TELEGRAM_LIMIT = 4000  # Telegram allows 4096 characters per message


def group(deals: list[Deal]) -> list[tuple[str, list[Deal]]]:
    """Europe first, then everything else; each sorted by price."""
    europe = sorted((d for d in deals if is_europe(d.country_code)), key=lambda d: (d.price, d.departure))
    rest = sorted((d for d in deals if not is_europe(d.country_code)), key=lambda d: (d.price, d.departure))
    return [(title, items) for title, items in [("🇪🇺 Europe", europe), ("🌍 Rest of the world", rest)] if items]


def line(d: Deal, as_html: bool) -> str:
    when = d.departure.strftime("%a %d %b %H:%M")
    place = f"{d.destination_name} ({d.destination}), {country_name(d.country_code)}"
    price = f"{d.price:.0f} {d.currency}" if d.price == int(d.price) else f"{d.price:.2f} {d.currency}"
    carrier = d.airline or d.provider
    if not as_html:
        return f"• {price} – {place} – {when} – {carrier}" + (f"\n  {d.link}" if d.link else "")
    text = f"<b>{html.escape(price)}</b> – {html.escape(place)} – {html.escape(when)} – {html.escape(carrier)}"
    return f"• <a href=\"{html.escape(d.link)}\">{text}</a>" if d.link else f"• {text}"


def render(deals: list[Deal], origin: str, max_price: float, currency: str, as_html: bool) -> list[str]:
    """Returns the message split into chunks that fit in one Telegram message."""
    header = f"✈️ {len(deals)} new flight deal{'s' if len(deals) != 1 else ''} from {origin} under {max_price:g} {currency}"
    lines = [f"<b>{html.escape(header)}</b>" if as_html else header]
    for title, items in group(deals):
        lines.append("")
        lines.append(f"<b>{title}</b>" if as_html else title)
        lines.extend(line(d, as_html) for d in items)
    chunks, current = [], ""
    for ln in lines:
        if current and len(current) + len(ln) + 1 > TELEGRAM_LIMIT:
            chunks.append(current)
            current = ""
        current = f"{current}\n{ln}" if current else ln
    if current:
        chunks.append(current)
    return chunks


def send_telegram(chunks: list[str], token: str, chat_id: str, post_json=http.post_json) -> None:
    for chunk in chunks:
        post_json(
            f"https://api.telegram.org/bot{token}/sendMessage",
            {"chat_id": chat_id, "text": chunk, "parse_mode": "HTML", "disable_web_page_preview": True},
        )


def send_email(subject: str, body: str) -> None:
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = os.environ.get("EMAIL_FROM") or os.environ["SMTP_USER"]
    msg["To"] = os.environ["EMAIL_TO"]
    msg.set_content(body)
    host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    port = int(os.environ.get("SMTP_PORT", "587"))
    with smtplib.SMTP(host, port, timeout=30) as smtp:
        smtp.starttls()
        smtp.login(os.environ["SMTP_USER"], os.environ["SMTP_PASSWORD"])
        smtp.send_message(msg)


def send(deals: list[Deal], origin: str, max_price: float, currency: str) -> list[str]:
    """Sends to every configured channel; returns the channels used."""
    used = []
    tg_token, tg_chat = os.environ.get("TELEGRAM_BOT_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")
    if tg_token and tg_chat:
        send_telegram(render(deals, origin, max_price, currency, as_html=True), tg_token, tg_chat)
        used.append("telegram")
    if os.environ.get("EMAIL_TO") and os.environ.get("SMTP_USER"):
        body = "\n".join(render(deals, origin, max_price, currency, as_html=False))
        send_email(body.splitlines()[0], body)
        used.append("email")
    if not used:
        print("\n".join(render(deals, origin, max_price, currency, as_html=False)))
        used.append("console")
    return used

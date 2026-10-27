# Flight alerts from Larnaca ✈️

Checks for **round trips from Larnaca (LCA) costing under 100 € in total** (staying 2–14 nights)
and sends you an alert by
**Telegram** and/or **email**. Deals in **Europe are listed first**, then every other country.
You're only alerted about a flight once — again only if its price drops by another 5 €.

Runs for free on GitHub Actions three times a day; no server needed. Plain Python 3.11+, no dependencies.

## Where the fares come from

| Source | Covers | Setup |
|---|---|---|
| Ryanair fare finder | Ryanair | none |
| [Travelpayouts / Aviasales](https://www.travelpayouts.com) | almost every airline (Wizz Air, Aegean, Cyprus Airways, easyJet, …) and every country | free token |

Travelpayouts prices come from recent searches by Aviasales users, so they can be a few hours or
days old — always confirm the price on the booking link before buying.

## Setup (≈10 minutes)

1. **Travelpayouts token** – sign up free at travelpayouts.com, then *Tools → API* (or *Profile → API token*) and copy the token.
2. **Telegram alerts** (recommended)
   - In Telegram, talk to **@BotFather**, send `/newbot`, and copy the bot token.
   - Send any message to your new bot, then open
     `https://api.telegram.org/bot<TOKEN>/getUpdates` and copy the `"chat":{"id": …}` number.
3. **Email alerts** (optional) – with Gmail, create an [App password](https://myaccount.google.com/apppasswords).
4. In GitHub: **Settings → Secrets and variables → Actions → New repository secret**, add:

   | Secret | Value |
   |---|---|
   | `TRAVELPAYOUTS_TOKEN` | your Travelpayouts token |
   | `TELEGRAM_BOT_TOKEN` | bot token from BotFather |
   | `TELEGRAM_CHAT_ID` | your chat id |
   | `SMTP_USER` / `SMTP_PASSWORD` | Gmail address / app password (optional) |
   | `EMAIL_TO` | where to send email alerts (optional) |

   For a non-Gmail provider also set `SMTP_HOST` and `SMTP_PORT` env vars in the workflow.
5. Merge to the default branch (scheduled workflows only run there), then **Actions → Flight alerts → Run workflow** to test.

## Settings

Edit [`config.toml`](config.toml): origin airport, price limit, how many days ahead to search
(default 90), round trip vs. one-way and the length of stay (`min_nights` / `max_nights`),
`europe_only` to drop non-European destinations, and re-alert/message-size limits.
To change how often it runs, edit the `cron` line in `.github/workflows/flight-alerts.yml`.

## Run locally

```sh
cd flight-alerts
export TRAVELPAYOUTS_TOKEN=...          # optional; without it only Ryanair is checked
python -m flight_alerts --dry-run       # print the deals, send nothing
python -m flight_alerts                 # send via Telegram/email if configured, else print
python -m unittest discover -s tests
```

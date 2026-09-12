# Positions — v5 (no yfinance / no curl_cffi)

Same Excel-matching layout as v4 (currency-grouped, subtotals, HKD→USD
equivalents, matching colors) — but the price-fetching underneath has been
rewritten to avoid the crash you hit.

## What changed and why

`yfinance` (which we were using to fetch prices) depends on a package
called `curl_cffi`. That package includes a compiled component that
deliberately mimics a real browser's network fingerprint, to get past
Yahoo Finance's bot-blocking. macOS's built-in security scanner (XProtect)
flags that specific technique heuristically — it's a known, recurring
false-positive with `curl_cffi` on macOS, not something specific to this
app, and not a sign your data was compromised.

Rather than override macOS security settings, **v5 removes that dependency
entirely.** It fetches prices with `requests` — a completely standard,
uncontroversial Python library — calling Yahoo Finance's public chart
endpoint directly. Same data source, no flagged binary.

## Setup (new folder — full steps)

**1. Create the folder:**
```
mkdir -p ~/Documents/portfolio_app_v5
```

**2. Copy in the three files** (click each file card in chat, Cmd+A to
select all the text in the preview, Cmd+C to copy), then in Terminal:

```
nano ~/Documents/portfolio_app_v5/app.py
```
Paste (Cmd+V), then Ctrl+O, Enter, Ctrl+X to save and exit.

```
nano ~/Documents/portfolio_app_v5/requirements.txt
```
Paste, Ctrl+O, Enter, Ctrl+X.

```
nano ~/Documents/portfolio_app_v5/README.md
```
Paste, Ctrl+O, Enter, Ctrl+X.

**3. Confirm all three are there:**
```
ls ~/Documents/portfolio_app_v5
```

**4. Set up the environment:**
```
cd ~/Documents/portfolio_app_v5
/opt/homebrew/bin/python3.14 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

This should install quickly — just Streamlit and `requests`, nothing
compiled, nothing flagged.

**5. Run it:**
```
streamlit run app.py
```

## Using it

- Click **🔄 Update Prices** to fetch every ticker in one pass.
- The **USD/HKD rate** field feeds the gold "Equiv. in USD" rows under
  each HKD section — update it any time.
- Manage positions (add/remove) at the bottom of the page.
- Data persists in `portfolio_data.json` next to this file.

## A note on data quality

Yahoo's chart endpoint doesn't require the browser-impersonation that
`curl_cffi` provided, but Yahoo does occasionally rate-limit or change
its response format without notice, since this isn't an officially
documented/supported API. If a ticker consistently shows "— (update)"
after clicking Update Prices, that specific symbol's quote may be
temporarily unavailable through this feed — try again in a bit, or
double-check the ticker format.

"""
Positions — mirrors your Excel workbook layout: sections grouped by
currency (USD / HKD / IDR), with subtotals per currency and an HKD-to-USD
equivalent, using the same color scheme as the spreadsheet.

RUN:
    pip install -r requirements.txt
    streamlit run app.py

Data is stored locally in portfolio_data.json (created automatically next
to this file) so any edits you make persist between runs.
"""

import json
import os
from datetime import datetime

import streamlit as st
import requests

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "portfolio_data.json")

# ---------- colors, matched to the Excel workbook ----------
C_BUY = "#DDEBF7"       # light blue  - buy price / qty / total purchased
C_CUR = "#FFF2CC"       # light amber - current price / total value
C_PL = "#FCE4E4"        # light pink  - profit/(loss)
C_TGT = "#E2EFDA"       # light green - target price / total value / profit(loss)
C_SUBTOTAL = "#D9D9D9"  # gray        - subtotal rows
C_GRANDTOTAL = "#FFD966"  # gold      - HKD -> USD equivalent rows
C_SECTION = "#203864"   # dark navy  - section banner
C_SECTION_TEXT = "#FFFFFF"


# ---------- default data, seeded from your spreadsheet ----------

def default_data():
    return {
        "to_sell": [
            {"ticker": "QS",      "name": "QuantumScape", "yahoo": "QS",      "currency": "USD", "buy": 48.44,   "qty": 2050,  "target": 10.00},
            {"ticker": "NIO",     "name": "",             "yahoo": "NIO",     "currency": "USD", "buy": 38.99,   "qty": 450,   "target": 6.40},
            {"ticker": "GRAB",    "name": "Grab Holdings","yahoo": "GRAB",    "currency": "USD", "buy": 12.49,   "qty": 5800,  "target": 4.00},
            {"ticker": "175.HK",  "name": "Geely",        "yahoo": "0175.HK", "currency": "HKD", "buy": 24.12,   "qty": 24660, "target": 24.50},
            {"ticker": "NIO",     "name": "(2nd lot)",    "yahoo": "NIO",     "currency": "HKD", "buy": 45.23,   "qty": 450,   "target": 6.00},
            {"ticker": "9888.HK", "name": "Baidu",        "yahoo": "9888.HK", "currency": "HKD", "buy": 177.12,  "qty": 1000,  "target": 130.00},
            {"ticker": "2318.HK", "name": "Ping An",      "yahoo": "2318.HK", "currency": "HKD", "buy": 71.47,   "qty": 27980, "target": 65.00},
            {"ticker": "BBHI.JK", "name": "",             "yahoo": "BBHI.JK", "currency": "IDR", "buy": 6128,    "qty": 175,   "target": 1000},
            {"ticker": "BEKS.JK", "name": "",             "yahoo": "BEKS.JK", "currency": "IDR", "buy": 30,      "qty": 500,   "target": 40},
            {"ticker": "KBRI",    "name": "",             "yahoo": "KBRI.JK", "currency": "IDR", "buy": 63.67,   "qty": 750,   "target": 50},
        ],
        "to_reduce": [
            {"ticker": "SEA",     "name": "Sea Ltd",      "yahoo": "SE",      "currency": "USD", "buy": 365.32,  "qty": 60,    "target": 180.00},
            {"ticker": "6618.HK", "name": "JD Health",    "yahoo": "6618.HK", "currency": "HKD", "buy": 137.59,  "qty": 2272,  "target": 48.00},
        ],
        "hold": [
            {"ticker": "SE",      "name": "Sea Ltd",      "yahoo": "SE",      "currency": "USD", "buy": 288.80,  "qty": 680,   "target": 180.00},
            {"ticker": "GLD",     "name": "SPDR Gold",    "yahoo": "GLD",     "currency": "USD", "buy": 367.00,  "qty": 60,    "target": 460.00},
            {"ticker": "0700.HK", "name": "Tencent",      "yahoo": "0700.HK", "currency": "HKD", "buy": 528.76,  "qty": 290,   "target": 484.00},
            {"ticker": "9988.HK", "name": "BABA",         "yahoo": "9988.HK", "currency": "HKD", "buy": 209.96,  "qty": 13594, "target": 160.00},
            {"ticker": "SCMA.JK", "name": "",             "yahoo": "SCMA.JK", "currency": "IDR", "buy": 222.15,  "qty": 37770, "target": 320},
            {"ticker": "EMTK.JK", "name": "",             "yahoo": "EMTK.JK", "currency": "IDR", "buy": 1082.59, "qty": 4075,  "target": 950},
            {"ticker": "AMMS.JK", "name": "",             "yahoo": "AMMS.JK", "currency": "IDR", "buy": 472.85,  "qty": 2188,  "target": 350},
            {"ticker": "BUKA.JK", "name": "",             "yahoo": "BUKA.JK", "currency": "IDR", "buy": 454.44,  "qty": 3650,  "target": 250},
            {"ticker": "AMOR.JK", "name": "",             "yahoo": "AMOR.JK", "currency": "IDR", "buy": 1351.00, "qty": 1348,  "target": 490},
            {"ticker": "AMMN.JK", "name": "",             "yahoo": "AMMN.JK", "currency": "IDR", "buy": 7950.00, "qty": 100,   "target": 9000},
            {"ticker": "MDKA.JK", "name": "",             "yahoo": "MDKA.JK", "currency": "IDR", "buy": 3500.00, "qty": 150,   "target": 4800},
            {"ticker": "DEWA.JK", "name": "",             "yahoo": "DEWA.JK", "currency": "IDR", "buy": 545.00,  "qty": 1000,  "target": 650},
            {"ticker": "RAJA.JK", "name": "",             "yahoo": "RAJA.JK", "currency": "IDR", "buy": 1015.00, "qty": 500,   "target": 1500},
        ],
        "build_core": [
            {"ticker": "MSFT",    "name": "",           "yahoo": "MSFT",    "currency": "USD", "target": 400},
            {"ticker": "TSM",     "name": "TSMC",       "yahoo": "TSM",     "currency": "USD", "target": 370},
            {"ticker": "BRK-B",   "name": "Berkshire",  "yahoo": "BRK-B",   "currency": "USD", "target": 470},
            {"ticker": "GOOG",    "name": "Alphabet",   "yahoo": "GOOG",    "currency": "USD", "target": 290},
            {"ticker": "GLD",     "name": "SPDR Gold",  "yahoo": "GLD",     "currency": "USD", "target": 370},
            {"ticker": "BMRI.JK", "name": "",           "yahoo": "BMRI.JK", "currency": "IDR", "target": 4100},
            {"ticker": "BBRI.JK", "name": "",           "yahoo": "BBRI.JK", "currency": "IDR", "target": 3200},
        ],
        "tactical": [
            {"ticker": "XOM",     "name": "Exxon", "yahoo": "XOM",     "currency": "USD", "target": 135},
            {"ticker": "ORCL",    "name": "",      "yahoo": "ORCL",    "currency": "USD", "target": 120},
            {"ticker": "FTNT",    "name": "",      "yahoo": "FTNT",    "currency": "USD", "target": 100},
            {"ticker": "DEWA.JK", "name": "",      "yahoo": "DEWA.JK", "currency": "IDR", "target": 300},
            {"ticker": "SCMA.JK", "name": "",      "yahoo": "SCMA.JK", "currency": "IDR", "target": 190},
            {"ticker": "RAJA.JK", "name": "",      "yahoo": "RAJA.JK", "currency": "IDR", "target": 680},
            {"ticker": "ANTM.JK", "name": "",      "yahoo": "ANTM.JK", "currency": "IDR", "target": 3050},
            {"ticker": "MDKA.JK", "name": "",      "yahoo": "MDKA.JK", "currency": "IDR", "target": 2700},
            {"ticker": "EMTK.JK", "name": "",      "yahoo": "EMTK.JK", "currency": "IDR", "target": 470},
        ],
        "trading": [
            {"ticker": "TSM",     "name": "", "yahoo": "TSM",     "currency": "USD", "target": 370},
            {"ticker": "MSFT",    "name": "", "yahoo": "MSFT",    "currency": "USD", "target": 400},
            {"ticker": "AMD",     "name": "", "yahoo": "AMD",     "currency": "USD", "target": 350},
            {"ticker": "PTRO.JK", "name": "", "yahoo": "PTRO.JK", "currency": "IDR", "target": 4800},
            {"ticker": "WIFI.JK", "name": "", "yahoo": "WIFI.JK", "currency": "IDR", "target": 1700},
            {"ticker": "DEWA.JK", "name": "", "yahoo": "DEWA.JK", "currency": "IDR", "target": 400},
            {"ticker": "ANTM.JK", "name": "", "yahoo": "ANTM.JK", "currency": "IDR", "target": 2800},
            {"ticker": "MDKA.JK", "name": "", "yahoo": "MDKA.JK", "currency": "IDR", "target": 2700},
        ],
    }


# ---------- persistence ----------

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as f:
            return json.load(f)
    return default_data()


def save_data(data):
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=2)


if "data" not in st.session_state:
    st.session_state.data = load_data()
if "prices" not in st.session_state:
    st.session_state.prices = {}
if "last_updated" not in st.session_state:
    st.session_state.last_updated = None
if "hkd_rate" not in st.session_state:
    st.session_state.hkd_rate = 7.84


def lot_mult(currency):
    return 100 if currency == "IDR" else 1


# ---------- price fetching ----------
# Talks to Yahoo Finance's public chart endpoint directly via `requests` —
# no yfinance, no curl_cffi. This is the same endpoint style used by many
# lightweight tools (including the Google Sheets script from earlier) and
# avoids the browser-fingerprint-spoofing binary that some macOS security
# software flags as suspicious.

YAHOO_HEADERS = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"}


def fetch_one(symbol: str):
    try:
        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"
        resp = requests.get(url, headers=YAHOO_HEADERS, timeout=8)
        resp.raise_for_status()
        data = resp.json()
        result = data.get("chart", {}).get("result")
        if not result:
            return {"price": None}
        meta = result[0].get("meta", {})
        price = meta.get("regularMarketPrice")
        return {"price": price}
    except Exception:
        return {"price": None}


def update_all_prices():
    symbols = set()
    for section in st.session_state.data.values():
        for row in section:
            symbols.add(row["yahoo"])
    symbols = sorted(symbols)
    progress = st.progress(0.0, text="Fetching prices...")
    for i, sym in enumerate(symbols):
        st.session_state.prices[sym] = fetch_one(sym)
        progress.progress((i + 1) / len(symbols), text=f"Fetched {sym}")
    progress.empty()
    st.session_state.last_updated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_price(yahoo_symbol):
    return st.session_state.prices.get(yahoo_symbol, {}).get("price")


# ---------- formatting ----------

def fmt(n, decimals=2):
    if n is None:
        return "—"
    return f"{n:,.{decimals}f}"


def fmt_signed(n, decimals=0):
    if n is None:
        return "—"
    s = f"{abs(n):,.{decimals}f}"
    return f"({s})" if n < 0 else s


# ---------- HTML table builders ----------

def td(text, bg=None, bold=False, align="right", color=None):
    style = f"padding:5px 8px; border:1px solid #ddd; text-align:{align};"
    if bg:
        style += f"background:{bg};"
    if bold:
        style += "font-weight:600;"
    if color:
        style += f"color:{color};"
    return f'<td style="{style}">{text}</td>'


def render_currency_group(currency, rows):
    decimals = 0 if currency == "IDR" else 2
    header_cells = (
        td("Ticker", bold=True, align="left")
        + td("Name", bold=True, align="left")
        + td("Buy Price", bg=C_BUY, bold=True)
        + td("Qty", bg=C_BUY, bold=True)
        + td("Total Purchased", bg=C_BUY, bold=True)
        + td("Current Price", bg=C_CUR, bold=True)
        + td("Total Value", bg=C_CUR, bold=True)
        + td("Profit/(Loss)", bg=C_PL, bold=True)
        + td("Target Price", bg=C_TGT, bold=True)
        + td("Total Value", bg=C_TGT, bold=True)
        + td("Profit/(Loss)", bg=C_TGT, bold=True)
    )
    body_rows = ""
    tot_purch = tot_val = tot_pl = tot_tval = tot_tpl = 0.0
    any_price = False
    for row in rows:
        mult = lot_mult(currency)
        qty = float(row.get("qty") or 0)
        buy = float(row.get("buy") or 0)
        target = float(row.get("target") or 0)
        current = get_price(row["yahoo"])
        purchased = buy * qty * mult
        value = (current or 0) * qty * mult if current is not None else None
        pl = (value - purchased) if value is not None else None
        tval = target * qty * mult
        tpl = tval - purchased
        tot_purch += purchased
        tot_tval += tval
        tot_tpl += tpl
        if value is not None:
            any_price = True
            tot_val += value
            tot_pl += pl
        qty_label = f"{qty:,.0f} lots" if currency == "IDR" else f"{qty:,.0f}"
        pl_color = "#C00000" if (pl is not None and pl < 0) else ("#2E7D32" if pl is not None else None)
        body_rows += (
            "<tr>"
            + td(row["ticker"], align="left")
            + td(row.get("name", "") or "—", align="left")
            + td(fmt(buy, 2), bg=C_BUY)
            + td(qty_label, bg=C_BUY)
            + td(fmt(purchased, decimals), bg=C_BUY)
            + td(fmt(current, 2) if current is not None else "— (update)", bg=C_CUR)
            + td(fmt(value, decimals) if value is not None else "—", bg=C_CUR)
            + td(fmt_signed(pl, decimals) if pl is not None else "—", bg=C_PL, color=pl_color)
            + td(fmt(target, 2), bg=C_TGT)
            + td(fmt(tval, decimals), bg=C_TGT)
            + td(fmt_signed(tpl, decimals), bg=C_TGT)
            + "</tr>"
        )

    subtotal_row = (
        "<tr>"
        + td(f"Total ({currency})", bold=True, align="left", bg=C_SUBTOTAL)
        + td("", bg=C_SUBTOTAL)
        + td("", bg=C_SUBTOTAL)
        + td("", bg=C_SUBTOTAL)
        + td(fmt(tot_purch, decimals), bold=True, bg=C_SUBTOTAL)
        + td("", bg=C_SUBTOTAL)
        + td(fmt(tot_val, decimals) if any_price else "—", bold=True, bg=C_SUBTOTAL)
        + td(fmt_signed(tot_pl, decimals) if any_price else "—", bold=True, bg=C_SUBTOTAL)
        + td("", bg=C_SUBTOTAL)
        + td(fmt(tot_tval, decimals), bold=True, bg=C_SUBTOTAL)
        + td(fmt_signed(tot_tpl, decimals), bold=True, bg=C_SUBTOTAL)
        + "</tr>"
    )

    equiv_row = ""
    if currency == "HKD":
        rate = st.session_state.hkd_rate
        equiv_row = (
            "<tr>"
            + td("Equiv. in USD", bold=True, align="left", bg=C_GRANDTOTAL)
            + td("", bg=C_GRANDTOTAL) + td("", bg=C_GRANDTOTAL) + td("", bg=C_GRANDTOTAL)
            + td(fmt(tot_purch / rate, 0), bold=True, bg=C_GRANDTOTAL)
            + td("", bg=C_GRANDTOTAL)
            + td(fmt(tot_val / rate, 0) if any_price else "—", bold=True, bg=C_GRANDTOTAL)
            + td(fmt_signed(tot_pl / rate, 0) if any_price else "—", bold=True, bg=C_GRANDTOTAL)
            + td("", bg=C_GRANDTOTAL)
            + td(fmt(tot_tval / rate, 0), bold=True, bg=C_GRANDTOTAL)
            + td(fmt_signed(tot_tpl / rate, 0), bold=True, bg=C_GRANDTOTAL)
            + "</tr>"
        )

    table = f"""
    <table style="border-collapse:collapse; width:100%; font-size:13px; margin-bottom:14px;">
      <tr>{header_cells}</tr>
      {body_rows}
      {subtotal_row}
      {equiv_row}
    </table>
    """
    return table


def render_sell_section(section_key, title):
    rows = st.session_state.data[section_key]
    st.markdown(f"### {title}")
    if not rows:
        st.info("No positions here.")
        return
    for currency in ["USD", "HKD", "IDR"]:
        group = [r for r in rows if r.get("currency", "USD") == currency]
        if not group:
            continue
        st.markdown(f"**{currency}**")
        st.markdown(render_currency_group(currency, group), unsafe_allow_html=True)


def render_buy_group(currency, rows):
    decimals = 0 if currency == "IDR" else 2
    header_cells = (
        td("Ticker", bold=True, align="left")
        + td("Name", bold=True, align="left")
        + td("Current Price", bg=C_CUR, bold=True)
        + td("Target Buy Price", bg=C_TGT, bold=True)
        + td("Distance to Target", bg=C_TGT, bold=True)
    )
    body_rows = ""
    for row in rows:
        current = get_price(row["yahoo"])
        target = float(row.get("target") or 0)
        dist = ((current - target) / target * 100) if (current is not None and target) else None
        dist_color = "#2E7D32" if (dist is not None and dist <= 0) else ("#C00000" if dist is not None else None)
        body_rows += (
            "<tr>"
            + td(row["ticker"], align="left")
            + td(row.get("name", "") or "—", align="left")
            + td(fmt(current, 2) if current is not None else "— (update)", bg=C_CUR)
            + td(fmt(target, decimals), bg=C_TGT)
            + td(f"{dist:+.1f}%" if dist is not None else "—", bg=C_TGT, color=dist_color)
            + "</tr>"
        )
    return f"""
    <table style="border-collapse:collapse; width:100%; font-size:13px; margin-bottom:14px;">
      <tr>{header_cells}</tr>
      {body_rows}
    </table>
    """


def render_buy_section(section_key, title):
    rows = st.session_state.data[section_key]
    st.markdown(f"### {title}")
    if not rows:
        st.info("No positions here.")
        return
    for currency in ["USD", "IDR"]:
        group = [r for r in rows if r.get("currency", "USD") == currency]
        if not group:
            continue
        st.markdown(f"**{currency}**")
        st.markdown(render_buy_group(currency, group), unsafe_allow_html=True)


# ---------- page ----------

st.set_page_config(page_title="Positions", layout="wide")

# ---------- password gate ----------
# The real password lives in Streamlit Cloud's "Secrets" (Settings > Secrets),
# never in this file or in GitHub. Locally, put it in .streamlit/secrets.toml
# (a file that should NOT be committed to git):
#     app_password = "your-password-here"


def check_password():
    def password_entered():
        import hmac
        correct = st.secrets.get("app_password", None)
        if correct is not None and hmac.compare_digest(st.session_state.get("pw_input", ""), correct):
            st.session_state["password_correct"] = True
            st.session_state.pop("pw_input", None)
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    st.title("Positions")
    st.text_input("Password", type="password", key="pw_input", on_change=password_entered)
    if "password_correct" in st.session_state and not st.session_state["password_correct"]:
        st.error("Incorrect password.")
    return False


if not check_password():
    st.stop()

st.title("Positions")
st.caption("Mirrors your Excel workbook: grouped by currency, with subtotals and HKD→USD equivalents.")

top1, top2, top3 = st.columns([1, 1, 3])
with top1:
    if st.button("🔄 Update Prices", type="primary", use_container_width=True):
        update_all_prices()
with top2:
    st.session_state.hkd_rate = st.number_input("USD/HKD rate", min_value=1.0, value=st.session_state.hkd_rate, step=0.01, format="%.4f")
if st.session_state.last_updated:
    st.caption(f"Last updated: {st.session_state.last_updated}")
else:
    st.caption("Prices not fetched yet — click **Update Prices** above.")

tab1, tab2 = st.tabs(["📉 Existing Portfolio", "🧩 New Restructuring"])

with tab1:
    render_sell_section("to_sell", "To Sell")
    render_sell_section("to_reduce", "To Reduce")
    render_sell_section("hold", "Hold")

with tab2:
    render_buy_section("build_core", "Build Core Portfolio — To Buy")
    render_buy_section("tactical", "Tactical — To Buy")
    render_buy_section("trading", "Trading — To Buy")

st.divider()

# ---------- manage positions (add / remove) ----------
st.markdown("### ⚙️ Manage Positions")

manage_tab1, manage_tab2 = st.tabs(["Add", "Remove"])

with manage_tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Add to an existing-portfolio section**")
        section = st.selectbox("Section", ["to_sell", "to_reduce", "hold"], key="add_sell_section")
        ticker = st.text_input("Ticker (Yahoo format, e.g. 0700.HK or BBCA.JK)", key="add_sell_ticker")
        name = st.text_input("Name (optional)", key="add_sell_name")
        currency = st.selectbox("Currency", ["USD", "HKD", "IDR"], key="add_sell_currency")
        qty_label = "Quantity (lots)" if currency == "IDR" else "Quantity (shares)"
        qty = st.number_input(qty_label, min_value=0.0, step=1.0, key="add_sell_qty")
        buy = st.number_input("Buy price", min_value=0.0, step=1.0, key="add_sell_buy")
        target = st.number_input("Target sell price", min_value=0.0, step=1.0, key="add_sell_target")
        if st.button("Add position", key="add_sell_btn"):
            if ticker:
                st.session_state.data[section].append({
                    "ticker": ticker.strip().upper(), "name": name, "yahoo": ticker.strip().upper(),
                    "currency": currency, "buy": buy, "qty": qty, "target": target,
                })
                save_data(st.session_state.data)
                st.rerun()
    with col2:
        st.markdown("**Add to a to-buy section**")
        section2 = st.selectbox("Section", ["build_core", "tactical", "trading"], key="add_buy_section")
        ticker2 = st.text_input("Ticker (Yahoo format, e.g. AAPL or ANTM.JK)", key="add_buy_ticker")
        name2 = st.text_input("Name (optional)", key="add_buy_name")
        currency2 = st.selectbox("Currency", ["USD", "IDR"], key="add_buy_currency")
        target2 = st.number_input("Target buy price", min_value=0.0, step=1.0, key="add_buy_target")
        if st.button("Add to watchlist", key="add_buy_btn"):
            if ticker2:
                st.session_state.data[section2].append({
                    "ticker": ticker2.strip().upper(), "name": name2, "yahoo": ticker2.strip().upper(),
                    "currency": currency2, "target": target2,
                })
                save_data(st.session_state.data)
                st.rerun()

with manage_tab2:
    all_sections = ["to_sell", "to_reduce", "hold", "build_core", "tactical", "trading"]
    section_to_edit = st.selectbox("Section", all_sections, key="remove_section")
    rows = st.session_state.data[section_to_edit]
    if not rows:
        st.info("Nothing to remove here.")
    else:
        options = [f"{i}: {r['ticker']} ({r.get('currency','USD')}) — {r.get('name','')}" for i, r in enumerate(rows)]
        choice = st.selectbox("Position", options, key="remove_choice")
        if st.button("Remove this position", key="remove_btn"):
            idx = int(choice.split(":")[0])
            rows.pop(idx)
            save_data(st.session_state.data)
            st.rerun()

st.caption(f"Data file: {DATA_FILE}")

# v_3.2
import warnings
warnings.simplefilter("ignore")

import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import html as _html

EXCEL_FILE = "Holdings.xlsx"
SHEET_NAME = "Holdings"

# ---------- Helpers ----------
def norm_ticker(t: str, market: str) -> str:
    t = str(t).strip().upper()
    m = (market or "ASX").strip().upper()
    if "." in t:
        return t
    return f"{t}.AX" if m == "ASX" else t

def get_company_name(ticker_yf: str) -> str:
    try:
        info = yf.Ticker(ticker_yf).info or {}
        return info.get("longName") or info.get("shortName") or ticker_yf
    except Exception:
        return ticker_yf

def get_price_native(ticker_yf: str) -> float | None:
    tk = yf.Ticker(ticker_yf)
    # 1) fast_info
    try:
        fi = getattr(tk, "fast_info", None)
        if fi and getattr(fi, "last_price", None):
            return float(fi.last_price)
    except Exception:
        pass
    # 2) intraday (delayed)
    try:
        h1 = tk.history(period="1d", interval="1m")
        if not h1.empty and "Close" in h1:
            return float(h1["Close"].iloc[-1])
    except Exception:
        pass
    # 3) recent daily
    try:
        h2 = tk.history(period="5d", interval="1d")
        if not h2.empty and "Close" in h2:
            return float(h2["Close"].iloc[-1])
    except Exception:
        pass
    # 4) info fallbacks
    try:
        info = tk.info or {}
        for k in ("regularMarketPrice", "currentPrice", "previousClose"):
            v = info.get(k)
            if v is not None:
                return float(v)
    except Exception:
        pass
    return None

def get_aud_per_usd() -> float:
    try:
        fx = yf.download("AUDUSD=X", period="5d", interval="1d", progress=False)
        if not fx.empty and "Close" in fx:
            usd_per_aud = float(fx["Close"].iloc[-1])  # USD per 1 AUD
            return (1.0 / usd_per_aud) if usd_per_aud else 1.5  # AUD per 1 USD
    except Exception:
        pass
    return 1.5

def get_daily_hist_400d(ticker_yf: str) -> pd.DataFrame | None:
    try:
        hist = yf.download(ticker_yf, period="400d", interval="1d", progress=False, auto_adjust=False)
        return hist if not hist.empty else None
    except Exception:
        return None

def _get_close_series(hist: pd.DataFrame) -> pd.Series | None:
    if hist is None or hist.empty:
        return None
    cols = hist.columns
    # MultiIndex support
    if hasattr(cols, "levels"):
        for label in ("Close", "Adj Close", "close", "adjclose"):
            if label in cols.levels[0]:
                df = hist[label].dropna()
                if df.empty:
                    continue
                if isinstance(df, pd.DataFrame):
                    s = df.iloc[:, 0].dropna()
                    return s if not s.empty else None
                if isinstance(df, pd.Series):
                    return df if not s.empty else None
    # Single-level
    for label in ("Close", "Adj Close", "close", "adjclose"):
        if label in cols:
            s = hist[label].dropna()
            return s if not s.empty else None
    return None

def _window_return_from_series(s: pd.Series, months: int, min_days_required: int) -> float | None:
    if s is None or s.empty:
        return None
    s = s.sort_index()
    last = float(s.iloc[-1])
    if pd.isna(last):
        return None
    target_dt = datetime.now() - timedelta(days=int(months * 30.5))
    idx = s.index.searchsorted(pd.Timestamp(target_dt))
    if idx >= len(s):
        idx = len(s) - 1
    ref = float(s.iloc[idx])
    days_span = (s.index[-1] - s.index[0]).days
    if days_span < min_days_required or ref == 0 or pd.isna(ref):
        return None
    return (last - ref) / ref

def compute_individual_perf_from_hist(hist: pd.DataFrame):
    s = _get_close_series(hist)
    r1  = _window_return_from_series(s, 1, 20)
    r6  = _window_return_from_series(s, 6, 120)
    r12 = _window_return_from_series(s, 12, 250)

    def to_str(r): return "" if r is None else f"{r*100:.1f}%"
    return to_str(r1), to_str(r6), to_str(r12), r1, r6, r12

# ---------- Main ----------
def main():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    if "Ticker" not in df.columns or "Quantity" not in df.columns:
        print("Your sheet must have Ticker, Quantity (optional Market).")
        return
    if "Market" not in df.columns:
        df["Market"] = "ASX"

    df["Ticker"] = df["Ticker"].astype(str).str.strip().str.upper()
    df["Market"] = df["Market"].astype(str).str.strip().str.upper()
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    df["TickerYF"] = [norm_ticker(t, m) for t, m in zip(df["Ticker"], df["Market"])]

    aud_per_usd = get_aud_per_usd()
    names, native_price, price_aud = {}, {}, {}
    perf1_str, perf6_str, perf12_str = {}, {}, {}
    perf1_num, perf6_num, perf12_num = {}, {}, {}

    # ---- NEW: in-memory cache for 400d history per ticker_yf
    hist_cache: dict[str, pd.DataFrame | None] = {}

    # Pre-fetch histories once per unique YF ticker
    for t_yf in df["TickerYF"].unique():
        hist_cache[t_yf] = get_daily_hist_400d(t_yf)

    for t_plain, t_yf, mkt in zip(df["Ticker"], df["TickerYF"], df["Market"]):
        nm = get_company_name(t_yf)
        px = get_price_native(t_yf)
        names[t_plain] = nm
        if px is None:
            continue
        native_price[t_plain] = px
        price_aud[t_plain] = px * aud_per_usd if mkt == "US" else px

        hist = hist_cache.get(t_yf)
        p1s, p6s, p12s, p1n, p6n, p12n = compute_individual_perf_from_hist(hist)
        perf1_str[t_plain], perf6_str[t_plain], perf12_str[t_plain] = p1s, p6s, p12s
        perf1_num[t_plain], perf6_num[t_plain], perf12_num[t_plain] = p1n, p6n, p12n

    df["Name"] = df["Ticker"].map(names.get)
    df["PriceNative"] = df["Ticker"].map(native_price.get)
    df["PriceAUD"] = df["Ticker"].map(price_aud.get)
    df["MarketValueAUD"] = df["Quantity"] * df["PriceAUD"]
    df["Perf1M"] = df["Ticker"].map(perf1_str.get)
    df["Perf6M"] = df["Ticker"].map(perf6_str.get)
    df["Perf12M"] = df["Ticker"].map(perf12_str.get)

    # --- Actual performance (AvgCost is native currency per share: USD for US, AUD for ASX) ---
    if "AvgCost" in df.columns:
        df["AvgCost"] = pd.to_numeric(df["AvgCost"], errors="coerce")

        def cost_to_aud(row):
            if pd.isna(row["AvgCost"]):
                return pd.NA
            if row["Market"] == "US":
                return row["AvgCost"] * aud_per_usd  # USD -> AUD
            return row["AvgCost"]  # already AUD for ASX

        df["AvgCostAUD"] = df.apply(cost_to_aud, axis=1)
        df["ActualPerfNum"] = (df["PriceAUD"] - df["AvgCostAUD"]) / df["AvgCostAUD"]
        df["ActualPerf"] = df["ActualPerfNum"].apply(lambda r: "" if pd.isna(r) else f"{r*100:.1f}%")
    else:
        df["AvgCostAUD"] = pd.NA
        df["ActualPerfNum"] = pd.NA
        df["ActualPerf"] = ""

    shown = df.dropna(subset=["PriceAUD"]).copy()
    shown.sort_values("MarketValueAUD", ascending=False, inplace=True)

    total = shown["MarketValueAUD"].sum() if not shown.empty else 0.0
    ts = datetime.now().strftime("%d %b %Y %H:%M")
    fx_str = f"{aud_per_usd:.4f}"

    # Hypothetical portfolio returns (weights = start-of-window MV)
    def agg_with_start_weights(shown_df: pd.DataFrame, ret_dict: dict, months: int):
        usable = []
        total_start_mv = 0.0
        now_ts = datetime.now()
        target_dt = now_ts - timedelta(days=int(months * 30.5))

        for _, r in shown_df.iterrows():
            ticker = r["Ticker"]
            ret = ret_dict.get(ticker)
            if ret is None:
                continue

            t_yf = r.get("TickerYF")
            if not t_yf:
                continue
            hist = hist_cache.get(t_yf)  # reuse cached
            s = _get_close_series(hist)
            if s is None or s.empty:
                continue

            idx = s.index.searchsorted(pd.Timestamp(target_dt))
            if idx >= len(s):
                idx = len(s) - 1
            try:
                p_start = float(s.iloc[idx])
            except Exception:
                continue
            if pd.isna(p_start) or p_start == 0:
                continue

            p_start_aud = p_start * aud_per_usd if r["Market"] == "US" else p_start
            mv_start = float(r["Quantity"]) * p_start_aud
            if mv_start <= 0:
                continue

            usable.append((mv_start, ret))
            total_start_mv += mv_start

        if not usable or total_start_mv <= 0:
            return None
        return sum(mv * ret for mv, ret in usable) / total_start_mv

    port_r1 = agg_with_start_weights(shown, perf1_num, 1) if not shown.empty else None
    port_r6 = agg_with_start_weights(shown, perf6_num, 6) if not shown.empty else None
    port_r12 = agg_with_start_weights(shown, perf12_num, 12) if not shown.empty else None

    # --- Portfolio Actual Performance (from AvgCost) ---
    if not shown.empty:
        have_cost = shown.dropna(subset=["AvgCostAUD"]).copy()
        if have_cost.empty:
            port_actual = None
        else:
            have_cost["CostAUD"] = have_cost["AvgCostAUD"] * have_cost["Quantity"]
            total_cost_aud = have_cost["CostAUD"].sum()
            total_mv_with_cost = have_cost["MarketValueAUD"].sum()
            port_actual = ((total_mv_with_cost - total_cost_aud) / total_cost_aud) if total_cost_aud and total_cost_aud > 0 else None
    else:
        port_actual = None

    def fmt_badge_val(r):
        return "-" if r is None else f"{r*100:.1f}%"

    # Console (unchanged)
    if not shown.empty:
        print(f"\nTOTAL PORTFOLIO VALUE: {total:,.2f} AUD\n")
        print(f"Actual Portfolio Performance: {fmt_badge_val(port_actual)}")
        print(f"Hypothetical Portfolio Performance")
        print(f"  1M: {fmt_badge_val(port_r1)}   6M: {fmt_badge_val(port_r6)}   12M: {fmt_badge_val(port_r12)}\n")

        for _, r in shown.iterrows():
            ccy = "USD" if r["Market"] == "US" else "AUD"
            p1, p6, p12 = r["Perf1M"] or "-", r["Perf6M"] or "-", r["Perf12M"] or "-"
            pact = r["ActualPerf"] or "-"
            print(
                f"{r['Name']} ({r['Ticker']}): {int(r['Quantity'])} × {r['PriceAUD']:,.2f} AUD "
                f"(native {r['PriceNative']:.2f} {ccy}) = {r['MarketValueAUD']:,.2f} AUD "
                f"| 1M {p1} | 6M {p6} | 12M {p12} | Actual {pact}"
            )
    else:
        print("No priced rows. Check tickers/market/network and try again.")

    # ---------- HTML (navy, mobile-first) ----------
    def perf_span(s: str) -> str:
        if not s or s == "-": return '<span class="perf neu">-</span>'
        try:
            val = float(s.replace("%",""))
        except Exception:
            return f'<span class="perf neu">{s}</span>'
        cls = "pos" if val > 0 else ("neg" if val < 0 else "neu")
        return f'<span class="perf {cls}">{s}</span>'

    cards = []
    if not shown.empty:
        for _, r in shown.iterrows():
            # Escape user-facing strings
            name = _html.escape(str(r["Name"]))
            ticker = _html.escape(str(r["Ticker"]))

            qty = int(r["Quantity"])
            price_aud = f"{r['PriceAUD']:,.2f}"
            value_aud = f"{r['MarketValueAUD']:,.2f}"
            native_ccy = "USD" if r["Market"] == "US" else "AUD"
            native_price = f"{r['PriceNative']:.2f}"
            p1 = perf_span(r["Perf1M"] or "-")
            p6 = perf_span(r["Perf6M"] or "-")
            p12 = perf_span(r["Perf12M"] or "-")
            pact = perf_span(r["ActualPerf"] or "-")

            card = f"""
        <div class="card">
          <div class="row1">
            <div class="titlebox">
              <div class="name">{name}</div>
              <div class="value value-mobile">$ {value_aud}</div>
            </div>
            <div class="ticker">{ticker}</div>
          </div>
          <div class="row2">
            <div class="left">
              <div class="metric"><span class="label">Qty</span><span class="val">{qty}</span></div>
              <div class="metric"><span class="label">Price (AUD)</span><span class="val">$ {price_aud}</span></div>
              <div class="metric sub"><span class="label">Native</span><span class="val">{native_price} {native_ccy}</span></div>
            </div>
            <div class="right">
              <div class="value value-desktop">$ {value_aud}</div>
              <div class="perfs">
                <div class="chip">1M {p1}</div>
                <div class="chip">6M {p6}</div>
                <div class="chip">12M {p12}</div>
                <div class="chip">SI {pact}</div>
              </div>
            </div>
          </div>
        </div>
            """
            cards.append(card)

    def badge(val):
        if val is None: return '<span class="badge neu">-</span>'
        cls = "pos" if val > 0 else ("neg" if val < 0 else "neu")
        return f'<span class="badge {cls}">{val*100:.1f}%</span>'

    total_str = f"{total:,.2f}"
    fx_txt = f"USD→AUD {fx_str}"
    updated = ts

    # If no cards, show graceful panel
    holdings_block = ''.join(cards) if cards else """
      <div class="panel" role="status" aria-live="polite">
        No priced rows. Check tickers/market/network and try again.
      </div>
    """

    html = f"""<!doctype html>
    <html lang="en">
    <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
    <title>Portfolio</title>
    <link rel="apple-touch-icon" href="icon.png">
    <style>
    :root {{
        --navy: #0b1530;
        --navy-2: #0f1d49;
        --navy-3: #132866;
        --card: #101b3f;
        --border: rgba(255,255,255,0.12); /* bumped from 0.08 for contrast */
        --text: #eef2ff;
        --muted: #b7c1e3;
        --pos: #22c55e;
        --neg: #ef4444;
        --neu: #9aa3b2;
        --accent: #60a5fa;
    }}
    * {{ box-sizing: border-box; }}
    html {{
        margin:0; padding:0;
        background-color: var(--navy); /* fallback */
    }}
    body {{
        margin:0; padding:0;
        background: linear-gradient(
        180deg,
        var(--navy) 0%,
        var(--navy-2) 80%,
        var(--navy-2) 100%
        );
        color:var(--text);
        font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,'Noto Sans',sans-serif;
    }}

  .wrap {{ max-width: 760px; margin: 0 auto; padding: 16px 14px 28px; }}

  /* big centered total */
  .totalbox {{ text-align:center; padding:18px 12px 10px; margin-bottom:8px; }}
  .totalbox .label {{ color:var(--text); font-size:1.35rem; font-weight:600; letter-spacing:.06em; text-transform:uppercase; }}
  .totalbox .value {{ font-size:2.35rem; font-weight:800; margin-top:4px; }}

  .meta {{ color:var(--muted); text-align:center; font-size:.9rem; margin:6px 0 14px; }}

  /* panels */
  .panel, .hypo {{
    background:var(--card); border:1px solid var(--border); border-radius:14px;
    padding:12px; margin:10px 2px 14px; box-shadow:0 10px 24px rgba(0,0,0,.25);
  }}
  .panel {{ display:flex; align-items:center; justify-content:space-between; gap:10px; }}
  .panel .title {{ font-weight:700; font-size:1rem; }}
  .panel .badge {{ margin-left:auto; }}

  /* hypothetical badges */
  .hypo .title {{ font-weight:700; margin-bottom:8px; font-size:1rem; }}
  .badges {{ display:flex; gap:10px; justify-content:center; flex-wrap:wrap; }}
  .item {{ display:flex; align-items:center; gap:8px; }}
  .item .label {{ color:var(--muted); font-size:.88rem; }}
  .badge {{
    display:inline-block; min-width:70px; text-align:center;
    border-radius:999px; padding:6px 10px; font-weight:800; color:#fff;
  }}
  .badge.pos {{ background:var(--pos); }}
  .badge.neg {{ background:var(--neg); }}
  .badge.neu {{ background:var(--neu); }}

  /* holding cards */
  .section-title {{ color:var(--muted); text-transform:uppercase; letter-spacing:.12em; font-size:.8rem; margin:8px 2px; }}
  .card {{
    background:var(--card); border:1px solid var(--border); border-radius:16px; padding:12px;
    margin:10px 0; box-shadow:0 10px 24px rgba(0,0,0,.25);
  }}
  .row1 {{ display:flex; justify-content:space-between; align-items:center; gap:8px; }}
  .titlebox {{ display:flex; flex-direction:column; gap:2px; }}
  .name {{ font-weight:700; font-size:1.0rem; }}
  .ticker {{ color:var(--muted); font-size:.84rem; padding:2px 8px; border:1px solid var(--border); border-radius:999px; }}
  .row2 {{ display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-top:8px; }}
  .metric {{ display:flex; align-items:center; justify-content:space-between; padding:8px 0; border-bottom:1px dashed var(--border); }}
  .metric.sub {{ border-bottom:none; color:var(--muted); padding-top:8px; }}
  .label {{ color:var(--muted); font-size:.88rem; }}
  .val {{ font-weight:700; }}
  .right {{ text-align:right; }}
  .value {{ font-size:1.06rem; font-weight:800; color:rgba(233,131,0,1); }}
  .value-mobile {{ display:none; }} /* hidden by default (desktop/tablet) */
  .value-desktop {{ display:block; }}

  .perfs {{ display:flex; gap:10px; justify-content:flex-end; margin-top:6px; flex-wrap:wrap; }}

  /* perf chips */
  .chip {{
    display:inline-flex; align-items:center; justify-content:center; gap:6px;
    background:#0c1840; border:1px solid var(--border); border-radius:999px;
    height:36px; min-width:110px; padding:0 12px;
    font-size:.88rem; font-weight:700; color:var(--muted); white-space:nowrap;
  }}

  .perf {{ font-weight:800; }}
  .perf.pos {{ color:var(--pos); }}
  .perf.neg {{ color:var(--neg); }}
  .perf.neu {{ color:var(--neu); }}

  /* Reduced motion */
  @media (prefers-reduced-motion: reduce) {{
    * {{ animation: none !important; transition: none !important; }}
  }}

@media (max-width:480px){{
  .totalbox .label{{font-size:1.2rem;}}
  .totalbox .value{{font-size:2.05rem;}}
  .row1{{align-items:flex-start;}}
  .value-mobile{{display:block;}}
  .value-desktop{{display:none;}}
  .row2{{grid-template-columns:1fr 1fr;gap:8px;margin-top:6px;}}
  .left{{padding-right:6px;}}
  .right{{text-align:initial;}}

  /* HOLDINGS chips unchanged */
  .perfs{{display:grid;grid-template-columns:1fr 1fr;gap:8px;justify-items:stretch;}}
  .chip{{display:grid;grid-template-columns:1fr 50px;align-items:center;justify-content:space-between;width:100%;min-width:auto;height:36px;font-size:.9rem;padding:0 10px;font-weight:700;}}

  .card{{padding:10px;margin:8px 0;border-radius:14px;}}
  .row1{{margin-bottom:4px;}}
  .titlebox{{gap:1px;}}

  /* ORIGINAL spacing */
  .metric{{padding:4px 0; border-bottom:1px dashed var(--border);}}
  .metric.sub{{border-bottom:none; color:var(--muted); padding-top:8px;}}

  .panel{{gap:8px;}}

  /* Align Hypothetical badges: 1M left, 6M center, 12M right */
  .hypo .badges{{display:flex;justify-content:space-between;gap:10px;width:100%;}}
  .hypo .badges .item{{flex:1;display:flex;align-items:center;gap:8px;}}
  .hypo .badges .item:first-child{{justify-content:flex-start;text-align:left;}}
  .hypo .badges .item:nth-child(2){{justify-content:center;text-align:center;}}
  .hypo .badges .item:last-child{{justify-content:flex-end;text-align:right;}}

  /* Actual Portfolio Performance badge */
  .panel .badge{{
    min-width:113px;
    padding:6px 10px;
    border-radius:999px;
    font-weight:800;
    font-size:1rem;
    line-height:normal;
    height:auto;
    text-align:center;
    margin-right:0;
  }}
}}

  @media (min-width: 720px) {{
    .wrap {{ padding:22px 20px 40px; }}
  }}
</style>
</head>
<body>
  <div class="wrap">
    <div class="totalbox" role="heading" aria-level="1">
      <div class="label">Total Portfolio Value</div>
      <div class="value">$ {total_str}</div>
    </div>
    <div class="meta">{fx_txt} • Updated {updated}</div>

    <!-- Actual Portfolio Performance (single-line badge) -->
    <div class="panel">
      <div class="title" role="heading" aria-level="2">Since Inception Performance</div>
      {badge(port_actual)}
    </div>

    <div class="hypo">
      <div class="title" role="heading" aria-level="2">Hypothetical Past Performance</div>
      <div class="badges">
        <div class="item"><div class="label">1M</div>{badge(port_r1)}</div>
        <div class="item"><div class="label">6M</div>{badge(port_r6)}</div>
        <div class="item"><div class="label">12M</div>{badge(port_r12)}</div>
      </div>
    </div>

    <div class="section-title" role="heading" aria-level="2">Holdings</div>
    {holdings_block}
  </div>
</body>
</html>"""

    with open("report.html", "w", encoding="utf-8") as f:
        f.write(html)

    print("\nSaved HTML report → report.html")

if __name__ == "__main__":
    main()

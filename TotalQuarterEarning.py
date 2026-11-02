# TotalQuarterEarning.py v4.0 - GitHub-safe, no YF ban
import os, json, time, re, threading, requests, random
os.environ["YFINANCE_NO_CACHE"] = "1"
import yfinance as yf, openpyxl, gspread
from datetime import datetime, timezone, timedelta

FINNHUB_API_KEY = os.getenv("FINNHUB_API_KEY") or ""
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID") or ""
SGT = timezone(timedelta(hours=8))
FH_BASE = "https://finnhub.io/api/v1"

# === SET THIS FOR TEST ===
TEST_MODE = True  # True = only Growth tab, False = all 17 tabs
TEST_TAB_NAME = "Growth"  # change to your Growth tab exact name, e.g. "Growth" or "Growth 352"
# =========================

def get_gc():
    secret = os.environ.get("GOOGLE_SERVICE_ACCOUNT")
    if not secret: raise ValueError("GOOGLE_SERVICE_ACCOUNT missing")
    return gspread.service_account_from_dict(json.loads(secret), scopes=["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"])

def is_valid_ticker(v):
    if not v: return False
    s=str(v).strip().upper()
    if len(s)<1 or len(s)>7 or s in ["TICKER","SYMBOL","PRICE"]: return False
    return bool(re.match(r"^[A-Z]{1,5}([.-][A-Z]{1,2})?$", s))

class Limiter:
    def __init__(self,cpm): self.iv=60.0/cpm; self.lock=threading.Lock(); self.nxt=0
    def wait(self):
        with self.lock:
            now=time.monotonic(); at=max(now,self.nxt); self.nxt=at+self.iv
        d=at-time.monotonic()
        if d>0: time.sleep(d)
lim=Limiter(50)

def fh_get(path, params):
    p=dict(params); p["token"]=FINNHUB_API_KEY
    lim.wait()
    for _ in range(3):
        try:
            r=requests.get(FH_BASE+path, params=p, timeout=20)
            if r.status_code==200: return r.json()
            if r.status_code==429: time.sleep(5)
        except: time.sleep(2)
    return None

def get_tickers():
    print(f"TEST_MODE={TEST_MODE} TAB={TEST_TAB_NAME if TEST_MODE else 'ALL'}")
    gc = get_gc()
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    tickers=[]
    for ws in sh.worksheets():
        title=ws.title
        if TEST_MODE and title != TEST_TAB_NAME: continue
        if not TEST_MODE and title in ["Quarterly Data","Surprise Earning Data","Dashboard","DASHBOARD","Index"]: continue
        if not TEST_MODE and "Data (1.9)" in title: continue
        print(f" Scanning {title}...")
        col = ws.col_values(1)
        for v in col[1:]:
            if is_valid_ticker(v): tickers.append(str(v).strip().upper())
    tickers=sorted(set(tickers))
    print(f"Found {len(tickers)} tickers")
    return tickers

def get_price_slow(ticker):
    """Single ticker with retry - GitHub safe"""
    for attempt in range(3):
        try:
            # Random delay to avoid ban
            time.sleep(random.uniform(0.5, 1.5))
            tk = yf.Ticker(ticker)
            hist = tk.history(period="5d", auto_adjust=True)
            if not hist.empty:
                return float(hist['Close'].iloc[-1])
        except Exception as e:
            if "Rate limited" in str(e) or "Too Many" in str(e):
                wait = 10 + attempt*10
                print(f"  YF limited on {ticker}, wait {wait}s")
                time.sleep(wait)
            else:
                time.sleep(1)
    return None

def main():
    start=time.time()
    tickers=get_tickers()
    if not tickers: return
    print(f"Total: {len(tickers)} e.g. {tickers[:10]}")

    earnings={}
    for i,t in enumerate(tickers):
        data=fh_get("/stock/earnings", {"symbol":t})
        if data and isinstance(data,list):
            valid=[x for x in data if x.get("actual") is not None]
            if valid: earnings[t]=sorted(valid, key=lambda x: x.get("period") or "", reverse=True)[0]
        if i%20==0: print(f" Finnhub {i}/{len(tickers)} OK:{len(earnings)}")

    print("Starting YFinance price fetch (slow mode to avoid ban)...")
    price_cache={}
    for i,t in enumerate(tickers):
        price = get_price_slow(t)
        if price: price_cache[t]=price
        if i%10==0: 
            print(f" Price {i}/{len(tickers)} OK:{len(price_cache)}")
            time.sleep(2)  # extra pause every 10

    now_sgt=datetime.now(SGT); now_utc=datetime.now(timezone.utc)
    elapsed=(time.time()-start)/60
    wb_out=openpyxl.Workbook()
    ws=wb_out.active; ws.title="Quarterly Data"
    ws.append(["Ticker","EPS Last","Period","Price","Surprise%","Last Updated"])
    for t in tickers:
        e=earnings.get(t,{}); act=e.get("actual"); est=e.get("estimate")
        sur=round((act-est)/abs(est)*100,2) if act and est not in (None,0) else ""
        ws.append([t, act, e.get("period",""), price_cache.get(t,""), sur, now_sgt.strftime("%Y-%m-%d %H:%M:%S (UTC+8)")])

    ws2=wb_out.create_sheet("Surprise Earning Data")
    ws2.append(["Ticker","Date","EPS Actual","EPS Est","Surprise %","Matched","Source","Flag >70%","Notes","Status","Last Pulled"])
    for t in tickers:
        e=earnings.get(t,{})
        if e:
            sur=round((e.get("actual",0)-e.get("estimate",0))/abs(e.get("estimate",1))*100,2) if e.get("estimate") else ""
            ws2.append([t, e.get("period",""), e.get("actual"), e.get("estimate"), sur, "Y", "Finnhub", "Y" if isinstance(sur,float) and abs(sur)>70 else "N", "", "OK", now_sgt.strftime("%Y-%m-%d %H:%M")])

    ws3=wb_out.create_sheet("Dashboard")
    ws3["A1"].value="RUN STATUS:"; ws3["B1"].value="COMPLETED"; ws3["C1"].value="COMPLETED"; ws3["D1"].value="COMPLETED"
    ws3["A2"].value="DASHBOARD UPDATED:"; ws3["C2"].value=f"Merged: {now_utc.strftime('%Y-%m-%d %H:%M:%S (UTC)')}"
    ws3["D2"].value=f"Merged Surprise: {now_utc.strftime('%Y-%m-%d %H:%M:%S (UTC)')}"
    ws3["A4"].value="TOTAL RUN TIME:"; ws3["C4"].value=f"{elapsed:.1f} min | OK:{len(earnings)} | Price:{len(price_cache)}"
    ws3["D4"].value=f"{len(tickers)} tickers"

    out=f"Watchlist_MERGED_FULL_{now_sgt.strftime('%Y%m%d_%H%M')}.xlsx"
    wb_out.save(out)
    print(f"\n✅ Saved {out} - {elapsed:.1f} min - OK:{len(earnings)}/{len(tickers)} Price:{len(price_cache)}")

if __name__=="__main__": main()

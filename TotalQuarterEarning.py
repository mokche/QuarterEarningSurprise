"""
TotalEarning - Watchlist MERGED to GDrive / Google Sheet
v1.0 - Colab + GitHub ready
"""
import os, time, re, threading, requests
os.environ["YFINANCE_NO_CACHE"] = "1"
import yfinance as yf, openpyxl
from datetime import datetime, timezone, timedelta

# =======================================================
# === EDIT ONLY THIS SECTION - LEAVE REST UNTOUCHED ===
# =======================================================

import os
FINNHUB_API_KEY = os.getenv("FINNHUB_API_KEY") or "PASTE_YOUR_KEY_HERE"
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID") or "PASTE_YOUR_SHEET_ID_HERE"

# Tabs to scan: None = all 17 tabs, ["Growth"] = test only
TABS_TO_SCAN = ["Growth"]
# =======================================================

SGT = timezone(timedelta(hours=8))
FH_BASE = "https://finnhub.io/api/v1"

def is_valid_ticker(v):
    if not v: return False
    s=str(v).strip().upper()
    if len(s)<1 or len(s)>7 or s in ["TICKER","SYMBOL","PRICE"]: return False
    return bool(re.match(r"^[A-Z]{1,5}([.-][A-Z]{1,2})?$", s))

class Limiter:
    def __init__(self,cpm):
        self.iv=60.0/cpm; self.lock=threading.Lock(); self.nxt=0
    def wait(self):
        with self.lock:
            now=time.monotonic(); at=max(now,self.nxt); self.nxt=at+self.iv
        d=at-time.monotonic()
        if d>0: time.sleep(d)
lim = Limiter(55)

def fh_get(path, params):
    p=dict(params); p["token"]=FINNHUB_API_KEY
    lim.wait()
    try:
        r=requests.get(FH_BASE+path, params=p, timeout=20)
        if r.status_code==200: return r.json()
    except: pass
    return None

def main():
    start=time.time()
    print("Loading Watchlist.xlsx...")
    wb=openpyxl.load_workbook("Watchlist.xlsx")
    
    tickers=[]
    for name in wb.sheetnames:
        if name in ["Quarterly Data","Surprise Earning Data","Dashboard","DASHBOARD","Index","Sheet1"]: continue
        if TABS_TO_SCAN and name not in TABS_TO_SCAN: continue
        ws=wb[name]
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if row and is_valid_ticker(row[0]):
                tickers.append(str(row[0]).strip().upper())
    tickers=sorted(set(tickers))
    print(f"Found {len(tickers)} tickers")

    # Earnings
    earnings={}
    for i,t in enumerate(tickers):
        data=fh_get("/stock/earnings", {"symbol":t})
        if data and isinstance(data,list):
            valid=[x for x in data if x.get("actual") is not None]
            if valid:
                earnings[t]=sorted(valid, key=lambda x: x.get("period") or "", reverse=True)[0]
        if i%50==0:
            print(f"  {i}/{len(tickers)} OK:{len(earnings)}")

    # Prices
    price_cache={}
    BATCH=20
    for i in range(0,len(tickers),BATCH):
        chunk=tickers[i:i+BATCH]
        try:
            df=yf.download(chunk, period="5d", threads=True, progress=False, auto_adjust=True, group_by='ticker')
            if df is not None and not df.empty:
                for c in chunk:
                    try:
                        if len(chunk)==1:
                            last=df['Close'].dropna()
                            if not last.empty: price_cache[c]=float(last.iloc[-1])
                        elif c in df.columns.get_level_values(0):
                            last=df[c]['Close'].dropna()
                            if not last.empty: price_cache[c]=float(last.iloc[-1])
                    except: pass
        except: pass
        time.sleep(1)

    # Build merged
    now_sgt=datetime.now(SGT)
    now_str=now_sgt.strftime("%Y-%m-%d %H:%M:%S (UTC+8)")
    now_utc=datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S (UTC)")
    elapsed=(time.time()-start)/60

    for n in ["Quarterly Data","Surprise Earning Data"]:
        if n in wb.sheetnames: wb.remove(wb[n])
    
    ws=wb.create_sheet("Quarterly Data")
    ws.append(["Ticker","EPS Last","Period","Price","Surprise%","Last Updated"])
    for t in tickers:
        e=earnings.get(t,{}); act=e.get("actual"); est=e.get("estimate")
        sur=round((act-est)/abs(est)*100,2) if act and est not in (None,0) else ""
        ws.append([t, act, e.get("period",""), price_cache.get(t,""), sur, now_str])

    ws2=wb.create_sheet("Surprise Earning Data")
    ws2.append(["Ticker","Date","EPS Actual","EPS Est","Surprise %","Matched","Source","Flag >70%","Notes","Status","Last Pulled"])
    for t in tickers:
        e=earnings.get(t,{})
        if e:
            est=e.get("estimate"); act=e.get("actual")
            sur=round((act-est)/abs(est)*100,2) if act and est not in (None,0) else ""
            ws2.append([t, e.get("period",""), act, est, sur, "Y", "Finnhub", "Y" if isinstance(sur,(int,float)) and abs(sur)>70 else "N", "", "OK", now_sgt.strftime("%Y-%m-%d %H:%M")])

    # UPDATE DASHBOARD - ONLY C & D, NEVER B
    dash_name="Dashboard" if "Dashboard" in wb.sheetnames else "DASHBOARD" if "DASHBOARD" in wb.sheetnames else None
    if dash_name:
        dws=wb[dash_name]
        ok=len(earnings); fail=len(tickers)-ok
        # Column C
        dws["C1"].value="COMPLETED"
        dws["C2"].value=f"Merged: {now_utc}"
        dws["C3"].value="Watchlist_MERGED v1.0-TotalEarning"
        dws["C4"].value=f"{elapsed:.1f} min | OK:{ok} | Price:{len(price_cache)} | Failed:{fail} | ETF:0"
        # Column D
        dws["D1"].value="COMPLETED"
        dws["D2"].value=f"Merged Surprise: {now_utc}"
        dws["D3"].value="TotalEarning v1.0-Surprise"
        dws["D4"].value=f"{len(tickers)} tickers | {ok} matched | {fail} failed"
        print(f"Dashboard updated: B untouched ({dws['B2'].value})")

    out=f"Watchlist_MERGED_FULL_{now_sgt.strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(out)
    print(f"Saved {out} - {elapsed:.1f} min - OK:{len(earnings)}/{len(tickers)}")
    # GDrive save happens in Colab via drive.mount, not here

if __name__=="__main__":
    main()

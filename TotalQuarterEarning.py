# TotalQuarterEarning.py v2.0 - GitHub Actions ready
import os, json, time, re, threading, requests
os.environ["YFINANCE_NO_CACHE"] = "1"
import yfinance as yf, openpyxl
from datetime import datetime, timezone, timedelta

FINNHUB_API_KEY = os.getenv("FINNHUB_API_KEY") or "PASTE_KEY_HERE"
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID") or "PASTE_SHEET_ID_HERE"
SGT = timezone(timedelta(hours=8))
FH_BASE = "https://finnhub.io/api/v1"

def is_valid_ticker(v):
    if not v: return False
    s=str(v).strip().upper()
    if len(s)<1 or len(s)>7 or s in ["TICKER","SYMBOL","PRICE","DATE","TIME"]: return False
    return bool(re.match(r"^[A-Z]{1,5}([.-][A-Z]{1,2})?$", s))

class Limiter:
    def __init__(self,cpm): self.iv=60.0/cpm; self.lock=threading.Lock(); self.nxt=0
    def wait(self):
        with self.lock:
            now=time.monotonic(); at=max(now,self.nxt); self.nxt=at+self.iv
        d=at-time.monotonic()
        if d>0: time.sleep(d)
lim=Limiter(55)

def fh_get(path, params):
    p=dict(params); p["token"]=FINNHUB_API_KEY
    lim.wait()
    try:
        r=requests.get(FH_BASE+path, params=p, timeout=20)
        if r.status_code==200: return r.json()
    except: pass
    return None

def get_tickers():
    print(f"Reading Google Sheet: {GOOGLE_SHEET_ID[:10]}***")
    # Try Service Account first (GitHub), then Colab
    gc = None
    try:
        # GitHub path - GOOGLE_CREDENTIALS secret
        creds_json = os.getenv("GOOGLE_CREDENTIALS")
        if creds_json:
            print("Using GOOGLE_CREDENTIALS from GitHub Secrets")
            import gspread
            from google.oauth2.service_account import Credentials
            creds_dict = json.loads(creds_json)
            scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly","https://www.googleapis.com/auth/drive.readonly"]
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            gc = gspread.authorize(creds)
        else:
            # Colab fallback
            from google.colab import auth
            import gspread
            auth.authenticate_user()
            from google.auth import default
            creds,_ = default()
            gc = gspread.authorize(creds)
    except Exception as e:
        print(f"Auth failed: {e}")
        return []

    try:
        sh = gc.open_by_key(GOOGLE_SHEET_ID)
        tickers=[]
        for ws in sh.worksheets():
            if ws.title in ["Quarterly Data","Surprise Earning Data","Dashboard","DASHBOARD","Index","Sheet1"]: continue
            if "Data (1.9)" in ws.title or "Dashboard" in ws.title: continue
            col = ws.col_values(1)
            for v in col[1:]:
                if is_valid_ticker(v): tickers.append(str(v).strip().upper())
        tickers=sorted(set(tickers))
        print(f"Found {len(tickers)} tickers")
        return tickers
    except Exception as e:
        print(f"Sheet read failed: {e}")
        return []

def main():
    start=time.time()
    tickers=get_tickers()
    if not tickers:
        print("❌ 0 tickers - Did you share sheet with service account email?")
        return
    print(f"Total: {len(tickers)} e.g. {tickers[:10]}")

    earnings={}
    for i,t in enumerate(tickers):
        data=fh_get("/stock/earnings", {"symbol":t})
        if data and isinstance(data,list):
            valid=[x for x in data if x.get("actual") is not None]
            if valid: earnings[t]=sorted(valid, key=lambda x: x.get("period") or "", reverse=True)[0]
        if i%50==0: print(f" {i}/{len(tickers)} OK:{len(earnings)}")

    price_cache={}
    for i in range(0,len(tickers),20):
        chunk=tickers[i:i+20]
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
    ws3["A2"].value="DASHBOARD UPDATED:"; ws3["B2"].value="KEEP v5.2.1"; ws3["C2"].value=f"Merged: {now_utc.strftime('%Y-%m-%d %H:%M:%S (UTC)')}"
    ws3["D2"].value=f"Merged Surprise: {now_utc.strftime('%Y-%m-%d %H:%M:%S (UTC)')}"
    ws3["A3"].value="SCRIPT VERSION:"; ws3["C3"].value="TotalEarning v1.0"; ws3["D3"].value="TotalEarning Surprise"
    ws3["A4"].value="TOTAL RUN TIME:"; ws3["C4"].value=f"{elapsed:.1f} min | OK:{len(earnings)} | Price:{len(price_cache)} | Failed:{len(tickers)-len(earnings)}"
    ws3["D4"].value=f"{len(tickers)} tickers | {len(earnings)} matched"

    out=f"Watchlist_MERGED_FULL_{now_sgt.strftime('%Y%m%d_%H%M')}.xlsx"
    wb_out.save(out)
    print(f"\n✅ Saved {out} - OK:{len(earnings)}/{len(tickers)}")

if __name__=="__main__": main()

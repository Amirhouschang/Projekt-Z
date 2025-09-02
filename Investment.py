import yfinance as yf
import openpyxl
from datetime import datetime


tickers = {
    "6AQQ.DE": "B2",     # Amundi Nasdaq 100 EUR (LU1681038243) ✅ 
    "EGLN.L": "B3",      # iShares Physical Gold ETC ✅
    "SCWX.DE": "B4",     # Scalable MSCI AC World (LU2903252349) ✅
    "DFEN.DE": "B5",     # VanEck Defense (Xetra) ✅
    "ASML.AS": "B6",     # ASML Holding N. V. ✅
    "RHM.DE": "B7",      # Rheinmetall AG ✅
    "XMLD.DE": "B8"      # L&G Artificial Intelligence (deutscher Ticker) ✅
}

dateiname = "/Users/amirhoushang/Desktop/Python/Investment_Plan_2025.xlsx"

wb = openpyxl.load_workbook(dateiname)
sheet = wb["Kurse"]

now = datetime.now()
datum = now.strftime("%d.%m.%Y")
uhrzeit = now.strftime("%H:%M")

for ticker, zelle in tickers.items():
    try:
        stock = yf.Ticker(ticker)
        preis = stock.fast_info.last_price
        
        sheet[zelle] = preis
        
        zeile = zelle[1:]  
        sheet[f"D{zeile}"] = datum  
        sheet[f"E{zeile}"] = uhrzeit
        
        print(f"✅ {ticker}: {preis:.2f} EUR (Zeile {zeile} aktualisiert)")
        
    except Exception as e:
        zeile = zelle[1:]
        print(f"❌ Fehler bei {ticker} (Zeile {zeile}): {e}")

wb.save(dateiname)
print("NUR Preise/Datum/Zeit wurden aktualisiert!")
print("Alle anderen Daten blieben unverändert!")
print("Excel-Datei sicher gespeichert.")

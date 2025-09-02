import requests
import openpyxl
from datetime import datetime

url = "https://api.coingecko.com/api/v3/simple/price"
params = {
    "ids": "bitcoin,ethereum,solana,tezos,polkadot,celo",
    "vs_currencies": "eur"
}

response = requests.get(url, params=params)
preise = response.json()

btc = preise["bitcoin"]["eur"]
eth = preise["ethereum"]["eur"]
sol = preise["solana"]["eur"]
xtz = preise["tezos"]["eur"]
dot = preise["polkadot"]["eur"]
celo = preise["celo"]["eur"]

dateiname = "/Users/amirhoushang/Desktop/Meine Z/01_Crypto_Art_Portfolio_2025.xlsx"
wb = openpyxl.load_workbook(dateiname)
sheet = wb["Kurse"]

now = datetime.now()
datum = now.strftime("%d.%m.%Y")
uhrzeit = now.strftime("%H:%M")

sheet["B2"] = btc
sheet["B3"] = eth
sheet["B4"] = sol
sheet["B5"] = xtz
sheet["B9"] = dot
sheet["B10"] = celo

sheet["C2"] = datum
sheet["C3"] = datum
sheet["C4"] = datum
sheet["C5"] = datum
sheet["C9"] = datum
sheet["C10"] = datum

sheet["D2"] = uhrzeit
sheet["D3"] = uhrzeit
sheet["D4"] = uhrzeit
sheet["D5"] = uhrzeit
sheet["D9"] = uhrzeit
sheet["D10"] = uhrzeit

wb.save(dateiname)
print("Preise und Zeit wurden automatisch aktualisiert.")
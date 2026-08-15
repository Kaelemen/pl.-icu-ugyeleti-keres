# -*- coding: utf-8 -*-
"""Ha a mai nap a hónap utolsó napja: elmenti az adott hónapra addig beküldött összes
kérést egy nyomtatható Excel-táblázatba (archivum/ mappa), majd törli azokat a
Firestore-ból, hogy a következő hónap üresen induljon. Nyílt Firestore-szabályokat
és a nyilvános Firebase API-kulcsot használja, service account nélkül."""
import datetime
import os
import sys
import requests
import openpyxl

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"

today = datetime.date.today()
holnap = today + datetime.timedelta(days=1)
kenyszeritett_teszt = os.environ.get("FORCE_TEST") == "1"
if holnap.day != 1 and not kenyszeritett_teszt:
    print(f"Ma ({today}) nem a hónap utolsó napja, nincs teendő.")
    sys.exit(0)

ev, honap = today.year, today.month
if kenyszeritett_teszt and os.environ.get("FORCE_HONAP"):
    ev, honap = int(os.environ["FORCE_EV"]), int(os.environ["FORCE_HONAP"])
print(f"Ma a hónap utolsó napja - archiválás és nullázás: {ev}.{honap:02d}")


def fs_ertek(v):
    if "stringValue" in v:
        return v["stringValue"]
    if "integerValue" in v:
        return int(v["integerValue"])
    if "booleanValue" in v:
        return v["booleanValue"]
    if "doubleValue" in v:
        return v["doubleValue"]
    return None


def doc_dict(doc):
    return {k: fs_ertek(v) for k, v in doc.get("fields", {}).items()}


osszes = []
page_token = None
while True:
    params = {"key": API_KEY, "pageSize": 100}
    if page_token:
        params["pageToken"] = page_token
    r = requests.get(f"{BASE}/keresek", params=params, timeout=30)
    r.raise_for_status()
    data = r.json()
    osszes.extend(data.get("documents", []))
    page_token = data.get("nextPageToken")
    if not page_token:
        break

honap_dokok = []
for doc in osszes:
    d = doc_dict(doc)
    if d.get("ev") == ev and d.get("honap") == honap:
        honap_dokok.append((doc["name"], d))

print(f"{len(honap_dokok)} kérés található {ev}.{honap:02d}-re.")
if not honap_dokok:
    print("Nincs mit archiválni.")
    sys.exit(0)

honap_dokok.sort(key=lambda x: x[1].get("nev", ""))

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Kérések"
ws.append(["Név", "Szabadság", "Nem szeretné", "Szeretné", "8 óra (nappali)",
           "Kért ügy.", "Külsős gyakorlat", "Megjegyzés"])
for _, d in honap_dokok:
    if d.get("nincs_keres"):
        ws.append([d.get("nev", ""), "", "", "", "", "", "", "Nincs kérése"])
    else:
        ws.append([
            d.get("nev", ""), d.get("szabadsag", ""), d.get("nem_szeretne", ""),
            d.get("szeretne", ""), d.get("nyolc_ora_nappal", ""), d.get("kert_ugyeletszam", ""),
            "Igen" if d.get("kulsos_gyakorlat") else "", d.get("megjegyzes", ""),
        ])
for col_letter in "ABCDEFGH":
    ws.column_dimensions[col_letter].width = 20

os.makedirs("archivum", exist_ok=True)
kimenet = f"archivum/kerelek_{ev}_{honap:02d}.xlsx"
wb.save(kimenet)
print(f"Mentve: {kimenet}")

for name, d in honap_dokok:
    if os.environ.get("DRY_RUN") == "1":
        print(f"(PRÓBAFUTTATÁS - nincs törlés) {d.get('nev','?')}")
        continue
    r = requests.delete(f"https://firestore.googleapis.com/v1/{name}", params={"key": API_KEY}, timeout=30)
    print(f"Törölve ({d.get('nev','?')}): {r.status_code}")

print("Kész - a hónap kérései archiválva" + (" (próbafuttatás, törlés nélkül)." if os.environ.get("DRY_RUN") == "1" else " és törölve."))

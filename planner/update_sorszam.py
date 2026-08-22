import json
import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"

UJ_SORREND = ["Kelemen József", "Katona József", "Berek Sarolta", "Gulya Réka", "Kovácsevics József",
              "Róbert Beáta", "Balogh László", "Berkes Tíbor", "Hajnal Csilla", "Kajner Krisztina",
              "Korompai Máté", "Pintér Enikő", "Sorbán József", "Szász Kornélia", "Sztermen Márton",
              "Zöldréti Anikó", "Daku Zsuzsa", "Karczub János", "Lambertus Iván", "Enyedi Marcián", "Rákóczi Réka"]


def fs_ertek(v):
    if "stringValue" in v:
        return v["stringValue"]
    if "integerValue" in v:
        return int(v["integerValue"])
    return None


r = requests.get(f"{BASE}/dolgozok", params={"key": API_KEY, "pageSize": 200}, timeout=30)
r.raise_for_status()
nev_to_docid = {}
for doc in r.json().get("documents", []):
    nev = fs_ertek(doc["fields"].get("nev", {}))
    doc_id = doc["name"].split("/")[-1]
    if nev:
        nev_to_docid[nev] = doc_id

print("Talalt dolgozok:", len(nev_to_docid))

hianyzik = [n for n in UJ_SORREND if n not in nev_to_docid]
if hianyzik:
    print("FIGYELEM - nem talalt nevek:", hianyzik)

for i, nev in enumerate(UJ_SORREND):
    doc_id = nev_to_docid.get(nev)
    if not doc_id:
        continue
    url = f"{BASE}/dolgozok/{doc_id}"
    params = {"key": API_KEY, "updateMask.fieldPaths": "sorszam"}
    body = {"fields": {"sorszam": {"integerValue": str(i)}}}
    resp = requests.patch(url, params=params, json=body, timeout=30)
    if resp.status_code == 200:
        print(f"OK: {nev} -> sorszam={i}")
    else:
        print(f"HIBA {nev}: {resp.status_code} {resp.text}")

import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"


def fs_ertek(v):
    if "stringValue" in v:
        return v["stringValue"]
    return None


r = requests.get(f"{BASE}/jelszavak", params={"key": API_KEY, "pageSize": 200}, timeout=30)
r.raise_for_status()
docs = r.json().get("documents", [])
print(f"Osszesen {len(docs)} bejegyzes:\n")
for doc in docs:
    doc_id = doc["name"].split("/")[-1]
    nev = fs_ertek(doc["fields"].get("nev", {}))
    hossz = len(doc_id)
    hex_e = all(c in "0123456789abcdef" for c in doc_id)
    allapot = "HASH (biztonsagos)" if hossz == 64 and hex_e else "FIGYELEM - NEM HASH FORMATUM"
    print(f"  {nev}: azonosito hossza={hossz}, {allapot}")

import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"


def fs_ertek(v):
    if "stringValue" in v:
        return v["stringValue"]
    if "integerValue" in v:
        return int(v["integerValue"])
    return None


r = requests.get(f"{BASE}/dolgozok", params={"key": API_KEY, "pageSize": 200}, timeout=30)
r.raise_for_status()
lista = []
for doc in r.json().get("documents", []):
    nev = fs_ertek(doc["fields"].get("nev", {}))
    sorszam = fs_ertek(doc["fields"].get("sorszam", {}))
    lista.append((sorszam, nev))
lista.sort(key=lambda x: (x[0] is None, x[0]))
for s, n in lista:
    print(s, n)

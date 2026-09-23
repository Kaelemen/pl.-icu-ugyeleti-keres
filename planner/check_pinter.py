import requests
import json

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"


def fs_ertek(v):
    if "stringValue" in v:
        return v["stringValue"]
    if "integerValue" in v:
        return int(v["integerValue"])
    if "doubleValue" in v:
        return v["doubleValue"]
    if "booleanValue" in v:
        return v["booleanValue"]
    if "timestampValue" in v:
        return v["timestampValue"]
    return None


# 1. Osszes "keresek" dokumentum, aminek a neve Pinter-t tartalmazza (barmilyen honapra)
r = requests.get(f"{BASE}/keresek", params={"key": API_KEY, "pageSize": 300}, timeout=30)
r.raise_for_status()
docs = r.json().get("documents", [])
print(f"Osszesen {len(docs)} kereses dokumentum\n")

talalatok = []
for doc in docs:
    doc_id = doc["name"].split("/")[-1]
    fields = doc["fields"]
    nev = fs_ertek(fields.get("nev", {}))
    if nev and "Pint" in nev:
        talalatok.append((doc_id, {k: fs_ertek(v) for k, v in fields.items()}))

print(f"Pinterhez kapcsolodo dokumentumok: {len(talalatok)}\n")
for doc_id, adat in talalatok:
    print(f"--- dokumentum ID: {doc_id} ---")
    for k, v in adat.items():
        print(f"  {k}: {v}")
    print()

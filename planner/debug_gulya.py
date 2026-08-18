import json
import sys
import requests

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
    if "nullValue" in v:
        return None
    if "arrayValue" in v:
        return [fs_ertek(x) for x in v["arrayValue"].get("values", [])]
    if "mapValue" in v:
        return {k: fs_ertek(x) for k, x in v["mapValue"].get("fields", {}).items()}
    return None


def doc_dict(doc):
    return {k: fs_ertek(v) for k, v in doc.get("fields", {}).items()}


print("=== GULYA ÉLŐ ADATOK A FIRESTORE-BÓL ===")

r = requests.get(f"{BASE}/keresek", params={"key": API_KEY, "pageSize": 200}, timeout=30)
r.raise_for_status()
for doc in r.json().get("documents", []):
    d = doc_dict(doc)
    if d.get("nev") == "Gulya Réka":
        print(f"keresek dokumentum (ev={d.get('ev')}, honap={d.get('honap')}):")
        print(json.dumps(d, ensure_ascii=False, indent=2))

r2 = requests.get(f"{BASE}/reszmunkaido_periodusok", params={"key": API_KEY, "pageSize": 50}, timeout=30)
r2.raise_for_status()
print("\nreszmunkaido_periodusok dokumentumok:")
for doc in r2.json().get("documents", []):
    d = doc_dict(doc)
    print(json.dumps(d, ensure_ascii=False, indent=2))

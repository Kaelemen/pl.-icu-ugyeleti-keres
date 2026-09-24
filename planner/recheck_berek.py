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
    return None


r = requests.get(f"{BASE}/keresek", params={"key": API_KEY, "pageSize": 300}, timeout=30)
r.raise_for_status()
for doc in r.json().get("documents", []):
    fields = doc["fields"]
    nev = fs_ertek(fields.get("nev", {}))
    honap = fs_ertek(fields.get("honap", {}))
    if nev and "Berek" in nev and honap == 10:
        doc_id = doc["name"].split("/")[-1]
        print(f"--- {doc_id} ---")
        for k, v in fields.items():
            print(f"  {k}: {fs_ertek(v)}")

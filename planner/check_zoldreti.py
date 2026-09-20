import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"


def fs_ertek(v):
    if "stringValue" in v:
        return v["stringValue"]
    if "integerValue" in v:
        return int(v["integerValue"])
    if "booleanValue" in v:
        return v["booleanValue"]
    return None


r = requests.get(f"{BASE}/keresek", params={"key": API_KEY, "pageSize": 300}, timeout=30)
r.raise_for_status()
docs = r.json().get("documents", [])
print(f"Osszesen {len(docs)} kereses dokumentum")
for doc in docs:
    fields = doc["fields"]
    nev = fs_ertek(fields.get("nev", {}))
    ev = fs_ertek(fields.get("ev", {}))
    honap = fs_ertek(fields.get("honap", {}))
    if honap == 10 and "ldr" in (nev or "").lower().replace("ö", "o").replace("é", "e") or (nev and "Zöldréti" in nev and honap == 10):
        print("\n--- TALALAT ---")
        for k, v in fields.items():
            print(f"  {k}: {fs_ertek(v)}")

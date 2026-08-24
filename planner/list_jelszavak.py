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
print(f"Osszesen {len(docs)} egyedi (mar megvaltoztatott) jelszo van elmentve:\n")
for doc in docs:
    jelszo = doc["name"].split("/")[-1]
    nev = fs_ertek(doc["fields"].get("nev", {}))
    print(f"  {nev}: {jelszo}")

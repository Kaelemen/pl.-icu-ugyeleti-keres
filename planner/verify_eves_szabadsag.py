import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"

r = requests.get(f"{BASE}/eves_szabadsag", params={"key": API_KEY, "pageSize": 200}, timeout=30)
r.raise_for_status()
docs = r.json().get("documents", [])
print(f"Osszesen {len(docs)} dokumentum az eves_szabadsag gyujtemenyben\n")
for doc in docs:
    fields = doc["fields"]
    nev = fields.get("nev", {}).get("stringValue")
    napok = fields.get("napok", {}).get("arrayValue", {}).get("values", [])
    print(f"  {nev}: {len(napok)} nap")

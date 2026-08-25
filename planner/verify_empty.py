import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"

r = requests.get(f"{BASE}/jelszavak", params={"key": API_KEY, "pageSize": 200}, timeout=30)
r.raise_for_status()
docs = r.json().get("documents", [])
print(f"Jelenlegi bejegyzesek szama: {len(docs)}")
for doc in docs:
    print(" -", doc["name"].split("/")[-1])

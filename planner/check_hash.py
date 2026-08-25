import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"

VART_HASH = "dddcb9058dfeaea53a70e6002a365a03ef2c499bec0804be63decacc11969849"

r = requests.get(f"{BASE}/jelszavak/{VART_HASH}", params={"key": API_KEY}, timeout=30)
print("Status:", r.status_code)
print(r.text)

print("\n--- Osszes bejegyzes listaja ---")
r2 = requests.get(f"{BASE}/jelszavak", params={"key": API_KEY, "pageSize": 200}, timeout=30)
for doc in r2.json().get("documents", []):
    doc_id = doc["name"].split("/")[-1]
    nev = doc["fields"].get("nev", {}).get("stringValue")
    egyezik = "IGEN <-----" if doc_id == VART_HASH else ""
    print(f"  {doc_id} -> {nev}  {egyezik}")

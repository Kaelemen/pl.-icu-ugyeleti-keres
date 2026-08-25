import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"

r = requests.get(f"{BASE}/jelszavak", params={"key": API_KEY, "pageSize": 200}, timeout=30)
r.raise_for_status()
docs = r.json().get("documents", [])
print(f"Talalt {len(docs)} bejegyzes - torles kezdodik, hogy mindenki alaperetelmezett jelszora alljon vissza...")

for doc in docs:
    doc_id = doc["name"].split("/")[-1]
    nev = doc["fields"].get("nev", {}).get("stringValue")
    url = f"{BASE}/jelszavak/{doc_id}"
    resp = requests.delete(url, params={"key": API_KEY}, timeout=30)
    print(f"  {nev}: torles {'OK' if resp.status_code == 200 else 'HIBA ' + str(resp.status_code)}")

print("\nKesz - mindenki most az alapertelmezett (vezeteknev+1) jelszavaval tud belepni,")
print("es a belepeskor a rendszer automatikusan felajanlja az uj jelszo beallitasat.")

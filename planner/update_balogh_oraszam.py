import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"

r = requests.get(f"{BASE}/dolgozok", params={"key": API_KEY, "pageSize": 200}, timeout=30)
r.raise_for_status()
doc_id = None
for doc in r.json().get("documents", []):
    nev = doc["fields"].get("nev", {}).get("stringValue")
    if nev == "Balogh László":
        doc_id = doc["name"].split("/")[-1]

if not doc_id:
    print("NEM TALALHATO Balogh Laszlo dokumentuma")
else:
    url = f"{BASE}/dolgozok/{doc_id}"
    params = {"key": API_KEY, "updateMask.fieldPaths": "napi_munkaido"}
    body = {"fields": {"napi_munkaido": {"doubleValue": 7.2}}}
    resp = requests.patch(url, params=params, json=body, timeout=30)
    print(f"Frissites: {'OK -> 7.2' if resp.status_code == 200 else 'HIBA ' + str(resp.status_code) + ' ' + resp.text}")

    # ellenorzo visszaolvasas
    r2 = requests.get(url, params={"key": API_KEY}, timeout=30)
    print("Ellenorzes:", r2.json().get("fields", {}).get("napi_munkaido"))

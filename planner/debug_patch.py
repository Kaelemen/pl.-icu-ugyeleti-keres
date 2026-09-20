import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"

r = requests.get(f"{BASE}/dolgozok", params={"key": API_KEY, "pageSize": 200}, timeout=30)
docs = r.json().get("documents", [])
for doc in docs:
    fields = doc["fields"]
    nev = fields.get("nev", {}).get("stringValue")
    if nev == "Zöldréti Anikó":
        doc_id = doc["name"].split("/")[-1]
        print("doc_id:", doc_id)
        print("teljes dokumentum:", doc)

        url = f"{BASE}/dolgozok/{doc_id}"
        params = {"key": API_KEY, "updateMask.fieldPaths": "napi_munkaido"}
        body = {"fields": {"napi_munkaido": {"doubleValue": 6.4}}}
        resp = requests.patch(url, params=params, json=body, timeout=30)
        print("\nPATCH statusz:", resp.status_code)
        print("PATCH valasz:", resp.text)

        # azonnali ujraolvasas
        r2 = requests.get(url, params={"key": API_KEY}, timeout=30)
        print("\nUjraolvasas:", r2.text)

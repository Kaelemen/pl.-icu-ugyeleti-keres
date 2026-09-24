import requests
import json

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"


def dolgozok_doc_id(nev):
    return nev.strip().replace(" ", "_")


eredmeny = json.load(open("eves_szabadsag_import.json", encoding="utf-8"))

for nev, napok in eredmeny.items():
    doc_id = dolgozok_doc_id(nev)
    url = f"{BASE}/eves_szabadsag/{doc_id}"
    body = {
        "fields": {
            "nev": {"stringValue": nev},
            "napok": {"arrayValue": {"values": [{"stringValue": nap} for nap in napok]}},
        }
    }
    resp = requests.patch(url, params={"key": API_KEY}, json=body, timeout=30)
    print(f"{nev}: {'OK -> ' + str(len(napok)) + ' nap' if resp.status_code == 200 else 'HIBA ' + str(resp.status_code) + ' ' + resp.text[:200]}")

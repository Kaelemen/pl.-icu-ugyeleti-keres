import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"

# a korabban rogzitett, helyes napi oraszamok (a beszelgetes tortenete alapjan)
JAVITANDO = {
    "Balogh László": 3.2,
    "Hajnal Csilla": 2.3,
    "Sorbán József": 1.52,
    "Szász Kornélia": 2.3,
    "Sztermen Márton": 1.84,
    "Zöldréti Anikó": 6.4,
    "Pintér Enikő": 4.6,
}


def dolgozok_doc_id(nev):
    return nev.lower().replace(" ", "_").replace("é", "e").replace("á", "a").replace("í", "i") \
        .replace("ó", "o").replace("ö", "o").replace("ő", "o").replace("ú", "u").replace("ü", "u").replace("ű", "u")


# elobb lekerjuk a dolgozok lista tenyleges dokumentum-azonositoit (nev alapjan, nem felteve az id-t)
r = requests.get(f"{BASE}/dolgozok", params={"key": API_KEY, "pageSize": 200}, timeout=30)
r.raise_for_status()
docs = r.json().get("documents", [])
nev_to_docid = {}
for doc in docs:
    fields = doc["fields"]
    nev = fields.get("nev", {}).get("stringValue")
    doc_id = doc["name"].split("/")[-1]
    if nev:
        nev_to_docid[nev] = doc_id

for nev, ertek in JAVITANDO.items():
    doc_id = nev_to_docid.get(nev)
    if not doc_id:
        print(f"NEM TALALHATO: {nev}")
        continue
    url = f"{BASE}/dolgozok/{doc_id}"
    params = {"key": API_KEY, "updateMask.fieldPaths": "napi_munkaido"}
    body = {"fields": {"napi_munkaido": {"doubleValue": ertek}}}
    resp = requests.patch(url, params=params, json=body, timeout=30)
    print(f"{nev}: {'OK -> ' + str(ertek) if resp.status_code == 200 else 'HIBA ' + str(resp.status_code) + ' ' + resp.text}")

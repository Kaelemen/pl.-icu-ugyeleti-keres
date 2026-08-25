import hashlib
import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"
PEPPER = "icu-ugyeleti-2026-sofL9x"


def hash_jelszo(jelszo):
    adat = (PEPPER + jelszo.strip().lower()).encode("utf-8")
    return hashlib.sha256(adat).hexdigest()


def fs_ertek(v):
    if "stringValue" in v:
        return v["stringValue"]
    return None


# 1. lekerdezzuk a jelenlegi (nyilt szoveges) bejegyzeseket
r = requests.get(f"{BASE}/jelszavak", params={"key": API_KEY, "pageSize": 200}, timeout=30)
r.raise_for_status()
docs = r.json().get("documents", [])
print(f"Talalt {len(docs)} nyilt szoveges bejegyzes - migralas kezdodik...")

migralt = 0
for doc in docs:
    regi_id = doc["name"].split("/")[-1]
    nev = fs_ertek(doc["fields"].get("nev", {}))
    if not nev:
        continue
    uj_hash = hash_jelszo(regi_id)

    # 2. letrehozzuk az uj, hash-elt dokumentumot
    url_uj = f"{BASE}/jelszavak/{uj_hash}"
    body = {"fields": {"nev": {"stringValue": nev}}}
    resp = requests.patch(url_uj, params={"key": API_KEY}, json=body, timeout=30)
    if resp.status_code != 200:
        print(f"HIBA uj letrehozasakor ({nev}): {resp.status_code} {resp.text}")
        continue

    # 3. toroljuk a regi, nyilt szoveges dokumentumot
    url_regi = f"{BASE}/jelszavak/{regi_id}"
    del_resp = requests.delete(url_regi, params={"key": API_KEY}, timeout=30)
    if del_resp.status_code != 200:
        print(f"HIBA regi torlesekor ({nev}): {del_resp.status_code} {del_resp.text}")
        continue

    print(f"OK: {nev} migralva (hash: {uj_hash[:12]}...)")
    migralt += 1

print(f"\nOsszesen {migralt}/{len(docs)} sikeresen migralva.")

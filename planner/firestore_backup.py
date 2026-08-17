# -*- coding: utf-8 -*-
"""Napi biztonsági mentés: az összes Firestore-gyűjteményt (keresek, jelszavak,
dolgozok) kiolvassa és elmenti egy dátumozott JSON fájlba a repóban. A nyílt
Firestore-szabályokat és a nyilvános Firebase API-kulcsot használja, service
account nélkül - ugyanaz a módszer, mint a havi archiválásnál.

A régi mentéseket 14 napnál tovább nem tartjuk meg, hogy a repó ne dagadjon fel
feleslegesen - mindig elég egy friss, közelmúltbeli állapot a visszaállításhoz."""
import datetime
import glob
import os
import requests

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"
GYUJTEMENYEK = ["keresek", "jelszavak", "dolgozok"]
MEGTARTASI_NAPOK = 14


def fs_ertek(v):
    if "stringValue" in v:
        return v["stringValue"]
    if "integerValue" in v:
        return int(v["integerValue"])
    if "doubleValue" in v:
        return v["doubleValue"]
    if "booleanValue" in v:
        return v["booleanValue"]
    if "nullValue" in v:
        return None
    if "arrayValue" in v:
        return [fs_ertek(x) for x in v["arrayValue"].get("values", [])]
    if "mapValue" in v:
        return {k: fs_ertek(x) for k, x in v["mapValue"].get("fields", {}).items()}
    return None


def doc_dict(doc):
    return {k: fs_ertek(v) for k, v in doc.get("fields", {}).items()}


def gyujtemeny_letoltese(nev):
    dokumentumok = []
    page_token = None
    while True:
        params = {"key": API_KEY, "pageSize": 200}
        if page_token:
            params["pageToken"] = page_token
        r = requests.get(f"{BASE}/{nev}", params=params, timeout=30)
        r.raise_for_status()
        data = r.json()
        for doc in data.get("documents", []):
            doc_id = doc["name"].split("/")[-1]
            dokumentumok.append({"id": doc_id, **doc_dict(doc)})
        page_token = data.get("nextPageToken")
        if not page_token:
            break
    return dokumentumok


def main():
    import json
    ma = datetime.date.today().isoformat()
    os.makedirs("backups", exist_ok=True)

    osszes = {}
    for gyujtemeny in GYUJTEMENYEK:
        try:
            dokumentumok = gyujtemeny_letoltese(gyujtemeny)
            osszes[gyujtemeny] = dokumentumok
            print(f"{gyujtemeny}: {len(dokumentumok)} dokumentum mentve.")
        except Exception as e:
            print(f"HIBA ({gyujtemeny} mentése közben): {e}")
            osszes[gyujtemeny] = None

    kimenet = f"backups/firestore_{ma}.json"
    with open(kimenet, "w", encoding="utf-8") as f:
        json.dump(osszes, f, ensure_ascii=False, indent=2)
    print(f"Mentve: {kimenet}")

    # régi mentések törlése (csak MEGTARTASI_NAPOK napnál régebbiek)
    hatarido = datetime.date.today() - datetime.timedelta(days=MEGTARTASI_NAPOK)
    for fajl in glob.glob("backups/firestore_*.json"):
        try:
            datum_str = os.path.basename(fajl).replace("firestore_", "").replace(".json", "")
            datum = datetime.date.fromisoformat(datum_str)
            if datum < hatarido:
                os.remove(fajl)
                print(f"Törölve (túl régi): {fajl}")
        except Exception:
            pass


if __name__ == "__main__":
    main()

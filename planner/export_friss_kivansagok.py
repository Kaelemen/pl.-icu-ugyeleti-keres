import requests
import json
import sys
import re

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"


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
    if "mapValue" in v:
        return {k: fs_ertek(val) for k, val in v["mapValue"].get("fields", {}).items()}
    if "arrayValue" in v:
        return [fs_ertek(item) for item in v["arrayValue"].get("values", [])]
    return None


def fs_doc_to_dict(fields):
    return {k: fs_ertek(v) for k, v in fields.items()}


def parse_nap_lista(s):
    if not s:
        return []
    napok = []
    for resz in str(s).split(","):
        resz = resz.strip()
        if not resz:
            continue
        if "-" in resz:
            a, b = resz.split("-")
            napok.extend(range(int(a), int(b) + 1))
        else:
            napok.append(int(resz))
    return napok


def main(celzott_fajlnev):
    # a fajlnevbol (pl. "kivansagok_2026_10.json") kiolvassuk az ev/honap erteket
    illeszkedes = re.search(r"kivansagok_(\d{4})_(\d{1,2})\.json", celzott_fajlnev)
    if not illeszkedes:
        print(f"A fajlnev nem 'kivansagok_EEEE_HH.json' mintazatu, kihagyva: {celzott_fajlnev}")
        return
    cel_ev, cel_honap = int(illeszkedes.group(1)), int(illeszkedes.group(2))
    print(f"Friss export keszul: {cel_ev}.{cel_honap:02d}")

    # 1. dolgozok
    r = requests.get(f"{BASE}/dolgozok", params={"key": API_KEY, "pageSize": 200}, timeout=30)
    r.raise_for_status()
    dolgozok_docs = [fs_doc_to_dict(d["fields"]) for d in r.json().get("documents", [])]
    dolgozok_docs.sort(key=lambda d: d.get("sorszam", 999))

    # 2. keresek (a celzott honapra)
    r = requests.get(f"{BASE}/keresek", params={"key": API_KEY, "pageSize": 300}, timeout=30)
    r.raise_for_status()
    all_requests = [fs_doc_to_dict(d["fields"]) for d in r.json().get("documents", [])]
    celzott_keresek = [req for req in all_requests if req.get("ev") == cel_ev and req.get("honap") == cel_honap]
    print(f"Talalt keresek a celzott honapra: {len(celzott_keresek)}")

    # 3. reszmunkaido_periodusok
    r = requests.get(f"{BASE}/reszmunkaido_periodusok", params={"key": API_KEY, "pageSize": 200}, timeout=30)
    r.raise_for_status()
    periodusok_docs = [fs_doc_to_dict(d["fields"]) for d in r.json().get("documents", [])]
    reszmunkaido_periodusok = {}
    for p in periodusok_docs:
        nev = p.get("nev")
        if nev:
            reszmunkaido_periodusok[nev] = {k: v for k, v in p.items() if k != "nev"}

    # 4. elozo_honap_tulora (ha kulon gyujtemenyben van tarolva, opcionalis)
    elozo_honap_tulora = {}
    try:
        r = requests.get(f"{BASE}/elozo_honap_tulora", params={"key": API_KEY, "pageSize": 200}, timeout=30)
        if r.status_code == 200:
            for doc in r.json().get("documents", []):
                fields = fs_doc_to_dict(doc["fields"])
                nev = fields.get("nev") or doc["name"].split("/")[-1]
                if "ora" in fields:
                    elozo_honap_tulora[nev] = fields["ora"]
    except Exception:
        pass

    # 5. eves_szabadsag - az eves szabadsagtervezobol azok a napok, amik a celzott
    # honapra esnek, minden dolgozonal osszefuzve a havi kereses sajat szabadsag-napjaival.
    eves_szabadsag_honapra = {}
    try:
        r = requests.get(f"{BASE}/eves_szabadsag", params={"key": API_KEY, "pageSize": 200}, timeout=30)
        if r.status_code == 200:
            honap_elotag = f"{cel_ev}-{cel_honap:02d}-"
            for doc in r.json().get("documents", []):
                fields = fs_doc_to_dict(doc["fields"])
                nev = fields.get("nev")
                napok = fields.get("napok") or []
                honapra_eso = sorted(int(nap.split("-")[2]) for nap in napok if nap.startswith(honap_elotag))
                if nev and honapra_eso:
                    eves_szabadsag_honapra[nev] = honapra_eso
    except Exception:
        pass

    kivansagok = {}
    nyolc_ora_nappal = {}
    mindenkeppen_szeretne = {}
    kert_ugyeletszam = {}
    for req in celzott_keresek:
        if req.get("nincs_keres"):
            continue
        nev = req["nev"]
        szeret_napok = set(parse_nap_lista(req.get("szeretne"))) | set(parse_nap_lista(req.get("mindenkeppen_szeretne")))
        szabadsag_napok = set(parse_nap_lista(req.get("szabadsag"))) | set(eves_szabadsag_honapra.get(nev, []))
        kivansagok[nev] = {
            "szabadsag": sorted(szabadsag_napok),
            "nem": parse_nap_lista(req.get("nem_szeretne")),
            "szeret": sorted(szeret_napok),
        }
        if req.get("nyolc_ora_nappal"):
            nyolc_ora_nappal[nev] = parse_nap_lista(req.get("nyolc_ora_nappal"))
        if req.get("mindenkeppen_szeretne"):
            mindenkeppen_szeretne[nev] = parse_nap_lista(req.get("mindenkeppen_szeretne"))
        if req.get("kert_ugyeletszam"):
            kert_ugyeletszam[nev] = req["kert_ugyeletszam"]

    # Azok, akiknek van éves szabadság-napjuk erre a hónapra, de egyáltalán nem küldtek be
    # havi kérést (vagy azt jelezték, hogy nincs kérésük) - nekik is fel kell venni legalább
    # a szabadság-napjaikat, különben a generálás nem tudna róla.
    for nev, honapra_eso in eves_szabadsag_honapra.items():
        if nev not in kivansagok:
            kivansagok[nev] = {"szabadsag": honapra_eso, "nem": [], "szeret": []}

    kulsos_gyakorlaton = [req["nev"] for req in celzott_keresek if req.get("kulsos_gyakorlat") is True]

    kivansagok_payload = {
        "ev": cel_ev,
        "honap": cel_honap,
        "kivansagok": kivansagok,
        "kert_ugyeletszam": kert_ugyeletszam,
        "nyolc_ora_nappal": nyolc_ora_nappal,
        "mindenkeppen_szeretne": mindenkeppen_szeretne,
        "kulsos_gyakorlaton": kulsos_gyakorlaton,
        "elozo_honap_lelepok": [],
        "reszmunkaido_periodusok": reszmunkaido_periodusok,
        "elozo_honap_tulora": elozo_honap_tulora,
        "dolgozok": dolgozok_docs,
    }

    with open(celzott_fajlnev, "w", encoding="utf-8") as f:
        json.dump(kivansagok_payload, f, ensure_ascii=False, indent=2)

    print(f"Sikeresen frissitve/mentve: {celzott_fajlnev} ({len(kivansagok)} fo kivansagaval)")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Hasznalat: python3 export_friss_kivansagok.py kivansagok_EEEE_HH.json")
        sys.exit(0)
    main(sys.argv[1])

import requests
import json

PROJECT_ID = "icu-ugyeleti-keres"
API_KEY = "AIzaSyCJRiAjLwtxnTeskC88YK9dPkN7JLK-Oqk"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"


def fs_ertek(v):
    if "stringValue" in v:
        return v["stringValue"]
    if "integerValue" in v:
        return int(v["integerValue"])
    if "booleanValue" in v:
        return v["booleanValue"]
    if "doubleValue" in v:
        return v["doubleValue"]
    if "nullValue" in v:
        return None
    return None


def fs_doc_to_dict(fields):
    return {k: fs_ertek(v) for k, v in fields.items()}


def parse_nap_lista(s):
    if not s:
        return []
    napok = []
    for resz in s.split(","):
        resz = resz.strip()
        if not resz:
            continue
        if "-" in resz:
            a, b = resz.split("-")
            napok.extend(range(int(a), int(b) + 1))
        else:
            napok.append(int(resz))
    return napok


# 1. dolgozok
r = requests.get(f"{BASE}/dolgozok", params={"key": API_KEY, "pageSize": 200}, timeout=30)
dolgozok_docs = [fs_doc_to_dict(d["fields"]) for d in r.json().get("documents", [])]
dolgozok_docs.sort(key=lambda d: d.get("sorszam", 999))

# 2. keresek (2026 oktober)
r = requests.get(f"{BASE}/keresek", params={"key": API_KEY, "pageSize": 300}, timeout=30)
all_requests = [fs_doc_to_dict(d["fields"]) for d in r.json().get("documents", [])]
oktoberi = [r for r in all_requests if r.get("ev") == 2026 and r.get("honap") == 10]
print(f"Oktoberi kercesek szama: {len(oktoberi)}")

# 3. reszmunkaido_periodusok
r = requests.get(f"{BASE}/reszmunkaido_periodusok", params={"key": API_KEY, "pageSize": 200}, timeout=30)
periodusok_docs = r.json().get("documents", [])

kivansagok = {}
nyolc_ora_nappal = {}
mindenkeppen_szeretne = {}
for req in oktoberi:
    if req.get("nincs_keres"):
        continue
    nev = req["nev"]
    szeret_napok = set(parse_nap_lista(req.get("szeretne"))) | set(parse_nap_lista(req.get("mindenkeppen_szeretne")))
    kivansagok[nev] = {
        "szabadsag": parse_nap_lista(req.get("szabadsag")),
        "nem": parse_nap_lista(req.get("nem_szeretne")),
        "szeret": sorted(szeret_napok),
    }
    if req.get("nyolc_ora_nappal"):
        nyolc_ora_nappal[nev] = parse_nap_lista(req.get("nyolc_ora_nappal"))
    if req.get("mindenkeppen_szeretne"):
        mindenkeppen_szeretne[nev] = parse_nap_lista(req.get("mindenkeppen_szeretne"))

kert_ugyeletszam = {}
for req in oktoberi:
    if req.get("kert_ugyeletszam"):
        kert_ugyeletszam[req["nev"]] = req["kert_ugyeletszam"]

kulsos_gyakorlaton = [req["nev"] for req in oktoberi if req.get("kulsos_gyakorlat") is True]

kivansagok_payload = {
    "ev": 2026,
    "honap": 10,
    "kivansagok": kivansagok,
    "kert_ugyeletszam": kert_ugyeletszam,
    "nyolc_ora_nappal": nyolc_ora_nappal,
    "mindenkeppen_szeretne": mindenkeppen_szeretne,
    "kulsos_gyakorlaton": kulsos_gyakorlaton,
    "elozo_honap_lelepok": [],
    "reszmunkaido_periodusok": {},
    "elozo_honap_tulora": {},
    "dolgozok": dolgozok_docs,
}

with open("kivansagok_2026_10.json", "w", encoding="utf-8") as f:
    json.dump(kivansagok_payload, f, ensure_ascii=False, indent=2)

print("Zoldreti kivansagok:", kivansagok.get("Zöldréti Anikó"))
print("mentve: kivansagok_2026_10.json")

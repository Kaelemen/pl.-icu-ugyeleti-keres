# -*- coding: utf-8 -*-
"""Beleírja az automatikusan (Google Drive-ból) kiolvasott "ki lépett le 1-jén"
listát a kívánságok fájlba - ha talált ilyet, felülírja az admin felületen
kézzel megadott listát (mivel a Drive-os adat a tényleges, megtörtént beosztásból
származik, megbízhatóbb, mint egy kézi becslés)."""
import sys
import json

kiv_fajl = sys.argv[1] if len(sys.argv) > 1 else "kivansagok_input.json"

with open(kiv_fajl, encoding="utf-8") as f:
    kiv = json.load(f)

try:
    with open("elozo_honap_auto.json", encoding="utf-8") as f:
        auto = json.load(f)
except FileNotFoundError:
    auto = {"elozo_honap_lelepok": []}

if auto.get("elozo_honap_lelepok"):
    kiv["elozo_honap_lelepok"] = auto["elozo_honap_lelepok"]
    print(f"Automatikusan felismert leszállók beépítve: {auto['elozo_honap_lelepok']}")
else:
    print(f"Nincs automatikusan felismert leszálló - marad a kézzel megadott: {kiv.get('elozo_honap_lelepok', [])}")

with open(kiv_fajl, "w", encoding="utf-8") as f:
    json.dump(kiv, f, ensure_ascii=False)

import json
import sys

kiv_fajl = sys.argv[1]
with open(kiv_fajl, encoding="utf-8") as f:
    kiv = json.load(f)

print("=== GULYA DIAGNOSZTIKA ===")
print("reszmunkaido_periodusok (teljes):", json.dumps(kiv.get("reszmunkaido_periodusok", {}), ensure_ascii=False))
print("Gulya kivansaga:", json.dumps(kiv.get("kivansagok", {}).get("Gulya Réka"), ensure_ascii=False))
print("Gulya kert_ugyeletszam:", kiv.get("kert_ugyeletszam", {}).get("Gulya Réka"))
dolgozok = kiv.get("dolgozok", [])
gulya_dolgozo = next((d for d in dolgozok if d.get("nev") == "Gulya Réka"), None)
print("Gulya dolgozo-rekord:", json.dumps(gulya_dolgozo, ensure_ascii=False))

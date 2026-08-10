# -*- coding: utf-8 -*-
"""
ICU ügyeleti beosztás generátor - szabály-adatbázisból (szabalyok.json) és
hónap-specifikus kívánság-adatbázisból (kivansagok_ÉÉÉÉ_HH.json) dolgozik.
Használat: python3 generate_schedule.py [kivansagok_fajl.json] [sablon.xlsx] [kimenet.xlsx]
"""
import sys
import json
import datetime
import random
import openpyxl
from openpyxl.styles import PatternFill, Font as XLFont

import os
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SZABALYOK_PATH = os.path.join(SCRIPT_DIR, "szabalyok.json")
KIVANSAGOK_PATH = sys.argv[1] if len(sys.argv) > 1 else os.path.join(SCRIPT_DIR, "kivansagok_2026_08.json")
SABLON_PATH = sys.argv[2] if len(sys.argv) > 2 else "ICU_ugyeleti_beosztas.xlsx"
KIMENET_PATH = sys.argv[3] if len(sys.argv) > 3 else "ICU_ugyeleti_beosztas_probabeosztas.xlsx"
SEED = int(sys.argv[4]) if len(sys.argv) > 4 else 0
rng = random.Random(SEED)

with open(SZABALYOK_PATH, encoding="utf-8") as f:
    SZAB = json.load(f)
with open(KIVANSAGOK_PATH, encoding="utf-8") as f:
    KIV = json.load(f)

YEAR, MONTH = KIV["ev"], KIV["honap"]
first_day = datetime.date(YEAR, MONTH, 1)
next_month = datetime.date(YEAR, MONTH + 1, 1) if MONTH != 12 else datetime.date(YEAR + 1, 1, 1)
num_days = (next_month - first_day).days

WEEKDAY_HU = {"hetfo": 0, "kedd": 1, "szerda": 2, "csutortok": 3, "pentek": 4, "szombat": 5, "vasarnap": 6}

# ---------------------------------------------------------------------------
# Dolgozók betöltése a szabalyok.json-ból
# ---------------------------------------------------------------------------
staff = []
for d in SZAB["dolgozok"]:
    staff.append((d["nev"], d["kategoria"], d["napi_munkaido"], KIV.get("kert_ugyeletszam", {}).get(d["nev"]), d["tipus"]))
staff_order_all = [d["nev"] for d in SZAB["dolgozok"]]
keret_tipus_map = {d["nev"]: ("Havi" if d["szerzodes_tipus"] == "Részmunkaidő - havi órakeret" else "Napi")
                    for d in SZAB["dolgozok"]}
havi_oraszam_map = {d["nev"]: d["havi_oraszam"] for d in SZAB["dolgozok"] if d["havi_oraszam"]}
HAVI_KERETESEK = {nev for nev, kt in keret_tipus_map.items() if kt == "Havi"}
REZIDENSEK = {d["nev"] for d in SZAB["dolgozok"] if d["tipus"] == "Rezidens"}
KULSOS_GYAKORLATON = set(KIV.get("kulsos_gyakorlaton", []))  # rezidensek, akik ebben a hónapban külsős gyakorlaton vannak

def keret_of(name):
    return keret_tipus_map.get(name, "Napi")

ALT = SZAB["altalanos_szabalyok"]
MIN_PIHENO = ALT["min_pihenonap_ket_ugyelet_kozott"]
SZOMBAT_NINCS_STROKE = ALT["szombaton_nincs_stroke"]
MIN_SZAKORVOS = ALT["min_szakorvos_naponta_ugyeletben"]
O1_ALAP = ALT["o1_alapertelmezett_szemely"]
O2_ALAP = ALT["o2_alapertelmezett_szemely"]
MUTO_MIN = ALT["muto_minimum_letszam"]
MUTO_PADLO = ALT["muto_padlo_letszam"]
O1_O2_TILTOTT = set(ALT.get("o1_o2_tiltott_szemelyek", []))

NAPI_KOTELEZO_ORA = SZAB["orakeret_konstansok"]["napi_kotelezo_ora"]
munkanapok_a_honapban = sum(1 for d in range(num_days)
                             if (first_day + datetime.timedelta(days=d)).weekday() < 5)
RESZ_NAPI_ORASZAMOS = {d["nev"]: d["napi_munkaido"] for d in SZAB["dolgozok"]
                        if d["szerzodes_tipus"] == "Részmunkaidő - napi óraszám"}
RESZ_NAPI_KAPACITAS = {nev: napi * munkanapok_a_honapban for nev, napi in RESZ_NAPI_ORASZAMOS.items()}
kotelezo_ora_used = {nev: 0 for nev in RESZ_NAPI_ORASZAMOS}

# ---------------------------------------------------------------------------
# Kívánságok betöltése (hónap-specifikus), majd a személyi (tartós) szabályok
# alkalmazása dátumokra bontva erre a hónapra
# ---------------------------------------------------------------------------
kivansagok = {name: {"szabadsag": [], "nem": [], "szeret": []} for name in staff_order_all}
for name, v in KIV["kivansagok"].items():
    kivansagok[name] = {"szabadsag": list(v.get("szabadsag", [])),
                         "nem": list(v.get("nem", [])),
                         "szeret": list(v.get("szeret", []))}

def days_by_weekday(weekday_idx):
    out = []
    for d in range(num_days):
        if (first_day + datetime.timedelta(days=d)).weekday() == weekday_idx:
            out.append(d + 1)
    return out

CSAK_JELOLT_NAPOKON = set()
RENDES_NAP_CSAK_HETENTE = {}   # name -> set(weekday_idx), csak ezeken a napokon lehet "m"
HETI_FIX_ESEMENY = {}   # name -> {"weekday": int, "kod": str}
RESZMUNKAIDO_TOL = {}   # name -> első nap (int), amitől "rendesen" jelen van

for szab in SZAB["szemelyi_megkotesek"]:
    name = szab["nev"]
    tipus = szab["tipus"]
    if tipus == "tiltott_napok_hetente":
        days = []
        for wd in szab["napok"]:
            days += days_by_weekday(WEEKDAY_HU[wd])
        for d in days:
            if d not in kivansagok[name]["nem"] and d not in kivansagok[name]["szeret"]:
                kivansagok[name]["nem"].append(d)
    elif tipus == "tiltott_kezdes_hetente":
        days = []
        for wd in szab["napok"]:
            days += days_by_weekday(WEEKDAY_HU[wd])
        for d in days:
            if d not in kivansagok[name]["nem"] and d not in kivansagok[name]["szeret"]:
                kivansagok[name]["nem"].append(d)
    elif tipus == "nem_dolgozik_hetente":
        days = []
        for wd in szab["napok"]:
            days += days_by_weekday(WEEKDAY_HU[wd])
        for d in days:
            if d not in kivansagok[name]["nem"] and d not in kivansagok[name]["szeret"]:
                kivansagok[name]["nem"].append(d)
        HETI_FIX_ESEMENY.setdefault(name, {})["nem_dolgozik_weekday"] = [WEEKDAY_HU[w] for w in szab["napok"]]
    elif tipus == "csak_jelolt_napokon_dolgozik":
        CSAK_JELOLT_NAPOKON.add(name)
    elif tipus == "heti_fix_esemeny":
        HETI_FIX_ESEMENY.setdefault(name, {})["kod"] = szab["kod"]
        HETI_FIX_ESEMENY[name]["weekday"] = WEEKDAY_HU[szab["nap"]]
    elif tipus == "rendes_nap_csak_hetente":
        RENDES_NAP_CSAK_HETENTE[name] = {WEEKDAY_HU[w] for w in szab["napok"]}

# Minden "Részmunkaidő - napi óraszám" dolgozó alapból csak a saját jelölt ("Szeretne")
# napjain lehet jelen (bármilyen szerepben) - kivéve, akinek konkrét heti mintája van
# (pl. Korompai Máté: rendes_nap_csak_hetente).
for _d in SZAB["dolgozok"]:
    if _d["szerzodes_tipus"] == "Részmunkaidő - napi óraszám" and _d["nev"] not in RENDES_NAP_CSAK_HETENTE:
        CSAK_JELOLT_NAPOKON.add(_d["nev"])

# Részmunkaidő-periódusok beolvasása ELŐBB kell, mint a "csak jelölt napokon" szigorítás,
# hogy tudjuk: kinek van strukturális időszak-korlátja (pl. Gulya: csak a hónap 2. felében
# dolgozik) - az ilyen embereknél a "jó napok" lista NEM egy szűkítő plusz-korlát, hanem
# csak ügyelet-preferencia; a jelenlétüket kizárólag az időszak-korlát szabja meg, a teljes
# időszakon belül egyébként elérhetők.
for name, v in KIV.get("reszmunkaido_periodusok", {}).items():
    tipus = v.get("tipus")
    erinti = set(v.get("erinti", ["jelenlet"]))
    if tipus == "csak_ettol_a_naptol":
        RESZMUNKAIDO_TOL[name] = {"kezdo_nap": v["kezdo_nap"], "erinti": erinti}
    elif tipus == "csak_eddig_a_napig":
        RESZMUNKAIDO_TOL[name] = {"utolso_nap": v["utolso_nap"], "erinti": erinti}

# "csak jelölt napokon dolgozik": minden más nap (ami nincs szabin/szeretve/nemben) -> Nem szeretne.
# Kivéve: (a) ha valakinek egyáltalán nincs megadva egyetlen "jó nap" sem - az nem azt jelenti,
# hogy soha nem osztható be, hanem hogy nincs konkrét kérése (bármelyik nap jó neki); (b) ha
# valakinek strukturális időszak-korlátja van a jelenlétére nézve (pl. Gulya) - annál a "jó
# napok" lista csak ügyelet-preferencia, nem szűkíti tovább a jelenlétét az időszakon belül.
EREDETI_KIFEJEZETT_NEM = {name: set(kivansagok[name]["nem"]) for name in CSAK_JELOLT_NAPOKON}
for name in CSAK_JELOLT_NAPOKON:
    p = kivansagok[name]
    if not p["szeret"]:
        continue
    info = RESZMUNKAIDO_TOL.get(name)
    if info and "jelenlet" in info.get("erinti", ()):
        continue
    covered = set(p["szabadsag"]) | set(p["szeret"]) | set(p["nem"])
    for d in range(1, num_days + 1):
        if d not in covered:
            p["nem"].append(d)

def nap_engedelyezett(name, day_nap_szam, hatas):
    """hatas: 'jelenlet' vagy 'ugyelet' - engedélyezett-e ez a nap erre a hatásra nézve"""
    info = RESZMUNKAIDO_TOL.get(name)
    if not info or hatas not in info.get("erinti", ()):
        return True
    if "kezdo_nap" in info:
        return day_nap_szam >= info["kezdo_nap"]
    if "utolso_nap" in info:
        return day_nap_szam <= info["utolso_nap"]
    return True

prefs = {}
for name, p in kivansagok.items():
    for d in p["szabadsag"]:
        prefs[(name, first_day + datetime.timedelta(days=d - 1))] = "Szabadság"
    for d in p["nem"]:
        prefs[(name, first_day + datetime.timedelta(days=d - 1))] = "Nem szeretne"
    for d in p["szeret"]:
        prefs[(name, first_day + datetime.timedelta(days=d - 1))] = "Szeretne"

def is_szabadsag(name, day_date):
    return prefs.get((name, day_date)) == "Szabadság"

def kotelezo_delta_ha_ma_ugyel(name, day_date):
    delta = NAPI_KOTELEZO_ORA
    if (day_date + datetime.timedelta(days=1)).weekday() < 5:
        delta += RESZ_NAPI_ORASZAMOS.get(name, NAPI_KOTELEZO_ORA)
    return delta

TULLEPES_TURESHATAR = 7  # óra - a havi kötelezőt inkább kicsit túl, mint alul, de max ennyivel lépheti túl

def would_exceed_resz_kapacitas(name, day_date):
    if name not in RESZ_NAPI_ORASZAMOS:
        return False
    delta = kotelezo_delta_ha_ma_ugyel(name, day_date)
    return kotelezo_ora_used[name] + delta > RESZ_NAPI_KAPACITAS[name] + TULLEPES_TURESHATAR

def heti_fix_esemeny_ma(name, day_date):
    info = HETI_FIX_ESEMENY.get(name)
    if info and "kod" in info and day_date.weekday() == info["weekday"]:
        return info["kod"]
    return None

def nem_dolgozik_hetente_ma(name, day_date):
    info = HETI_FIX_ESEMENY.get(name)
    return bool(info and day_date.weekday() in info.get("nem_dolgozik_weekday", []))

def jelenlet_tiltott(name, day_date):
    day_nap = (day_date - first_day).days + 1
    return not nap_engedelyezett(name, day_nap, "jelenlet")

def ugyelet_tiltott(name, day_date):
    day_nap = (day_date - first_day).days + 1
    return not nap_engedelyezett(name, day_nap, "ugyelet")

def eligible(cat, duty):
    if duty == "Intenzív":
        return cat == "I"
    if duty == "Aneszt":
        return cat in ("I", "A")
    if duty == "Stroke":
        return cat in ("I", "A", "S")
    return False

tipus_of = {name: tipus for name, _, _, _, tipus in staff}
assigned_count = {name: 0 for name, *_ in staff}
target_weight = {name: (req if req else hrs) for name, _, hrs, req, _ in staff}
last_duty_date = {name: None for name, *_ in staff}

def would_exceed_havi_kvota(name, extra_duties=1):
    if name not in HAVI_KERETESEK:
        return False
    kvota = havi_oraszam_map.get(name)
    if not kvota:
        return False
    return (assigned_count[name] + extra_duties) * 24 > kvota

duty_types = ["Intenzív", "Aneszt", "Stroke"]
schedule = {}

for d in range(num_days):
    day_date = first_day + datetime.timedelta(days=d)
    is_saturday = day_date.weekday() == 5
    today_assigned = set()
    schedule[d] = {}
    for duty in duty_types:
        if duty == "Stroke" and is_saturday and SZOMBAT_NINCS_STROKE:
            continue
        candidates = []
        for name, cat, hrs, req, tipus in staff:
            if not eligible(cat, duty):
                continue
            if name in today_assigned:
                continue
            pref = prefs.get((name, day_date))
            if pref in ("Szabadság", "Nem szeretne"):
                continue
            ld = last_duty_date[name]
            if ld is not None and (day_date - ld).days < MIN_PIHENO:
                continue
            if would_exceed_havi_kvota(name):
                continue
            if ugyelet_tiltott(name, day_date):
                continue
            if would_exceed_resz_kapacitas(name, day_date):
                continue
            ratio = assigned_count[name] / target_weight[name]
            bonus = -0.3 if pref == "Szeretne" else 0.0
            jitter = (rng.random() - 0.5) * 0.06
            candidates.append((ratio + bonus + jitter, assigned_count[name], name))
        if not candidates:
            schedule[d][duty] = None
            continue
        candidates.sort(key=lambda x: (x[0], x[1], x[2]))
        chosen = candidates[0][2]
        schedule[d][duty] = chosen
        today_assigned.add(chosen)
        assigned_count[chosen] += 1
        last_duty_date[chosen] = day_date
        if chosen in RESZ_NAPI_ORASZAMOS:
            kotelezo_ora_used[chosen] += kotelezo_delta_ha_ma_ugyel(chosen, day_date)

    if MIN_SZAKORVOS > 0 and any(schedule[d].values()) and not any(
        tipus_of.get(nm) == "Szakorvos" for nm in schedule[d].values() if nm
    ):
        for duty in ("Stroke", "Aneszt", "Intenzív"):
            current = schedule[d].get(duty)
            if current is None:
                continue
            best = None
            for name, cat, hrs, req, tipus in staff:
                if tipus != "Szakorvos" or not eligible(cat, duty):
                    continue
                if name in today_assigned and name != current:
                    continue
                pref = prefs.get((name, day_date))
                if pref in ("Szabadság", "Nem szeretne"):
                    continue
                if name != current and would_exceed_havi_kvota(name):
                    continue
                if name != current and ugyelet_tiltott(name, day_date):
                    continue
                if name != current and would_exceed_resz_kapacitas(name, day_date):
                    continue
                ld = last_duty_date[name]
                temp_ld = None if name == current else ld
                if temp_ld is not None and (day_date - temp_ld).days < MIN_PIHENO:
                    continue
                ratio = assigned_count[name] / target_weight[name]
                if best is None or ratio < best[0]:
                    best = (ratio, name)
            if best is not None and best[1] != current:
                new_name = best[1]
                assigned_count[current] -= 1
                today_assigned.discard(current)
                if current in RESZ_NAPI_ORASZAMOS:
                    kotelezo_ora_used[current] -= kotelezo_delta_ha_ma_ugyel(current, day_date)
                schedule[d][duty] = new_name
                today_assigned.add(new_name)
                assigned_count[new_name] += 1
                last_duty_date[new_name] = day_date
                if new_name in RESZ_NAPI_ORASZAMOS:
                    kotelezo_ora_used[new_name] += kotelezo_delta_ha_ma_ugyel(new_name, day_date)
                break

# ---------------------------------------------------------------------------
# lelépő nap: bármelyik duty utáni nap MINDENKINÉL (napi és havi keretesnél is)
# ---------------------------------------------------------------------------
worked_days = {name: set() for name, *_ in staff}
for d in range(num_days):
    day_date = first_day + datetime.timedelta(days=d)
    for nm in schedule[d].values():
        if nm:
            worked_days[nm].add(day_date)

def is_lelepo(name, day_date):
    return (day_date - datetime.timedelta(days=1)) in worked_days[name]

# ---------------------------------------------------------------------------
# O1/O2 osztályos kiosztás
# ---------------------------------------------------------------------------
o_assigned_count = {name: 0 for name, *_ in staff}
o1_o2 = {}

def is_available_for_O(name, day_date, exclude=()):
    if name in exclude:
        return False
    if name in schedule[(day_date - first_day).days].values():
        return False
    if is_szabadsag(name, day_date):
        return False
    if is_lelepo(name, day_date):
        return False
    if heti_fix_esemeny_ma(name, day_date):
        return False
    if name in HAVI_KERETESEK:
        return False
    if name in O1_O2_TILTOTT:
        return False
    if prefs.get((name, day_date)) == "Nem szeretne" and name in CSAK_JELOLT_NAPOKON:
        return False
    if jelenlet_tiltott(name, day_date):
        return False
    if nem_dolgozik_hetente_ma(name, day_date):
        return False
    if name in RESZ_NAPI_ORASZAMOS and not kivansagok[name]["szeret"]:
        return False
    if name in RENDES_NAP_CSAK_HETENTE and day_date.weekday() not in RENDES_NAP_CSAK_HETENTE[name]:
        return False
    if name in KULSOS_GYAKORLATON:
        return False
    return True

def is_available_for_O2(name, day_date, exclude=()):
    """Az O2-nél az St és A ügyeletesek is szóba jöhetnek (nappal még dolgoznak,
    az ügyeletük csak este kezdődik) - csak az Intenzív-ügyeletes van kizárva."""
    if name in exclude:
        return False
    if schedule[(day_date - first_day).days].get("Intenzív") == name:
        return False
    if is_szabadsag(name, day_date):
        return False
    if is_lelepo(name, day_date):
        return False
    if heti_fix_esemeny_ma(name, day_date):
        return False
    if name in HAVI_KERETESEK:
        return False
    if name in O1_O2_TILTOTT:
        return False
    if prefs.get((name, day_date)) == "Nem szeretne" and name in CSAK_JELOLT_NAPOKON:
        return False
    if jelenlet_tiltott(name, day_date):
        return False
    if nem_dolgozik_hetente_ma(name, day_date):
        return False
    if name in RESZ_NAPI_ORASZAMOS and not kivansagok[name]["szeret"]:
        return False
    if name in RENDES_NAP_CSAK_HETENTE and day_date.weekday() not in RENDES_NAP_CSAK_HETENTE[name]:
        return False
    if name in KULSOS_GYAKORLATON:
        return False
    return True

def find_substitute(d, day_date, exclude):
    if d + 1 < num_days:
        next_ito = schedule[d + 1].get("Intenzív")
        if next_ito and is_available_for_O(next_ito, day_date, exclude):
            return next_ito
    candidates = [(o_assigned_count[nm], nm) for nm in staff_order_all
                  if is_available_for_O(nm, day_date, exclude)]
    if candidates:
        candidates.sort()
        return candidates[0][1]
    return None

def find_substitute_o2(day_date, exclude):
    d = (day_date - first_day).days
    on_duty_today = set(schedule[d].values())
    # elsőként a nem-ügyeletes (sima "m") jelöltek közül választunk - az St/A ügyeletes
    # csak akkor jön szóba, ha egyáltalán nincs más elérhető, VAGY ha ő a heti
    # folytonosság megtartásához kell (azt a hívó fél külön kezeli, mielőtt idejutna).
    plain_candidates = [(o_assigned_count[nm], nm) for nm in staff_order_all
                         if nm not in on_duty_today and is_available_for_O2(nm, day_date, exclude)]
    if plain_candidates:
        plain_candidates.sort()
        return plain_candidates[0][1]
    duty_candidates = [(o_assigned_count[nm], nm) for nm in staff_order_all
                        if is_available_for_O2(nm, day_date, exclude)]
    if duty_candidates:
        duty_candidates.sort()
        return duty_candidates[0][1]
    return None

def pick_weekly_o2_anchor(d_start):
    """A hét hátralévő hétköznapjaira megnézi, ki elérhető a legtöbb napon O2-re -
    ő lesz a heti "anchor", hogy lehetőleg végig ő maradjon, ne töredezzen szét a hét."""
    het_szam = (first_day + datetime.timedelta(days=d_start)).isocalendar()[1]
    het_napok = [dd for dd in range(d_start, num_days)
                 if (first_day + datetime.timedelta(days=dd)).isocalendar()[1] == het_szam
                 and (first_day + datetime.timedelta(days=dd)).weekday() < 5]
    scores = []
    for nm in staff_order_all:
        if nm in O1_O2_TILTOTT or nm in HAVI_KERETESEK:
            continue
        count = sum(1 for dd in het_napok
                    if is_available_for_O2(nm, first_day + datetime.timedelta(days=dd), exclude=set()))
        if count > 0:
            scores.append((-count, o_assigned_count[nm], nm))
    if scores:
        scores.sort()
        return scores[0][2]
    return None

heti_o2_szemely = None
heti_o2_hetszam = None
heti_anchor_by_het = {}

for d in range(num_days):
    day_date = first_day + datetime.timedelta(days=d)
    if day_date.weekday() >= 5 or not ALT["o1_o2_szukseges_hetkoznap"]:
        continue

    o1 = O1_ALAP if is_available_for_O(O1_ALAP, day_date) else find_substitute(d, day_date, exclude=())
    chosen = {o1} if o1 else set()

    present_count = 0
    ito_present = 1 if schedule[d].get("Intenzív") else 0
    mr_present = 1 if any(heti_fix_esemeny_ma(nm, day_date) for nm in staff_order_all) else 0
    for name in staff_order_all:
        if jelenlet_tiltott(name, day_date) and name not in schedule[d].values():
            continue
        if nem_dolgozik_hetente_ma(name, day_date) and name not in schedule[d].values():
            continue
        if (name in RENDES_NAP_CSAK_HETENTE and day_date.weekday() not in RENDES_NAP_CSAK_HETENTE[name]
                and name not in schedule[d].values()):
            continue
        if (name in RESZ_NAPI_ORASZAMOS and not kivansagok[name]["szeret"]
                and name not in schedule[d].values()):
            continue
        if (name in schedule[d].values() or heti_fix_esemeny_ma(name, day_date) or
                (keret_of(name) == "Napi" and not is_szabadsag(name, day_date) and not is_lelepo(name, day_date)
                 and not (name in CSAK_JELOLT_NAPOKON and prefs.get((name, day_date)) == "Nem szeretne"))):
            present_count += 1

    # heti folytonosság: ugyanaz az O2 személy próbál maradni egy héten belül,
    # csak akkor váltunk, ha új hét kezdődik, vagy ha aznap nem elérhető
    het_szam = day_date.isocalendar()[1]
    if het_szam != heti_o2_hetszam:
        heti_o2_hetszam = het_szam
        heti_o2_szemely = pick_weekly_o2_anchor(d)
        heti_anchor_by_het[het_szam] = heti_o2_szemely

    o2_candidate = None
    if O2_ALAP and is_available_for_O2(O2_ALAP, day_date, exclude=chosen):
        o2_candidate = O2_ALAP
    elif heti_o2_szemely and is_available_for_O2(heti_o2_szemely, day_date, exclude=chosen):
        o2_candidate = heti_o2_szemely
    else:
        o2_candidate = find_substitute_o2(day_date, exclude=chosen)

    o2 = None
    if o2_candidate:
        mar_ugyel = schedule[d].get("Stroke") == o2_candidate or schedule[d].get("Aneszt") == o2_candidate
        letszam_csokkenes = 0 if mar_ugyel else 1  # ha már úgyis St/A ügyeletben van, nem vesz el új főt a jelenléti körből
        or_pool_with_o2 = present_count - ito_present - (1 if o1 else 0) - letszam_csokkenes - mr_present
        if or_pool_with_o2 >= MUTO_PADLO:
            o2 = o2_candidate
            if het_szam == heti_o2_hetszam:
                heti_o2_szemely = o2

    if o1:
        o_assigned_count[o1] += 1
    if o2:
        o_assigned_count[o2] += 1
    o1_o2[d] = {"O1": o1, "O2": o2}

# ---------------------------------------------------------------------------
# sanity check
# ---------------------------------------------------------------------------
violations = []
for d in range(num_days):
    day_date = first_day + datetime.timedelta(days=d)
    names_today = [nm for nm in schedule[d].values() if nm]
    if len(names_today) != len(set(names_today)):
        violations.append(f"double-book on {day_date}")
    for duty, nm in schedule[d].items():
        if nm is None:
            continue
        pref = prefs.get((nm, day_date))
        if pref in ("Szabadság", "Nem szeretne"):
            violations.append(f"PREF VIOLATION {nm} {day_date} {pref} -> {duty}")
for name, *_ in staff:
    days_for_name = sorted(first_day + datetime.timedelta(days=d) for d in range(num_days)
                            for duty, nm in schedule[d].items() if nm == name)
    for i in range(1, len(days_for_name)):
        gap = (days_for_name[i] - days_for_name[i - 1]).days
        if gap < MIN_PIHENO:
            violations.append(f"{name} gap {gap} {days_for_name[i-1]}->{days_for_name[i]}")

print("Napok:", num_days)
print("Violations:", violations if violations else "NONE")
for name, *_ in staff:
    print(f"{name}: {assigned_count[name]} ügyelet")

# ---------------------------------------------------------------------------
# Excel írás
# ---------------------------------------------------------------------------
wb = openpyxl.load_workbook(SABLON_PATH)
ws = wb["Beosztás"]
ws["B3"] = first_day  # a naptár/hét-formulák ehhez igazodnak - a generált hónapra kell állítani
DATA_START = 6
col_map = {"Intenzív": 3, "Stroke": 4, "Aneszt": 5}
for d in range(num_days):
    row = DATA_START + d
    for duty, col in col_map.items():
        ws.cell(row=row, column=col, value=schedule[d].get(duty))
    ws.cell(row=row, column=6, value=o1_o2.get(d, {}).get("O1"))
    ws.cell(row=row, column=7, value=o1_o2.get(d, {}).get("O2"))

# Kívánságok lap feltöltése (a végleges, szabályokkal kiegészített prefs alapján)
ws_kiv = wb["Kívánságok"]
KIV_START = 5
kiv_row_of = {name: KIV_START + i for i, name in enumerate(staff_order_all)}
for (name, day_date), value in prefs.items():
    row = kiv_row_of.get(name)
    if row is None:
        continue
    col = 2 + (day_date - first_day).days
    ws_kiv.cell(row=row, column=col, value=value)

# Dolgozók lap "Kért ügyeletszám" oszlopának feltöltése a havi kívánság-fájlból
ws_staff = wb["Dolgozók"]
STAFF_START = 4
kert = KIV.get("kert_ugyeletszam", {})
for i, name in enumerate(staff_order_all):
    ws_staff.cell(row=STAFF_START + i, column=4, value=kert.get(name))

ws_print = wb["Nyomtatási beosztás"]
PRINT_HEADER_ROW = 4
PRINT_START = 5
staff_order = staff_order_all
row_of = {name: PRINT_START + i for i, name in enumerate(staff_order)}
muto_row = PRINT_START + len(staff_order)

WEEKEND_FILL = PatternFill("solid", fgColor="C6E0B4")
SZABADSAG_FILL = PatternFill("solid", fgColor="000000")
SZABADSAG_FONT = XLFont(name="Arial", size=9, color="FFFFFF")

m_count = {name: 0 for name in staff_order}
aktiv_nap_count = {name: 0 for name in staff_order}
hetvegi_ugyelet_count = {name: 0 for name in staff_order}
hetkoznapi_ugyelet_count = {name: 0 for name in staff_order}
lelepo_hetkoznap_count = {name: 0 for name in staff_order}
szabadsag_hetkoznap_count = {name: 0 for name in staff_order}
kulsos_hetkoznap_count = {name: 0 for name in staff_order}  # rezidens külsős gyakorlaton töltött hétköznapja (8 óra/nap jóváírás)
visszahivas_count = {name: 0 for name in REZIDENSEK}
napi_rate_of = {d["nev"]: d["napi_munkaido"] for d in SZAB["dolgozok"]}
raw_present_count_by_day = {}  # d (0-alapú) -> jelenlévők száma aznap, a fő ciklusból

for d in range(num_days):
    day_date = first_day + datetime.timedelta(days=d)
    col = 2 + d
    is_weekend = day_date.weekday() >= 5
    duty_today = {nm: code for code, nm in
                  {"I": schedule[d].get("Intenzív"), "St": schedule[d].get("Stroke"),
                   "A": schedule[d].get("Aneszt")}.items() if nm}
    o_today = o1_o2.get(d, {"O1": None, "O2": None})

    present_count = 0
    ito_present = 1 if schedule[d].get("Intenzív") else 0
    o1_present = 1 if o_today["O1"] else 0
    o2_present = 1 if o_today["O2"] else 0
    mr_present = 0

    if is_weekend:
        ws_print.cell(row=PRINT_HEADER_ROW, column=col).fill = WEEKEND_FILL

    for name in staff_order:
        parts = []
        if name in duty_today:
            parts.append(duty_today[name])
        if name == o_today["O1"]:
            parts.append("O1")
        if name == o_today["O2"]:
            parts.append("O2")
        fix_kod = heti_fix_esemeny_ma(name, day_date)
        if fix_kod and not is_szabadsag(name, day_date):
            parts.append(fix_kod)
            if day_date.weekday() < 5:
                mr_present = 1

        on_szabadsag = is_szabadsag(name, day_date)

        if parts:
            code = "/".join(parts)
        elif is_lelepo(name, day_date):
            code = "el"
        elif on_szabadsag:
            code = ""
        elif is_weekend:
            code = ""
        elif name in CSAK_JELOLT_NAPOKON and prefs.get((name, day_date)) == "Nem szeretne":
            code = ""
        elif name in RENDES_NAP_CSAK_HETENTE and day_date.weekday() not in RENDES_NAP_CSAK_HETENTE[name]:
            code = ""  # fix heti rendes napja van, ezen a hétköznapon nincs bent
        elif jelenlet_tiltott(name, day_date):
            code = ""
        elif nem_dolgozik_hetente_ma(name, day_date):
            code = ""
        elif name in RESZ_NAPI_ORASZAMOS and not kivansagok[name]["szeret"]:
            code = ""  # nincs megadott jó napja - a havi kerete kizárólag ügyeletből teljesül, "m" nap nélkül
        elif name in KULSOS_GYAKORLATON:
            code = ""  # rezidens külsős gyakorlaton - csak visszahívással kaphat "m"-et, ld. lentebb
        elif keret_of(name) != "Napi":
            code = ""
        else:
            code = "m"

        if code and code != "el":
            present_count += 1
            aktiv_nap_count[name] += 1
        if name in duty_today:
            if is_weekend:
                hetvegi_ugyelet_count[name] += 1
            else:
                hetkoznapi_ugyelet_count[name] += 1
        if code == "m":
            m_count[name] += 1
        if code == "el" and day_date.weekday() < 5:
            lelepo_hetkoznap_count[name] += 1
        if on_szabadsag and day_date.weekday() < 5:
            szabadsag_hetkoznap_count[name] += 1
        if name in KULSOS_GYAKORLATON and code == "" and day_date.weekday() < 5:
            kulsos_hetkoznap_count[name] += 1
        cell = ws_print.cell(row=row_of[name], column=col, value=code if code else None)
        if on_szabadsag:
            cell.fill = SZABADSAG_FILL
            cell.font = SZABADSAG_FONT
        elif is_weekend:
            cell.fill = WEEKEND_FILL

    if is_weekend:
        ws_print.cell(row=muto_row, column=col).fill = WEEKEND_FILL
    else:
        muto_val = present_count - ito_present - o1_present - o2_present - mr_present - MUTO_MIN
        ws_print.cell(row=muto_row, column=col, value=muto_val)
        raw_present_count_by_day[d] = present_count - ito_present - o1_present - o2_present - mr_present

# Pótlólagos "m" napok hozzáadása azoknak, akiknél a megadott "jó napok" nem elegek a
# havi kötelező óraszám eléréséhez - a jelölt napjaik előnyt élveznek (azokat a fő ciklus
# már beírta), de ha ez nem elég, bevonunk plusz, kifejezetten nem tiltott hétköznapokat
# is, amíg el nem éri a kapacitást (vagy el nem fogynak a lehetséges napok). Akinek fix
# heti mintája van (pl. Korompai: csak kedd-csütörtök), azt ez nem érinti - az egy más
# típusú, szándékosan korlátozott szabály.
for name in RESZ_NAPI_ORASZAMOS:
    if name in RENDES_NAP_CSAK_HETENTE:
        continue
    napi_rate = RESZ_NAPI_ORASZAMOS[name]
    kapacitas = RESZ_NAPI_KAPACITAS[name]
    current_nappali = (NAPI_KOTELEZO_ORA * (aktiv_nap_count[name] - hetvegi_ugyelet_count[name]) +
                        napi_rate * lelepo_hetkoznap_count[name] +
                        napi_rate * szabadsag_hetkoznap_count[name])
    if current_nappali >= kapacitas:
        continue
    for d in sorted(range(num_days), key=lambda x: raw_present_count_by_day.get(x, 999)):
        if current_nappali >= kapacitas:
            break
        day_date = first_day + datetime.timedelta(days=d)
        if day_date.weekday() >= 5:
            continue
        day_nap = d + 1
        if day_nap in EREDETI_KIFEJEZETT_NEM.get(name, set()):
            continue
        if is_szabadsag(name, day_date):
            continue
        col = 2 + d
        cell = ws_print.cell(row=row_of[name], column=col)
        if cell.value:
            continue
        cell.value = "m"
        m_count[name] += 1
        aktiv_nap_count[name] += 1
        current_nappali += NAPI_KOTELEZO_ORA
        raw_present_count_by_day[d] = raw_present_count_by_day.get(d, 0) + 1
        muto_cell = ws_print.cell(row=muto_row, column=col)
        if muto_cell.value is not None:
            muto_cell.value += 1

# Rezidens-visszahívás: aki külsős gyakorlaton van, azt csak akkor hívjuk vissza "m"-nek,
# ha a műtői minimum (6 fő) máshogy nem teljesülne - méltányosan elosztva a visszahívásokat
# a külsős rezidensek között (mindig a legkevesebbszer visszahívottat választva).
for d in range(num_days):
    day_date = first_day + datetime.timedelta(days=d)
    if day_date.weekday() >= 5:
        continue
    hiany = MUTO_PADLO - raw_present_count_by_day.get(d, 0)
    if hiany <= 0:
        continue
    col = 2 + d
    jeloltek = [nm for nm in REZIDENSEK if nm in KULSOS_GYAKORLATON]
    jeloltek = [nm for nm in jeloltek if ws_print.cell(row=row_of[nm], column=col).value is None
                and not is_szabadsag(nm, day_date) and not is_lelepo(nm, day_date)]
    jeloltek.sort(key=lambda nm: visszahivas_count[nm])
    for nm in jeloltek:
        if hiany <= 0:
            break
        cell = ws_print.cell(row=row_of[nm], column=col)
        cell.value = "m"
        m_count[nm] += 1
        aktiv_nap_count[nm] += 1
        visszahivas_count[nm] += 1
        kulsos_hetkoznap_count[nm] -= 1  # ezen a napon már nem "hiányzó", hanem "visszahívott"
        raw_present_count_by_day[d] = raw_present_count_by_day.get(d, 0) + 1
        muto_cell = ws_print.cell(row=muto_row, column=col)
        if muto_cell.value is not None:
            muto_cell.value += 1
        hiany -= 1

# Utólagos O2-pótlás: törekedni kell arra, hogy lehetőleg mindennap két osztályos (O1 és
# O2) is legyen, nem csak egy - ezt a fenti pótlólagos "m" feltöltés UTÁN kell megnézni,
# mert lehet, hogy csak az odakerült extra emberek miatt lesz elég a jelenléti létszám.
for d in range(num_days):
    day_date = first_day + datetime.timedelta(days=d)
    if day_date.weekday() >= 5:
        continue
    if o1_o2.get(d, {}).get("O2"):
        continue  # már van O2
    o1_today = o1_o2.get(d, {}).get("O1")
    pool_if_o2 = raw_present_count_by_day.get(d, 0) - 1
    if pool_if_o2 < MUTO_PADLO:
        continue
    exclude = {o1_today} if o1_today else set()
    het_szam_ma = day_date.isocalendar()[1]
    heti_anchor_ma = heti_anchor_by_het.get(het_szam_ma)
    if O2_ALAP and is_available_for_O2(O2_ALAP, day_date, exclude=exclude):
        candidate = O2_ALAP
    elif heti_anchor_ma and is_available_for_O2(heti_anchor_ma, day_date, exclude=exclude):
        candidate = heti_anchor_ma
    else:
        candidate = find_substitute_o2(day_date, exclude=exclude)
    if not candidate:
        continue
    o1_o2.setdefault(d, {"O1": o1_today, "O2": None})["O2"] = candidate
    o_assigned_count[candidate] += 1
    col = 2 + d
    cell = ws_print.cell(row=row_of[candidate], column=col)
    was_blank = cell.value in (None, "")
    if cell.value in (None, "", "m"):
        cell.value = "O2"
    else:
        cell.value = f"{cell.value}/O2"
    if was_blank:
        aktiv_nap_count[candidate] += 1
    raw_present_count_by_day[d] = raw_present_count_by_day.get(d, 0) - 1
    muto_cell = ws_print.cell(row=muto_row, column=col)
    if muto_cell.value is not None:
        muto_cell.value -= 1

for name in staff_order:
    duty_count = assigned_count.get(name, 0)
    napi_rate = napi_rate_of.get(name, NAPI_KOTELEZO_ORA)
    if keret_of(name) == "Napi":
        # hétvégi ügyelet: a teljes 24 óra ügyeletibe megy, nincs külön 8 órás nappali rész
        nappali_aktiv_napok = aktiv_nap_count[name] - hetvegi_ugyelet_count[name]
        nappali_total = (NAPI_KOTELEZO_ORA * nappali_aktiv_napok +
                          napi_rate * lelepo_hetkoznap_count[name] +
                          napi_rate * szabadsag_hetkoznap_count[name] +
                          napi_rate * kulsos_hetkoznap_count.get(name, 0))
        ugyeleti_total = 16 * hetkoznapi_ugyelet_count[name] + 24 * hetvegi_ugyelet_count[name]
        ws_print.cell(row=row_of[name], column=33, value=nappali_total)
    else:
        ugyeleti_total = 24 * duty_count
    ws_print.cell(row=row_of[name], column=34, value=ugyeleti_total)

    # Óraelszámolás lap szinkronban tartása ugyanezzel a (helyes, teljes) számítással
    ws_hr = wb["Óraelszámolás"]
    hr_row = row_of[name]  # ugyanaz a sor-illeszkedés, mint a Dolgozók/Nyomtatási beosztás lapokon
    if keret_of(name) == "Napi":
        ws_hr.cell(row=hr_row, column=4, value=nappali_total)
        ws_hr.cell(row=hr_row, column=5, value=ugyeleti_total)
    else:
        ws_hr.cell(row=hr_row, column=6, value=ugyeleti_total)

wb.save(KIMENET_PATH)
print("saved", KIMENET_PATH, f"(seed={SEED})")

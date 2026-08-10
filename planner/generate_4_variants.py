#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
4 beosztás-változat generálása egyszerre - mindegyik ugyanabból a szabalyok.json +
kivansagok_ÉÉÉÉ_HH.json alapból indul, de más-más magszámmal (seed), így a döntetlen
jelöltek közötti választás máshogy dől el, valóban eltérő (de mind szabálykövető)
beosztásokat eredményezve.

Használat: python3 generate_4_variants.py [kivansagok_fajl.json] [sablon.xlsx] [kimenet_elotag]
"""
import sys
import subprocess
import os

KIVANSAGOK = sys.argv[1] if len(sys.argv) > 1 else "/home/claude/kivansagok_2026_08.json"
SABLON = sys.argv[2] if len(sys.argv) > 2 else "ICU_ugyeleti_beosztas.xlsx"
ELOTAG = sys.argv[3] if len(sys.argv) > 3 else "ICU_ugyeleti_beosztas_probabeosztas"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
GENERATOR = os.path.join(SCRIPT_DIR, "generate_schedule.py")

for seed in (1, 2, 3, 4):
    kimenet = f"{ELOTAG}_valtozat{seed}.xlsx"
    print(f"\n=== {seed}. változat ===")
    subprocess.run(
        [sys.executable, GENERATOR, KIVANSAGOK, SABLON, kimenet, str(seed)],
        check=True,
    )

print("\nKész: 4 változat elmentve.")

import json
import glob

osszes = []
for fajl in glob.glob("*_javasolt_kivetelek.json"):
    osszes.extend(json.load(open(fajl, encoding="utf-8")))

egyedi = {json.dumps(x, sort_keys=True): x for x in osszes}
json.dump(list(egyedi.values()), open("javasolt_kivetelek_osszevont.json", "w", encoding="utf-8"), ensure_ascii=False, indent=2)
print(f"{len(egyedi)} egyedi javasolt kivétel összegyűjtve.")

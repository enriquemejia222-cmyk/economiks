"""Revisa que los datos nuevos tengan sentido antes de publicarlos. Sale con error si no."""
import json, subprocess, sys, os
d = json.load(open("data.json"))
h = json.load(open("historial_riesgo.json"))
try:
    old = json.loads(subprocess.run(["git", "show", "HEAD:historial_riesgo.json"], capture_output=True, text=True, check=True).stdout)
except Exception:
    old = {}
problems = []
if os.path.getsize("data.json") < 80_000: problems.append("data.json demasiado pequeño")
if missing := set(old) - set(h): problems.append(f"el historial perdió {len(missing)} días")
for k in ("DOM", "LAT"):
    if not isinstance(d["risk"].get(k, {}).get("last"), (int, float)): problems.append(f"falta el riesgo país de {k}")
if sum(1 for k in d["risk"] if k in d["countries"]) < 15: problems.append("faltan países en riesgo país")
if problems:
    print("NO SE PUBLICA:", "; ".join(problems)); sys.exit(1)
print("Datos OK:", d["riskDate"], "RD", d["risk"]["DOM"]["last"], "Latam", d["risk"]["LAT"]["last"], "| días en historial:", len(h))

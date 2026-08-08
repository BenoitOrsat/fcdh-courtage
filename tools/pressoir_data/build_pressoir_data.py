#!/usr/bin/env python3
"""
Pipeline de génération des données pressoirs pour pressoir.html, à partir du
fichier source "V.2026 Répartion des pressoirs par courtiers.xlsx" (colonnes :
Sect, Courtier1, Courtier2, NouveauCAP, Ville, Pressoir).

Usage :
    python3 build_pressoir_data.py <source.xlsx> <pressoir.html> [--source-modified-time ISO8601]

Ce que fait le script :
  1. Lit et nettoie les lignes du xlsx (ignore les lignes vides/placeholder,
     déduplique les lignes strictement identiques).
  2. Normalise les noms de commune (espaces/tirets/CEDEX) et résout les cas
     particuliers via overrides.json (lieux-dits rattachés à une commune,
     variantes standalone).
  3. Géocode chaque commune via geo.api.gouv.fr (filtré aux départements
     champenois, préférence à la correspondance exacte), en réutilisant
     geocode_cache.json pour ne jamais re-géocoder une commune déjà résolue.
     Toute commune qui nécessite un NOUVEAU géocodage (absente du cache) est
     signalée en fin d'exécution pour relecture humaine.
  4. Applique la règle métier "Nouveau CAP" : si renseigné, le nouveau
     courtier devient courtier1, l'ancien passe en courtier2.
  5. Génère le bloc JS `var communes = [...]` et le remplace dans
     pressoir.html.
  6. Met à jour geocode_cache.json et affiche un résumé (communes/pressoirs
     totaux, nouveautés, avertissements).

Ce script NE COMMIT NI NE POUSSE RIEN sur git — il modifie seulement le
fichier local pressoir.html. Relire le diff (git diff pressoir.html) avant
de committer, en particulier les communes signalées comme nouvellement
géocodées ou en échec.
"""

import sys
import os
import re
import json
import time
import unicodedata
import argparse

try:
    import requests
except ImportError:
    sys.exit("Le module 'requests' est requis (pip install requests).")

try:
    import openpyxl
except ImportError:
    sys.exit("Le module 'openpyxl' est requis (pip install openpyxl).")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OVERRIDES_PATH = os.path.join(SCRIPT_DIR, "overrides.json")
CACHE_PATH = os.path.join(SCRIPT_DIR, "geocode_cache.json")
STATE_PATH = os.path.join(SCRIPT_DIR, "last_import_state.json")

CHAMPAGNE_DEPTS = {"02", "08", "10", "51", "52", "77"}
COURTIERS = ["BO", "ASM", "JPNC", "EB", "JBC", "EC", "EH", "MAA"]


# =============================================================================
# Normalisation des noms de commune
# =============================================================================
def strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def norm_key(s):
    s = s.strip()
    s = re.sub(r"(?i)\bcedex\b\.?\s*\d*", "", s)
    s = strip_accents(s).lower()
    s = re.sub(r"[-']", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_cell(x):
    if x is None:
        return None
    if isinstance(x, str):
        x = x.strip()
        return x if x else None
    return x


# =============================================================================
# Lecture du xlsx
# =============================================================================
def read_rows(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Tableau"]

    # La zone de données utile s'arrête avant le bloc de synthèse (tableaux
    # croisés en bas de la feuille) : on s'arrête à la première ligne
    # entièrement vide suivie d'un long passage vide, en pratique on
    # s'arrête dès qu'on rencontre une ligne dont la colonne Sect contient
    # un intitulé de synthèse ("Courtier", "TOTAL", ...) plutôt qu'un vrai
    # secteur, ou une longue série de lignes vides.
    SUMMARY_MARKERS = {"Courtier", "TOTAL", "Nouveau CAP \\ Origine",
                       "Courtier1 (perd) \\ Nouveau CAP (gagne)",
                       "Courtier1 (perd) \\ Courtier2 (gagne)"}

    rows = []
    empty_streak = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        sect = clean_cell(r[0])
        if sect in SUMMARY_MARKERS:
            break
        if all(clean_cell(v) is None for v in r):
            empty_streak += 1
            if empty_streak > 20:
                break
            continue
        empty_streak = 0

        c1, c2, cap, ville, pressoir = (clean_cell(r[1]), clean_cell(r[2]),
                                        clean_cell(r[3]), clean_cell(r[4]), clean_cell(r[5]))
        if not ville or not pressoir or pressoir in ("...", "..", "."):
            continue
        rows.append({
            "sect": sect, "courtier1": c1, "courtier2": c2, "nouveauCap": cap,
            "ville_raw": ville, "pressoir": pressoir,
        })

    # Dédup des lignes strictement identiques (erreur de saisie dans le
    # fichier source, déjà rencontrée une fois).
    seen = set()
    deduped = []
    dupes_removed = 0
    for r in rows:
        sig = (r["ville_raw"], r["pressoir"].strip().upper(), r["courtier1"],
               r["courtier2"], r["nouveauCap"], r["sect"])
        if sig in seen:
            dupes_removed += 1
            continue
        seen.add(sig)
        deduped.append(r)

    return deduped, dupes_removed


# =============================================================================
# Résolution des communes (overrides + regroupement)
# =============================================================================
def load_json(path, default):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return default


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, sort_keys=True)


def resolve_search_string(ville_raw, overrides):
    raw_key = norm_key(ville_raw)
    dash_overrides = overrides.get("dash_overrides", {})
    standalone_overrides = overrides.get("standalone_overrides", {})
    if raw_key in dash_overrides:
        return dash_overrides[raw_key]
    if raw_key in standalone_overrides:
        return standalone_overrides[raw_key]
    return re.sub(r"(?i)\bcedex\b\.?\s*\d*", "", ville_raw).strip()


# =============================================================================
# Géocodage
# =============================================================================
def geocode_commune(search):
    """Géocode une commune via geo.api.gouv.fr, filtré aux départements
    champenois, en préférant une correspondance exacte du nom."""
    r = requests.get(
        "https://geo.api.gouv.fr/communes",
        params={"nom": search, "fields": "nom,code,centre,codeDepartement", "boost": "population"},
        timeout=10,
    )
    r.raise_for_status()
    candidates = r.json()
    in_region = [c for c in candidates if c.get("codeDepartement") in CHAMPAGNE_DEPTS]
    if not in_region:
        return None
    qn = norm_key(search)
    exact = [c for c in in_region if norm_key(c["nom"]) == qn]
    chosen = exact[0] if exact else in_region[0]
    return {
        "label": chosen["nom"],
        "lat": chosen["centre"]["coordinates"][1],
        "lng": chosen["centre"]["coordinates"][0],
        "ambiguous": not exact and len(in_region) > 1,
    }


def resolve_all_communes(rows, overrides, cache):
    """Retourne {canon_key: {label, lat, lng}} pour toutes les communes des
    lignes fournies, en réutilisant le cache et en géocodant les nouvelles."""
    geo = {}
    new_geocoded = []
    failed = []

    search_by_key = {}
    for r in rows:
        key = norm_key(resolve_search_string(r["ville_raw"], overrides))
        r["canon_key"] = key
        search_by_key.setdefault(key, r["ville_raw"])

    for key, raw_variant in search_by_key.items():
        if key in cache:
            geo[key] = cache[key]
            continue
        search = resolve_search_string(raw_variant, overrides)
        try:
            result = geocode_commune(search)
        except Exception as e:
            failed.append((key, search, str(e)))
            continue
        if result is None:
            failed.append((key, search, "aucune commune trouvée dans les départements champenois"))
            continue
        entry = {"label": result["label"], "lat": result["lat"], "lng": result["lng"]}
        geo[key] = entry
        cache[key] = entry
        new_geocoded.append((key, search, result["label"], result["ambiguous"]))
        time.sleep(0.03)

    return geo, new_geocoded, failed


# =============================================================================
# Construction des données finales (règle Nouveau CAP incluse)
# =============================================================================
def build_communes(rows, geo):
    communes_map = {}
    for r in rows:
        key = r["canon_key"]
        g = geo.get(key)
        if g is None:
            continue  # commune en échec de géocodage : exclue, signalée séparément
        if key not in communes_map:
            communes_map[key] = {"nom": g["label"], "lat": g["lat"], "lng": g["lng"], "pressoirs": []}

        if r["nouveauCap"]:
            c1, c2 = r["nouveauCap"], r["courtier1"]
        else:
            c1, c2 = r["courtier1"], r["courtier2"]

        pressoir = {"nom": r["pressoir"].strip(), "courtier1": c1}
        if c2:
            pressoir["courtier2"] = c2
        if r["sect"]:
            pressoir["secteur"] = r["sect"]
        communes_map[key]["pressoirs"].append(pressoir)

    return sorted(communes_map.values(), key=lambda c: c["nom"])


# =============================================================================
# Génération du bloc JS et injection dans pressoir.html
# =============================================================================
def esc(s):
    return s.replace("\\", "\\\\").replace('"', '\\"')


def generate_js_block(communes):
    lines = ["  var communes = ["]
    for ci, c in enumerate(communes):
        lines.append("    {")
        lines.append('      nom: "%s", lat: %s, lng: %s,' % (esc(c["nom"]), c["lat"], c["lng"]))
        lines.append("      pressoirs: [")
        for pi, p in enumerate(c["pressoirs"]):
            parts = ['nom: "%s"' % esc(p["nom"]), 'courtier1: "%s"' % p["courtier1"]]
            if p.get("courtier2"):
                parts.append('courtier2: "%s"' % p["courtier2"])
            if p.get("secteur"):
                parts.append('secteur: "%s"' % p["secteur"])
            comma = "," if pi < len(c["pressoirs"]) - 1 else ""
            lines.append("        { %s }%s" % (", ".join(parts), comma))
        lines.append("      ]")
        comma = "," if ci < len(communes) - 1 else ""
        lines.append("    }%s" % comma)
    lines.append("  ];")
    return "\n".join(lines)


def splice_into_html(html_path, new_block):
    with open(html_path, encoding="utf-8") as f:
        lines = f.readlines()

    start_idx = end_idx = None
    for i, line in enumerate(lines):
        if line.strip() == "var communes = [":
            start_idx = i
            break
    if start_idx is None:
        sys.exit("Impossible de trouver 'var communes = [' dans " + html_path)
    for i in range(start_idx, len(lines)):
        if lines[i].strip() == "];":
            end_idx = i
            break
    if end_idx is None:
        sys.exit("Impossible de trouver la fin du tableau communes dans " + html_path)

    new_lines = lines[:start_idx] + [new_block + "\n"] + lines[end_idx + 1:]
    with open(html_path, "w", encoding="utf-8") as f:
        f.writelines(new_lines)


# =============================================================================
# Main
# =============================================================================
def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("xlsx_path", help="Chemin vers le fichier source .xlsx")
    parser.add_argument("html_path", help="Chemin vers pressoir.html à mettre à jour")
    parser.add_argument("--source-modified-time", default=None,
                         help="Horodatage ISO8601 de dernière modification du fichier source (pour last_import_state.json)")
    parser.add_argument("--source-file-id", default=None,
                         help="Identifiant du fichier source (ex. Google Drive fileId), pour last_import_state.json")
    args = parser.parse_args()

    overrides = load_json(OVERRIDES_PATH, {"dash_overrides": {}, "standalone_overrides": {}})
    cache = load_json(CACHE_PATH, {})

    rows, dupes_removed = read_rows(args.xlsx_path)
    geo, new_geocoded, failed = resolve_all_communes(rows, overrides, cache)
    communes = build_communes(rows, geo)

    save_json(CACHE_PATH, cache)

    excluded_rows = [r for r in rows if r["canon_key"] not in geo]
    total_pressoirs = sum(len(c["pressoirs"]) for c in communes)

    courtier_counts = {k: 0 for k in COURTIERS}
    for c in communes:
        for p in c["pressoirs"]:
            if p["courtier1"] in courtier_counts:
                courtier_counts[p["courtier1"]] += 1

    js_block = generate_js_block(communes)
    splice_into_html(args.html_path, js_block)

    if args.source_modified_time or args.source_file_id:
        state = load_json(STATE_PATH, {})
        if args.source_modified_time:
            state["lastImportedModifiedTime"] = args.source_modified_time
        if args.source_file_id:
            state["sourceFileId"] = args.source_file_id
        state["lastImportedAt"] = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
        save_json(STATE_PATH, state)

    print("=" * 70)
    print("IMPORT TERMINÉ")
    print("=" * 70)
    print("Communes : %d" % len(communes))
    print("Pressoirs : %d" % total_pressoirs)
    print("Lignes dupliquées supprimées : %d" % dupes_removed)
    print("Répartition par courtier1 : %s" % courtier_counts)

    if new_geocoded:
        print("\n⚠️  %d commune(s) NOUVELLEMENT géocodée(s) (à relire) :" % len(new_geocoded))
        for key, search, label, ambiguous in new_geocoded:
            flag = "  [AMBIGU — pas de correspondance exacte, vérifier]" if ambiguous else ""
            print("   - %-30s -> %s%s" % (search, label, flag))

    if failed:
        print("\n❌ %d commune(s) NON géocodée(s) (exclues des données, à corriger manuellement) :" % len(failed))
        for key, search, reason in failed:
            print("   - %-30s : %s" % (search, reason))

    if excluded_rows:
        print("\n%d pressoir(s) exclu(s) faute de géocodage de leur commune." % len(excluded_rows))

    if not new_geocoded and not failed:
        print("\nAucune nouvelle commune à géocoder — toutes déjà en cache.")

    print("\nRelire le diff avant de committer : git diff -- " + os.path.relpath(args.html_path))


if __name__ == "__main__":
    main()

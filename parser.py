#!/usr/bin/env python3
"""
POMA Industrial Drawing PDF Parser - VERSION ADAPTÉE DYNAMIQUE
Extrait le cartouche et la nomenclature (BOM) depuis un PDF vectoriel POMA.
Aucun traitement d'image — extraction de texte pur par coordonnées.
Supporte les formats A0, A1, A2, A3, A4 et les templates nouveau/ancien.
ADAPTATIONS:
- Parsing dynamique des colonnes BOM basé sur les labels du header (pas de hardcoded ranges ou noms)
- Les noms des champs sont lus depuis les labels du PDF et utilisés comme clés dans le JSON
- Nettoyage automatique des noms de colonnes
- Conversion automatique des valeurs (int/float/string)
- Pas besoin d'IA pour détection, car PDF vectoriel avec texte extractible
Usage:
    python poma_parser_dynamic.py <fichier.pdf> [output.json]
"""
import sys
import json
import pdfplumber
from collections import defaultdict

# ─────────────────────────────────────────────────────────────────
# Dimensions des formats (en points @ 72 dpi)
# ─────────────────────────────────────────────────────────────────
PAGE_FORMATS = {
    "A0": (3370, 2384),
    "A1": (2384, 1684),
    "A2": (1684, 1191),
    "A3": (1191, 842),
    "A4": (842, 595),
}
FORMAT_TOLERANCE = 50

# ─────────────────────────────────────────────────────────────────
# Détection du format et du template
# ─────────────────────────────────────────────────────────────────
def detect_page_format(page):
    """Détecte le format de la page (A0 à A4)."""
    width = page.width
    height = page.height
    for fmt, (w, h) in PAGE_FORMATS.items():
        # Paysage
        if abs(width - w) < FORMAT_TOLERANCE and abs(height - h) < FORMAT_TOLERANCE:
            return fmt, "landscape"
        # Portrait
        if abs(width - h) < FORMAT_TOLERANCE and abs(height - w) < FORMAT_TOLERANCE:
            return fmt, "portrait"
    return f"INCONNU ({width:.0f}x{height:.0f})", "unknown"

def detect_template_type(words):
    """
    Détecte le type de template:
    - 'new': Nouveau format avec SECTEUR D'ACTIVITÉ
    - 'legacy': Ancien format bilingue avec ORDER/DATE/PIECE
    """
    texts = [w["text"].upper() for w in words]
    if any("SECTEUR" in t for t in texts):
        return "new"
    if any("ORDER" in t for t in texts) or any("COMMANDE" in t for t in texts):
        return "legacy"
    return "new"  # Par défaut

def detect_company(words):
    """
    Détecte le nom de la société depuis le texte du copyright.
    Cherche 'PROPERTY OF XXXX.' ou 'XXXX' comme logo.
    """
    for i, w in enumerate(words):
        if w["text"].upper() == "OF" and i + 1 < len(words):
            next_word = words[i + 1]["text"]
            if next_word.endswith("."):
                return next_word[:-1].upper()
            return next_word.upper()
    # Chercher POMA directement
    for w in words:
        if w["text"].upper() in ("POMA", "POMA."):
            return "POMA"
    return "INCONNU"

# ─────────────────────────────────────────────────────────────────
# Utilitaires
# ─────────────────────────────────────────────────────────────────
def group_into_lines(words, y_tol=4):
    """Regroupe les mots par ligne (même Y ± y_tol)."""
    buckets = defaultdict(list)
    for w in words:
        key = round(w["top"] / y_tol) * y_tol
        buckets[key].append(w)
    return {y: sorted(ws, key=lambda w: w["x0"]) for y, ws in sorted(buckets.items())}

def words_in_box(words, x0, y0, x1, y1):
    """Un mot est dans la boîte si son coin supérieur-gauche est dans la zone."""
    return [w for w in words
            if w["x0"] >= x0 and w["x0"] <= x1
            and w["top"] >= y0 and w["top"] <= y1]

def find_word(words, text):
    t = text.upper()
    return next((w for w in words if w["text"].upper() == t), None)

def text_in_box(words, x0, y0, x1, y1):
    ws = words_in_box(words, x0, y0, x1, y1)
    return " ".join(w["text"] for w in sorted(ws, key=lambda w: w["x0"])).strip()

def assign_column(x, column_ranges):
    for name, x0, x1 in column_ranges:
        if x0 <= x < x1:
            return name
    return None

def find_bom_header_y(words):
    rep_w = find_word(words, "REP.") or find_word(words, "REP")
    return rep_w["top"] if rep_w else None

def parse_bom(words, header_y, page_format, template_type, page_width):
    """Parse la nomenclature (BOM) de manière dynamique en détectant les labels des colonnes."""
    # Déterminer y_tol et merge_dist basés sur le format et template
    if page_format == "A4":
        y_tol = 3
        merge_dist = 5
    elif page_format in ["A0", "A1"]:
        y_tol = 5
        merge_dist = 10
    else:  # A2, A3
        y_tol = 4
        merge_dist = 8
    if template_type == "legacy":
        y_tol += 2  # Augmenté pour legacy comme dans l'exemple A3

    # Extraire les mots du header (ligne contenant REP.)
    header_words = [w for w in words if abs(w["top"] - header_y) < y_tol / 2]
    sorted_header_words = sorted(header_words, key=lambda w: w["x0"])

    # Fusionner les mots proches pour les headers multi-mots (ex: "MASSE (kg)")
    merged_headers = []
    for w in sorted_header_words:
        if merged_headers and w["x0"] - merged_headers[-1]["x1"] < merge_dist:
            merged_headers[-1]["text"] += " " + w["text"]
            merged_headers[-1]["x1"] = w["x1"]
        else:
            merged_headers.append({"text": w["text"], "x0": w["x0"], "x1": w["x1"]})

    # Créer les noms de colonnes normalisés
    column_names = []
    for mh in merged_headers:
        name = mh["text"].lower().rstrip(".").replace(" (kg)", "_kg").replace("(", "").replace(")", "").replace(" ", "_")
        column_names.append(name)

    # Calculer les plages de colonnes basées sur les positions des headers
    ranges = []
    for i in range(len(merged_headers)):
        x0 = merged_headers[i]["x0"]
        x1 = merged_headers[i + 1]["x0"] if i + 1 < len(merged_headers) else page_width + 10
        ranges.append((column_names[i], x0, x1))

    # Mots dans la zone BOM (au-dessus du header, comme dans la version originale)
    bom_words = [w for w in words
                 if w["top"] < header_y - 2
                 and w["x0"] >= ranges[0][1] - 10
                 and w["x0"] <= ranges[-1][2] + 10]

    lines = group_into_lines(bom_words, y_tol=y_tol)

    rows = []
    for y, line_words in lines.items():
        # Vérifier qu'un entier est dans la plage de la première colonne (rep)
        rep_range = ranges[0]
        rep_candidates = [w for w in line_words
                          if rep_range[1] - 10 <= w["x0"] < rep_range[2] + 10
                          and w["text"].isdigit()]
        if not rep_candidates:
            continue
        rep_text = rep_candidates[0]["text"].strip()

        # Affecter chaque mot à sa colonne
        row_data = defaultdict(list)
        for w in line_words:
            col = assign_column(w["x0"], ranges)
            if col:
                row_data[col].append(w["text"])

        # Construire la ligne avec conversion automatique des types
        row = {}
        for key in column_names:
            val_str = " ".join(row_data.get(key, [])).strip()
            if val_str:
                try:
                    row[key] = int(val_str)
                    continue
                except ValueError:
                    pass
                try:
                    row[key] = float(val_str.replace(",", "."))
                    continue
                except ValueError:
                    pass
                row[key] = val_str

        # Filtrer les lignes fantômes (sans article si présent)
        if "article" in row and not row["article"]:
            continue

        # Nettoyage spécifique pour niveau_de_securite si présent
        if "niveau_de_securite" in row:
            parts = row["niveau_de_securite"].split()
            clean = [p for p in parts if not (len(p) == 1 and p.isupper() and p.isalpha())]
            row["niveau_de_securite"] = " ".join(clean)

        rows.append(row)

    # Trier par la première colonne (rep) décroissant
    first_col = column_names[0]
    rows.sort(key=lambda r: -(r.get(first_col, 0) if isinstance(r.get(first_col), (int, float)) else 0))

    return rows

# ─────────────────────────────────────────────────────────────────
# Cartouche - Nouveau template (SECTEUR D'ACTIVITÉ)
# ─────────────────────────────────────────────────────────────────
TITLE_BLOCK_COORDS = {
    "A0": {
        "secteur": (2773, 2199, 3100, 2215),
        "tolerances": (3130, 2199, 3355, 2215),
        "designation": (2773, 2222, 3100, 2240),
        "orientation": (2773, 2238, 3100, 2260),
        "util_principale": (2773, 2262, 3100, 2280),
        "carac1": (3130, 2220, 3355, 2240),
        "carac2": (3130, 2244, 3355, 2260),
        "carac3": (3130, 2264, 3355, 2280),
        "echelle": (2815, 2283, 2858, 2298),
        "format": (2856, 2283, 2892, 2298),
        "pages": (2885, 2283, 2918, 2298),
        "execution": (2918, 2283, 2998, 2298),
        "date_valid": (3030, 2283, 3190, 2298),
        "masse": (3185, 2283, 3260, 2298),
        "copie_de": (3268, 2283, 3355, 2298),
        "niveau_sec": (2930, 2312, 2975, 2325),
        "cree_par": (3025, 2304, 3165, 2320),
        "approuve_par": (3020, 2325, 3165, 2345),
        "numero_art": (3165, 2310, 3355, 2325),
    },
    "A2": {
        "secteur": (1091, 1008, 1250, 1022),
        "tolerances": (1445, 1008, 1580, 1022),
        "designation": (1091, 1030, 1250, 1045),
        "orientation": (1091, 1045, 1250, 1065),
        "util_principale": (1091, 1072, 1250, 1088),
        "carac1": (1445, 1030, 1580, 1044),
        "carac2": (1445, 1052, 1580, 1065),
        "carac3": (1445, 1074, 1580, 1086),
        "echelle": (1130, 1092, 1170, 1105),
        "format": (1170, 1092, 1200, 1105),
        "pages": (1200, 1092, 1230, 1105),
        "execution": (1260, 1092, 1340, 1105),
        "date_valid": (1340, 1092, 1420, 1105),
        "masse": (1495, 1092, 1545, 1105),
        "copie_de": (1585, 1092, 1650, 1105),
        "niveau_sec": (1250, 1118, 1295, 1135),
        "cree_par": (1320, 1112, 1430, 1128),
        "approuve_par": (1340, 1132, 1430, 1150),
        "numero_art": (1480, 1115, 1620, 1130),
    },
    "A4": {
        "secteur": (15, 673, 200, 682),
        "tolerances": (370, 673, 450, 682),
        "designation": (15, 695, 200, 710),
        "orientation": (15, 710, 200, 725),
        "util_principale": (15, 740, 60, 752),
        "carac1": (370, 693, 450, 708),
        "carac2": (370, 715, 450, 728),
        "carac3": (370, 736, 450, 750),
        "echelle": (60, 758, 95, 770),
        "format": (95, 758, 128, 770),
        "pages": (128, 758, 155, 770),
        "execution": (158, 758, 250, 770),
        "date_valid": (265, 758, 345, 770),
        "masse": (420, 758, 475, 770),
        "copie_de": (510, 758, 565, 770),
        "niveau_sec": (175, 785, 225, 800),
        "cree_par": (265, 778, 345, 792),
        "approuve_par": (260, 800, 345, 815),
        "numero_art": (400, 782, 550, 795),
    },
}

def interpolate_coords(coords_high, ratio):
    """Interpole les coordonnées entre deux formats."""
    result = {}
    for key in coords_high:
        x0_h, y0_h, x1_h, y1_h = coords_high[key]
        result[key] = (
            int(x0_h * ratio),
            int(y0_h * ratio),
            int(x1_h * ratio),
            int(y1_h * ratio),
        )
    return result

def get_title_block_coords(page_format, page_width):
    """Retourne les coordonnées du cartouche pour le format donné."""
    if page_format in TITLE_BLOCK_COORDS:
        return TITLE_BLOCK_COORDS[page_format]
    if page_format == "A1":
        ratio = page_width / 3370
        return interpolate_coords(TITLE_BLOCK_COORDS["A0"], ratio)
    if page_format == "A3":
        ratio = page_width / 1684
        return interpolate_coords(TITLE_BLOCK_COORDS["A2"], ratio)
    return TITLE_BLOCK_COORDS["A4"]

def parse_title_block_new(words, page_format, page_width, company):
    """Parse le cartouche du nouveau template."""
    coords = get_title_block_coords(page_format, page_width)
    def box(key):
        x0, y0, x1, y1 = coords[key]
        return text_in_box(words, x0, y0, x1, y1)

    secteur = box("secteur")
    tolerances = box("tolerances")
    designation = box("designation")
    orientation = box("orientation")
    util_principale = box("util_principale")
    carac1 = box("carac1")
    carac2 = box("carac2")
    carac3 = box("carac3")
    echelle = box("echelle")
    format_ = box("format")
    pages = box("pages")
    execution = box("execution")
    date_valid = box("date_valid")
    masse_raw = box("masse")
    copie_de = box("copie_de")
    niveau_sec = box("niveau_sec")
    cree_par = box("cree_par")
    approuve_par = box("approuve_par")
    numero_art = box("numero_art")

    # Masse → décomposer valeur + unité
    masse_parts = masse_raw.split()
    try:
        masse_valeur = float(masse_parts[0].replace(",", "."))
        masse_unite = masse_parts[1] if len(masse_parts) > 1 else "kg"
    except Exception:
        masse_valeur = masse_raw
        masse_unite = "kg"

    return {
        "societe": company,
        "secteur_activite": secteur,
        "tolerances_generales": tolerances,
        "description_piece": {
            "designation": designation,
            "orientation": orientation,
            "utilisation_principale": util_principale,
            "caracteristiques": {
                "caracteristique_1": carac1,
                "caracteristique_2": carac2,
                "caracteristique_3": carac3,
            },
        },
        "infos_plan": {
            "echelle": echelle,
            "format": format_,
            "pages": pages,
            "execution": execution,
            "date_de_validation": date_valid,
            "niveau_de_securite": niveau_sec,
        },
        "validation": {
            "cree_par": cree_par,
            "approuve_par": approuve_par,
        },
        "identifiants": {
            "numero_article": numero_art,
            "copie_de": copie_de,
        },
        "masse_totale": {
            "valeur": masse_valeur,
            "unite": masse_unite,
        },
    }

# ─────────────────────────────────────────────────────────────────
# Cartouche - Ancien template A3 bilingue (gardé en hardcoded pour le moment)
# ─────────────────────────────────────────────────────────────────
def parse_title_block_legacy(words, page_format, page_width, company):
    """Parse le cartouche de l'ancien template bilingue A3."""
    def box(x0, y0, x1, y1):
        return text_in_box(words, x0, y0, x1, y1)

    # Format A3 paysage (1191 x 842) - Coordonnées corrigées
    secteur_en = box(595, 650, 680, 665)  # "CABLE WAY"
    secteur_fr = box(635, 660, 750, 680)  # "INSTALLATION A CABLE"
    secteur = secteur_fr if secteur_fr else secteur_en

    designation_en = box(635, 677, 850, 695)  # "REINFORCEMENT-SUPPORT ARM"
    designation_fr = box(635, 697, 850, 715)  # "RENFORT-BRAS SUPPORT"
    designation = designation_en if designation_en else designation_fr

    dim_value = ""
    for w in words:
        if 'X' in w["text"] and w["text"].replace('X', '').replace('x', '').isdigit():
            if 595 <= w["x0"] <= 700 and 715 <= w["top"] <= 735:
                dim_value = w["text"]
                break

    spec1 = box(735, 715, 850, 730)
    spec5 = box(950, 715, 1000, 730)

    format_ = box(705, 745, 730, 760)
    if format_:
        format_ = format_.replace("FORMAT", "").replace("/", "").strip()

    date_valid = ""
    for w in words:
        if "-" in w["text"] and len(w["text"]) >= 9:
            parts = w["text"].split("-")
            if len(parts) == 3 and parts[0].isdigit():
                date_valid = w["text"]
                break

    cree_par = box(835, 745, 945, 760)
    echelle = box(970, 745, 1000, 760)
    masse_raw = box(1030, 745, 1080, 760)
    copie_de = box(1100, 745, 1160, 760)
    approuve_par = box(838, 765, 945, 780)
    numero_art = box(980, 772, 1120, 805)
    if numero_art:
        numero_art = ''.join(c for c in numero_art if c.isdigit())

    niveau_sec = box(756, 745, 785, 760)
    if not niveau_sec:
        for w in words:
            if '--' in w["text"] or (len(w["text"]) <= 4 and '-' in w["text"]):
                if 730 <= w["x0"] <= 800 and 745 <= w["top"] <= 800:
                    niveau_sec = w["text"]
                    break

    tolerances = ""
    for w in words:
        if "ISO" in w["text"].upper() or "2768" in w["text"] or "9013" in w["text"]:
            if 850 <= w["x0"] <= 950 and 778 <= w["top"] <= 800:
                tolerances += w["text"] + " "
    tolerances = tolerances.strip()
    if not tolerances:
        tolerances = box(850, 778, 950, 800)

    masse_parts = masse_raw.split() if masse_raw else []
    try:
        masse_valeur = float(masse_parts[0].replace(",", "."))
        masse_unite = masse_parts[1] if len(masse_parts) > 1 else "kg"
    except Exception:
        masse_valeur = masse_raw if masse_raw else None
        masse_unite = "kg"

    return {
        "societe": company,
        "secteur_activite": secteur,
        "tolerances_generales": tolerances,
        "description_piece": {
            "designation": designation,
            "designation_en": designation_en,
            "designation_fr": designation_fr,
            "orientation": "",
            "utilisation_principale": "",
            "caracteristiques": {
                "caracteristique_1": dim_value if isinstance(dim_value, str) else "",
                "caracteristique_2": spec1.replace("SPECIFICATION-1", "").strip() if spec1 else "",
                "caracteristique_3": spec5.replace("SPECIFICATION-5", "").strip() if spec5 else "",
            },
        },
        "infos_plan": {
            "echelle": echelle,
            "format": format_.replace("A3", "A3").strip() if format_ else "A3",
            "pages": "",
            "execution": "",
            "date_de_validation": date_valid,
            "niveau_de_securite": niveau_sec if niveau_sec else "1--",
        },
        "validation": {
            "cree_par": cree_par,
            "approuve_par": approuve_par,
        },
        "identifiants": {
            "numero_article": numero_art,
            "copie_de": copie_de,
        },
        "masse_totale": {
            "valeur": masse_valeur,
            "unite": masse_unite,
        },
    }

def parse_title_block(words, page_format, page_width, template_type, company):
    """Dispatch vers le parser approprié."""
    if template_type == "legacy":
        return parse_title_block_legacy(words, page_format, page_width, company)
    return parse_title_block_new(words, page_format, page_width, company)

# ─────────────────────────────────────────────────────────────────
# Révisions
# ─────────────────────────────────────────────────────────────────
def parse_revisions(words, page_format):
    rev_label = find_word(words, "RÉV.") or find_word(words, "REV.")
    if not rev_label:
        return []
    rev_y = rev_label["top"]
    if page_format == "A4":
        margin_min, margin_max, y_tol = 1, 32, 3
    elif page_format == "A2":
        margin_min, margin_max, y_tol = 1, 20, 4
    else:
        margin_min, margin_max, y_tol = 1, 35, 5
    data_words = [w for w in words
                  if w["top"] < rev_y - margin_min
                  and w["top"] > rev_y - margin_max
                  and w["x0"] >= rev_label["x0"] - 5]
    lines = group_into_lines(data_words, y_tol=y_tol)
    revisions = []
    for y, lw in lines.items():
        tokens = [w["text"] for w in lw]
        date_idx = next(
            (i for i, t in enumerate(tokens) if len(t) >= 9 and t.count("-") == 2),
            None
        )
        if date_idx is None:
            continue
        after = tokens[date_idx + 1 :]
        revisions.append({
            "rev": tokens[0],
            "modification": " ".join(tokens[1:date_idx]),
            "date_validation": tokens[date_idx],
            "modifie_par": " ".join(after[:2]) if len(after) >= 2 else "",
            "approuve_par": " ".join(after[2:]) if len(after) > 2 else "",
        })
    return revisions

# ─────────────────────────────────────────────────────────────────
# Spécifications (ligne SPEC.)
# ─────────────────────────────────────────────────────────────────
def parse_spec(words, page_format):
    spec_w = find_word(words, "SPEC.")
    if not spec_w:
        return {}
    spec_y = spec_w["top"]
    y_tol = 3 if page_format == "A4" else 5
    merge_dist = 5 if page_format == "A4" else 8
    row = sorted(
        [w for w in words if abs(w["top"] - spec_y) < y_tol and w["x0"] > spec_w["x1"]],
        key=lambda w: w["x0"]
    )
    merged = []
    for w in row:
        if merged and w["x0"] - merged[-1]["x1"] < merge_dist:
            merged[-1] = {**merged[-1], "text": merged[-1]["text"] + w["text"],
                          "x1": w["x1"]}
        else:
            merged.append(dict(w))
    return {"spec": "SPEC.", "valeurs": [w["text"] for w in merged]}

# ─────────────────────────────────────────────────────────────────
# Point d'entrée
# ─────────────────────────────────────────────────────────────────
def parse_poma_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        page_format, orientation = detect_page_format(page)
        page_width = page.width
        page_height = page.height
        words = page.extract_words(
            x_tolerance=2,
            y_tolerance=2,
            keep_blank_chars=False,
            use_text_flow=False,
        )
        template_type = detect_template_type(words)
        company = detect_company(words)
        print(f"📄 Format: {page_format} ({orientation}) - {page_width:.1f}x{page_height:.1f}", file=sys.stderr)
        print(f"📋 Template: {template_type}, Société: {company}", file=sys.stderr)

        header_y = find_bom_header_y(words)
        if header_y is None:
            raise ValueError("Header BOM introuvable (colonne REP. absente)")

        return {
            "format_detecte": page_format,
            "template": template_type,
            "document": parse_title_block(words, page_format, page_width, template_type, company),
            "spec": parse_spec(words, page_format),
            "revisions": parse_revisions(words, page_format),
            "nomenclature": parse_bom(words, header_y, page_format, template_type, page_width),
        }

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python poma_parser_dynamic.py <input.pdf> [output.json]")
        sys.exit(1)
    pdf_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    result = parse_poma_pdf(pdf_path)
    output = json.dumps(result, ensure_ascii=False, indent=2)
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(output)
        print(f"✅ JSON sauvegardé → {output_path}")
    else:
        print(output)
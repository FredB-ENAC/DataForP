# -*- coding: utf-8 -*-
"""FICHIER CONSOLIDE DES PARCOURS - version 1.0

Produit un NOUVEAU classeur, tres proche de REEF_REBUILD.xlsx :
  - une ligne par dossier et periode, toutes les colonnes REEF_rebuild conservees telles quelles ;
  - apres chaque Degree, les valeurs complementaires de Parcours, Notes, Sessions et ASTRAL ;
  - a droite, le suivi KLIF du dossier : affectation, MU, durees, issue de la premiere affectation,
    second centre.
REEF_rebuild et les autres sources sont lus uniquement (empreinte verifiee avant/apres).
Les correspondances proposees (variantes de nom) ne sont utilisees qu'apres decision OUI dans
DECISIONS_CORRESPONDANCES.xlsx ; ces decisions sont conservees d'une execution a l'autre.
Usage : double-clic sur LANCER_WINDOWS.bat | python consolidation.py --autotest
"""
import argparse, datetime as dt, difflib, hashlib, json, os, re, shutil, sys, unicodedata, uuid
from collections import Counter, defaultdict, OrderedDict
from pathlib import Path

VERSION = "1.0"
REVISION = "2026-11-13-lot-c1-c8-r1-provenance"
DEGREES = ("BTWR", "BSURV", "ADI", "APS1", "APS2", "ACS1", "ACS2")
ROLES = {"BTWR": ("TWR",), "BSURV": ("CO", "CR"), "ADI": ("LOC",), "APS1": ("PL", "AP"),
         "APS2": ("PL", "AP"), "ACS1": ("CO", "CR"), "ACS2": ("CO", "CR")}
PROMO_MIN_KLIF = (2012, "A")          # feuilles KLIF anterieures a 12A : non etudiees
REEF_REQUIS = ("ID_DOSSIER_LOCAL", "NOM", "PRENOM", "ORDRE_PERIODE", "PROMOTION")
DECISIONS_NOM = "DECISIONS_CORRESPONDANCES.xlsx"
FEUILLES_IGNOREES = {
    "REEF": frozenset(("REPRISES_A_PRECISER", "LIRE_MOI", "_META_LOCAL")),
    "NOTES": frozenset(("AUDIT_DEGREES",)),
}
DOSSIERS_18A_C4 = frozenset(("Dd43dd2048f958378", "D4d8e8aa756149655"))
FAMILLES_HORS_CIBLE = frozenset(("ICNA_P", "TSEEAC", "DEF", "EXPORT"))


class Erreur(Exception):
    pass


# ---------------------------------------------------------------- utilitaires
def sans_accent(s):
    s = unicodedata.normalize("NFKD", str(s))
    return "".join(c for c in s if not unicodedata.combining(c))


def norm(v):
    return re.sub(r"[^a-z0-9]", "", sans_accent("" if v is None else v).lower())


def vide(v):
    return v is None or (isinstance(v, str) and not v.strip())


def propre(v):
    return "" if v is None else re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()


def cle_nom(*valeurs):
    """Nom + prenom -> jetons tries (accents, casse, traits d'union, ordre nom/prenom)."""
    t = []
    for v in valeurs:
        if vide(v):
            continue
        t += [x for x in re.split(r"[^a-z0-9]+", sans_accent(v).lower()) if x]
    return " ".join(sorted(t))


def promo(brut):
    """(famille, code) : ICNA_LETTRE '22C', ICNA_P '22P', TSEEAC, DEF, EXPORT, AUTRE, ABSENTE."""
    if vide(brut):
        return "ABSENTE", None
    t = re.sub(r"\s+", "", str(brut)).upper()
    t = re.sub(r"[-_]?BIS$", "", t)
    if t.startswith("DEF"):
        return "DEF", t
    if t.startswith("ATCO"):
        return "EXPORT", t
    m = re.fullmatch(r"(MCTA|ICNA)?(\d{2})([A-Z])", t)
    if m:
        code = m.group(2) + m.group(3)
        return ("ICNA_P" if m.group(3) == "P" else "ICNA_LETTRE"), code
    m = re.fullmatch(r"(GSEA|FCTLA|SARA)(\d{2})([A-Z])", t)
    if m:
        return "TSEEAC", t
    return "AUTRE", t


def ordre_promo(code):
    m = re.fullmatch(r"(\d{2})([A-Z])", code or "")
    if not m:
        return None
    an = int(m.group(1))
    # Annees historiques : 50-99 = 1950-1999 ; 00-49 = 2000-2049.
    return (1900 + an if an >= 50 else 2000 + an, m.group(2))


def promotion_avant_12a(brut):
    famille, code = promo(brut)
    ordre = ordre_promo(code) if famille in ("ICNA_LETTRE", "ICNA_P") else None
    return ordre is not None and ordre < PROMO_MIN_KLIF


def ecarter_anciennes_periodes_registre(registre):
    """Mise a l'ecart reversible, sans renumerotation des periodes conservees."""
    exclus = registre.setdefault("hors_perimetre_avant_12A", {"dossiers": {}, "periodes": {}})
    nd = np = 0
    for did, periodes in list(registre["periodes"].items()):
        for pr in list(periodes):
            if promotion_avant_12a(pr):
                exclus["periodes"].setdefault(did, {})[pr] = periodes.pop(pr)
                np += 1
        if not periodes:
            del registre["periodes"][did]
    for did in list(registre["dossiers"]):
        if not registre["periodes"].get(did) and did in exclus["periodes"]:
            exclus["dossiers"][did] = registre["dossiers"].pop(did)
            nd += 1
    return {"dossiers_mis_a_ecart_avant_12A": nd, "periodes_mises_a_ecart_avant_12A": np}


def degree_de(texte):
    t = sans_accent(texte or "").upper()
    m = re.search(r"(?<![A-Z0-9])(BASIC[ _-]*TOWER|BASIC[ _-]*ACC|BTWR|BSURV|ADI|APS1|APS2|ACS1|ACS2)(?=$|[^A-Z0-9])", t)
    if not m:
        return None
    return {"BASICTOWER": "BTWR", "BASICACC": "BSURV"}.get(norm(m.group(1)).upper(), m.group(1))


def degree_entete_astral(valeur):
    """Seuls les en-tetes Degree exacts comptent ; ABES_ADI n'est pas ADI."""
    t = propre(valeur).upper().replace(" ", "")
    if t.startswith("ABES_"):
        return None
    return {"BASIC_TOWER": "BTWR", "BASIC_ACC": "BSURV"}.get(t, t if t in DEGREES else None)


def famille_explicitement_exclue(brut):
    return promo(brut)[0] in FAMILLES_HORS_CIBLE


def est_moyenne_klif(valeur):
    """C3 : egalite stricte apres normalisation de la casse et des espaces seulement."""
    return isinstance(valeur, str) and "".join(valeur.split()).casefold() == "moyenne"


def role_de(degree, brut):
    s = norm(brut)
    table = {"co": "CO", "pc": "CO", "organique": "CO", "cr": "CR", "rc": "CR", "radariste": "CR",
             "pl": "PL", "planner": "PL", "cordon": "PL", "ap": "AP", "app": "AP", "approche": "AP",
             "twr": "TWR", "tour": "TWR", "loc": "LOC", "local": "LOC"}
    r = table.get(s)
    if r and degree and degree.startswith("APS"):
        r = {"CO": "PL", "CR": "AP"}.get(r, r)
    if r and degree and r not in ROLES[degree]:
        return None
    return r


def nombre(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(propre(v).replace(",", "."))
    except ValueError:
        return None


def note(v):
    x = nombre(v)
    return x if x is not None and 0 <= x <= 20 else None


def resultat(v):
    s = norm(v)
    return {"go": "GO", "nogo": "NOGO", "ok": "GO", "ko": "NOGO", "admis": "GO",
            "ajourne": "NOGO"}.get(s)


def date(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        for f in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
            try:
                return dt.datetime.strptime(v.strip(), f)
            except ValueError:
                pass
    return None


def mois(d1, d2):
    if d1 is None or d2 is None:
        return None
    return round((d2 - d1).days / 30.4375, 1)


def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def dossier_synchronise(p):
    p = str(Path(p).resolve()).lower()
    for n in ("OneDrive", "OneDriveConsumer", "OneDriveCommercial", "Dropbox"):
        q = os.environ.get(n)
        if q and p.startswith(str(Path(q).resolve()).lower()):
            return True
    return "onedrive" in p or "dropbox" in p


def charger(p):
    from openpyxl import load_workbook
    return load_workbook(p, read_only=True, data_only=True)


# ---------------------------------------------------------------- REEF_rebuild
def lire_reef(chemin):
    wb = charger(chemin)
    try:
        autorisees = [w for w in wb.worksheets if w.title.upper() not in FEUILLES_IGNOREES["REEF"]]
        ws = next((w for w in autorisees if norm(w.title) == "reefrebuild"),
                  autorisees[0] if autorisees else None)
        if ws is None:
            raise Erreur("REEF_REBUILD : aucune feuille de donnees hors feuilles ignorees.")
        it = ws.iter_rows(values_only=True)
        entetes = [propre(h) for h in next(it, ())]
        while entetes and not entetes[-1]:
            entetes.pop()
        manquants = [c for c in REEF_REQUIS if c not in entetes]
        if manquants:
            raise Erreur("REEF_REBUILD : colonnes absentes " + ", ".join(manquants))
        lignes = []
        for no, row in enumerate(it, 2):
            row = list(row)[:len(entetes)] + [None] * max(0, len(entetes) - len(row))
            if not any(not vide(x) for x in row):
                continue
            lignes.append({"_ligne": no, "valeurs": OrderedDict(zip(entetes, row))})
    finally:
        wb.close()
    dossiers = OrderedDict()
    for l in lignes:
        v = l["valeurs"]
        did = propre(v["ID_DOSSIER_LOCAL"])
        d = dossiers.setdefault(did, {"id": did, "nom": v["NOM"], "prenom": v["PRENOM"],
                                      "cle": cle_nom(v["NOM"], v["PRENOM"]), "periodes": []})
        try:
            ordre = int(nombre(v["ORDRE_PERIODE"]))
        except (TypeError, ValueError):
            ordre = len(d["periodes"]) + 1
        d["periodes"].append({"ordre": ordre, "promo": promo(v["PROMOTION"])[1], "ligne": l})
        l["dossier"], l["ordre"] = did, ordre
    for d in dossiers.values():
        d["periodes"].sort(key=lambda p: p["ordre"])
        d["promos"] = {p["promo"] for p in d["periodes"] if p["promo"]}
    return ws.title, entetes, lignes, dossiers


# ---------------------------------------------------------------- sources
def obs(source, fichier, feuille, ligne, nom_brut, cle, promo_brute, degree=None, **data):
    fam, code = promo(promo_brute)
    return {"source": source, "fichier": Path(fichier).name, "feuille": feuille, "ligne": ligne,
            "nom_brut": nom_brut, "cle": cle, "promo_brute": promo_brute, "famille": fam,
            "promo": code, "degree": degree, "data": data}


def lire_parcours(chemin, lecture):
    alias = {"nom": "nom", "prenom": "prenom", "promotion": "promo", "promo": "promo",
             "degre": "degree", "degree": "degree", "position1": "role1", "role1": "role1",
             "noteposition1": "note1", "note1": "note1", "resultat1": "res1", "rattrapage1": "ratt1",
             "position2": "role2", "role2": "role2", "noteposition2": "note2", "note2": "note2",
             "resultat2": "res2", "rattrapage2": "ratt2", "note": "note", "noteglobale": "note",
              "resultat": "res", "resultatglobal": "res",
              "notereconstituee": "note_reconst", "originenotereconstituee": "origine_reconst",
              "statutnotereconstituee": "statut_reconst", "traitementposition2": "traitement2",
              "position2interpretee": "position2_interpretee"}
    out = []
    wb = charger(chemin)
    try:
        for ws in wb.worksheets:
            rows = ws.iter_rows(values_only=True)
            h = next(rows, ())
            carte = {}
            for j, x in enumerate(h):
                k = alias.get(norm(x))
                if k and k not in carte:
                    carte[k] = j
            if not {"nom", "prenom", "promo", "degree"} <= set(carte):
                lecture["PARCOURS"]["feuilles_non_utilisees"] += 1
                continue
            lecture["PARCOURS"]["feuilles_lues"] += 1
            nh = {norm(x) for x in h}
            niveau = ("DETAIL_EVALUATION" if norm(ws.title) == "detailseval" or
                      {"simulation", "role"} <= nh else
                      "SYNTHESE_PARCOURS" if norm(ws.title) == "parcours" or "role1" in carte else
                      "NIVEAU_INDETERMINE")
            lecture["PARCOURS"]["feuilles_" + niveau] += 1
            for no, row in enumerate(rows, 2):
                g = lambda k: row[carte[k]] if k in carte and carte[k] < len(row) else None
                if all(vide(x) for x in row):
                    continue
                d = degree_de(propre(g("degree")))
                brut_degree = norm(g("degree"))
                if d in ("BTWR", "BSURV") and brut_degree.startswith(("basictower", "basicacc")):
                    lecture["PARCOURS"]["alias_" + ("BASIC_TOWER" if d == "BTWR" else "BASIC_ACC") + "_vers_" + d] += 1
                roles = {}
                for i in ("1", "2"):
                    traitement = norm(g("traitement2")) + " " + norm(g("position2_interpretee")) if i == "2" else ""
                    if i == "2" and any(m in traitement for m in ("neutralis", "ignore", "nonretenu")):
                        lecture["PARCOURS"]["position_2_neutralisee_explicitement"] += 1
                        continue  # la ligne source et son traitement restent dans observation_brute
                    r = role_de(d, g("role" + i)) if d else None
                    if d in ("BTWR", "ADI") and r is None and not vide(g("note" + i)) and i == "1":
                        r = ROLES[d][0]
                    if r:
                        roles[r] = {"note": note(g("note" + i)), "res": resultat(g("res" + i)),
                                    "ratt": propre(g("ratt" + i))}
                        if r == "PL":
                            roles[r]["res"] = resultat(g("res" + i)) or propre(g("res" + i)) or None
                out.append(obs("PARCOURS", chemin, ws.title, no, "%s %s" % (propre(g("nom")), propre(g("prenom"))),
                               cle_nom(g("nom"), g("prenom")), g("promo"), d,
                               note=note(g("note")), res=resultat(g("res")) or propre(g("res")) or None, roles=roles,
                               note_reconstituee=(note(g("note_reconst")) if norm(g("note")) == "nogo" else None),
                               note_reconstituee_brute=propre(g("note_reconst")),
                               origine_note_reconstituee=propre(g("origine_reconst")),
                               statut_note_reconstituee=propre(g("statut_reconst")),
                               position_2_neutralisee=any(m in (norm(g("traitement2")) + " " + norm(g("position2_interpretee")))
                                                         for m in ("neutralis", "ignore", "nonretenu")),
                               degree_brut=propre(g("degree")), identite_nom=propre(g("nom")),
                               identite_prenom=propre(g("prenom")), niveau=niveau,
                               observation_brute={propre(k): row[j] if j < len(row) else None
                                                  for j, k in enumerate(h) if not vide(k)}))
    finally:
        wb.close()
    return out


def lire_blocs(chemin, source, lecture, max_col):
    """Exports Socrates : blocs commencant par une ligne 'Class:' en colonne B."""
    blocs = []
    wb = charger(chemin)
    try:
        for ws in wb.worksheets:
            if ws.title.upper() in FEUILLES_IGNOREES.get(source, ()):
                lecture[source]["feuilles_ignorees_par_decision"] += 1
                continue
            courant, n = None, 0
            for i, row in enumerate(ws.iter_rows(min_col=1, max_col=max_col, values_only=True), 1):
                row = tuple(row) + (None,) * (max_col - len(row))
                if all(vide(v) for v in row):
                    courant = None
                    continue
                b = propre(row[1])
                if b.casefold().startswith("class:"):
                    if source in ("NOTES", "SESSIONS") and famille_explicitement_exclue(b[6:].strip()):
                        courant = None
                        lecture[source]["blocs_hors_perimetre_exclus"] += 1
                        continue
                    courant = {"feuille": ws.title, "ligne": i, "row": row, "classe": b[6:].strip(),
                               "degree": None, "sims": []}
                    blocs.append(courant)
                    n += 1
                elif courant is not None:
                    if b.casefold().startswith("degree:") and courant["degree"] is None:
                        courant["degree"] = b[7:].strip()
                    elif b.casefold().startswith("sim:"):
                        courant["sims"].append(b[4:].strip())
            lecture[source]["feuilles_lues" if n else "feuilles_non_utilisees"] += 1
    finally:
        wb.close()
    return blocs


def lire_notes(chemin, lecture):
    out = []
    for b in lire_blocs(chemin, "NOTES", lecture, 7):
        a, _, statut, nb, e, f, g = b["row"]
        d = degree_de(b["degree"]) or degree_de(b["feuille"])
        roles = {}
        for cell in (e, f):
            m = re.match(r"^\s*([A-Za-z]+)\s*:\s*(-?[\d.,]+)\s*$", propre(cell))
            if m and d:
                r = role_de(d, m.group(1))
                if r:
                    roles[r] = {"note": note(m.group(2)), "res": None, "ratt": ""}
        out.append(obs("NOTES", chemin, b["feuille"], b["ligne"], propre(a), cle_nom(a), b["classe"], d,
                       note=note(g), res=None, roles=roles, statut=propre(statut)))
    return out


def lire_sessions(chemin, lecture):
    out = []
    for b in lire_blocs(chemin, "SESSIONS", lecture, 6):
        a, _, _, dte, _, res = b["row"]
        d = degree_de(b["degree"]) or degree_de(b["feuille"])
        quand = date(dte)
        if quand is None and isinstance(dte, (int, float)):
            try:
                from openpyxl.utils.datetime import from_excel
                quand = from_excel(dte)
            except Exception:
                quand = None
        out.append(obs("SESSIONS", chemin, b["feuille"], b["ligne"], propre(a), cle_nom(a), b["classe"], d,
                       date=quand, res=resultat(res)))
    return out


RE_REM = re.compile(r"^\s*(\d+)\s*remedial\(s\)\s*$", re.I)
RE_PCT = re.compile(r"^\s*\d{1,3}\s*%\s*done\s*$", re.I)


def lire_astral(chemin, lecture):
    out = []
    wb = charger(chemin)
    try:
        for ws in wb.worksheets:
            rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
            h, col = None, {}
            for i, r in enumerate(rows[:15]):
                noms = [degree_entete_astral(c) for c in r]
                if any(noms):
                    h, col = i, {x: j for j, x in enumerate(noms) if x in DEGREES}
                    break
            if h is None:
                lecture["ASTRAL"]["feuilles_non_utilisees"] += 1
                continue
            lecture["ASTRAL"]["feuilles_lues"] += 1
            courant = None

            def cell(r, m):
                return propre(r[col[m]]) if col[m] < len(r) else ""

            def terminer():
                if courant is not None:
                    if courant["data"]["remedials"]:
                        out.append(courant)
                    else:
                        lecture["ASTRAL"]["blocs_sans_remedial_explicite"] += 1

            for i in range(h + 1, len(rows)):
                r = rows[i]
                label = propre(r[1]) if len(r) > 1 else ""
                if label.upper() == "TOTAL":
                    terminer()
                    courant = None
                    break
                etiquette = bool(re.match(r"^(ENAC|ENAC\s*/.*|[A-Za-z]{2,12}\s*/\s*[A-Za-z0-9]{2,8})$", label, re.I))
                donnees = {m: cell(r, m) for m in col if cell(r, m)}
                if label and not etiquette and len(cle_nom(label).split()) >= 2:
                    terminer()
                    courant = obs("ASTRAL", chemin, ws.title, i + 1, label, cle_nom(label), ws.title, None,
                                  remedials=Counter())
                elif label and not etiquette:
                    terminer()
                    courant = None
                if courant is not None and donnees:
                    for m, v in donnees.items():
                        q = RE_REM.match(v)
                        if q:
                            courant["data"]["remedials"][m] += int(q.group(1))
            else:
                terminer()
    finally:
        wb.close()
    return out


KLIF_LIBELLES = {
    "promo": "promo", "datedentreealenac": "date_entree", "nomdustagiaire": "nom",
    "prenomdustagiaire": "prenom", "affectation": "affectation", "zone": "zone",
    "datedaffectation": "date_affectation", "dateobtentiondelamuzz": "date_mu",
    "arretformation": "arret", "nouveaucentre": "nouveau_centre", "datedarrivee": "date_arrivee",
    "nombredechecs": "nb_echecs", "datedefindeformation": "date_fin",
    "modederecrutement": "recrutement", "origine": "origine",
    "nombredemoisdequalificationenacexclus": "duree_enac_exclus"}
EPISODE_2 = {"date_mu", "arret", "nb_echecs", "date_fin", "date_arrivee"}


def lire_klif(chemin, lecture):
    out = []
    wb = charger(chemin)
    try:
        for ws in wb.worksheets:
            fam, code = promo(ws.title)
            if fam == "ICNA_P":
                lecture["KLIF"]["feuilles_voie_p_exclues"] += 1
                continue
            o = ordre_promo(code[:3] if code else None) if fam in ("ICNA_LETTRE", "ICNA_P") else None
            if o is None:
                lecture["KLIF"]["feuilles_non_promotion"] += 1
                continue
            if o < PROMO_MIN_KLIF:
                lecture["KLIF"]["feuilles_anterieures_12A_ignorees"] += 1
                continue
            rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
            k = next((i for i, r in enumerate(rows[:6]) if any(norm(v) == "promo" for v in r if v is not None)), None)
            if k is None:
                lecture["KLIF"]["feuilles_sans_en_tete"] += 1
                continue
            lecture["KLIF"]["feuilles_lues"] += 1
            roles = [KLIF_LIBELLES.get(norm(v)) if not vide(v) else None for v in rows[k]]
            inc = roles.index("nouveau_centre") if "nouveau_centre" in roles else None
            carte = {}
            for j, r in enumerate(roles):
                if r is None:
                    continue
                ep = 2 if inc is not None and j > inc and r in EPISODE_2 else 1
                carte.setdefault((r, ep), j)
            for i in range(k + 1, len(rows)):
                row = rows[i]
                g = lambda r, ep=1: row[carte[(r, ep)]] if (r, ep) in carte and carte[(r, ep)] < len(row) else None
                moyenne = any(est_moyenne_klif(x) for x in row)
                if moyenne and vide(g("nom")) and vide(g("prenom")):
                    lecture["KLIF"]["lignes_moyenne_ignorees"] += 1
                    continue
                if vide(g("nom")) and vide(g("prenom")):
                    if any(not vide(x) for x in row):
                        lecture["KLIF"]["lignes_sans_identite"] += 1
                    continue
                v = {"date_entree": g("date_entree"), "affectation": g("affectation"), "zone": g("zone"),
                     "origine": g("origine"), "duree_enac_exclus_brute": g("duree_enac_exclus"),
                     "date_affectation": g("date_affectation"), "date_mu_1": g("date_mu"), "arret_1": g("arret"),
                     "nouveau_centre": g("nouveau_centre"), "date_arrivee_2": g("date_arrivee", 2),
                     "date_mu_2": g("date_mu", 2), "arret_2": g("arret", 2),
                     "nb_echecs": g("nb_echecs") if g("nb_echecs") is not None else g("nb_echecs", 2),
                     "date_fin": g("date_fin") if g("date_fin") is not None else g("date_fin", 2),
                     "conflit_moyenne_identite": moyenne}
                out.append(obs("KLIF", chemin, ws.title, i + 1, "%s %s" % (propre(g("nom")), propre(g("prenom"))),
                               cle_nom(g("nom"), g("prenom")), ws.title, None, identite_nom=propre(g("nom")), identite_prenom=propre(g("prenom")), **v))
    finally:
        wb.close()
    return out


# ---------------------------------------------------------------- KLIF : interpretation
MOTIFS_SORTIE = ("deced", "inapt", "demission", "abandon", "licenci", "radiation", "radie", "reform")


def categorie_arret(v):
    if vide(v):
        return None
    x = nombre(v)
    if x is not None:
        return "COMPTEUR_%d" % int(x) if x > 0 else "NON"
    s = norm(v)
    if s == "oui":
        return "OUI"
    if s == "non":
        return "NON"
    if s == "dispo":
        return "DISPO"
    if any(c in s for c in ("ieeac", "tseeac", "iessa")):
        return "AUTRE_CURSUS"
    if any(m in s for m in MOTIFS_SORTIE):
        return "SORTIE"
    return "AUTRE"


def interpreter_klif(d):
    alertes = []
    entree, daff, mu1 = date(d["date_entree"]), date(d["date_affectation"]), date(d["date_mu_1"])
    darr, mu2 = date(d["date_arrivee_2"]), date(d["date_mu_2"])
    a1, a2 = categorie_arret(d["arret_1"]), categorie_arret(d["arret_2"])
    texte_sortie = any(isinstance(d[c], str) and any(m in norm(d[c]) for m in MOTIFS_SORTIE)
                       for c in ("affectation", "date_mu_1"))
    m1 = mois(daff, mu1)
    if m1 is not None and m1 < 0:
        alertes.append("MU_1_AVANT_AFFECTATION")
        m1 = None
    m2 = mois(darr, mu2)
    if m2 is not None and m2 < 0:
        alertes.append("MU_2_AVANT_ARRIVEE")
        m2 = None
    mu_finale = mu2 or mu1
    mt = mois(entree, mu_finale)
    if mt is not None and mt < 0:
        alertes.append("MU_AVANT_ENTREE_ENAC")
        mt = None
    if mu1:
        issue = "QUALIFIEE"
        if a1 and (a1 == "OUI" or a1.startswith("COMPTEUR")):
            alertes.append("MU_1_ET_ARRET_1_RENSEIGNES")
    elif a1 == "OUI" or (a1 or "").startswith("COMPTEUR"):
        issue = "ECHEC_ARRET_DE_FORMATION"
    elif a1 == "SORTIE" or texte_sortie:
        issue = "SORTIE"
    elif a1 == "AUTRE_CURSUS":
        issue = "AUTRE_CURSUS"
    elif a1 == "DISPO":
        issue = "DISPONIBILITE"
    elif not vide(d["affectation"]):
        issue = "NON_QUALIFIEE_A_LA_DATE_DU_FICHIER"
    else:
        issue = "NON_DETERMINEE"
    if not vide(d["nouveau_centre"]) and issue == "QUALIFIEE":
        alertes.append("NOUVEAU_CENTRE_APRES_QUALIFICATION")
    if mu1 is None and not vide(d["date_mu_1"]) and not texte_sortie:
        alertes.append("DATE_MU_1_ILLISIBLE")
    duree_brute = d.get("duree_enac_exclus_brute")
    duree_source = nombre(duree_brute)
    if not vide(duree_brute):
        if duree_source is None:
            alertes.append("DUREE_ENAC_EXCLUS_NON_NUMERIQUE")
        elif duree_source < 0:
            alertes.append("DUREE_ENAC_EXCLUS_NEGATIVE")
            duree_source = None
    return {"KLIF_MOIS_AFFECTATION_MU_1": m1, "KLIF_ISSUE_AFFECTATION_1": issue,
            "KLIF_MOIS_ARRIVEE_MU_2": m2, "KLIF_MOIS_ENTREE_ENAC_MU": mt,
            "KLIF_MOIS_QUALIFICATION_ENAC_EXCLUS": duree_source,
            "KLIF_ARRET_1_CATEGORIE": a1 or "", "KLIF_ARRET_2_CATEGORIE": a2 or ""}, alertes


KLIF_COLONNES = ["KLIF_PORTEE", "KLIF_LIEN", "KLIF_FEUILLE", "KLIF_DATE_ENTREE_ENAC",
                 "KLIF_ORIGINE", "KLIF_AFFECTATION_1", "KLIF_ZONE", "KLIF_DATE_AFFECTATION_1", "KLIF_DATE_MU_1",
                 "KLIF_ARRET_1", "KLIF_ARRET_1_CATEGORIE", "KLIF_MOIS_AFFECTATION_MU_1",
                 "KLIF_ISSUE_AFFECTATION_1", "KLIF_NOUVEAU_CENTRE", "KLIF_DATE_ARRIVEE_2", "KLIF_DATE_MU_2",
                 "KLIF_ARRET_2", "KLIF_ARRET_2_CATEGORIE", "KLIF_MOIS_ARRIVEE_MU_2", "KLIF_NB_ECHECS",
                 "KLIF_DATE_FIN_FORMATION", "KLIF_MOIS_ENTREE_ENAC_MU",
                 "KLIF_MOIS_QUALIFICATION_ENAC_EXCLUS", "KLIF_DUREE_ENAC_EXCLUS_BRUTE", "KLIF_ALERTES"]
KLIF_BRUT = {"KLIF_DATE_ENTREE_ENAC": "date_entree", "KLIF_ORIGINE": "origine",
             "KLIF_DUREE_ENAC_EXCLUS_BRUTE": "duree_enac_exclus_brute",
             "KLIF_AFFECTATION_1": "affectation", "KLIF_ZONE": "zone",
             "KLIF_DATE_AFFECTATION_1": "date_affectation", "KLIF_DATE_MU_1": "date_mu_1",
             "KLIF_ARRET_1": "arret_1", "KLIF_NOUVEAU_CENTRE": "nouveau_centre",
             "KLIF_DATE_ARRIVEE_2": "date_arrivee_2", "KLIF_DATE_MU_2": "date_mu_2", "KLIF_ARRET_2": "arret_2",
             "KLIF_NB_ECHECS": "nb_echecs", "KLIF_DATE_FIN_FORMATION": "date_fin"}


def colonnes_degree(d):
    c = [d + "_PARCOURS_RESULTAT", d + "_PARCOURS_NOTE"]
    c += [d + "_PARCOURS_NOTE_RECONSTITUEE", d + "_PARCOURS_ORIGINE_NOTE_RECONSTITUEE",
          d + "_PARCOURS_STATUT_NOTE_RECONSTITUEE"]
    if len(ROLES[d]) > 1:
        c += ["%s_%s_PARCOURS_NOTE" % (d, r) for r in ROLES[d]]
    if d in ("APS1", "APS2"):
        c += [d + "_PL_PARCOURS_RESULTAT"]
    c += [d + "_SOCRATES_NOTE"]
    if len(ROLES[d]) > 1:
        c += ["%s_%s_SOCRATES_NOTE" % (d, r) for r in ROLES[d]]
    c += [d + "_SESSIONS_NB", d + "_SESSIONS_NOGO", d + "_SESSIONS_PREMIERE", d + "_SESSIONS_DERNIERE",
          d + "_ASTRAL_REMEDIALS", d + "_ASTRAL_PRESENCE_REMEDIAL", d + "_ETAT_SOURCES"]
    return c


# ---------------------------------------------------------------- decisions
def lire_decisions(dossier_sortie):
    p = Path(dossier_sortie) / DECISIONS_NOM
    dec = {}
    if not p.exists():
        return dec, []
    wb = charger(p)
    lignes = []
    try:
        ws = wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        h = [propre(x) for x in next(it, ())]
        for row in it:
            v = dict(zip(h, row))
            if vide(v.get("SOURCE")):
                continue
            lignes.append(v)
            d = norm(v.get("DECISION"))
            if d in ("oui", "non"):
                dec[(propre(v["SOURCE"]), propre(v["CLE_SOURCE"]), propre(v["PROMO_SOURCE"]),
                     propre(v["ID_DOSSIER_LOCAL"]))] = (d.upper(), propre(v.get("CLE_REEF")))
    finally:
        wb.close()
    return dec, lignes


# ---------------------------------------------------------------- rapprochement
class Moteur:
    def __init__(self, dossiers, decisions):
        self.dossiers = dossiers
        self.decisions = decisions
        self.decisions_utilisees = set()
        self.par_cle = defaultdict(list)
        for d in dossiers.values():
            if d["cle"]:
                self.par_cle[d["cle"]].append(d)
        self.promos_reef = {p for d in dossiers.values() for p in d["promos"]}
        self.propositions = {}

    def _decision(self, o):
        ok, non = None, set()
        for (src, cle, pr, did), (dec, cle_reef) in self.decisions.items():
            if src != o["source"] or cle != o["cle"] or pr != (o["promo"] or ""):
                continue
            d = self.dossiers.get(did)
            if d is None and cle_reef:  # identifiant regenere : retrouver par la cle REEF
                c = [x for x in self.par_cle.get(cle_reef, [])]
                d = c[0] if len(c) == 1 else None
            if d is None:
                continue
            self.decisions_utilisees.add((src, cle, pr, did))
            if dec == "OUI":
                ok = d
            else:
                non.add(d["id"])
        return ok, non

    def identifier(self, o):
        """-> (dossier ou None, statut)"""
        if not o["cle"] or len(o["cle"].split()) < 2:
            return None, "IDENTITE_INSUFFISANTE"
        valide, rejets = self._decision(o)
        if valide:
            return valide, "DECISION_VALIDEE"
        if o["famille"] in ("AUTRE", "ABSENTE"):
            return None, "PROMOTION_A_QUALIFIER"
        if o.get("_propositions_population"):
            for did, motif in o["_propositions_population"]:
                if did in self.dossiers and did not in rejets:
                    self._proposer(o, self.dossiers[did], motif)
            return None, "PROPOSITION_A_DECIDER"
        cands = [d for d in self.par_cle.get(o["cle"], []) if d["id"] not in rejets]
        if cands:
            avec_promo = [d for d in cands if o["promo"] in d["promos"]]
            if len(avec_promo) == 1:
                return avec_promo[0], "NOM_ET_PROMOTION"
            if len(avec_promo) > 1:
                return None, "HOMONYMES_MEME_PROMOTION"
            if len(cands) == 1 and o["source"] != "KLIF":
                return cands[0], "NOM_EXACT_PROMOTION_DIFFERENTE"
            if len(cands) == 1:
                self._proposer(o, cands[0], "NOM_EXACT_PROMOTION_KLIF_DIFFERENTE")
                return None, "PROPOSITION_A_DECIDER"
            return None, "HOMONYMES"
        if o["famille"] in FAMILLES_HORS_CIBLE:
            return None, "HORS_CIBLE_CURSUS_OU_VOIE"
        if o["promo"] not in self.promos_reef:
            return None, "PROMOTION_NON_REPRESENTEE_DANS_REEF"
        props = self._approchants(o, rejets)
        if props:
            for d, motif in props:
                self._proposer(o, d, motif)
            return None, "PROPOSITION_A_DECIDER"
        return None, "SANS_CANDIDAT"

    def _approchants(self, o, rejets):
        t = set(o["cle"].split())
        res = []
        for d in self.dossiers.values():
            if d["id"] in rejets or o["promo"] not in d["promos"] or not d["cle"]:
                continue
            u = set(d["cle"].split())
            if (t < u or u < t) and len(t & u) >= 2:
                res.append((d, "NOM_COMPOSE_OU_PARTIEL"))
            elif difflib.SequenceMatcher(None, o["cle"], d["cle"]).ratio() >= 0.88:
                res.append((d, "ORTHOGRAPHE_PROCHE"))
        return res[:3]

    def _proposer(self, o, d, motif):
        k = (o["source"], o["cle"], o["promo"] or "", d["id"])
        p = self.propositions.setdefault(k, {"obs": 0, "motif": motif, "dossier": d, "o": o})
        p["obs"] += 1

    @staticmethod
    def periode(d, o):
        if (d and d.get("id") in DOSSIERS_18A_C4 and o.get("source") == "PARCOURS"
                and o.get("degree") in ("ACS1", "ACS2")):
            brut = norm(o["promo_brute"])
            cible = (1 if brut == "mcta18a" else
                     2 if brut == "mcta18abis" and o["degree"] == "ACS2" else None)
            if cible is not None:
                candidates = [p for p in d["periodes"] if p["ordre"] == cible and p["promo"] == "18A"]
                if len(candidates) == 1:
                    return candidates[0], "PERIODE_C4_18A_EXPLICITE"
                return None, "PERIODE_C4_ABSENTE" if not candidates else "PERIODE_C4_AMBIGUE"
        ps = [p for p in d["periodes"] if p["promo"] == o["promo"]]
        if len(ps) == 1:
            return ps[0], "PERIODE_UNIQUE"
        if len(ps) > 1:
            return None, "PERIODE_AMBIGUE"
        return None, "PERIODE_SANS_VIS_A_VIS"



REGISTRE_POPULATION = "REGISTRE_POPULATION_LOCAL.json"
COLONNES_POPULATION = ("ORIGINE_LIGNE_CONSOLIDEE", "NATURE_ORDRE_PERIODE")

def lire_registre_population(sortie):
    p = Path(sortie) / REGISTRE_POPULATION
    if not p.exists():
        return {"version": 1, "dossiers": {}, "periodes": {}}
    registre = json.loads(p.read_text(encoding="utf-8"))
    if registre.get("version") != 1 or not isinstance(registre.get("dossiers"), dict) or not isinstance(registre.get("periodes"), dict):
        raise Erreur("Registre de population incompatible : conserver le fichier et verifier sa version.")
    return registre

def ecrire_registre_population(sortie, registre):
    p = Path(sortie) / REGISTRE_POPULATION
    if p.exists():
        shutil.copy2(p, p.with_name("REGISTRE_POPULATION_sauvegarde_%s.json" % dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")))
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(registre, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, p)

def etendre_population(entetes, lignes, dossiers, observations, decisions, registre):
    """Ajouts issus des sources, sans changer une ligne ni une periode REEF."""
    ids_reef = set(dossiers)
    cles_reef = {d["cle"] for d in dossiers.values()}
    archive = registre.get("hors_perimetre_avant_12A", {})
    anciens = {did: saved.get("cle") for did, saved in {**registre.get("dossiers", {}),
                **archive.get("dossiers", {})}.items()}
    deja_ecartes = {(anciens.get(did), pr) for did, ps in archive.get("periodes", {}).items()
                    for pr in ps}
    nouveaux = periodes_nouvelles = 0
    def nouvelle_ligne(d, pr, ordre, origine, nature):
        v = OrderedDict((c, None) for c in entetes)
        v.update(ID_DOSSIER_LOCAL=d["id"], NOM=d["nom"], PRENOM=d["prenom"], ORDRE_PERIODE=ordre, PROMOTION=pr)
        l = {"_ligne": None, "valeurs": v, "dossier": d["id"], "ordre": ordre,
             "_origine_population": origine, "_nature_ordre": nature}
        lignes.append(l)
        d["periodes"].append({"ordre": ordre, "promo": pr, "ligne": l})
        d["promos"].add(pr)
    # Restituer les identifiants et periodes deja attribues dans ce dossier de resultats.
    for did, saved in registre["dossiers"].items():
        if did in dossiers:
            raise Erreur("Collision entre un identifiant REEF et le registre des dossiers supplementaires.")
        if saved["cle"] in cles_reef:
            raise Erreur("Un dossier supplementaire est maintenant aussi present dans REEF : correspondance a verifier avant migration du registre.")
        dossiers[did] = {"id": did, "nom": saved["nom"], "prenom": saved["prenom"], "cle": saved["cle"],
                         "periodes": [], "promos": set(), "_hors_reef": True}
    for did, periods in registre["periodes"].items():
        if did not in dossiers:
            raise Erreur("Registre : dossier de periode absent.")
        d = dossiers[did]
        for pr, saved in periods.items():
            ordre = saved["ordre"]
            if any(p["ordre"] == ordre for p in d["periodes"]):
                raise Erreur("Registre : ordre de periode supplementaire en collision avec REEF.")
            nouvelle_ligne(d, pr, ordre, saved["origine"], "TECHNIQUE_SANS_CHRONOLOGIE_CERTIFIEE")
    def index_exact(cle):
        return [d for d in dossiers.values() if d["cle"] == cle]
    def variantes(o, rejets):
        t = set(o["cle"].split()); out = []
        for d in dossiers.values():
            if d["id"] in rejets or d["cle"] == o["cle"] or not d["cle"]:
                continue
            u = set(d["cle"].split())
            if ((t < u or u < t) and len(t & u) >= 2) or difflib.SequenceMatcher(None, o["cle"], d["cle"]).ratio() >= .88:
                out.append((d["id"], "VARIANTE_AVANT_CREATION_POPULATION"))
        return out[:3]
    def periode_supplementaire(d, o):
        nonlocal periodes_nouvelles
        pr = o["promo"]
        if any(p["promo"] == pr for p in d["periodes"]):
            return
        # KLIF ne prouve pas une nouvelle periode pedagogique d'un dossier REEF.
        if o["source"] == "KLIF" and d["id"] in ids_reef:
            return
        ordre = max((p["ordre"] for p in d["periodes"]), default=0) + 1
        origine = "CONTEXTE_KLIF_SANS_PASSAGE_DEGREE_DEDUIT" if o["source"] == "KLIF" else "PERIODE_DOCUMENTEE_SOURCE"
        registre["periodes"].setdefault(d["id"], {})[pr] = {"ordre": ordre, "origine": origine}
        nouvelle_ligne(d, pr, ordre, origine, "TECHNIQUE_SANS_CHRONOLOGIE_CERTIFIEE")
        periodes_nouvelles += 1
    par_cle = defaultdict(list)
    for o in observations:
        par_cle[o["cle"]].append(o)
    # Ordre deterministe ; les variantes restent des propositions et ne sont jamais fusionnees.
    ordre_cles = sorted(par_cle, key=lambda k: (-len({o["source"] for o in par_cle[k]}), -len(k.split()), k))
    for cle in ordre_cles:
        for o in par_cle[cle]:
            if (o["famille"] != "ICNA_LETTRE" or not ordre_promo(o["promo"])
                    or (o["cle"], o["promo"]) in deja_ecartes):
                continue
            if len(cle.split()) < 2:
                continue
            moteur = Moteur(dossiers, decisions)
            accepte, rejets = moteur._decision(o)
            exacts = [d for d in index_exact(cle) if d["id"] not in rejets]
            avec_promo = [d for d in exacts if o["promo"] in d["promos"]]
            if avec_promo:
                exacts = avec_promo
            d = accepte
            if d is None and len(exacts) == 1:
                d = exacts[0]
                # Correspondance KLIF avec une autre promotion : conserver la decision requise.
                if o["source"] == "KLIF" and d["id"] in ids_reef and o["promo"] not in d["promos"]:
                    continue
            if d is None and len(exacts) > 1:
                continue
            if d is None:
                proches = variantes(o, rejets)
                if proches:
                    o["_propositions_population"] = proches
                    continue
                did = "N" + uuid.uuid4().hex[:20]
                while did in dossiers:
                    did = "N" + uuid.uuid4().hex[:20]
                n = o["data"].get("identite_nom") or o["nom_brut"]
                p = o["data"].get("identite_prenom") or None
                d = {"id": did, "nom": n, "prenom": p, "cle": cle, "periodes": [], "promos": set(), "_hors_reef": True}
                dossiers[did] = d
                registre["dossiers"][did] = {"cle": cle, "nom": n, "prenom": p}
                nouveaux += 1
            if d.get("_hors_reef") and o["data"].get("identite_nom") and o["data"].get("identite_prenom"):
                d["nom"], d["prenom"] = o["data"]["identite_nom"], o["data"]["identite_prenom"]
                registre["dossiers"][d["id"]].update(nom=d["nom"], prenom=d["prenom"])
                for p in d["periodes"]:
                    p["ligne"]["valeurs"].update(NOM=d["nom"], PRENOM=d["prenom"])
            periode_supplementaire(d, o)
    for d in dossiers.values():
        d["periodes"].sort(key=lambda p:p["ordre"])
    return {"dossiers_crees_cette_execution": nouveaux, "periodes_creees_cette_execution": periodes_nouvelles,
            "dossiers_supplementaires_conserves": len(registre["dossiers"]),
            "periodes_supplementaires_conservees": sum(len(v) for v in registre["periodes"].values())}


# ---------------------------------------------------------------- consolidation
def consolider(reef_path, sources, sortie, journal=print):
    import couverture
    sortie = Path(sortie)
    sortie.mkdir(parents=True, exist_ok=True)
    empreintes = {str(reef_path): sha(reef_path)}
    for p in sources.values():
        if p:
            empreintes[str(p)] = sha(p)
    journal("Lecture de REEF_rebuild (lecture seule)...")
    titre, entetes, lignes, dossiers = lire_reef(reef_path)
    lecture = defaultdict(Counter)
    lecteurs = {"PARCOURS": lire_parcours, "NOTES": lire_notes, "SESSIONS": lire_sessions,
                "ASTRAL": lire_astral, "KLIF": lire_klif}
    toutes = []
    for src, fn in lecteurs.items():
        if sources.get(src):
            journal("Lecture %s..." % src)
            o = fn(sources[src], lecture)
            lecture[src]["observations"] = len(o)
            toutes += o
    decisions, anciennes = lire_decisions(sortie)
    observations_lues = list(toutes)
    journal("Decisions lues : %d (OUI : %d, NON : %d). Fichier : %s" %
            (len(decisions), sum(v[0] == "OUI" for v in decisions.values()),
             sum(v[0] == "NON" for v in decisions.values()), sortie / DECISIONS_NOM))
    nb_dossiers_reef, nb_lignes_reef = len(dossiers), len(lignes)
    retenues = []
    for o in toutes:
        if o["source"] == "KLIF" and promotion_avant_12a(o["promo_brute"]):
            lecture[o["source"]]["observations_anterieures_12A_ignorees"] += 1
        elif o["source"] in ("NOTES", "SESSIONS", "KLIF") and famille_explicitement_exclue(o["promo_brute"]):
            lecture[o["source"]]["observations_hors_cible_ignorees"] += 1
        else:
            retenues.append(o)
    toutes = retenues
    registre = lire_registre_population(sortie)
    exclusions = ecarter_anciennes_periodes_registre(registre)
    population = etendre_population(entetes, lignes, dossiers,
                                    [o for o in toutes
                                     if not (o["source"] == "KLIF" and o["data"].get("conflit_moyenne_identite"))
                                     and (o["source"] != "PARCOURS" or
                                          o["data"].get("niveau") == "SYNTHESE_PARCOURS")],
                                    decisions, registre)
    population.update(exclusions)
    m = Moteur(dossiers, decisions)
    journal("Rapprochement des eleves...")
    statuts = defaultdict(Counter)
    examiner = OrderedDict()
    corresp = OrderedDict()
    par_periode = defaultdict(list)       # (dossier, ordre, degree) -> obs
    par_dossier_klif = defaultdict(list)
    observations_detail = []
    bilan_niveaux = Counter()
    rattachements_couverture = {}

    def signaler(cause, cle, o, dossier=None):
        g = examiner.setdefault((cause,) + cle, {"cause": cause, "dossier": dossier, "o": o, "n": 0,
                                                 "sources": set(), "degrees": set(), "reperes": []})
        g["n"] += 1
        g["sources"].add(o["source"])
        if o["degree"]:
            g["degrees"].add(o["degree"])
        if True:
            g["reperes"].append("%s!%s:%s" % (o["fichier"], o["feuille"], o["ligne"]))

    for o in toutes:
        if o["source"] == "KLIF" and o["data"].get("conflit_moyenne_identite"):
            rattachements_couverture[id(o)] = {
                "dossier": "", "ordre": "", "statut": "KLIF_MOYENNE_AVEC_IDENTITE"}
            statuts["KLIF"]["KLIF_MOYENNE_AVEC_IDENTITE"] += 1
            signaler("KLIF_MOYENNE_AVEC_IDENTITE", (o["feuille"], o["ligne"]), o)
            continue
        d, st = m.identifier(o)
        lien_couverture = {"dossier": d["id"] if d else "", "ordre": "", "statut": st}
        rattachements_couverture[id(o)] = lien_couverture
        statuts[o["source"]][st] += 1
        ck = (o["source"], o["cle"], o["promo"] or "")
        c = corresp.setdefault(ck, {"o": o, "statut": st, "dossier": d, "n": 0})
        c["n"] += 1
        if o["source"] == "PARCOURS" and o["data"].get("niveau") != "SYNTHESE_PARCOURS":
            pd, std = m.periode(d, o) if d else (None, st)
            observations_detail.append((d, pd, o, std))
            bilan_niveaux["observations_" + o["data"].get("niveau", "NIVEAU_INDETERMINE")] += 1
            bilan_niveaux["observations_hors_synthese_rattachees" if pd else
                          "observations_hors_synthese_rattachement_non_resolu"] += 1
        if d is None:
            if st in ("HOMONYMES", "HOMONYMES_MEME_PROMOTION", "SANS_CANDIDAT"):
                signaler(st, ck, o)
            elif st == "IDENTITE_INSUFFISANTE" and o["source"] != "KLIF":
                signaler(st, (o["source"], o["feuille"]), o)
            continue
        if o["source"] == "KLIF":
            lien_couverture["ordre"] = max(p["ordre"] for p in d["periodes"])
            par_dossier_klif[d["id"]].append((o, st))
            continue
        p, stp = m.periode(d, o)
        statuts[o["source"]][stp] += 1
        if p is None:
            signaler(stp, (d["id"], o["source"]), o, d)
            continue
        lien_couverture.update(ordre=p["ordre"], statut=stp)
        if o["degree"]:
            par_periode[(d["id"], p["ordre"], o["degree"])].append(o)
        elif o["source"] == "ASTRAL":
            for dg, q in o["data"]["remedials"].items():
                par_periode[(d["id"], p["ordre"], dg)].append(dict(o, degree=dg, rem=q))

    # ------------------------------------------------ colonnes de sortie
    colonnes = []
    ajout = {}
    derniere = {d: max((i for i, h in enumerate(entetes) if h.startswith(d + "_")), default=None) for d in DEGREES}
    for i, h in enumerate(entetes):
        colonnes.append(h)
        for d in DEGREES:
            if derniere[d] == i:
                colonnes += colonnes_degree(d)
                ajout.update({c: d for c in colonnes_degree(d)})
    for d in DEGREES:
        if derniere[d] is None:
            colonnes += colonnes_degree(d)
            ajout.update({c: d for c in colonnes_degree(d)})
    colonnes += [c for c in COLONNES_POPULATION if c not in colonnes]
    colonnes += KLIF_COLONNES
    provenance = []
    compte = Counter()

    def contexte_observation(o):
        brut = o["data"].get("observation_brute", {})
        return {norm(k): v for k, v in brut.items()}

    def extras(o, cas="", role="", statut="", champ="", valeur=None):
        x = contexte_observation(o)
        brut_json = json.dumps(o["data"].get("observation_brute", {}), ensure_ascii=False, default=str)
        if len(brut_json) > couverture.LIMITE_JSON:
            brut_json = json.dumps({
                "contenu_long_conserve_dans": "COUVERTURE_DONNEES",
                "ID_LIGNE_IMPORT": couverture.identifiant_ligne(o["source"], o["feuille"], o["ligne"]),
                "lecture": "Reunir CELLULES_SOURCE_JSON par PARTIE_JSON ; valeurs indexees par colonne source."
            }, ensure_ascii=False)
        return [cas, o["data"].get("niveau", "SOURCE_" + o["source"]), role or role_de(o["degree"], x.get("role", "")) or x.get("role", ""),
                o["promo_brute"], o["degree"], x.get("resultat", ""), x.get("rattrapage", ""),
                x.get("passage", ""), x.get("evalid", ""), x.get("roleid", ""), x.get("simulation", ""),
                x.get("source", ""), x.get("feuillesource", ""), x.get("lignesource", ""), statut,
                brut_json,
                champ, valeur]

    def tracer(l, col, val, liste, regle):
        # Une ligne par observation, pas une liste tronquee de references.
        for o in liste:
            provenance.append([l["dossier"], l["ordre"], col, val, o["source"],
                               "%s!%s:%s" % (o["fichier"], o["feuille"], o["ligne"]), regle] + extras(o))

    def conflit(l, degree, source, col, vals, groupe, cause="PLUSIEURS_VALEURS_MEME_PERIODE", contexte=None):
        ident = "CAS_" + hashlib.sha256(json.dumps([l["dossier"], l["ordre"], degree, source, col, cause, contexte]).encode()).hexdigest()[:20]
        role = next((rr for rr in ROLES[degree] if ("_" + rr + "_") in col), "")
        candidats = [(o, val) for o, val in zip(groupe, vals) if val is not None]
        raison = ("Valeurs distinctes entre observations comparables ; aucun choix automatique."
                  if cause in ("PLUSIEURS_VALEURS_MEME_PERIODE", "DETAIL_EVALUATION_DIVERGENTE") else
                  "La combinaison note/resultat n'est portee par aucune observation unique ; lien a etablir.")
        signaler(cause, (ident,), groupe[0], dossiers[l["dossier"]])
        gcase = examiner[(cause, ident)]
        gcase.update(cas=ident, periode=l["ordre"], champ=col, role=role,
                     candidats=candidats, raison=raison)
        for o, val in candidats:
            provenance.append([l["dossier"], l["ordre"], col, val, o["source"],
                               "%s!%s:%s" % (o["fichier"], o["feuille"], o["ligne"]),
                               "CANDIDAT_NON_ARBITRE"] + extras(o, ident, role, cause, col, val))
        compte["champs_ambigus_" + source] += 1

    comparables_details = defaultdict(list)
    for did, pd, o, statut in observations_detail:
        x = contexte_observation(o)
        cles = ("source", "feuillesource", "lignesource", "evalid", "roleid", "passage", "rattrapage", "simulation")
        # Ces reperes identifient une observation technique, pas une tentative metier.
        if did and pd and o["degree"] and o["data"].get("niveau") == "DETAIL_EVALUATION" and all(not vide(x.get(k)) for k in cles):
            contexte = tuple(str(x[k]) for k in cles)
            comparables_details[(did["id"], pd["ordre"], o["degree"], contexte)].append(o)
    for (did, ordre, dg, contexte), groupe in comparables_details.items():
        vals = [o["data"].get("res") for o in groupe]
        if len({v for v in vals if v is not None}) > 1:
            conflit({"dossier": did, "ordre": ordre}, dg, "PARCOURS", dg + "_DETAIL_RESULTAT", vals, groupe,
                    "DETAIL_EVALUATION_DIVERGENTE", contexte)
            bilan_niveaux["champs_detail_resultat_divergents"] += 1

    for did, pd, o, statut in observations_detail:
        provenance.append([did["id"] if did else "", pd["ordre"] if pd else "", "OBSERVATION_DETAILLEE",
                           o["data"].get("res"), o["source"],
                           "%s!%s:%s" % (o["fichier"], o["feuille"], o["ligne"]),
                           "DETAIL_CONSERVE_SANS_RESULTAT_GLOBAL_DEDUIT"] + extras(o, statut=statut))

    def unique(vals):
        s = {round(v, 2) if isinstance(v, float) else v for v in vals if v is not None}
        return (next(iter(s)), True) if len(s) == 1 else (None, len(s) > 1)

    sortie_lignes = []
    for l in lignes:
        v = l["valeurs"]
        r = OrderedDict((c, v.get(c)) for c in colonnes)
        if "ORIGINE_LIGNE_CONSOLIDEE" not in entetes:
            r["ORIGINE_LIGNE_CONSOLIDEE"] = l.get("_origine_population", "REEF_REBUILD")
        if "NATURE_ORDRE_PERIODE" not in entetes:
            r["NATURE_ORDRE_PERIODE"] = l.get("_nature_ordre", "ORDRE_REEF_CONSERVE")
        if l.get("_origine_population"):
            justifs = [o for o in toutes if o["promo"] == promo(v.get("PROMOTION"))[1] and
                       (o["cle"] == dossiers[l["dossier"]]["cle"] or
                        (o["source"], o["cle"], o["promo"] or "", l["dossier"]) in decisions and
                        decisions[(o["source"], o["cle"], o["promo"] or "", l["dossier"])][0] == "OUI")]
            tracer(l, "ORIGINE_LIGNE_CONSOLIDEE", r["ORIGINE_LIGNE_CONSOLIDEE"], justifs,
                   "PERIODE_SOURCE_ORDRE_TECHNIQUE_PAS_CHRONOLOGIE_DE_PASSAGES")
        for d in DEGREES:
            etat = []
            os_ = par_periode.get((l["dossier"], l["ordre"], d), [])
            groupes = defaultdict(list)
            for o in os_:
                groupes[o["source"]].append(o)
            for src, cible in (("PARCOURS", "PARCOURS"), ("NOTES", "SOCRATES")):
                g = groupes.get(src, [])
                if src == "PARCOURS":
                    avant = {x["data"].get("res") for x in g if x["data"].get("res") is not None}
                    if len(avant) > 1:
                        bilan_niveaux["champs_resultat_global_mixtes_avant"] += 1
                    g = [x for x in g if x["data"].get("niveau") == "SYNTHESE_PARCOURS"]
                    apres = {x["data"].get("res") for x in g if x["data"].get("res") is not None}
                    if len(apres) > 1:
                        bilan_niveaux["champs_resultat_global_synthese_divergents_apres"] += 1
                    if len(avant) > 1 and len(apres) <= 1:
                        bilan_niveaux["champs_resultat_global_separes"] += 1
                        categorie = ("champs_resultat_global_retablis_depuis_synthese" if apres else
                                     "champs_resultat_global_sans_resultat_synthese")
                        bilan_niveaux[categorie] += 1
                if not g:
                    continue
                champs = [(d + "_%s_NOTE" % cible, [x["data"]["note"] for x in g])]
                if src == "PARCOURS":
                    champs.append((d + "_PARCOURS_RESULTAT", [x["data"]["res"] for x in g]))
                    champs.extend([
                        (d + "_PARCOURS_NOTE_RECONSTITUEE", [x["data"].get("note_reconstituee") for x in g]),
                        (d + "_PARCOURS_ORIGINE_NOTE_RECONSTITUEE",
                         [x["data"].get("origine_note_reconstituee") or None
                          if x["data"].get("note_reconstituee") is not None else None for x in g]),
                        (d + "_PARCOURS_STATUT_NOTE_RECONSTITUEE",
                         [x["data"].get("statut_note_reconstituee") or None
                          if x["data"].get("note_reconstituee") is not None else None for x in g]),
                    ])
                    if d in ("APS1", "APS2"):
                        champs.append((d + "_PL_PARCOURS_RESULTAT",
                                       [x["data"]["roles"].get("PL", {}).get("res") for x in g]))
                for role in ROLES[d]:
                    vals = [x["data"]["roles"].get(role, {}).get("note") for x in g]
                    if len(ROLES[d]) > 1:
                        champs.append(("%s_%s_%s_NOTE" % (d, role, cible), vals))
                    elif (all(x is None for x in champs[0][1])
                          and not any(norm(x["data"].get("observation_brute", {}).get("Note")) == "nogo"
                                      for x in g)):
                        champs[0] = (champs[0][0], vals)
                # Ne pas fabriquer une combinaison note/resultat depuis des lignes disjointes.
                uniques = [(col, unique(vals)[0], vals) for col, vals in champs if unique(vals)[0] is not None]
                coherent = not uniques or any(all(vals[i] is not None and unique([vals[i]])[0] == value
                                                   for col, value, vals in uniques) for i in range(len(g)))
                for col, vals in champs:
                    val, ok = unique(vals)
                    if val is not None and not coherent:
                        r[col] = "PLUSIEURS_VALEURS"
                        conflit(l, d, src, col, vals, g, "SYNTHESE_COMPOSITE_NON_JUSTIFIEE")
                        etat.append(src + ":SYNTHESE_COMPOSITE_NON_JUSTIFIEE")
                        continue
                    if val is not None:
                        r[col] = val
                        tracer(l, col, val, [o for o, vv in zip(g, vals)
                                           if vv is not None and unique([vv])[0] == val],
                               "VALEUR_UNIQUE_DANS_LA_PERIODE")
                        compte["valeurs_ajoutees_" + src] += 1
                    elif ok:
                        r[col] = "PLUSIEURS_VALEURS"
                        etat.append("%s:VALEURS_MULTIPLES_NON_ARBITREES" % src)
                        conflit(l, d, src, col, vals, g)
                # comparaison avec REEF (notes par role)
                for role in ROLES[d]:
                    rc = d + "_NOTE" if d in ("BTWR", "ADI") else "%s_%s_NOTE" % (d, role)
                    cc = (d + "_%s_NOTE" % cible) if d in ("BTWR", "ADI") else "%s_%s_%s_NOTE" % (d, role, cible)
                    a, b = note(v.get(rc)), r.get(cc)
                    if a is not None and isinstance(b, float):
                        if abs(a - b) > 0.011:
                            etat.append("%s:ECART_NOTE_%s" % (src, role))
                            compte["ecarts_note"] += 1
                        else:
                            compte["concordances_note"] += 1
            g = groupes.get("SESSIONS", [])
            if g:
                ds = sorted(x["data"]["date"] for x in g if x["data"]["date"])
                r[d + "_SESSIONS_NB"] = len(g)
                r[d + "_SESSIONS_NOGO"] = sum(1 for x in g if x["data"]["res"] == "NOGO")
                r[d + "_SESSIONS_PREMIERE"] = ds[0] if ds else None
                r[d + "_SESSIONS_DERNIERE"] = ds[-1] if ds else None
                tracer(l, d + "_SESSIONS_NB", len(g), g, "SESSIONS_DE_LA_PERIODE")
                compte["valeurs_ajoutees_SESSIONS"] += 1
            g = groupes.get("ASTRAL", [])
            if g:
                val, ok = unique([x["rem"] for x in g])
                r[d + "_ASTRAL_REMEDIALS"] = val if val is not None else "PLUSIEURS_VALEURS"
                tracer(l, d + "_ASTRAL_REMEDIALS", r[d + "_ASTRAL_REMEDIALS"], g,
                       "COMPTEUR_ASTRAL_EXPLICITEMENT_SOURCE")
                if val is not None and val > 0:
                    r[d + "_ASTRAL_PRESENCE_REMEDIAL"] = "OUI"
                    tracer(l, d + "_ASTRAL_PRESENCE_REMEDIAL", "OUI", g,
                           "MENTION_EXPLICITE_REMEDIAL_SOURCE")
                compte["valeurs_ajoutees_ASTRAL"] += 1
            if groupes:
                r[d + "_ETAT_SOURCES"] = "; ".join(etat) if etat else "SOURCES:" + "+".join(sorted(groupes))
        sortie_lignes.append(r)

    # ------------------------------------------------ KLIF (une fois par dossier, derniere periode)
    dernieres = {}
    for r, l in zip(sortie_lignes, lignes):
        if l["dossier"] not in dernieres or l["ordre"] >= dernieres[l["dossier"]][1]["ordre"]:
            dernieres[l["dossier"]] = (r, l)
    for did, (r, l) in dernieres.items():
        cands = par_dossier_klif.get(did, [])
        for rr, ll in zip(sortie_lignes, lignes):
            if ll["dossier"] == did and ll is not l:
                rr["KLIF_PORTEE"] = "VOIR_PERIODE_%s" % l["ordre"]
        if not cands:
            r["KLIF_PORTEE"] = "AUCUNE_LIGNE_KLIF_RACCORDEE"
            continue
        if len(cands) > 1:
            affectes = [c for c in cands if not vide(c[0]["data"]["affectation"])]
            signature = {tuple(propre(c[0]["data"][k]) for k in ("affectation", "date_affectation", "date_mu_1"))
                         for c in affectes}
            if len(signature) == 1:
                cands = [affectes[0]]
            elif not affectes:
                cands = [cands[-1]]
            else:
                r["KLIF_PORTEE"] = "PLUSIEURS_LIGNES_KLIF_DIVERGENTES"
                signaler("KLIF_PLUSIEURS_LIGNES_DIVERGENTES", (did,), cands[0][0], dossiers[did])
                continue
        o, st = cands[0]
        r["KLIF_PORTEE"] = "SUIVI_KLIF_DU_DOSSIER"
        r["KLIF_LIEN"] = st
        r["KLIF_FEUILLE"] = o["feuille"]
        for col, k in KLIF_BRUT.items():
            r[col] = o["data"][k]
        calc, alertes = interpreter_klif(o["data"])
        r.update(calc)
        r["KLIF_ALERTES"] = "; ".join(alertes)
        tracer(l, "KLIF_*", o["data"]["affectation"], [o], "KLIF_" + st)
        for col in ("KLIF_ORIGINE", "KLIF_MOIS_QUALIFICATION_ENAC_EXCLUS"):
            if not vide(r[col]):
                tracer(l, col, r[col], [o], "VALEUR_SOURCE_KLIF_NON_SUBSTITUEE")
        compte["dossiers_avec_suivi_klif"] += 1
        compte["issue_" + calc["KLIF_ISSUE_AFFECTATION_1"]] += 1

    # ------------------------------------------------ ecriture
    horo = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    local = sortie / "LOCAL"
    part = sortie / "PARTAGEABLE"
    local.mkdir(exist_ok=True)
    part.mkdir(exist_ok=True)
    fichier = local / ("PARCOURS_CONSOLIDES_%s.xlsx" % horo)
    n = 2
    while fichier.exists():
        fichier = local / ("PARCOURS_CONSOLIDES_%s_%d.xlsx" % (horo, n))
        n += 1
    journal("Ecriture du fichier consolide...")
    journal("Inventaire independant et controle de couverture des sources...")
    controle = couverture.construire(
        reef_path, sources, observations_lues, rattachements_couverture, provenance,
        sortie_lignes, entetes, {k: globals()[k] for k in
            ("norm", "promo", "degree_de", "degree_entete_astral", "role_de", "promotion_avant_12a",
             "famille_explicitement_exclue", "est_moyenne_klif", "sha", "note", "nombre",
             "resultat", "date", "FEUILLES_IGNOREES")})
    ecrire_classeur(fichier, colonnes, entetes, ajout, sortie_lignes, provenance, examiner, corresp, reef_path,
                    sources, lecture, controle)
    journal("Verification des cellules du fichier Excel enregistre...")
    verification = verifier_classeur(fichier, colonnes, sortie_lignes, entetes, len(provenance))
    verification.update(verifier_tracabilite(fichier, provenance, examiner))
    bilan_couverture = controle.verifier(fichier)
    journal("Couverture : %s ; exploitation : %s." %
            (bilan_couverture["collecte"], bilan_couverture["exploitation"]))
    for src, n in verification["cellules_non_vides_par_source"].items():
        journal("%s : %d cellules non vides relues dans le fichier Excel." % (src, n))
    ecrire_decisions(sortie, m.propositions, anciennes)
    for p, h in empreintes.items():
        if sha(p) != h:
            raise Erreur("Une source a change pendant le traitement : " + Path(p).name)
    couverts = defaultdict(set)
    for (did, ordre, dg), os_ in par_periode.items():
        for o in os_:
            couverts[o["source"]].add(did)
    for did in par_dossier_klif:
        couverts["KLIF"].add(did)
    bilan = {
        "version": VERSION, "revision": REVISION, "execution": "TERMINEE",
        "fichier_cible": fichier.name, "empreinte_sha256_fichier": sha(fichier),
        "reef_rebuild_modifie": False, "sources_inchangees": True,
        "dossiers_reef": nb_dossiers_reef, "lignes_periodes_reef": nb_lignes_reef,
        "dossiers_consolides": len(dossiers), "lignes_periodes": len(lignes), "population_ouverte": population, "separation_niveaux": dict(bilan_niveaux),
        "lecture": {k: dict(v) for k, v in lecture.items()},
        "statuts_rapprochement": {k: dict(v) for k, v in statuts.items()},
        "dossiers_enrichis_par_source": {k: len(v) for k, v in sorted(couverts.items())},
        "valeurs_restituees": dict(compte),
        "points_a_examiner_par_cause": dict(Counter(g["cause"] for g in examiner.values())),
        "propositions_de_correspondance": len(m.propositions),
        "decisions_lues": len(decisions),
        "decisions_oui_lues": sum(v[0] == "OUI" for v in decisions.values()),
        "decisions_non_lues": sum(v[0] == "NON" for v in decisions.values()),
        "decisions_appliquees": len(m.decisions_utilisees),
        "decisions_sans_observation_retrouvee": len(decisions) - len(m.decisions_utilisees),
        "verification_fichier_enregistre": verification,
        "couverture_donnees": bilan_couverture,
        "regles": ["REEF_REBUILD_LU_SEULEMENT", "KLIF_FEUILLES_ANTERIEURES_12A_IGNOREES",
                   "AUCUNE_VALEUR_REEF_ECRASEE", "PROPOSITIONS_UTILISEES_APRES_DECISION_OUI",
                   "VALEUR_VIDE_NON_CONVERTIE_EN_ECHEC", "PLUSIEURS_PASSAGES_NON_FUSIONNES",
                   "KLIF_UNE_FOIS_PAR_DOSSIER_SUR_SA_DERNIERE_PERIODE",
                   "PROMOTIONS_P_EXCLUES_NOTES_SESSIONS_KLIF", "AVANT_12A_EXCLU_KLIF_SEULEMENT",
                   "ASTRAL_PRESENCE_EXPLICITE", "NOTE_RECONSTITUEE_DISTINCTE",
                   "KLIF_DUREE_ENAC_EXCLUS_SOURCE"]}
    (part / ("BILAN_CONSOLIDATION_%s.json" % horo)).write_text(json.dumps(bilan, ensure_ascii=False, indent=2),
                                                                encoding="utf-8")
    ecrire_registre_population(sortie, registre)
    journal("Termine : %s" % fichier)
    return fichier, bilan


def verifier_classeur(fichier, colonnes, attendues, entetes_reef, nb_provenance):
    """Relit le fichier produit, compare chaque cellule et compte les ajouts reels."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from itertools import zip_longest
    from math import isclose

    def source_colonne(c):
        if c in entetes_reef:
            return None
        for src, marqueur in (("PARCOURS", "_PARCOURS_"), ("NOTES", "_SOCRATES_"),
                              ("SESSIONS", "_SESSIONS_"), ("ASTRAL", "_ASTRAL_")):
            if marqueur in c:
                return src
        return None

    compteurs = {s: 0 for s in ("PARCOURS", "NOTES", "SESSIONS", "ASTRAL")}
    premiers = {}
    compteurs_degree = defaultdict(Counter)
    wb = load_workbook(fichier, read_only=True, data_only=False)
    try:
        ws = wb["PARCOURS_CONSOLIDES"]
        rows = ws.iter_rows(values_only=True)
        if list(next(rows, ())) != colonnes:
            raise Erreur("Verification Excel : les entetes enregistrees different des entetes attendues.")
        nb_lignes = 0
        for no, (row, attendu) in enumerate(zip_longest(rows, attendues), 2):
            if row is None or attendu is None:
                raise Erreur("Verification Excel : nombre de lignes incorrect.")
            nb_lignes += 1
            for j, c in enumerate(colonnes):
                a, b = row[j], valeur(attendu.get(c))
                egaux = a == b or (vide(a) and vide(b))
                if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                    egaux = isclose(a, b, rel_tol=1e-12, abs_tol=1e-9)
                if isinstance(a, dt.datetime) and isinstance(b, dt.datetime):
                    egaux = abs((a - b).total_seconds()) < 0.001
                adresse = "%s%d" % (get_column_letter(j + 1), no)
                if not egaux:
                    raise Erreur("Verification Excel : cellule %s differente de la valeur attendue (%s)." %
                                 (adresse, c))
                src = source_colonne(c)
                if src and not vide(a):
                    compteurs[src] += 1
                    premiers.setdefault(src, {"cellule": adresse, "colonne": c})
                    dg = next((d for d in DEGREES if c.startswith(d + "_")), None)
                    if dg:
                        compteurs_degree[src][dg] += 1
        n_prov = sum(1 for _ in wb["PROVENANCE"].iter_rows(min_row=2, values_only=True))
        if n_prov != nb_provenance:
            raise Erreur("Verification Excel : nombre de lignes de provenance incorrect.")
    finally:
        wb.close()
    return {"statut": "CONFORME_APRES_RELECTURE", "lignes_verifiees": nb_lignes,
            "cellules_non_vides_par_source": compteurs, "premieres_cellules_par_source": premiers,
            "lignes_provenance_verifiees": n_prov,
            "cellules_non_vides_par_source_et_degree": {s: dict(v) for s, v in compteurs_degree.items()}}


COLONNES_TRACABILITE = ["ID_CAS", "NIVEAU", "ROLE", "PROMOTION_SOURCE", "DEGREE", "RESULTAT_OBSERVE",
    "RATTRAPAGE_SOURCE", "PASSAGE_TECHNIQUE_SOURCE", "EVALID_SOURCE", "ROLEID_SOURCE", "SIMULATION",
    "ORIGINE_HISTORIQUE", "FEUILLE_HISTORIQUE", "LIGNE_HISTORIQUE", "STATUT_RATTACHEMENT",
    "OBSERVATION_SOURCE_JSON", "CHAMP_CANDIDAT", "VALEUR_CANDIDATE"]


def valeur_provenance(v):
    """Type Excel sans perte pour les metadonnees de tracabilite.

    Les donnees numeriques a plus de 15 chiffres significatifs et les dates
    sous-milliseconde restent du texte EXACT dans PROVENANCE. Une valeur
    ordinaire conserve son type Excel. Les controles ne sont pas assouplis.
    """
    import math
    v = valeur(v)
    if isinstance(v, str):
        if len(v) > 32767:
            raise Erreur("Provenance trop longue pour une cellule Excel : "
                         "%d caracteres ; conservation impossible sans reference explicite." % len(v))
        return v
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return str(v) if len(str(abs(v))) > 15 else v
    if isinstance(v, float):
        if not math.isfinite(v) or float(format(v, ".15g")) != v:
            return repr(v)
        return v
    if isinstance(v, dt.datetime):
        if v.microsecond % 1000:
            return v.isoformat(sep=" ")
        return v
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, dt.time):
        return v.isoformat()
    return v


def verifier_tracabilite(fichier, provenance, examiner):
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from itertools import zip_longest
    wb = load_workbook(fichier, read_only=True, data_only=False)
    try:
        feuille = wb["PROVENANCE"]
        titres = list(next(feuille.iter_rows(min_row=1, max_row=1, values_only=True), ()))
        attendus_titres = ["ID_DOSSIER_LOCAL", "ORDRE_PERIODE", "COLONNE", "VALEUR", "SOURCE",
                           "REPERES (fichier!feuille:ligne)", "REGLE"] + COLONNES_TRACABILITE
        if titres != attendus_titres:
            raise Erreur("Provenance : entetes differentes apres sauvegarde.")
        it = feuille.iter_rows(min_row=2, values_only=True)
        for no, (row, attendu) in enumerate(zip_longest(it, provenance), 2):
            if row is None or attendu is None:
                raise Erreur("Provenance incomplete apres sauvegarde : ligne %d." % no)
            exp = [valeur_provenance(v) for v in attendu]
            exp += [None] * (len(attendus_titres) - len(exp))
            if len(row) != len(exp):
                raise Erreur("Provenance : nombre de colonnes different ligne %d." % no)
            for j, (a, b) in enumerate(zip(row, exp)):
                if a == b or vide(a) and vide(b):
                    continue
                # Diagnostic sans affichage de la valeur source (donnees individuelles).
                def typ(v):
                    return "texte(%d caracteres)" % len(v) if isinstance(v, str) else type(v).__name__
                raise Erreur("Provenance differente apres sauvegarde : ligne %d, colonne %s (%s), "
                             "type enregistre %s, type attendu %s. Valeurs non affichees." % (
                                 no, get_column_letter(j + 1), titres[j], typ(a), typ(b)))
        canon = lambda v: ("N", float(v)) if isinstance(v, (int, float)) else ("T", str(v))
        attentes = Counter((g["cas"], canon(val), "%s!%s:%s" % (o["fichier"], o["feuille"], o["ligne"]))
                           for g in examiner.values() if g.get("cas") for o, val in g["candidats"])
        rows = wb["A_EXAMINER"].iter_rows(values_only=True); h = next(rows)
        vus = Counter()
        for row in rows:
            x = dict(zip(h, row))
            if x.get("ID_CAS"):
                vus[(x["ID_CAS"], canon(x["VALEUR_CANDIDATE"]), x["REPERES"])] += 1
        if attentes != vus:
            raise Erreur("Candidats A_EXAMINER differents apres sauvegarde.")
        return {"provenance_cellules_verifiees": True, "candidats_a_examiner_verifies": sum(vus.values())}
    finally:
        wb.close()


def ecrire_classeur(fichier, colonnes, entetes, ajout, lignes, provenance, examiner, corresp, reef, sources, lecture,
                    controle_couverture=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "PARCOURS_CONSOLIDES"
    ws.append(colonnes)
    vert = PatternFill("solid", fgColor="E2EFDA")
    bleu = PatternFill("solid", fgColor="DDEBF7")
    for j, c in enumerate(colonnes, 1):
        cell = ws.cell(1, j)
        cell.font = Font(bold=True)
        if c in ajout:
            cell.fill = vert
        elif c.startswith("KLIF_"):
            cell.fill = bleu
    for r in lignes:
        ws.append([valeur(r[c]) for c in colonnes])
    for j, c in enumerate(colonnes, 1):
        if c in ajout:
            ws.column_dimensions[get_column_letter(j)].outlineLevel = 1
        if isinstance(lignes[0].get(c) if lignes else None, dt.datetime) or "DATE" in c or "PREMIERE" in c or "DERNIERE" in c:
            for row in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                if isinstance(row[0].value, dt.datetime):
                    row[0].number_format = "DD/MM/YYYY"
    ws.freeze_panes = "G2"
    ws.auto_filter.ref = ws.dimensions
    p = wb.create_sheet("PROVENANCE")
    p.append(["ID_DOSSIER_LOCAL", "ORDRE_PERIODE", "COLONNE", "VALEUR", "SOURCE", "REPERES (fichier!feuille:ligne)", "REGLE"] + COLONNES_TRACABILITE)
    for x in provenance:
        p.append([valeur_provenance(y) for y in x])
    e = wb.create_sheet("A_EXAMINER")
    e.append(["CAUSE", "ID_DOSSIER_LOCAL", "NOM_REEF", "PRENOM_REEF", "NOM_SOURCE", "PROMO_SOURCE",
              "SOURCES", "DEGREES", "OBSERVATIONS", "REPERES", "AIDE", "ID_CAS", "ORDRE_PERIODE",
              "CHAMP", "ROLE", "VALEUR_CANDIDATE", "NIVEAU", "ORIGINE_HISTORIQUE", "FEUILLE_HISTORIQUE", "LIGNE_HISTORIQUE"])
    for g in examiner.values():
        d = g["dossier"]
        candidats = g.get("candidats") or [(g["o"], None)]
        for o, val in candidats:
            brut = {norm(k): v for k, v in o["data"].get("observation_brute", {}).items()}
            ref = "%s!%s:%s" % (o["fichier"], o["feuille"], o["ligne"])
            e.append([g["cause"], d["id"] if d else "", d["nom"] if d else "", d["prenom"] if d else "",
                      o["nom_brut"], o["promo_brute"], ", ".join(sorted(g["sources"])),
                      ", ".join(sorted(g["degrees"])), len(candidats) if g.get("candidats") else g["n"],
                      ref if g.get("candidats") else " ; ".join(g["reperes"]), g.get("raison", ""),
                      g.get("cas", ""), g.get("periode"), g.get("champ", ""), g.get("role", ""), val,
                      o["data"].get("niveau", "SOURCE_" + o["source"]), brut.get("source"),
                      brut.get("feuillesource"), brut.get("lignesource")])
    c = wb.create_sheet("CORRESPONDANCES")
    c.append(["SOURCE", "NOM_SOURCE", "PROMO_SOURCE", "STATUT", "ID_DOSSIER_LOCAL", "NOM_REEF", "PRENOM_REEF",
              "OBSERVATIONS"])
    for (src, _, _), x in corresp.items():
        d = x["dossier"]
        c.append([src, x["o"]["nom_brut"], valeur(x["o"]["promo_brute"]), x["statut"], d["id"] if d else "",
                  d["nom"] if d else "", d["prenom"] if d else "", x["n"]])
    li = wb.create_sheet("LISEZMOI")
    for t in [
        "Fichier consolide des parcours %s (%s) - genere le %s" %
        (VERSION, REVISION, dt.datetime.now().strftime("%d/%m/%Y %H:%M")),
        "REEF_rebuild source (lu seulement) : %s" % Path(reef).name,
        "Autres sources : " + ", ".join("%s=%s" % (k, Path(v).name) for k, v in sources.items() if v),
        "",
        "Colonnes REEF : valeurs originales preservees pour les lignes REEF ; cellules non documentees vides sur les nouvelles lignes.",
        "BASIC_TOWER = BTWR ; BASIC_ACC = BSURV. Cordon = PL pour APS1/APS2.",
        "Resultats PL Parcours separes du resultat global ; aucun resultat deduit d'une note.",
        "Syntheses Parcours seules dans les champs de synthese ; details complets dans PROVENANCE (OBSERVATION_DETAILLEE).",
        "Passage technique et mentions de rattrapage conserves tels que fournis, sans chronologie metier deduite.",
        "Les nouvelles periodes ont un ordre TECHNIQUE, pas une chronologie de tentatives demontree.",
        "Une ligne issue seulement de KLIF n'etablit aucun passage de Degree.",
        "Colonnes vertes (groupees, repliables) : complements Parcours / Notes Socrates / Sessions / ASTRAL pour la meme periode.",
        "Colonnes bleues : suivi KLIF du dossier, affiche une seule fois, sur la derniere periode du dossier.",
        "Une valeur n'est ajoutee que si l'eleve est retrouve ET la promotion de la donnee correspond a une seule periode.",
        "PLUSIEURS_VALEURS : valeurs de synthese non arbitrables ; ne prouve pas plusieurs passages (voir A_EXAMINER).",
        "ETAT_SOURCES : sources presentes, ou ecart avec la note REEF (aucune valeur REEF n'est remplacee).",
        "ISSUE_AFFECTATION_1 : QUALIFIEE (date MU), ECHEC_ARRET_DE_FORMATION (arret declare), SORTIE, AUTRE_CURSUS,",
        "  DISPONIBILITE, NON_QUALIFIEE_A_LA_DATE_DU_FICHIER (affecte sans MU ni arret), NON_DETERMINEE. Un vide n'est jamais un echec.",
        "Durees en mois (jours/30,4375) recalculees a partir des dates ; une duree negative n'est pas restituee (alerte).",
        "KLIF_MOIS_QUALIFICATION_ENAC_EXCLUS : duree SOURCE en mois, jamais remplacee par les durees recalculees.",
        "KLIF_ORIGINE : valeur telle que fournie, sans recodage de CE / CI / promotions / corps.",
        "ASTRAL : presence OUI seulement si mention remedial explicite ; vide ne signifie pas NON.",
        "NOTE=NOGO : note reconstituee source conservee distinctement ; NOGO et resultats conserves.",
        "Feuilles KLIF anterieures a 12A : non etudiees.",
        "Les promotions P sont exclues de Notes, Sessions et KLIF, quelle que soit l'annee.",
        "Avant 12A : seule la source KLIF est exclue automatiquement par date.",
        "Les variantes de nom ne sont pas utilisees automatiquement : decider OUI/NON dans DECISIONS_CORRESPONDANCES.xlsx,",
        "  puis relancer. Les decisions sont conservees quand les sources sont enrichies.",
        "",
        "Lecture des sources : " + json.dumps({k: dict(v) for k, v in lecture.items()}, ensure_ascii=False),
        "FICHIER NOMINATIF : ne pas transmettre. Seul le dossier PARTAGEABLE peut l'etre."]:
        li.append([t])
    if controle_couverture is not None:
        controle_couverture.ajouter_feuilles(wb)
    wb.save(fichier)


def ecrire_decisions(sortie, propositions, anciennes):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    p = Path(sortie) / DECISIONS_NOM
    h = ["SOURCE", "NOM_SOURCE", "PROMO_SOURCE", "CLE_SOURCE", "ID_DOSSIER_LOCAL", "NOM_REEF", "PRENOM_REEF",
         "PROMOTIONS_REEF", "CLE_REEF", "MOTIF", "OBSERVATIONS", "DECISION"]
    vus, lignes = set(), []
    for v in anciennes:
        k = (propre(v.get("SOURCE")), propre(v.get("CLE_SOURCE")), propre(v.get("PROMO_SOURCE")),
             propre(v.get("ID_DOSSIER_LOCAL")))
        if k not in vus:
            vus.add(k)
            lignes.append([v.get(x) for x in h])
    for k, x in propositions.items():
        if k in vus:
            continue
        d, o = x["dossier"], x["o"]
        lignes.append([k[0], o["nom_brut"], k[2], k[1], d["id"], d["nom"], d["prenom"],
                       ", ".join(p_["promo"] or "" for p_ in d["periodes"]), d["cle"], x["motif"], x["obs"], ""])
    backup = None
    if p.exists():
        backup = p.with_name("DECISIONS_CORRESPONDANCES_sauvegarde_%s.xlsx"
                             % dt.datetime.now().strftime("%Y%m%d_%H%M%S"))
        try:
            shutil.copy2(p, backup)
        except PermissionError:
            raise Erreur("Fermez DECISIONS_CORRESPONDANCES.xlsx dans Excel, puis relancez.")
    wb = Workbook()
    ws = wb.active
    ws.title = "DECISIONS"
    ws.append(h)
    for j in range(1, len(h) + 1):
        ws.cell(1, j).font = Font(bold=True)
    ws.cell(1, len(h)).fill = PatternFill("solid", fgColor="FFF2CC")
    for l in lignes:
        ws.append([valeur(x) for x in l])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    try:
        wb.save(p)
    except PermissionError:
        if backup is not None:
            try:
                backup.unlink()
            except OSError:
                pass
        raise Erreur("Fermez DECISIONS_CORRESPONDANCES.xlsx dans Excel, puis relancez.")


def valeur(v):
    if isinstance(v, (Counter, set, dict, list)):
        return str(v)
    return v


# ---------------------------------------------------------------- interface
def interface():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext
    import threading
    root = tk.Tk()
    root.title("Fichier consolide des parcours %s" % VERSION)
    champs = OrderedDict([("REEF", "REEF_REBUILD.xlsx (obligatoire, lu seulement)"),
                          ("PARCOURS", "Parcours agreges"), ("NOTES", "Export Notes Socrates"),
                          ("SESSIONS", "Export Sessions Socrates"), ("ASTRAL", "Export ASTRAL"),
                          ("KLIF", "KLIF ICNA"), ("SORTIE", "Dossier de resultats")])
    var = {k: tk.StringVar() for k in champs}
    cfg = Path(__file__).with_name("derniers_choix.json")
    if cfg.exists():
        try:
            for k, v in json.loads(cfg.read_text(encoding="utf-8")).items():
                if k in var:
                    var[k].set(v)
        except Exception:
            pass
    for i, (k, lib) in enumerate(champs.items()):
        tk.Label(root, text=lib, anchor="w", width=42).grid(row=i, column=0, sticky="w", padx=6, pady=2)
        tk.Entry(root, textvariable=var[k], width=70).grid(row=i, column=1, padx=4)

        def choisir(k=k):
            p = filedialog.askdirectory() if k == "SORTIE" else filedialog.askopenfilename(
                filetypes=[("Excel", "*.xlsx *.xlsm")])
            if p:
                var[k].set(p)
        tk.Button(root, text="Choisir...", command=choisir).grid(row=i, column=2, padx=4)
    log = scrolledtext.ScrolledText(root, width=110, height=14)
    log.grid(row=len(champs) + 1, column=0, columnspan=3, padx=6, pady=6)

    def ecrire(t):
        root.after(0, lambda: (log.insert("end", t + "\n"), log.see("end")))

    def lancer():
        v = {k: var[k].get().strip() for k in champs}
        if not v["REEF"] or not v["SORTIE"]:
            messagebox.showerror("Manquant", "Choisissez REEF_REBUILD.xlsx et le dossier de resultats.")
            return
        if any(dossier_synchronise(x) for x in v.values() if x) and not messagebox.askyesno(
                "Dossier synchronise", "Un emplacement semble synchronise (OneDrive/Dropbox). Le fichier produit est "
                "nominatif et serait synchronise. Continuer quand meme ?"):
            return
        cfg.write_text(json.dumps(v, ensure_ascii=False), encoding="utf-8")

        def travail():
            try:
                f, b = consolider(v["REEF"], {k: v[k] for k in ("PARCOURS", "NOTES", "SESSIONS", "ASTRAL", "KLIF")},
                                  v["SORTIE"], ecrire)
                ecrire(json.dumps(b["dossiers_enrichis_par_source"], ensure_ascii=False))
                resume = "\n".join("%s : %d cellules non vides" % (src, n) for src, n in
                                  b["verification_fichier_enregistre"]["cellules_non_vides_par_source"].items())
                message = ("Fichier produit et relu :\n%s\n\n%s\n\nDecisions lues : %d ; appliquees : %d.\n"
                           "Ouvrez ce nouveau fichier dans LOCAL (les anciens ne sont pas actualises).") % (
                               f, resume, b["decisions_lues"], b["decisions_appliquees"])
                message += ("\n\nCouverture : %s\nExploitation : %s\n"
                            "Voir les feuilles COUVERTURE_* et le bilan PARTAGEABLE.") % (
                                b["couverture_donnees"]["collecte"], b["couverture_donnees"]["exploitation"])
                root.after(0, lambda message=message: messagebox.showinfo("Termine", message))
            except Exception as e:
                ecrire("ERREUR : %s" % e)
                root.after(0, lambda e=e: messagebox.showerror("Erreur", str(e)))
        threading.Thread(target=travail, daemon=True).start()
    tk.Button(root, text="Produire le fichier consolide", command=lancer, bg="#DDEBF7").grid(
        row=len(champs), column=1, pady=6)
    root.mainloop()


def main():
    a = argparse.ArgumentParser()
    a.add_argument("--autotest", action="store_true")
    for k in ("reef", "parcours", "notes", "sessions", "astral", "klif", "sortie"):
        a.add_argument("--" + k)
    x = a.parse_args()
    if x.autotest:
        import test_consolidation, test_separation, test_couverture, test_lot_c1_c8
        historique = test_consolidation.main()
        separation = test_separation.main()
        couverture = test_couverture.main()
        lot_c1_c8 = test_lot_c1_c8.main()
        sys.exit(0 if all(x == 0 for x in (historique, separation, couverture, lot_c1_c8)) else 1)
    if x.reef and x.sortie:
        f, b = consolider(x.reef, {"PARCOURS": x.parcours, "NOTES": x.notes, "SESSIONS": x.sessions,
                                    "ASTRAL": x.astral, "KLIF": x.klif}, x.sortie)
        print(json.dumps(b, ensure_ascii=False, indent=2))
        return
    interface()


if __name__ == "__main__":
    main()

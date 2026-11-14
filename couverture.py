# -*- coding: utf-8 -*-
"""Controle de couverture local. Ne modifie ni les sources ni les regles metier."""
import datetime as dt
import hashlib
import json
import re
from collections import Counter, defaultdict
from itertools import zip_longest
from pathlib import Path

REVISION = "2026-11-13-lot-c1-c8"
LIMITE_JSON = 14000  # au plus 28000 unites UTF-16, meme avec caracteres hors BMP
ETATS = ("RESTITUEE_EN_SYNTHESE", "CONSERVEE_EN_DETAIL", "EN_ATTENTE_CONSERVEE")
H_FEUILLES = ["SOURCE", "FICHIER_LOCAL", "FEUILLE", "STATUT", "MOTIF", "LIGNES_EXAMINEES",
              "LIGNES_NON_VIDES", "LIGNES_CANDIDATES", "OBSERVATIONS_LECTEUR",
              "LIGNES_CANDIDATES_SANS_OBSERVATION", "FORMULES_SANS_CACHE"]
H_DONNEES = ["ID_LIGNE_IMPORT", "SOURCE", "FICHIER_LOCAL", "FEUILLE", "LIGNE", "NATURE",
             "STATUT", "MOTIF", "PROMOTION", "DEGREE", "ID_DOSSIER_LOCAL", "ORDRE_PERIODE",
             "DESTINATIONS", "CHAMPS_EN_CAUSE", "PARTIE_JSON", "NB_PARTIES_JSON", "CELLULES_SOURCE_JSON"]
H_CHAMPS = ["SOURCE", "FEUILLE", "COLONNE_SOURCE", "LIBELLE_SOURCE", "TRAITEMENT",
            "TYPE_ATTENDU", "DESTINATION_ET_CONTROLE", "VALEURS_NON_VIDES",
            "FORMULES_SANS_CACHE", "OBSERVATIONS_AVEC_ALERTE"]


def encoder(v):
    if isinstance(v, (dt.datetime, dt.date, dt.time)):
        return {"type": type(v).__name__, "valeur": v.isoformat()}
    return str(v)


def serie(v):
    return json.dumps(v, ensure_ascii=False, default=encoder, separators=(",", ":"))


def identifiant_ligne(source, feuille, ligne):
    return "L_" + hashlib.sha256(serie([source, feuille, ligne]).encode()).hexdigest()[:24]


def fige(v):
    """Normalise seulement les types pour une comparaison technique, pas metier."""
    if v is None or v == "":
        return None
    if isinstance(v, float):
        return round(v, 10)
    if isinstance(v, (dt.datetime, dt.date)):
        return v.isoformat()
    return v


def egal(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) <= 0.005000001  # sorties de synthese arrondies a deux decimales
    return fige(a) == fige(b)


class ControleCouverture:
    def __init__(self, regles):
        self.r = regles
        self.feuilles = []
        self.donnees = []
        self.champs = defaultdict(Counter)
        self.libelles = {}
        self.alertes = Counter()
        self.volumes = defaultdict(Counter)
        self.groupes = defaultdict(Counter)
        self.champs_degree = Counter()
        self.sources_absentes = []
        self.empreintes = {}
        self.bilan = {}

    def probleme(self, rec, code):
        if code not in rec["motifs"]:
            rec["motifs"].append(code)
            self.alertes[code] += 1

    def promotion_agregee(self, valeur):
        famille, code = self.r["promo"](valeur)
        return code if famille in ("ICNA_LETTRE", "ICNA_P", "TSEEAC") else "NON_RECONNUE"

    def contrat(self, src, niveau, label, colonne=None):
        n = self.r["norm"](label)
        if src == "REEF":
            return "VALEUR_REEF_CONSERVEE"
        if src == "PARCOURS":
            if n in {"nom", "prenom", "promotion", "promo", "degre", "degree"}:
                return "CONTEXTE_ET_NORMALISATION"
            if niveau == "DETAIL_EVALUATION":
                return "DETAIL_BRUT_DOCUMENTE_SANS_TENTATIVE_DEDUITE"
            if n in {"note", "noteglobale", "resultat", "resultatglobal",
                     "noteposition1", "noteposition2", "note1", "note2", "resultat1", "resultat2",
                     "notereconstituee", "originenotereconstituee", "statutnotereconstituee"}:
                return "SYNTHESE_OU_ROLE_A_CONTROLER"
            if n in {"position1", "position2", "role1", "role2"}:
                return "NORMALISATION_ROLE"
            if n in {"eleveid", "idaurion"}:
                return "IDENTIFIANT_CONSERVE_NON_UTILISE_COMME_CLE_UNIVERSELLE"
            if n in {"promotionfichier", "degrefichier", "test1", "test2", "rattrapage1", "rattrapage2",
                     "datefichier", "source", "feuillesource", "lignesource", "notereconstituee",
                     "originenotereconstituee", "position2interpretee", "statutnotereconstituee",
                     "traitementposition2"}:
                return "BRUT_DOCUMENTE_SANS_NOUVELLE_REGLE"
            return "CHAMP_NON_PREVU_CONSERVE_A_EXAMINER"
        if src in ("NOTES", "SESSIONS"):
            return "BLOC_SOURCE_CONSERVE_ET_CONTROLE"
        if src == "ASTRAL":
            return "BRUT_ASTRAL_CONSERVE_SANS_PREUVE_AUTOMATIQUE_DE_ZERO"
        if src == "KLIF":
            if colonne == "A" or n == "matricule":
                return "CHAMP_KLIF_IDENTIFIANT_IGNORE"
            if n == "origine":
                return "CHAMP_KLIF_ORIGINE_SOURCE_CONTROLE"
            if n == "nombredemoisdequalificationenacexclus":
                return "CHAMP_KLIF_DUREE_ENAC_EXCLUS_SOURCE_CONTROLE"
            if n in {"promo", "datedentreealenac", "nomdustagiaire", "prenomdustagiaire",
                     "affectation", "zone", "datedaffectation", "dateobtentiondelamuzz",
                     "arretformation", "nouveaucentre", "datedarrivee", "nombredechecs",
                     "datedefindeformation", "modederecrutement", "nombredemoisdequalificationenacinclus",
                     "nombredemoisdequalification", "nombredechecs", "prolongations",
                     "nombredeprolongations", "dureedeprolongation"}:
                return "CHAMP_KLIF_LU_OU_CONSERVE_BRUT"
            return "CHAMP_KLIF_NON_PREVU_CONSERVE_A_EXAMINER"
        return "BRUT_CONSERVE_A_EXAMINER"

    def description_champ(self, src, label, traitement):
        n = self.r["norm"](label)
        if "note" in n:
            typ = "NOMBRE_OU_VALEUR_NON_INTERPRETEE"
        elif "date" in n:
            typ = "DATE_SOURCE_OU_VALEUR_NON_INTERPRETEE"
        elif "resultat" in n:
            typ = "GO_NOGO_OU_TEXTE_SOURCE_SANS_DEDUCTION"
        elif src == "ASTRAL" and self.r["degree_entete_astral"](str(label)):
            typ = "REMEDIAL_EXPLICITE_OU_POURCENTAGE_OU_ABSENCE_DISTINCTS"
        elif traitement == "CHAMP_KLIF_DUREE_ENAC_EXCLUS_SOURCE_CONTROLE":
            typ = "NOMBRE_DE_MOIS_SOURCE_SANS_SUBSTITUTION"
        else:
            typ = "TYPE_SOURCE_CONSERVE"
        if traitement == "VALEUR_REEF_CONSERVEE":
            dest = "PARCOURS_CONSOLIDES : meme colonne ; comparaison independante + relecture."
        elif traitement in ("SYNTHESE_OU_ROLE_A_CONTROLER", "BLOC_SOURCE_CONSERVE_ET_CONTROLE"):
            dest = "Champs Degree PARCOURS/SOCRATES/SESSIONS + PROVENANCE ; brut dans COUVERTURE_DONNEES."
        elif traitement == "DETAIL_BRUT_DOCUMENTE_SANS_TENTATIVE_DEDUITE":
            dest = "PROVENANCE detail + COUVERTURE_DONNEES ; pas de resultat global derive."
        elif src == "ASTRAL":
            dest = "Brut dans COUVERTURE_DONNEES ; compteurs remedials controles, zeros derives non probants."
        elif traitement == "CHAMP_KLIF_ORIGINE_SOURCE_CONTROLE":
            dest = "KLIF_ORIGINE : valeur source non codee ; provenance et brute conservees."
        elif traitement == "CHAMP_KLIF_DUREE_ENAC_EXCLUS_SOURCE_CONTROLE":
            dest = "KLIF_MOIS_QUALIFICATION_ENAC_EXCLUS : valeur source distincte des durees recalculees."
        elif traitement == "CHAMP_KLIF_IDENTIFIANT_IGNORE":
            dest = "Identifiant administratif non exploite pour rapprocher les personnes ; brut conserve."
        elif src == "KLIF":
            dest = "Brut dans COUVERTURE_DONNEES ; champs KLIF retenus controles sur derniere periode."
        else:
            dest = "COUVERTURE_DONNEES : conservation brute verifiee ; aucune nouvelle interpretation."
        return typ, dest

    def inventorier(self, src, path):
        """Double lecture valeurs/formules, independante des listes des lecteurs."""
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
        path = Path(path)
        self.empreintes[src] = self.r["sha"](path)
        valeurs = load_workbook(path, read_only=True, data_only=True)
        formules = load_workbook(path, read_only=True, data_only=False)
        try:
            for ws in valeurs:
                fws = formules[ws.title]
                lignes = []
                caches = 0
                no = 0
                for no, pair in enumerate(zip_longest(ws.iter_rows(values_only=True),
                                                     fws.iter_rows(values_only=False), fillvalue=()), 1):
                    row, cells = pair
                    brut, formulas = {}, {}
                    for j, (v, cell) in enumerate(zip_longest(row, cells, fillvalue=None), 1):
                        coord = get_column_letter(j)
                        if v is not None:
                            brut[coord] = v
                        if cell is not None and cell.data_type == "f":
                            formulas[coord] = cell.value
                            if v is None:
                                caches += 1
                    if brut or formulas:
                        lignes.append({"ligne": no, "values": brut, "formules": formulas})
                info = {"source": src, "fichier": path.name, "feuille": ws.title,
                        "statut": "NON_RECONNUE_A_EXAMINER", "motif": "",
                        "lignes_examinees": max(ws.max_row or 0, no),
                        "non_vides": len(lignes), "candidates": 0, "lues": 0, "manquantes": 0,
                        "sans_cache": caches}
                self.classer(src, info, lignes)
                self.feuilles.append(info)
                if info["statut"] == "NON_RECONNUE_A_EXAMINER":
                    self.alertes["FEUILLE_NON_RECONNUE"] += 1
                if caches:
                    self.alertes["FORMULE_SANS_CACHE"] += caches
                for rec in lignes:
                    rec.update(source=src, fichier=path.name, feuille=ws.title, info=info,
                               motifs=[], destinations=[], did="", ordre="", statut="CONSERVEE_EN_DETAIL")
                    rec.setdefault("nature", "STRUCTURE_OU_TECHNIQUE")
                    rec.setdefault("niveau", "")
                    rec.setdefault("promo", "")
                    rec.setdefault("degree", "")
                    self.donnees.append(rec)
            if self.r["sha"](path) != self.empreintes[src]:
                raise ValueError("Une source a change pendant l'inventaire de couverture.")
        finally:
            valeurs.close()
            formules.close()

    def classer(self, src, info, rows):
        norm = self.r["norm"]
        if not rows:
            info.update(statut="VIDE_OU_ENTETES_SEULS", motif="Aucune cellule utile.")
            return
        if info["feuille"].upper() in self.r.get("FEUILLES_IGNOREES", {}).get(src, ()):
            info.update(statut="EXCLUE_PAR_REGLE", motif="Feuille auxiliaire ignoree par decision C1.")
            for rec in rows:
                rec.update(nature="EXCLUE_PAR_REGLE", exclusion="FEUILLE_AUXILIAIRE_C1")
            return
        headers, entete, niveau = {}, None, ""
        if src in ("REEF", "PARCOURS"):
            for rec in rows[:15]:
                h = {k: norm(v) for k, v in rec["values"].items() if isinstance(v, str)}
                req = {"nom", "prenom", "promotion"}
                if req <= set(h.values()):
                    headers, entete = rec["values"], rec["ligne"]
                    break
            if entete is not None:
                nh = {norm(v) for v in headers.values()}
                niveau = ("DETAIL_EVALUATION" if {"simulation", "role"} <= nh or
                          norm(info["feuille"]) == "detailseval" else
                          "SYNTHESE_PARCOURS")
                info.update(statut="LUE_DETAIL" if niveau == "DETAIL_EVALUATION" else "LUE_SYNTHESE",
                            motif="Structure inventoriee independamment du lecteur.")
            else:
                labels = {norm(v) for rec in rows[:8] for v in rec["values"].values()}
                techniques = {"fichiers", "alertes", "metadonneesagregation"}
                if src == "PARCOURS" and (norm(info["feuille"]) in techniques or
                                         {"etiquettesdelignes", "etiquettesdecolonnes"} <= labels):
                    info.update(statut="TECHNIQUE_NON_OBSERVATIONNELLE",
                                motif="Inventaire, metadonnees, alertes ou tableau croise ; donnees brutes conservees.")
        elif src in ("NOTES", "SESSIONS"):
            bloc = None
            hors_cible = False
            exclusions = 0
            bloc_refs, dg_bloc = [], ""
            for rec in rows:
                b = str(rec["values"].get("B", "")).strip()
                if b.lower().startswith("class:"):
                    hors_cible = self.r["famille_explicitement_exclue"](b[6:].strip())
                    if hors_cible:
                        bloc = None
                        exclusions += 1
                        rec.update(nature="EXCLUE_PAR_REGLE", exclusion="POPULATION_HORS_CIBLE_C2",
                                   promo=self.r["promo"](b[6:].strip())[1] or "")
                        continue
                    bloc = rec["ligne"]
                    rec.update(nature="CANDIDATE", porteur=bloc, niveau="BLOC_" + src,
                               promo=self.r["promo"](b[6:].strip())[1] or "")
                    bloc_refs, dg_bloc = [rec], ""
                    info["candidates"] += 1
                elif hors_cible:
                    rec.update(nature="EXCLUE_PAR_REGLE", exclusion="POPULATION_HORS_CIBLE_C2")
                elif bloc is not None:
                    rec["porteur"] = bloc
                    rec["degree"] = dg_bloc
                    bloc_refs.append(rec)
                    if b.lower().startswith("degree:"):
                        dg_bloc = self.r["degree_de"](b[7:].strip()) or ""
                        for prev in bloc_refs:
                            prev["degree"] = dg_bloc
                # Les lignes de bloc non prises en charge restent archivees.
            if info["candidates"]:
                info.update(statut="LUE_SYNTHESE", motif="Blocs Class: inventoriees independamment.")
                headers = {"A": "Identite", "B": "Contexte bloc", "C": "Statut ou information source",
                           "D": "Nombre ou date source", "E": "Role ou information source",
                           "F": "Role ou resultat source", "G": "Note globale source"}
            else:
                if exclusions:
                    info.update(statut="EXCLUE_PAR_REGLE", motif="Blocs hors cible explicitement exclus C2.")
                else:
                    info["motif"] = "Aucun bloc Class: ; contenu brut conserve pour examen."
        elif src == "ASTRAL":
            for rec in rows[:15]:
                h = {k: self.r["degree_entete_astral"](str(v)) for k, v in rec["values"].items()}
                if any(h.values()):
                    headers, entete = rec["values"], rec["ligne"]
                    info.update(statut="LUE_SYNTHESE", motif="En-tete Degree ASTRAL exact ; ABES_ ignore.")
                    break
            if entete is not None:
                courant = None
                suite = []
                explicite = False

                def terminer_bloc():
                    nonlocal courant, suite, explicite
                    if courant is not None:
                        if explicite:
                            courant.update(nature="CANDIDATE", niveau="SYNTHESE_ASTRAL",
                                           promo=self.r["promo"](info["feuille"])[1] or "")
                            info["candidates"] += 1
                            for x in suite:
                                x.update(nature="COMPLEMENT_ASTRAL", porteur=courant["ligne"])
                        else:
                            courant["nature"] = "STRUCTURE_OU_TECHNIQUE"
                            for x in suite:
                                x["nature"] = "STRUCTURE_OU_TECHNIQUE"
                    courant, suite, explicite = None, [], False

                for rec in rows:
                    if rec["ligne"] <= entete:
                        continue
                    label = str(rec["values"].get("B", "")).strip()
                    if label.upper() == "TOTAL":
                        terminer_bloc()
                        rec.update(nature="EXCLUE_PAR_REGLE", exclusion="TOTAL_ASTRAL_C5")
                        break
                    institution = bool(re.match(r"^(ENAC|ENAC\s*/.*|[A-Za-z]{2,12}\s*/\s*[A-Za-z0-9]{2,8})$",
                                                label, re.I))
                    # norm() colle les mots : le nom et prenom sont delimites dans le libelle brut.
                    identite_probable = len(re.findall(r"[A-Za-zÀ-ÿ0-9]+", label)) >= 2
                    if label and not institution and identite_probable:
                        terminer_bloc()
                        courant = rec
                    elif label and not institution:
                        terminer_bloc()
                        rec["nature"] = "DONNEE_ASTRAL_NON_RECONNUE"
                    elif courant is not None:
                        suite.append(rec)
                    rec["porteur"] = courant["ligne"] if courant is not None else None
                    if courant is not None:
                        for col, val in rec["values"].items():
                            if (self.r["degree_entete_astral"](str(headers.get(col, "")))
                                    and re.fullmatch(r"\s*\d+\s*remedial\(s\)\s*", str(val), re.I)):
                                explicite = True
                terminer_bloc()
        elif src == "KLIF":
            fam, pr = self.r["promo"](info["feuille"])
            if fam == "ICNA_P":
                info.update(statut="EXCLUE_PAR_REGLE", motif="KLIF voie P exclue C2.")
                for rec in rows:
                    rec.update(nature="EXCLUE_PAR_REGLE", exclusion="VOIE_P_C2", promo=pr or "")
            elif fam == "ICNA_LETTRE" and self.r["promotion_avant_12a"](info["feuille"]):
                info.update(statut="EXCLUE_PAR_REGLE", motif="KLIF : promotion anterieure a 12A.")
                for rec in rows:
                    rec.update(nature="EXCLUE_PAR_REGLE", exclusion="KLIF_AVANT_12A_C2", promo=pr or "")
            elif fam in ("TSEEAC", "DEF", "EXPORT"):
                info.update(statut="EXCLUE_PAR_REGLE",
                            motif="Cursus non ICNA non importe par le lecteur KLIF existant ; brut conserve.")
                for rec in rows:
                    rec.update(nature="EXCLUE_PAR_REGLE", exclusion="FAMILLE_HORS_CIBLE_C2", promo=pr or "")
            else:
                for rec in rows[:6]:
                    if "promo" in {norm(v) for v in rec["values"].values()}:
                        headers, entete = rec["values"], rec["ligne"]
                        info.update(statut="LUE_SYNTHESE", motif="En-tete stagiaires KLIF.")
                        break
                if entete is None:
                    info["motif"] = "Feuille KLIF sans structure interpretable ; conservation brute."
        if headers:
            for k, label in headers.items():
                traitement = self.contrat(src, niveau, label, k)
                ch = (src, info["feuille"], k, traitement)
                self.libelles[ch] = str(label)
                self.champs[ch]["non_vides"] += 0
            for rec in rows:
                if src in ("REEF", "PARCOURS", "KLIF") and entete is not None and rec["ligne"] > entete:
                    rec.update(nature="CANDIDATE", niveau=niveau or "SOURCE_" + src)
                    info["candidates"] += 1
                    if src == "KLIF":
                        identites = [v for k, v in rec["values"].items()
                                     if norm(headers.get(k, k)) in ("nomdustagiaire", "prenomdustagiaire") and v is not None]
                        premier = norm(rec["values"].get("A", ""))
                        moyenne = any(self.r["est_moyenne_klif"](v) for v in rec["values"].values())
                        if not identites and moyenne:
                            rec.update(nature="EXCLUE_PAR_REGLE", exclusion="MOYENNE_KLIF_C3")
                            info["candidates"] -= 1
                        elif not identites and premier in {"total", "totalgeneral", "mediane", "ecarttype"}:
                            rec["nature"] = "STRUCTURE_OU_TECHNIQUE"
                            info["candidates"] -= 1
                        elif not identites:
                            rec["nature"] = "CONTENU_NON_RECONNU"
                            info["candidates"] -= 1
                        elif moyenne:
                            rec["conflit_moyenne_identite"] = True
                if entete is not None and rec["ligne"] <= entete:
                    continue
                rec["headers"] = headers
                normal = {norm(headers.get(k, k)): v for k, v in rec["values"].items()}
                if src in ("REEF", "PARCOURS"):
                    rec["promo"] = self.r["promo"](normal.get("promotion", normal.get("promo")))[1] or ""
                    rec["degree"] = self.r["degree_de"](str(normal.get("degre", normal.get("degree", "")))) or ""
                elif src == "KLIF":
                    rec["promo"] = self.r["promo"](info["feuille"])[1] or ""
            if info["candidates"] == 0 and entete is not None:
                info.update(statut="VIDE_OU_ENTETES_SEULS", motif="En-tetes sans observation.")
        if info["statut"] == "NON_RECONNUE_A_EXAMINER":
            for rec in rows:
                rec["nature"] = "CONTENU_NON_RECONNU"

    def calculer(self, observations, rattachements, provenance, sorties, entetes_reef):
        norm = self.r["norm"]
        par_ref = defaultdict(list)
        for o in observations:
            par_ref[(o["source"], o["feuille"], o["ligne"])].append(o)
        traces = defaultdict(list)
        for p in provenance:
            traces[(p[4], p[5])].append(p)
        self._traces_klif = {p[5] for p in provenance if p[4] == "KLIF" and p[2] == "KLIF_*"}
        sorties_idx = {(str(r["ID_DOSSIER_LOCAL"]), r["ORDRE_PERIODE"]): r for r in sorties}
        records = {(r["source"], r["feuille"], r["ligne"]): r for r in self.donnees}
        for rec in self.donnees:
            key = (rec["source"], rec["feuille"], rec["ligne"])
            obs = par_ref.get(key, [])
            porteur = rec.get("porteur")
            support = records.get((rec["source"], rec["feuille"], porteur)) if porteur else None
            if not obs and support and support is not rec:
                obs = par_ref.get((rec["source"], rec["feuille"], porteur), [])
            if rec["source"] == "REEF" and rec["nature"] == "CANDIDATE":
                rec["statut"] = "RESTITUEE_EN_SYNTHESE"
                raw = {norm(rec.get("headers", {}).get(k, k)): v for k, v in rec["values"].items()}
                did, ordre = str(raw.get("iddossierlocal", "")), raw.get("ordreperiode")
                rec["did"], rec["ordre"] = did, ordre
                cible = sorties_idx.get((did, ordre), {})
                for k, valeur in rec["values"].items():
                    col = rec.get("headers", {}).get(k, k)
                    if col not in cible or fige(cible[col]) != fige(valeur):
                        rec.setdefault("champs_en_cause", []).append(str(col))
                        self.probleme(rec, "CHAMP_UTILE_NON_RESTITUE")
            if rec["nature"] == "CANDIDATE" and rec["source"] != "REEF":
                rec["info"]["lues"] += len(par_ref.get(key, []))
                if not par_ref.get(key):
                    rec["info"]["manquantes"] += 1
                    self.probleme(rec, "LIGNE_CANDIDATE_ABSENTE_DU_LECTEUR")
                    rec["statut"] = "EN_ATTENTE_CONSERVEE"
            if rec["nature"] == "CANDIDATE" and self.r["promo"](rec["promo"])[0] in ("AUTRE", "ABSENTE"):
                self.probleme(rec, "PROMOTION_NON_RECONNUE")
            if rec["nature"] in ("CONTENU_NON_RECONNU", "DONNEE_ASTRAL_NON_RECONNUE"):
                self.probleme(rec, "CONTENU_NON_RECONNU_CONSERVE")
                rec["statut"] = "EN_ATTENTE_CONSERVEE"
            if rec.get("conflit_moyenne_identite"):
                self.probleme(rec, "KLIF_MOYENNE_AVEC_IDENTITE")
                rec["statut"] = "EN_ATTENTE_CONSERVEE"
            if rec["formules"]:
                if any(k not in rec["values"] for k in rec["formules"]):
                    self.probleme(rec, "FORMULE_SANS_CACHE")
            for o in obs:
                lien = rattachements.get(id(o), {})
                rec["did"], rec["ordre"] = lien.get("dossier", ""), lien.get("ordre", "")
                ref = "%s!%s:%s" % (o["fichier"], o["feuille"], o["ligne"])
                pp = traces.get((o["source"], ref), [])
                rec["destinations"] += sorted({str(p[2]) for p in pp})
                if o["source"] == "KLIF" and self.r["promotion_avant_12a"](o["promo_brute"]):
                    rec["statut"] = "EN_ATTENTE_CONSERVEE"
                    self.probleme(rec, "EXCLUSION_AVANT_12A_KLIF")
                elif o["source"] in ("NOTES", "SESSIONS", "KLIF") and self.r["famille_explicitement_exclue"](o["promo_brute"]):
                    rec["statut"] = "EN_ATTENTE_CONSERVEE"
                    self.probleme(rec, "EXCLUSION_METIER_HORS_CIBLE_C2")
                elif not lien.get("dossier") or lien.get("ordre") in (None, ""):
                    rec["statut"] = "EN_ATTENTE_CONSERVEE"
                    self.probleme(rec, "RATTACHEMENT_EN_ATTENTE")
                elif o["source"] == "PARCOURS" and o["data"].get("niveau") == "DETAIL_EVALUATION":
                    # C8 : le niveau du lecteur prime sur une trace d'origine de ligne.
                    if any(p[2] == "OBSERVATION_DETAILLEE" for p in pp):
                        rec["statut"] = "CONSERVEE_EN_DETAIL"
                    else:
                        rec["statut"] = "EN_ATTENTE_CONSERVEE"
                        self.probleme(rec, "DETAIL_SANS_PROVENANCE")
                elif any(p[2] != "OBSERVATION_DETAILLEE" for p in pp):
                    rec["statut"] = "RESTITUEE_EN_SYNTHESE"
                else:
                    rec["statut"] = "EN_ATTENTE_CONSERVEE"
                    self.probleme(rec, "OBSERVATION_LUE_SANS_RESTITUTION_SYNTHETIQUE")
                if o["source"] not in ("ASTRAL", "KLIF") and not o["degree"]:
                    self.probleme(rec, "DEGREE_NON_RECONNU")
                if rec["nature"] == "CANDIDATE":
                    self.controler_champs(rec, o, lien, pp, sorties_idx)
            for k, value in rec["values"].items():
                if rec.get("headers") and rec["nature"] != "STRUCTURE_OU_TECHNIQUE":
                    label = rec["headers"].get(k, k)
                    traitement = self.contrat(rec["source"], rec["niveau"], label, k)
                    ch = (rec["source"], rec["feuille"], k, traitement)
                    self.libelles[ch] = str(label)
                    self.champs[ch]["non_vides"] += 1
                    dg = ((self.r["degree_entete_astral"](str(label)) if rec["source"] == "ASTRAL"
                           else self.r["degree_de"](str(label)))
                          if rec["source"] in ("ASTRAL", "REEF") else rec["degree"])
                    self.champs_degree[(rec["source"], self.promotion_agregee(rec["promo"]),
                                        dg or "NON_PRECISE")] += 1
                    self.champs[ch]["sans_cache"] += int(k in rec["formules"] and k not in rec["values"])
                    self.champs[ch]["alertes"] += bool(rec["motifs"])
                    if "NON_PREVU" in traitement:
                        if not rec["motifs"]:
                            self.champs[ch]["alertes"] += 1
                        self.probleme(rec, "CHAMP_NON_PREVU_CONSERVE")
            # Les formules sans cache ne figurent pas dans values.
            for k in set(rec["formules"]) - set(rec["values"]):
                label = rec.get("headers", {}).get(k, k)
                ch = (rec["source"], rec["feuille"], k, "FORMULE_SANS_CACHE")
                self.libelles[ch] = str(label)
                self.champs[ch]["sans_cache"] += 1
            self.volumes[rec["source"]]["lignes_non_vides_archivees"] += 1
            if rec["nature"] == "CANDIDATE":
                self.volumes[rec["source"]]["observations_candidates"] += 1
                self.volumes[rec["source"]][rec["statut"]] += 1
                g = (rec["source"], self.promotion_agregee(rec["promo"]), rec["degree"] or "NON_PRECISE",
                     rec["niveau"] or "SOURCE")
                self.groupes[g][rec["statut"]] += 1
        for key in par_ref:
            if key not in records:
                self.alertes["OBSERVATION_LECTEUR_SANS_LIGNE_SOURCE"] += len(par_ref[key])
        self.controler_sessions(records, par_ref, rattachements, sorties_idx)
        self.controler_astral_klif(records, par_ref, rattachements, sorties_idx)
        for v in self.champs.values():
            v["alertes"] = 0
        for rec in self.donnees:
            if rec["motifs"] and rec.get("headers") and rec["nature"] != "STRUCTURE_OU_TECHNIQUE":
                for col in rec["values"]:
                    label = rec["headers"].get(col, col)
                    ch = (rec["source"], rec["feuille"], col,
                          self.contrat(rec["source"], rec["niveau"], label, col))
                    self.champs[ch]["alertes"] += 1
        bloquantes = sum(v for k, v in self.alertes.items() if k in (
            "LIGNE_CANDIDATE_ABSENTE_DU_LECTEUR", "FEUILLE_NON_RECONNUE", "FORMULE_SANS_CACHE",
            "OBSERVATION_LECTEUR_SANS_LIGNE_SOURCE", "CHAMP_UTILE_NON_RESTITUE", "DETAIL_SANS_PROVENANCE",
            "VALEUR_NON_INTERPRETEE", "DEGREE_NON_RECONNU", "ROLE_NON_RECONNU",
            "CHAMP_NON_PREVU_CONSERVE", "CONTENU_NON_RECONNU_CONSERVE", "PROMOTION_NON_RECONNUE",
            "KLIF_MOYENNE_AVEC_IDENTITE"))
        candidates = sum(v["observations_candidates"] for v in self.volumes.values())
        classes = {st: sum(v[st] for v in self.volumes.values()) for st in ETATS}
        if candidates != sum(classes.values()):
            raise ValueError("Reconciliation de couverture desequilibree.")
        conclusion = ("COUVERTURE_NON_DEMONTREE" if self.alertes["FEUILLE_NON_RECONNUE"] else
                      "COUVERTURE_PARTIELLE_EXPLIQUEE" if bloquantes else
                      "COUVERTURE_COMPLETE_DU_PERIMETRE_DECLARE")
        self.bilan = {
            "revision": REVISION, "perimetre": "FICHIERS_SELECTIONNES_DE_CETTE_EXECUTION",
            "sources_selectionnees": sorted(self.empreintes),
            "sources_non_selectionnees": self.sources_absentes,
            "collecte": conclusion,
            "restitution": conclusion,
            "exploitation": "SOUS_RESTRICTIONS" if self.alertes else "RATTACHEMENTS_CONTROLES",
            "inventaire_independant_du_lecteur": True,
            "observations_candidates": candidates, "devenir_observations_candidates": classes,
            "reconciliation_equilibree": True, "alertes_par_cause": dict(sorted(self.alertes.items())),
            "par_source": {s: dict(c) for s, c in sorted(self.volumes.items())},
            "exclusions_lignes_non_vides_par_source": dict(Counter(
                r["source"] for r in self.donnees if r["nature"] == "EXCLUE_PAR_REGLE")),
            "exclusions_par_regle": dict(sorted(Counter(
                r.get("exclusion") for r in self.donnees if r.get("exclusion")).items())),
            "moyennes_klif_ignorees": sum(r.get("exclusion") == "MOYENNE_KLIF_C3" for r in self.donnees),
            "couverture_par_source_promotion_degree_niveau": [
                {"source": s, "promotion": pr, "degree": dg, "niveau": ni, "statuts": dict(c)}
                for (s, pr, dg, ni), c in sorted(self.groupes.items())],
            "champs_non_vides_par_source_promotion_degree": [
                {"source": s, "promotion": pr, "degree": dg, "cellules_sources_non_vides": n}
                for (s, pr, dg), n in sorted(self.champs_degree.items())],
            "limites": ["Ne certifie ni l'exhaustivite historique des sources ni l'identite des personnes.",
                       "Les champs documentaires et identifiants non exploites restent conserves en brut.",
                       "Les sources non selectionnees ne sont pas declarees couvertes.",
                       "Les formules sources ne sont pas recalculees.",
                       "Seules les mentions remedial ASTRAL explicites alimentent le compteur ; absence n'est pas NON."]}
        if self.sources_absentes:
            self.bilan["limites"].append("Conclusion limitee aux sources selectionnees, pas aux cinq sources potentielles.")
        return self

    def controler_champs(self, rec, o, lien, traces, sorties_idx):
        """Contrats independants des champs de sortie, a partir des cellules brutes."""
        norm = self.r["norm"]
        raw = {norm(k): v for k, v in o["data"].get("observation_brute", {}).items()}
        if rec["source"] == "PARCOURS":
            raw = {norm(rec.get("headers", {}).get(k, k)): v for k, v in rec["values"].items()}
        if rec["source"] not in ("PARCOURS", "NOTES", "SESSIONS"):
            return
        detail = rec["niveau"] == "DETAIL_EVALUATION"
        if detail:
            return
        dg = rec["degree"] or o["degree"]
        if not dg:
            return
        attendus = []
        if rec["source"] == "PARCOURS":
            for champ in ("note", "noteglobale"):
                if champ in raw:
                    x = self.r["note"](raw[champ])
                    if norm(raw[champ]) == "nogo":
                        raw_reconst = raw.get("notereconstituee")
                        nr = self.r["note"](raw_reconst)
                        if nr is not None:
                            attendus.append((dg + "_PARCOURS_NOTE_RECONSTITUEE", nr))
                            if raw.get("originenotereconstituee"):
                                attendus.append((dg + "_PARCOURS_ORIGINE_NOTE_RECONSTITUEE",
                                                raw["originenotereconstituee"]))
                            if raw.get("statutnotereconstituee"):
                                attendus.append((dg + "_PARCOURS_STATUT_NOTE_RECONSTITUEE",
                                                raw["statutnotereconstituee"]))
                        elif raw_reconst not in (None, ""):
                            self.probleme(rec, "VALEUR_NON_INTERPRETEE")
                    elif x is None and raw[champ] not in (None, ""):
                        self.probleme(rec, "VALEUR_NON_INTERPRETEE")
                    elif x is not None:
                        attendus.append((dg + "_PARCOURS_NOTE", x))
                    break
            res = raw.get("resultat", raw.get("resultatglobal"))
            if res is not None and res != "":
                attendus.append((dg + "_PARCOURS_RESULTAT", self.r["resultat"](res) or str(res).strip()))
            for i in ("1", "2"):
                trait = norm(raw.get("traitementposition2", "")) + " " + norm(raw.get("position2interpretee", ""))
                if i == "2" and any(t in trait for t in ("neutralis", "ignore", "nonretenu")):
                    continue  # neutralisation explicite source, non suppression de l'observation
                pos = raw.get("position" + i, raw.get("role" + i))
                n = raw.get("noteposition" + i, raw.get("note" + i))
                res = raw.get("resultat" + i)
                role = self.r["role_de"](dg, pos)
                if role is None and dg in ("BTWR", "ADI") and i == "1" and n is not None:
                    role = "TWR" if dg == "BTWR" else "LOC"
                if pos is not None and role is None and (n is not None or res is not None):
                    self.probleme(rec, "ROLE_NON_RECONNU")
                if n is not None and self.r["note"](n) is None:
                    self.probleme(rec, "VALEUR_NON_INTERPRETEE")
                if role and self.r["note"](n) is not None:
                    if dg not in ("BTWR", "ADI"):
                        attendus.append(("%s_%s_PARCOURS_NOTE" % (dg, role), self.r["note"](n)))
                    elif raw.get("note", raw.get("noteglobale")) is None:
                        attendus.append((dg + "_PARCOURS_NOTE", self.r["note"](n)))
                if role == "PL" and res is not None and res != "":
                    attendus.append((dg + "_PL_PARCOURS_RESULTAT",
                                     self.r["resultat"](res) or str(res).strip()))
        elif rec["source"] == "NOTES":
            v = rec["values"]
            if "G" in v:
                n = self.r["note"](v["G"])
                if n is None:
                    self.probleme(rec, "VALEUR_NON_INTERPRETEE")
                else:
                    attendus.append((dg + "_SOCRATES_NOTE", n))
            for k in ("E", "F"):
                if k not in v:
                    continue
                m = re.match(r"^\s*([A-Za-z]+)\s*:\s*(-?[\d.,]+)\s*$", str(v[k]))
                if not m and str(v[k]).strip():
                    self.probleme(rec, "VALEUR_NON_INTERPRETEE")
                if m:
                    role = self.r["role_de"](dg, m[1])
                    n = self.r["note"](m[2])
                    if not role:
                        self.probleme(rec, "ROLE_NON_RECONNU")
                    if n is None:
                        self.probleme(rec, "VALEUR_NON_INTERPRETEE")
                    if role and n is not None:
                        col = dg + "_SOCRATES_NOTE" if dg in ("BTWR", "ADI") else "%s_%s_SOCRATES_NOTE" % (dg, role)
                        if dg not in ("BTWR", "ADI") or "G" not in v:
                            attendus.append((col, n))
        if not lien.get("dossier") or lien.get("ordre") in (None, ""):
            return
        sortie = sorties_idx.get((str(lien["dossier"]), lien["ordre"]), {})
        for col, valeur in attendus:
            matches = [t for t in traces if t[2] == col and egal(t[3], valeur)]
            # Les candidats d'un conflit sont une restitution valide de l'incertitude.
            if not matches or col not in sortie or (
                    not egal(sortie[col], valeur) and sortie[col] != "PLUSIEURS_VALEURS"):
                rec.setdefault("champs_en_cause", []).append(col)
                self.probleme(rec, "CHAMP_UTILE_NON_RESTITUE")

    def controler_sessions(self, records, par_ref, liens, sorties):
        groupes = defaultdict(list)
        for key, oo in par_ref.items():
            if key[0] != "SESSIONS" or key not in records:
                continue
            rec = records[key]
            for o in oo:
                lien = liens.get(id(o), {})
                if lien.get("dossier") and lien.get("ordre") not in (None, "") and o["degree"]:
                    groupes[(str(lien["dossier"]), lien["ordre"], o["degree"])].append((rec, o))
        for (did, ordre, dg), items in groupes.items():
            dates, nogo = [], 0
            for rec, o in items:
                rawdate = rec["values"].get("D")
                date = self.r["date"](rawdate)
                if date is None and isinstance(rawdate, (int, float)):
                    from openpyxl.utils.datetime import from_excel
                    try:
                        date = from_excel(rawdate)
                    except (ValueError, OverflowError, TypeError):
                        pass
                if rawdate is not None and not isinstance(date, dt.datetime):
                    self.probleme(rec, "VALEUR_NON_INTERPRETEE")
                elif date is not None:
                    dates.append(date)
                rawres = rec["values"].get("F")
                if rawres is not None and self.r["resultat"](rawres) is None:
                    self.probleme(rec, "VALEUR_NON_INTERPRETEE")
                nogo += self.r["resultat"](rawres) == "NOGO"
            expected = {dg + "_SESSIONS_NB": len(items), dg + "_SESSIONS_NOGO": nogo,
                        dg + "_SESSIONS_PREMIERE": min(dates) if dates else None,
                        dg + "_SESSIONS_DERNIERE": max(dates) if dates else None}
            mauvais = [c for c, val in expected.items()
                       if not egal(sorties.get((did, ordre), {}).get(c), val)]
            if mauvais:
                for rec, _ in items:
                    rec.setdefault("champs_en_cause", []).extend(mauvais)
                    self.probleme(rec, "CHAMP_UTILE_NON_RESTITUE")

    def controler_astral_klif(self, records, par_ref, liens, sorties):
        """Controle les agregats ASTRAL et les valeurs KLIF depuis les cellules archivees."""
        norm = self.r["norm"]
        complements = defaultdict(list)
        for r in self.donnees:
            if r["source"] == "ASTRAL" and r.get("porteur") != r["ligne"]:
                complements[(r["feuille"], r.get("porteur"))].append(r)
        for key, oo in par_ref.items():
            if key[0] not in ("ASTRAL", "KLIF") or key not in records:
                continue
            rec = records[key]
            for o in oo:
                lien = liens.get(id(o), {})
                if not lien.get("dossier") or lien.get("ordre") in (None, ""):
                    continue
                did, ordre = str(lien["dossier"]), lien["ordre"]
                cible = sorties.get((did, ordre), {})
                expected = {}
                if key[0] == "ASTRAL":
                    counts = Counter()
                    parties = [rec] + complements[(rec["feuille"], rec["ligne"])]
                    for r in parties:
                        for col, v in r["values"].items():
                            if col in ("A", "B"):
                                continue
                            titre = str(rec.get("headers", {}).get(col, ""))
                            dg = self.r["degree_entete_astral"](titre)
                            match = re.fullmatch(r"\s*(\d+)\s*remedial\(s\)\s*", str(v), re.I)
                            if dg and match:
                                counts[dg] += int(match[1])
                            elif dg and not re.fullmatch(r"\s*\d{1,3}\s*%\s*done\s*", str(v), re.I):
                                self.probleme(r, "VALEUR_NON_INTERPRETEE")
                            elif not dg and titre and not norm(titre).startswith("abes"):
                                self.probleme(r, "DEGREE_NON_RECONNU")
                    expected = {dg + "_ASTRAL_REMEDIALS": n for dg, n in counts.items()}
                    expected.update({dg + "_ASTRAL_PRESENCE_REMEDIAL": "OUI"
                                     for dg, n in counts.items() if n > 0})
                else:
                    if cible.get("KLIF_PORTEE") != "SUIVI_KLIF_DU_DOSSIER":
                        continue  # les divergences sont conservees brutes, sans nouveau choix
                    # Une autre ligne KLIF a pu etre retenue par la regle historique.
                    repere = "%s!%s:%s" % (o["fichier"], o["feuille"], o["ligne"])
                    if repere not in self._traces_klif:
                        continue
                    mapping = {"datedentreealenac": "KLIF_DATE_ENTREE_ENAC", "affectation": "KLIF_AFFECTATION_1",
                               "zone": "KLIF_ZONE", "datedaffectation": "KLIF_DATE_AFFECTATION_1",
                               "origine": "KLIF_ORIGINE",
                               "nouveaucentre": "KLIF_NOUVEAU_CENTRE", "datedarrivee": "KLIF_DATE_ARRIVEE_2",
                               "nombredechecs": "KLIF_NB_ECHECS", "datedefindeformation": "KLIF_DATE_FIN_FORMATION"}
                    episode = 1
                    for col, label in rec.get("headers", {}).items():
                        n = norm(label)
                        if n == "nouveaucentre":
                            episode = 2
                        ciblecol = mapping.get(n)
                        if n == "dateobtentiondelamuzz":
                            ciblecol = "KLIF_DATE_MU_" + str(episode)
                        if n == "arretformation":
                            ciblecol = "KLIF_ARRET_" + str(episode)
                        if n == "nombredemoisdequalificationenacexclus" and col in rec["values"]:
                            expected["KLIF_DUREE_ENAC_EXCLUS_BRUTE"] = rec["values"][col]
                            nb = self.r["nombre"](rec["values"][col])
                            if nb is not None and nb >= 0:
                                expected["KLIF_MOIS_QUALIFICATION_ENAC_EXCLUS"] = nb
                            elif rec["values"][col] not in (None, ""):
                                self.probleme(rec, "VALEUR_NON_INTERPRETEE")
                        if ciblecol and col in rec["values"]:
                            expected[ciblecol] = rec["values"][col]
                    dernier = max(p for d, p in sorties if d == did)
                    if ordre != dernier or any(r.get("KLIF_PORTEE") == "SUIVI_KLIF_DU_DOSSIER" and p != ordre
                                              for (d, p), r in sorties.items() if d == did):
                        self.probleme(rec, "CHAMP_UTILE_NON_RESTITUE")
                mauvais = [col for col, v in expected.items()
                           if col not in cible or (not egal(cible[col], v) and cible[col] != "PLUSIEURS_VALEURS")]
                if mauvais:
                    rec.setdefault("champs_en_cause", []).extend(mauvais)
                    self.probleme(rec, "CHAMP_UTILE_NON_RESTITUE")

    def tables(self):
        if hasattr(self, "_tables"):
            return self._tables
        feuilles = [[s["source"], s["fichier"], s["feuille"], s["statut"], s["motif"],
                     s["lignes_examinees"], s["non_vides"], s["candidates"], s["lues"],
                     s["manquantes"], s["sans_cache"]] for s in self.feuilles]
        donnees = []
        for r in self.donnees:
            ident = identifiant_ligne(r["source"], r["feuille"], r["ligne"])
            payload = serie({"valeurs": r["values"], "formules": r["formules"]})
            parts = [payload[i:i + LIMITE_JSON] for i in range(0, len(payload), LIMITE_JSON)] or [""]
            for i, chunk in enumerate(parts, 1):
                donnees.append([ident, r["source"], r["fichier"], r["feuille"], r["ligne"], r["nature"],
                                r["statut"], "; ".join(r["motifs"]), r["promo"], r["degree"],
                                r["did"], r["ordre"], serie(sorted(set(r["destinations"]))),
                                serie(sorted(set(r.get("champs_en_cause", [])))),
                                i, len(parts), chunk])
        champs = [[s, f, col, self.libelles[key], traitement,
                   *self.description_champ(s, self.libelles[key], traitement),
                   v["non_vides"], v["sans_cache"], v["alertes"]]
                  for key, v in sorted(self.champs.items())
                  for s, f, col, traitement in [key]]
        self._tables = {"COUVERTURE_FEUILLES": (H_FEUILLES, feuilles),
                        "COUVERTURE_DONNEES": (H_DONNEES, donnees),
                        "COUVERTURE_CHAMPS": (H_CHAMPS, champs)}
        return self._tables

    def ajouter_feuilles(self, wb):
        from openpyxl.styles import Font
        for name, (head, rows) in self.tables().items():
            ws = wb.create_sheet(name)
            ws.append(head)
            for c in ws[1]:
                c.font = Font(bold=True)
            for row in rows:
                ws.append(row)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            from openpyxl.utils import get_column_letter
            for j, label in enumerate(head, 1):
                ws.column_dimensions[get_column_letter(j)].width = min(46, max(16, len(label) + 2))
        for src, empreinte in sorted(self.empreintes.items()):
            wb["LISEZMOI"].append(["EMPREINTE SOURCE " + src, empreinte])
        li = wb["LISEZMOI"]
        li.append(["CONTROLE DE COUVERTURE : " + self.bilan["collecte"]])
        li.append(["Les 3 feuilles COUVERTURE_* sont locales et nominatives. Donnees brutes conservees, pas seulement les observations acceptees par le lecteur."])
        li.append(["COUVERTURE_DONNEES : une ligne par ligne source non vide ; JSON long reparti en parties ordonnees a reunir par ID_LIGNE_IMPORT."])
        li.append(["Les identifiants L_ reperent une ligne dans cette execution ; ils ne sont ni une identite d'eleve ni un identifiant stable apres deplacement des lignes sources."])
        li.append(["Le contrat BRUT_DOCUMENTE ne signifie pas qu'un champ alimente une colonne metier ; voir COUVERTURE_CHAMPS."])
        li.append(["Les motifs et CHAMPS_EN_CAUSE de COUVERTURE_DONNEES donnent les anomalies exactes ; OBSERVATIONS_AVEC_ALERTE porte sur la ligne source, pas uniquement ce champ."])
        li.append(["Alias conserves : BASIC_TOWER=BTWR, BASIC_ACC=BSURV ; Cordon/Planner=PL, Approche=AP ; CO/CR selon le Degree. EleveId/id_Aurion conserves, sans nouvelle cle universelle."])
        li.append(["C2 : l'exclusion par date avant 12A s'applique a KLIF seulement ; les anciennes exclusions du registre restent archivees."])
        li.append(["C2 : promotions P exclues de KLIF, Notes et Sessions sans ajout de population."])
        li.append(["C3 : les lignes Moyenne KLIF sont ignorees par contenu si l'identite est vide ; conflit avec identite conserve a examiner."])
        li.append(["C8 : les details restent classes en detail, meme avec une trace supplementaire d'origine de population."])
        li.append(["Le bilan agrege distingue collecte, restitution et exploitation. Une ambiguite conservee n'est pas une perte."])

    def verifier(self, fichier):
        from openpyxl import load_workbook
        wb = load_workbook(fichier, read_only=True, data_only=False)
        total = 0
        try:
            for name, (head, rows) in self.tables().items():
                it = wb[name].iter_rows(values_only=True)
                if list(next(it, ())) != head:
                    raise ValueError("Entetes de couverture incorrectes : " + name)
                for no, (got, exp) in enumerate(zip_longest(it, rows), 2):
                    if got is None or exp is None or len(got) != len(exp) or any(
                            fige(a) != fige(b) for a, b in zip(got, exp)):
                        raise ValueError("Conservation de couverture incorrecte : %s ligne %d." % (name, no))
                    total += 1
        finally:
            wb.close()
        self.bilan["conservation_brute_verifiee_apres_sauvegarde"] = True
        self.bilan["lignes_de_controle_verifiees"] = total
        return self.bilan


def construire(reef, sources, observations, rattachements, provenance, sorties, entetes, regles):
    controle = ControleCouverture(regles)
    controle.inventorier("REEF", reef)
    for src in ("PARCOURS", "NOTES", "SESSIONS", "ASTRAL", "KLIF"):
        if sources.get(src):
            controle.inventorier(src, sources[src])
        else:
            controle.sources_absentes.append(src)
    return controle.calculer(observations, rattachements, provenance, sorties, entetes)

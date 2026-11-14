# -*- coding: utf-8 -*-
"""Reception de couverture sur sources fictives uniquement."""
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook, load_workbook
import consolidation as C
import couverture as V
import test_consolidation as H


class Couverture(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.reef = self.root / "reef.xlsx"
        H.reef(self.reef)
        self.source = self.root / "parcours.xlsx"
        self.head = ["Nom", "Prénom", "EleveId", "Promotion", "Degré", "Position 1",
                     "Note Position 1", "Resultat 1", "Note", "Résultat"]
        self.row = ["MARTIN", "Paul", "E_FICTIF", "20A", "BTWR", "TWR", 12, "GO", 12, "GO"]
        self.regles = {k: getattr(C, k) for k in ("norm", "promo", "degree_de", "role_de",
                       "promotion_avant_12a", "sha", "note", "resultat", "date")}
        self.controle = None

    def tearDown(self):
        self.tmp.cleanup()

    def ecrire(self, rows=None, head=None):
        w = Workbook()
        ws = w.active
        ws.title = "Parcours"
        ws.append(head or self.head)
        for r in [self.row] if rows is None else rows:
            ws.append(r)
        w.save(self.source)
        w.close()

    def executer(self, sources=None):
        if sources is None:
            sources = {"PARCOURS": self.source}
        original = V.construire
        def retenir(*args, **kwargs):
            self.controle = original(*args, **kwargs)
            return self.controle
        empreintes = {str(p): C.sha(p) for p in [self.reef] + list(sources.values())}
        with patch.object(V, "construire", side_effect=retenir):
            f, b = C.consolider(self.reef, sources, self.root / "resultats", journal=lambda *_: None)
        self.assertEqual(empreintes, {p: C.sha(p) for p in empreintes})
        return f, b["couverture_donnees"]

    def archives(self, fichier):
        w = load_workbook(fichier, read_only=True, data_only=False)
        try:
            it = w["COUVERTURE_DONNEES"].iter_rows(values_only=True)
            h = next(it)
            rows = [dict(zip(h, r)) for r in it]
        finally:
            w.close()
        return rows

    def test_01_reference_complete_et_sources_absentes_explicites(self):
        self.ecrire()
        f, b = self.executer()
        self.assertEqual(b["collecte"], "COUVERTURE_COMPLETE_DU_PERIMETRE_DECLARE")
        self.assertEqual(b["exploitation"], "RATTACHEMENTS_CONTROLES")
        self.assertEqual(b["sources_non_selectionnees"], ["NOTES", "SESSIONS", "ASTRAL", "KLIF"])
        self.assertTrue(b["conservation_brute_verifiee_apres_sauvegarde"])
        self.assertTrue(b["reconciliation_equilibree"])

    def test_02_ligne_oubliee_par_lecteur_detectee(self):
        self.ecrire()
        with patch.object(C, "lire_parcours", return_value=[]):
            f, b = self.executer()
        self.assertEqual(b["alertes_par_cause"]["LIGNE_CANDIDATE_ABSENTE_DU_LECTEUR"], 1)
        r = next(r for r in self.archives(f) if r["SOURCE"] == "PARCOURS" and r["LIGNE"] == 2)
        self.assertEqual(r["STATUT"], "EN_ATTENTE_CONSERVEE")
        self.assertIn("E_FICTIF", r["CELLULES_SOURCE_JSON"])

    def test_03_feuille_utile_inconnue_interdit_certification(self):
        self.ecrire()
        w = load_workbook(self.source)
        w.create_sheet("Nouveau format").append(["Notes utiles", 14])
        w.save(self.source); w.close()
        f, b = self.executer()
        self.assertEqual(b["collecte"], "COUVERTURE_NON_DEMONTREE")
        self.assertEqual(b["alertes_par_cause"]["FEUILLE_NON_RECONNUE"], 1)
        self.assertTrue(any("Notes utiles" in r["CELLULES_SOURCE_JSON"] for r in self.archives(f)))

    def test_04_feuilles_vides_techniques_entetes(self):
        self.ecrire()
        w = load_workbook(self.source)
        w.create_sheet("Vide")
        w.create_sheet("Fichiers").append(["Fichier", "archive.xlsx"])
        w.create_sheet("Parcours vide").append(self.head)
        w.save(self.source); w.close()
        f, b = self.executer()
        etats = {r["feuille"]: r["statut"] for r in self.controle.feuilles}
        self.assertEqual(etats["Vide"], "VIDE_OU_ENTETES_SEULS")
        self.assertEqual(etats["Fichiers"], "TECHNIQUE_NON_OBSERVATIONNELLE")
        self.assertEqual(etats["Parcours vide"], "VIDE_OU_ENTETES_SEULS")
        self.assertEqual(b["collecte"], "COUVERTURE_COMPLETE_DU_PERIMETRE_DECLARE")

    def test_05_champ_nouveau_conserve_et_signale(self):
        self.ecrire([self.row + ["information inedite"]], self.head + ["Champ metier nouveau"])
        f, b = self.executer()
        self.assertIn("CHAMP_NON_PREVU_CONSERVE", b["alertes_par_cause"])
        self.assertEqual(b["collecte"], "COUVERTURE_PARTIELLE_EXPLIQUEE")
        self.assertTrue(any("information inedite" in r["CELLULES_SOURCE_JSON"] for r in self.archives(f)))

    def test_06_champ_entierement_vide_inventorie(self):
        self.ecrire([self.row + [None]], self.head + ["Champ vide"])
        f, b = self.executer()
        self.assertTrue(any(self.controle.libelles[k] == "Champ vide" and v["non_vides"] == 0
                            for k, v in self.controle.champs.items()))

    def test_07_note_invalide_conservee(self):
        r = self.row.copy(); r[6] = r[8] = "illisible"
        self.ecrire([r])
        f, b = self.executer()
        self.assertIn("VALEUR_NON_INTERPRETEE", b["alertes_par_cause"])
        self.assertTrue(any("illisible" in r["CELLULES_SOURCE_JSON"] for r in self.archives(f)))

    def test_08_formule_sans_cache_preservee_sans_recalcul(self):
        r = self.row.copy(); r[8] = "=6+6"
        self.ecrire([r])
        f, b = self.executer()
        self.assertIn("FORMULE_SANS_CACHE", b["alertes_par_cause"])
        r = next(r for r in self.archives(f) if r["SOURCE"] == "PARCOURS" and r["LIGNE"] == 2)
        self.assertEqual(json.loads(r["CELLULES_SOURCE_JSON"])["formules"]["I"], "=6+6")

    def test_09_resultat_pl_oublie_par_lecteur_detecte(self):
        r = self.row.copy(); r[3:8] = ["20C", "APS1", "Cordon", 12, "NOGO"]
        self.ecrire([r])
        original = C.lire_parcours
        def supprimer(*args):
            oo = original(*args)
            oo[0]["data"]["roles"]["PL"]["res"] = None
            return oo
        with patch.object(C, "lire_parcours", side_effect=supprimer):
            f, b = self.executer()
        self.assertIn("CHAMP_UTILE_NON_RESTITUE", b["alertes_par_cause"])
        self.assertTrue(any("APS1_PL_PARCOURS_RESULTAT" in r["CHAMPS_EN_CAUSE"] for r in self.archives(f)))

    def test_10_degree_role_inconnus_signales(self):
        a = self.row.copy(); a[4] = "DEGREE_INCONNU"
        b = self.row.copy(); b[4:6] = ["APS1", "ROLE_INCONNU"]
        self.ecrire([a, b])
        _, bilan = self.executer()
        self.assertIn("DEGREE_NON_RECONNU", bilan["alertes_par_cause"])
        self.assertIn("ROLE_NON_RECONNU", bilan["alertes_par_cause"])

    def test_11_attente_conservee_sans_fausse_perte(self):
        self.ecrire()
        w = load_workbook(self.source)
        ws = w.create_sheet("Détails Eval")
        ws.append(["Nom", "Prénom", "Promotion", "Degré", "Simulation", "Rôle", "Résultat", "Passage"])
        ws.append(["INCONNU", "Fictif", "24A", "BTWR", "TEST", "TWR", "NOGO", "99"])
        w.save(self.source); w.close()
        f, b = self.executer()
        self.assertEqual(b["collecte"], "COUVERTURE_COMPLETE_DU_PERIMETRE_DECLARE")
        self.assertEqual(b["exploitation"], "SOUS_RESTRICTIONS")
        self.assertEqual(b["devenir_observations_candidates"]["EN_ATTENTE_CONSERVEE"], 1)
        self.assertEqual(sum(b["devenir_observations_candidates"].values()), b["observations_candidates"])

    def test_12_repetitions_non_fusionnees(self):
        self.ecrire([self.row] * 9)
        f, b = self.executer()
        rows = [r for r in self.archives(f) if r["SOURCE"] == "PARCOURS" and r["NATURE"] == "CANDIDATE"]
        self.assertEqual(len(rows), 9)
        self.assertEqual(len({r["ID_LIGNE_IMPORT"] for r in rows}), 9)

    def test_13_json_long_decoupe_et_reconstructible(self):
        textes = ["é" * 24000, "b" * 24000, "c" * 24000]
        self.ecrire([self.row + textes], self.head + ["Champ A", "Champ B", "Champ C"])
        f, b = self.executer()
        rows = [r for r in self.archives(f) if r["SOURCE"] == "PARCOURS" and r["LIGNE"] == 2]
        self.assertGreater(len(rows), 2)
        self.assertEqual(len({r["ID_LIGNE_IMPORT"] for r in rows}), 1)
        self.assertTrue(all(len(r["CELLULES_SOURCE_JSON"]) <= V.LIMITE_JSON for r in rows))
        payload = json.loads("".join(r["CELLULES_SOURCE_JSON"] for r in sorted(rows, key=lambda r:r["PARTIE_JSON"])))
        self.assertEqual([payload["valeurs"][col] for col in ("K", "L", "M")], textes)

    def test_14_alteration_archive_detectee(self):
        self.ecrire()
        f, _ = self.executer()
        w = load_workbook(f)
        w["COUVERTURE_DONNEES"].cell(2, len(V.H_DONNEES), "ALTERE")
        w.save(f); w.close()
        with self.assertRaisesRegex(ValueError, "Conservation de couverture incorrecte"):
            self.controle.verifier(f)

    def test_15_sessions_dates_et_nogo_controles(self):
        p = self.root / "sessions.xlsx"; H.sessions(p)
        _, b = self.executer({"SESSIONS": p})
        self.assertNotIn("CHAMP_UTILE_NON_RESTITUE", b["alertes_par_cause"])
        original = C.lire_sessions
        def modifier(*args):
            oo = original(*args); oo[0]["data"]["res"] = "GO"; return oo
        with patch.object(C, "lire_sessions", side_effect=modifier):
            _, b = self.executer({"SESSIONS": p})
        self.assertIn("CHAMP_UTILE_NON_RESTITUE", b["alertes_par_cause"])

    def test_16_date_illisible_conservee(self):
        p = self.root / "sessions.xlsx"; H.sessions(p)
        w = load_workbook(p); w.active["D1"] = "DATE_ILLISIBLE"; w.save(p); w.close()
        f, b = self.executer({"SESSIONS": p})
        self.assertIn("VALEUR_NON_INTERPRETEE", b["alertes_par_cause"])
        self.assertTrue(any("DATE_ILLISIBLE" in r["CELLULES_SOURCE_JSON"] for r in self.archives(f)))

    def test_17_chronologie_klif_et_exclusions(self):
        p = self.root / "klif.xlsx"; H.klif(p)
        w = load_workbook(p)
        for titre in ("99A", "00A", "12A"):
            ws = w.copy_worksheet(w["20A"]); ws.title = titre
        w.save(p); w.close()
        _, b = self.executer({"KLIF": p})
        self.assertLess(C.ordre_promo("99A"), C.ordre_promo("00A"))
        self.assertLess(C.ordre_promo("00A"), C.ordre_promo("12A"))
        self.assertLess(C.ordre_promo("12A"), C.ordre_promo("12B"))
        etats = {r["feuille"]: r["statut"] for r in self.controle.feuilles}
        for titre in ("99A", "00A", "11A"):
            self.assertEqual(etats[titre], "EXCLUE_PAR_REGLE")
        self.assertEqual(etats["12A"], "LUE_SYNTHESE")
        self.assertGreater(b["exclusions_lignes_non_vides_par_source"]["KLIF"], 0)

    def test_18_filtre_avant_12a_limite_a_klif_par_c2(self):
        r = self.row.copy(); r[3] = "99A"
        self.ecrire([r])
        f, b = self.executer()
        self.assertNotIn("EXCLUSION_AVANT_12A_HORS_KLIF_A_CONFIRMER", b["alertes_par_cause"])
        w = load_workbook(f, read_only=True, data_only=True)
        try:
            rows = list(w["PARCOURS_CONSOLIDES"].iter_rows(values_only=True))
            promo_col = rows[0].index("PROMOTION")
            self.assertIn("99A", [row[promo_col] for row in rows[1:]])
        finally:
            w.close()

    def test_19_bilan_partageable_sans_donnees_individuelles(self):
        self.ecrire()
        _, b = self.executer()
        texte = json.dumps(b, ensure_ascii=False)
        for interdit in ("MARTIN", "Paul", "E_FICTIF", str(self.root), "parcours.xlsx", "D1"):
            self.assertNotIn(interdit, texte)

    def test_20_population_recente_monosource_et_reexecution(self):
        r = self.row.copy(); r[:4] = ["NOUVEAU", "Fictif", "N_FICTIF", "26B"]
        self.ecrire([r])
        _, a = self.executer()
        f, b = self.executer()
        self.assertEqual(a["observations_candidates"], b["observations_candidates"])
        self.assertTrue(any(g["promotion"] == "26B" and g["source"] == "PARCOURS"
                            for g in b["couverture_par_source_promotion_degree_niveau"]))
        w = load_workbook(f, read_only=True)
        rows = list(w["PARCOURS_CONSOLIDES"].iter_rows(values_only=True)); w.close()
        h = rows[0]
        self.assertEqual(sum(r[h.index("NOM")] == "NOUVEAU" for r in rows[1:]), 1)

    def test_21_notes_role_non_interpretable_signale(self):
        p = self.root / "notes.xlsx"; H.notes(p)
        w = load_workbook(p); w.active["E1"] = "PL: INVALIDE"; w.save(p); w.close()
        _, b = self.executer({"NOTES": p})
        self.assertIn("VALEUR_NON_INTERPRETEE", b["alertes_par_cause"])

    def test_22_astral_compteur_controle_independamment(self):
        p = self.root / "astral.xlsx"; H.astral(p)
        original = C.lire_astral
        def modifier(*args):
            oo = original(*args); oo[0]["data"]["remedials"]["APS1"] = 9; return oo
        with patch.object(C, "lire_astral", side_effect=modifier):
            _, b = self.executer({"ASTRAL": p})
        self.assertIn("CHAMP_UTILE_NON_RESTITUE", b["alertes_par_cause"])

    def test_23_klif_valeur_source_oubliee_detectee(self):
        p = self.root / "klif.xlsx"; H.klif(p)
        original = C.lire_klif
        def modifier(*args):
            oo = original(*args)
            for o in oo: o["data"]["zone"] = None
            return oo
        with patch.object(C, "lire_klif", side_effect=modifier):
            _, b = self.executer({"KLIF": p})
        self.assertIn("CHAMP_UTILE_NON_RESTITUE", b["alertes_par_cause"])

    def test_24_reef_omission_de_champ_detectee_independamment(self):
        self.ecrire()
        original = C.lire_reef
        def modifier(*args):
            titre, h, lignes, dossiers = original(*args)
            # Simule une omission dans le lecteur, avant le calcul des sorties.
            for ligne in lignes:
                ligne["valeurs"]["COMMENTAIRE"] = None
            return titre, h, lignes, dossiers
        with patch.object(C, "lire_reef", side_effect=modifier):
            _, b = self.executer()
        self.assertIn("CHAMP_UTILE_NON_RESTITUE", b["alertes_par_cause"])

    def test_25_promotion_inconnue_non_divulguee_dans_bilan(self):
        r = self.row.copy(); r[3] = "PROMO_INCONNUE_NOM_CONFIDENTIEL"
        self.ecrire([r])
        f, b = self.executer()
        self.assertIn("PROMOTION_NON_RECONNUE", b["alertes_par_cause"])
        self.assertNotIn(r[3], json.dumps(b))
        self.assertTrue(any(r[3] in x["CELLULES_SOURCE_JSON"] for x in self.archives(f)))


def main():
    result = unittest.TextTestRunner(verbosity=2).run(
        unittest.defaultTestLoader.loadTestsFromTestCase(Couverture))
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())

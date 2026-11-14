# -*- coding: utf-8 -*-
"""Reception C1-C8 sur sources exclusivement fictives. Aucune source utilisateur lue."""
import datetime as dt
import hashlib
import json
import tempfile
import unittest
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook, load_workbook

import consolidation as C
import test_consolidation as H

@contextmanager
def open_read(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        yield wb
    finally:
        wb.close()


def rows_of(wb, sheet):
    it = wb[sheet].iter_rows(values_only=True)
    h = next(it)
    return [dict(zip(h, row)) for row in it]


class LotC1C8(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.reef = self.root / 'REEF_REBUILD.xlsx'
        H.reef(self.reef)
        self.resultats = self.root / 'resultats'

    def tearDown(self):
        self.temp.cleanup()

    def run_app(self, sources):
        all_sources = {'REEF': self.reef, **sources}
        before = {s: C.sha(p) for s, p in all_sources.items()}
        out, bilan = C.consolider(self.reef, sources, self.resultats, journal=lambda *_: None)
        self.assertEqual(before, {s: C.sha(p) for s, p in all_sources.items()})
        self.assertEqual(bilan['verification_fichier_enregistre']['statut'], 'CONFORME_APRES_RELECTURE')
        self.assertTrue(bilan['couverture_donnees']['conservation_brute_verifiee_apres_sauvegarde'])
        self.assertTrue(bilan['couverture_donnees']['reconciliation_equilibree'])
        return out, bilan

    def new_parcours(self, name='parcours.xlsx'):
        p = self.root / name
        wb = Workbook()
        ws = wb.active
        ws.title = 'Parcours'
        ws.append(['Nom', 'Prénom', 'Promotion', 'Degré', 'Position 1', 'Note Position 1',
                   'Resultat 1', 'Note', 'Résultat'])
        wb.save(p)
        wb.close()
        return p

    def add_reef(self, did, nom, prenom, promos):
        wb = load_workbook(self.reef)
        ws = wb['REEF_REBUILD']
        h = [c.value for c in ws[1]]
        for ordre, pr in enumerate(promos, 1):
            values = {'ID_DOSSIER_LOCAL': did, 'NOM': nom, 'PRENOM': prenom,
                      'NUMERO_PARCOURS': 1, 'ORDRE_PERIODE': ordre, 'PROMOTION': pr}
            ws.append([values.get(c) for c in h])
        wb.save(self.reef)
        wb.close()

    def test_c1_feuilles_ignorees_et_inconnue(self):
        wr = load_workbook(self.reef)
        for name in ('REPRISES_A_PRECISER', 'LIRE_MOI', '_META_LOCAL'):
            wr.create_sheet(name).append(['Information auxiliaire', 'aucune observation'])
        wr.create_sheet('FEUILLE_INCONNUE').append(['Contenu potentiellement utile', 13])
        wr.save(self.reef)
        wr.close()
        p = self.root / 'notes.xlsx'
        H.notes(p)
        wn = load_workbook(p)
        wn.create_sheet('AUDIT_DEGREES').append(['Notes', 11])
        wn.save(p)
        wn.close()
        file, bilan = self.run_app({'NOTES': p})
        with open_read(file) as w:
            feuilles = rows_of(w, 'COUVERTURE_FEUILLES')
        statuses = {(r['SOURCE'], r['FEUILLE']): r['STATUT'] for r in feuilles}
        for n in ('REPRISES_A_PRECISER', 'LIRE_MOI', '_META_LOCAL'):
            self.assertEqual(statuses['REEF', n], 'EXCLUE_PAR_REGLE')
        self.assertEqual(statuses['NOTES', 'AUDIT_DEGREES'], 'EXCLUE_PAR_REGLE')
        self.assertEqual(statuses['REEF', 'FEUILLE_INCONNUE'], 'NON_RECONNUE_A_EXAMINER')
        self.assertEqual(bilan['couverture_donnees']['alertes_par_cause'].get('FEUILLE_NON_RECONNUE'), 1)
        self.assertEqual(bilan['couverture_donnees']['exclusions_par_regle']['FEUILLE_AUXILIAIRE_C1'], 4)

    def test_c2_promotions_p_toutes_annees_et_avant12a_hors_klif(self):
        p = self.new_parcours()
        w = load_workbook(p)
        ws = w['Parcours']
        for pr in ('99A', '00A', '11C', '12A'):
            ws.append(['ELEVE' + pr, 'Fictif', pr, 'BTWR', 'TWR', 15, 'GO', 15, 'GO'])
        ws.append(['INCONNUE', 'Fictif', 'PROMOTION_INCONNUE', 'BTWR', 'TWR', 12, 'GO', 12, 'GO'])
        w.save(p)
        w.close()
        pn, ps, pk = self.root / 'notes.xlsx', self.root / 'sessions.xlsx', self.root / 'klif.xlsx'
        H.notes(pn); H.sessions(ps); H.klif(pk)
        for source, path in [('NOTES', pn), ('SESSIONS', ps)]:
            wb = load_workbook(path)
            ws = wb.active
            for pr in ('11P', '25P', '20P'):
                r = ws.max_row + 2
                ws.cell(r, 1, 'MARTIN Paul' if pr == '20P' else 'HORS P Fictif')
                ws.cell(r, 2, 'Class: MCTA' + pr)
                ws.cell(r + 1, 2, 'Degree: BTWR')
                if source == 'NOTES':
                    ws.cell(r, 5, 'TWR: 15'); ws.cell(r, 7, 15)
                else:
                    ws.cell(r, 4, dt.datetime(2026, 3, 1)); ws.cell(r, 6, 'GO')
            wb.save(path); wb.close()
        wb = load_workbook(pk)
        for pr in ('12P', '25P'):
            nw = wb.copy_worksheet(wb['20A']); nw.title = pr
            nw.cell(2, 1, pr)
        wb.save(pk); wb.close()
        file, bilan = self.run_app({'PARCOURS': p, 'NOTES': pn, 'SESSIONS': ps, 'KLIF': pk})
        with open_read(file) as w:
            data = rows_of(w, 'PARCOURS_CONSOLIDES')
        promos = [str(r['PROMOTION']) for r in data]
        for pr in ('99A', '00A', '11C', '12A'):
            self.assertIn(pr, promos)
        self.assertFalse(any(pr in promos for pr in ('11P', '12P', '20P', '25P')))
        self.assertEqual(bilan['lecture']['NOTES']['blocs_hors_perimetre_exclus'], 3)
        self.assertEqual(bilan['lecture']['SESSIONS']['blocs_hors_perimetre_exclus'], 3)
        self.assertEqual(bilan['lecture']['KLIF']['feuilles_voie_p_exclues'], 2)
        self.assertIn('PROMOTION_NON_RECONNUE', bilan['couverture_donnees']['alertes_par_cause'])
        self.assertNotIn('EXCLUSION_AVANT_12A_HORS_KLIF_A_CONFIRMER',
                         bilan['couverture_donnees']['alertes_par_cause'])

    def test_c2_archive_ancienne_conservee(self):
        registre = {'version': 1, 'dossiers': {}, 'periodes': {},
                    'hors_perimetre_avant_12A': {'dossiers': {
                        'OLD': {'cle': 'archive fictif', 'nom': 'ARCHIVE', 'prenom': 'Fictif'}},
                        'periodes': {'OLD': {'99A': {'ordre': 1, 'origine': 'HISTORIQUE'}}}}}
        self.resultats.mkdir()
        (self.resultats / C.REGISTRE_POPULATION).write_text(json.dumps(registre), encoding='utf8')
        p = self.new_parcours()
        w = load_workbook(p)
        w.active.append(['ARCHIVE', 'Fictif', '99A', 'BTWR', 'TWR', 14, 'GO', 14, 'GO'])
        w.save(p); w.close()
        file, b = self.run_app({'PARCOURS': p})
        apres = json.loads((self.resultats / C.REGISTRE_POPULATION).read_text('utf8'))
        self.assertEqual(apres['hors_perimetre_avant_12A'], registre['hors_perimetre_avant_12A'])
        with open_read(file) as w:
            self.assertFalse(any(r['NOM'] == 'ARCHIVE' for r in rows_of(w, 'PARCOURS_CONSOLIDES')))

    def test_c3_moyenne_toute_colonne_dernier_eleve_et_conflit(self):
        p = self.root / 'klif.xlsx'
        wb = Workbook(); ws = wb.active; ws.title = '20A'
        h = ['PROMO', 'Nom du stagiaire', 'Prénom du stagiaire', 'Affectation',
             "Date d'affectation", 'Zone']
        ws.append(h)
        ws.append(['20A', 'MARTIN', 'Paul', 'LFRR', dt.datetime(2023, 1, 1), 'W'])
        ws.append(['20A', None, None, None, None, '  MoYenne  '])
        ws.append(['20A', 'CONFLIT', 'Fictif', None, None, 'MOYENNE'])
        ws.append(['20A', None, None, 42])
        ws.append(['20A', None, None, None, None, 'Moyenne!'])
        ws.append(['20A', 'DERNIER', 'Fictif', 'LFXX', dt.datetime(2024, 1, 1), 'Z'])
        wb.save(p); wb.close()
        file, b = self.run_app({'KLIF': p})
        self.assertEqual(b['lecture']['KLIF']['lignes_moyenne_ignorees'], 1)
        self.assertEqual(b['couverture_donnees']['moyennes_klif_ignorees'], 1)
        self.assertEqual(b['points_a_examiner_par_cause'].get('KLIF_MOYENNE_AVEC_IDENTITE'), 1)
        self.assertIn('CONTENU_NON_RECONNU_CONSERVE', b['couverture_donnees']['alertes_par_cause'])
        with open_read(file) as w:
            data = rows_of(w, 'PARCOURS_CONSOLIDES')
            detail = rows_of(w, 'COUVERTURE_DONNEES')
        self.assertTrue(any(r['NOM'] == 'DERNIER' for r in data))
        self.assertFalse(any(r['NOM'] == 'CONFLIT' for r in data))
        self.assertTrue(any(r['SOURCE'] == 'KLIF' and r['LIGNE'] == 3
                            and r['NATURE'] == 'EXCLUE_PAR_REGLE' for r in detail))

    def test_c4_deux_dossiers_18a_trente_observations(self):
        noms = [('Dd43dd2048f958378', 'ALPHA'), ('D4d8e8aa756149655', 'BETA')]
        for did, nom in noms:
            self.add_reef(did, nom, 'Fictif', ['18A', '18A'])
        self.add_reef('AUTRE', 'GAMMA', 'Fictif', ['18A', '18A'])
        p = self.root / 'parcours.xlsx'
        wb = Workbook(); ws = wb.active; ws.title = 'Parcours'
        ws.append(['Nom', 'Prénom', 'Promotion', 'Degré', 'Position 1', 'Note Position 1',
                   'Resultat 1', 'Note', 'Résultat'])
        ds = wb.create_sheet('Détails Eval')
        ds.append(['Nom', 'Prénom', 'Promotion', 'Degré', 'Passage', 'EvalId', 'RoleId',
                   'Rôle', 'Résultat', 'Rattrapage', 'Simulation', 'Source', 'Feuille source', 'Ligne source'])
        for did, nom in noms:
            for pr, dg, note, qty in [('MCTA18A', 'ACS1', 11, 4), ('MCTA18A', 'ACS2', 12, 4),
                                     ('MCTA18Abis', 'ACS2', 13, 4)]:
                ws.append([nom, 'Fictif', pr, dg, 'CO', note, 'GO', note, 'GO'])
                for j in range(qty):
                    ds.append([nom, 'Fictif', pr, dg, str(j), 'Eval' + str(j), 'Role' + str(j),
                               'CO', 'GO', '0', 'Sim', 'test', 'Feuille', j + 2])
        ws.append(['GAMMA', 'Fictif', 'MCTA18Abis', 'ACS2', 'CO', 16, 'GO', 16, 'GO'])
        wb.save(p); wb.close()
        file, b = self.run_app({'PARCOURS': p})
        self.assertEqual(b['statuts_rapprochement']['PARCOURS']['PERIODE_C4_18A_EXPLICITE'], 30)
        self.assertIn('PERIODE_AMBIGUE', b['points_a_examiner_par_cause'])
        with open_read(file) as w:
            data = rows_of(w, 'PARCOURS_CONSOLIDES')
            prov = rows_of(w, 'PROVENANCE')
        for did, _ in noms:
            par = [r for r in data if r['ID_DOSSIER_LOCAL'] == did]
            self.assertEqual(len(par), 2)
            a, z = sorted(par, key=lambda x: x['ORDRE_PERIODE'])
            self.assertEqual(a['ACS1_CO_PARCOURS_NOTE'], 11)
            self.assertEqual(a['ACS2_CO_PARCOURS_NOTE'], 12)
            self.assertEqual(z['ACS2_CO_PARCOURS_NOTE'], 13)
        dp = [r for r in prov if r['COLONNE'] == 'OBSERVATION_DETAILLEE'
              and r['ID_DOSSIER_LOCAL'] in dict(noms)]
        self.assertEqual(len(dp), 24)
        self.assertEqual(sum(x['ORDRE_PERIODE'] == 1 for x in dp), 16)
        self.assertEqual(sum(x['ORDRE_PERIODE'] == 2 for x in dp), 8)

    def test_c4_absente_ou_dupliquee_sans_repli(self):
        o = {'source': 'PARCOURS', 'degree': 'ACS2', 'promo': '18A', 'promo_brute': 'MCTA18Abis'}
        d = {'id': 'Dd43dd2048f958378', 'periodes': [{'ordre': 1, 'promo': '18A'}]}
        self.assertEqual(C.Moteur.periode(d, o), (None, 'PERIODE_C4_ABSENTE'))
        d['periodes'] += [{'ordre': 2, 'promo': '18A'}, {'ordre': 2, 'promo': '18A'}]
        self.assertEqual(C.Moteur.periode(d, o), (None, 'PERIODE_C4_AMBIGUE'))
        d['id'] = 'AUTRE'
        self.assertEqual(C.Moteur.periode(d, o), (None, 'PERIODE_AMBIGUE'))

    def test_c5_astral_abes_enac_pourcentage_et_total(self):
        p = self.root / 'astral.xlsx'
        wb = Workbook(); ws = wb.active; ws.title = '20A'
        ws.append(['', 'Apprenant', 'ABES_ADI', 'BTWR', 'ADI', 'APS1'])
        ws.append(['', 'MARTIN Paul', '9 remedial(s)', '100% done', '100% done', '100% done'])
        ws.append(['', 'ENAC', '', '', '2 remedial(s)', '1 remedial(s)'])
        ws.append(['', 'DURAND Léa', '', '100% done', '100% done', '100% done'])
        ws.append(['', 'TOTAL', '', '', '999 remedial(s)', '999 remedial(s)'])
        wb.save(p); wb.close()
        file, b = self.run_app({'ASTRAL': p})
        with open_read(file) as w:
            martin = next(r for r in rows_of(w, 'PARCOURS_CONSOLIDES')
                          if r['ID_DOSSIER_LOCAL'] == 'D1' and r['ORDRE_PERIODE'] == 1)
        self.assertEqual(martin['ADI_ASTRAL_REMEDIALS'], 2)
        self.assertEqual(martin['APS1_ASTRAL_REMEDIALS'], 1)
        self.assertEqual(martin['ADI_ASTRAL_PRESENCE_REMEDIAL'], 'OUI')
        self.assertIsNone(martin['BTWR_ASTRAL_REMEDIALS'])
        self.assertIsNone(martin['BTWR_ASTRAL_PRESENCE_REMEDIAL'])
        self.assertNotIn('DEGREE_NON_RECONNU', b['couverture_donnees']['alertes_par_cause'])
        self.assertEqual(b['lecture']['ASTRAL']['observations'], 1)

    def test_c5_reprise_astral_periodes_distinctes(self):
        self.add_reef('REP24', 'REPRISE', 'Fictif', ['24A', '24B'])
        p = self.root / 'astral.xlsx'
        wb = Workbook(); w1 = wb.active; w1.title = '24A'
        for pr, count, ws in [('24A', 1, w1), ('24B', 3, wb.create_sheet('24B'))]:
            ws.append(['', 'Apprenant', 'APS1', 'ACS2'])
            ws.append(['', 'REPRISE Fictif', '100% done', None])
            ws.append(['', 'ENAC', '%d remedial(s)' % count, None])
            ws.append(['', 'TOTAL', '999 remedial(s)', None])
        wb.save(p); wb.close()
        file, b = self.run_app({'ASTRAL': p})
        with open_read(file) as w:
            data = [r for r in rows_of(w, 'PARCOURS_CONSOLIDES') if r['ID_DOSSIER_LOCAL'] == 'REP24']
        self.assertEqual({r['PROMOTION']: r['APS1_ASTRAL_REMEDIALS'] for r in data}, {'24A': 1, '24B': 3})
        self.assertTrue(all(r['ACS2_ASTRAL_REMEDIALS'] is None for r in data))

    def test_c6_klif_origine_duree_source_et_identifiants_ignores(self):
        p = self.root / 'klif.xlsx'
        H.klif(p)
        wb = load_workbook(p)
        for pr, orig, duree in [('20A', 'CE', 8.5), ('21B', 'CI', -1), ('23A', 'MCTA18A', '?')]:
            ws = wb[pr]
            c = ws.max_column + 1
            for j, n in enumerate(['Origine', 'Nombre de mois de qualification ENAC exclus',
                                   'Matricule', 'Prolongations']):
                ws.cell(1, c + j, n)
            for i in range(2, ws.max_row + 1):
                if ws.cell(i, 3).value:
                    ws.cell(i, c, orig)
                    ws.cell(i, c + 1, duree)
                    ws.cell(i, c + 2, 'MAT_FICTIF')
                    ws.cell(i, c + 3, '2 mois')
        wb.save(p); wb.close()
        file, b = self.run_app({'KLIF': p})
        with open_read(file) as w:
            data = rows_of(w, 'PARCOURS_CONSOLIDES')
            prov = rows_of(w, 'PROVENANCE')
        martin = next(r for r in data if r['ID_DOSSIER_LOCAL'] == 'D1' and r['ORDRE_PERIODE'] == 2)
        durand = next(r for r in data if r['ID_DOSSIER_LOCAL'] == 'D2')
        self.assertEqual(martin['KLIF_ORIGINE'], 'CE')
        self.assertEqual(martin['KLIF_MOIS_QUALIFICATION_ENAC_EXCLUS'], 8.5)
        self.assertEqual(martin['KLIF_DUREE_ENAC_EXCLUS_BRUTE'], 8.5)
        self.assertNotEqual(martin['KLIF_MOIS_QUALIFICATION_ENAC_EXCLUS'], martin['KLIF_MOIS_AFFECTATION_MU_1'])
        self.assertEqual(durand['KLIF_ORIGINE'], 'CI')
        self.assertIsNone(durand['KLIF_MOIS_QUALIFICATION_ENAC_EXCLUS'])
        self.assertEqual(durand['KLIF_DUREE_ENAC_EXCLUS_BRUTE'], -1)
        self.assertIn('DUREE_ENAC_EXCLUS_NEGATIVE', durand['KLIF_ALERTES'])
        self.assertTrue(any(r['COLONNE'] == 'KLIF_ORIGINE' and r['VALEUR'] == 'CE' for r in prov))
        self.assertTrue(any(r['COLONNE'] == 'KLIF_MOIS_QUALIFICATION_ENAC_EXCLUS' and r['VALEUR'] == 8.5 for r in prov))
        self.assertNotIn('CHAMP_NON_PREVU_CONSERVE', b['couverture_donnees']['alertes_par_cause'])

    def test_c6_origines_de_tous_types_et_mois_virgule(self):
        p = self.root / 'klif_types.xlsx'
        w = Workbook(); ws = w.active; ws.title = '24A'
        ws.append(['PROMO', 'Nom du stagiaire', 'Prénom du stagiaire', 'Origine',
                   'Nombre de mois de qualification ENAC exclus'])
        for i, (origine, duree) in enumerate((('CE', 10), ('CI', 11), ('MCTA21B', '12,5'),
                                              ('TSEEAC', 13)), 1):
            ws.append(['24A', 'ORIGINE' + str(i), 'Fictif', origine, duree])
        w.save(p); w.close()
        from collections import defaultdict, Counter
        o = C.lire_klif(p, defaultdict(Counter))
        self.assertEqual([x['data']['origine'] for x in o], ['CE', 'CI', 'MCTA21B', 'TSEEAC'])
        self.assertEqual([C.nombre(x['data']['duree_enac_exclus_brute']) for x in o],
                         [10.0, 11.0, 12.5, 13.0])

    def test_c7_note_nogo_reconstituee_et_position2_neutralisee(self):
        p = self.root / 'parcours.xlsx'
        wb = Workbook(); ws = wb.active; ws.title = 'Parcours'
        ws.append(['Nom', 'Prénom', 'Promotion', 'Degré', 'Position 1', 'Résultat 1',
                   'Position 2', 'Note Position 2', 'Resultat 2', 'Traitement Position 2',
                   'Note', 'Résultat', 'Note reconstituée', 'Origine note reconstituée', 'Statut note reconstituée'])
        ws.append(['MARTIN', 'Paul', '20A', 'BTWR', 'TWR', 'NOGO', '2', 12, 'GO',
                   'neutralisée', 'NOGO', 'NOGO', 11.5, 'RECONSTRUCTION_SOURCE', 'VALIDEE'])
        wb.save(p); wb.close()
        file, b = self.run_app({'PARCOURS': p})
        with open_read(file) as w:
            row = next(r for r in rows_of(w, 'PARCOURS_CONSOLIDES') if r['ID_DOSSIER_LOCAL'] == 'D1' and r['ORDRE_PERIODE'] == 1)
            prov = rows_of(w, 'PROVENANCE')
        self.assertEqual(row['BTWR_PARCOURS_RESULTAT'], 'NOGO')
        self.assertIsNone(row['BTWR_PARCOURS_NOTE'])
        self.assertEqual(row['BTWR_PARCOURS_NOTE_RECONSTITUEE'], 11.5)
        self.assertEqual(row['BTWR_PARCOURS_ORIGINE_NOTE_RECONSTITUEE'], 'RECONSTRUCTION_SOURCE')
        self.assertEqual(row['BTWR_PARCOURS_STATUT_NOTE_RECONSTITUEE'], 'VALIDEE')
        self.assertNotIn('ROLE_NON_RECONNU', b['couverture_donnees']['alertes_par_cause'])
        self.assertTrue(any(r['COLONNE'] == 'BTWR_PARCOURS_NOTE_RECONSTITUEE' and r['VALEUR'] == 11.5 for r in prov))

    def test_c7_valeur_invalide_conflits_et_role_inconnu(self):
        p = self.root / 'parcours.xlsx'
        wb = Workbook(); ws = wb.active; ws.title = 'Parcours'
        ws.append(['Nom', 'Prénom', 'Promotion', 'Degré', 'Position 1', 'Position 2', 'Note Position 2',
                   'Note', 'Résultat', 'Note reconstituée', 'Origine note reconstituée'])
        for nr in (10.0, 11.0):
            ws.append(['MARTIN', 'Paul', '20A', 'BTWR', 'TWR', 'MYSTERE', 12,
                       'NOGO', 'NOGO', nr, 'SOURCE_FICTIVE'])
        ws.append(['DURAND', 'Léa', '21B', 'ADI', 'LOC', None, None,
                   'NOGO', 'NOGO', 'INVALIDE', 'SOURCE_FICTIVE'])
        wb.save(p); wb.close()
        file, b = self.run_app({'PARCOURS': p})
        with open_read(file) as w:
            data = rows_of(w, 'PARCOURS_CONSOLIDES')
            examiner = rows_of(w, 'A_EXAMINER')
        martin = next(r for r in data if r['ID_DOSSIER_LOCAL'] == 'D1' and r['ORDRE_PERIODE'] == 1)
        self.assertEqual(martin['BTWR_PARCOURS_NOTE_RECONSTITUEE'], 'PLUSIEURS_VALEURS')
        self.assertEqual(martin['BTWR_PARCOURS_RESULTAT'], 'NOGO')
        self.assertEqual(len([r for r in examiner if r['CHAMP'] == 'BTWR_PARCOURS_NOTE_RECONSTITUEE']), 2)
        self.assertIn('ROLE_NON_RECONNU', b['couverture_donnees']['alertes_par_cause'])
        self.assertIn('VALEUR_NON_INTERPRETEE', b['couverture_donnees']['alertes_par_cause'])

    def test_c7_vingt_sept_observations_fictives_conservees(self):
        p = self.root / 'parcours_27_fictifs.xlsx'
        wb = Workbook(); ws = wb.active; ws.title = 'Parcours'
        ws.append(['Nom', 'Prénom', 'Promotion', 'Degré', 'Note', 'Résultat',
                   'Note reconstituée', 'Origine note reconstituée', 'Statut note reconstituée'])
        for i in range(27):
            nom = 'F' + hashlib.sha256(str(i).encode()).hexdigest()[:10]
            ws.append([nom, 'Fictif', '24A', 'BTWR', 'NOGO', 'NOGO',
                       10.5, 'FOURNIE', 'CONFIRMEE'])
        wb.save(p); wb.close()
        file, b = self.run_app({'PARCOURS': p})
        with open_read(file) as w:
            base = rows_of(w, 'PARCOURS_CONSOLIDES')
            prov = rows_of(w, 'PROVENANCE')
        self.assertEqual(b['lecture']['PARCOURS']['observations'], 27)
        self.assertEqual(sum(r['BTWR_PARCOURS_NOTE_RECONSTITUEE'] == 10.5 for r in base), 27)
        self.assertEqual(sum(r['COLONNE'] == 'BTWR_PARCOURS_NOTE_RECONSTITUEE' for r in prov), 27)

    def test_c8_origine_trace_detail_et_rattachement_indecis(self):
        p = self.root / 'parcours.xlsx'
        wb = Workbook(); ws = wb.active; ws.title = 'Parcours'
        ws.append(['Nom', 'Prénom', 'Promotion', 'Degré', 'Position 1', 'Note Position 1',
                   'Résultat 1', 'Note', 'Résultat'])
        ws.append(['NOUVEAU', 'Fictif', '24A', 'BTWR', 'TWR', 15, 'GO', 15, 'GO'])
        ds = wb.create_sheet('Détails Eval')
        ds.append(['Nom', 'Prénom', 'Promotion', 'Degré', 'Simulation', 'Rôle', 'Résultat', 'Passage'])
        ds.append(['NOUVEAU', 'Fictif', '24A', 'BTWR', 'SIM', 'TWR', 'NOGO', 0])
        ds.append(['NOUVEAU', 'Fictif', '24A', 'BTWR', 'SIM', 'TWR', 'GO', 1])
        ds.append(['INCONNU', 'Fictif', '24A', 'BTWR', 'SIM', 'TWR', 'GO', 2])
        wb.save(p); wb.close()
        file, b = self.run_app({'PARCOURS': p})
        with open_read(file) as w:
            data = rows_of(w, 'PARCOURS_CONSOLIDES')
            details = [r for r in rows_of(w, 'COUVERTURE_DONNEES') if r['SOURCE'] == 'PARCOURS' and r['FEUILLE'] == 'Détails Eval']
            prov = rows_of(w, 'PROVENANCE')
        n = next(r for r in data if r['NOM'] == 'NOUVEAU')
        self.assertEqual(n['BTWR_PARCOURS_RESULTAT'], 'GO')
        self.assertEqual(len(details), 4)  # entete + 3 details
        self.assertEqual([r['STATUT'] for r in details if r['LIGNE'] in (2, 3)], ['CONSERVEE_EN_DETAIL'] * 2)
        self.assertEqual(next(r for r in details if r['LIGNE'] == 4)['STATUT'], 'EN_ATTENTE_CONSERVEE')
        self.assertEqual(sum(r['COLONNE'] == 'OBSERVATION_DETAILLEE' for r in prov), 3)
        self.assertEqual(b['couverture_donnees']['devenir_observations_candidates']['CONSERVEE_EN_DETAIL'], 2)

    def test_c8_omission_simulee_detectee(self):
        p = self.root / 'parcours.xlsx'
        wb = Workbook(); ws = wb.active; ws.title = 'Parcours'
        ws.append(['Nom', 'Prénom', 'Promotion', 'Degré', 'Note', 'Résultat'])
        ws.append(['MARTIN', 'Paul', '20A', 'BTWR', 15, 'GO'])
        ws = wb.create_sheet('Détails Eval')
        ws.append(['Nom', 'Prénom', 'Promotion', 'Degré', 'Simulation', 'Rôle', 'Résultat'])
        ws.append(['MARTIN', 'Paul', '20A', 'BTWR', 'SIM', 'TWR', 'NOGO'])
        wb.save(p); wb.close()
        original = C.lire_parcours
        def oublier(*args):
            return [r for r in original(*args) if r['data']['niveau'] != 'DETAIL_EVALUATION']
        with patch.object(C, 'lire_parcours', side_effect=oublier):
            file, b = self.run_app({'PARCOURS': p})
        self.assertIn('LIGNE_CANDIDATE_ABSENTE_DU_LECTEUR', b['couverture_donnees']['alertes_par_cause'])
        with open_read(file) as w:
            recs = rows_of(w, 'COUVERTURE_DONNEES')
        cible = next(r for r in recs if r['SOURCE'] == 'PARCOURS' and r['FEUILLE'] == 'Détails Eval' and r['LIGNE'] == 2)
        self.assertEqual(cible['STATUT'], 'EN_ATTENTE_CONSERVEE')


    # Regressions de la validation de PROVENANCE (aucune donnee nominative).
    def _provenance_fictive(self, modifications=None):
        modifications = modifications or {}
        enregistrement = ['D1', 1, 'OBSERVATION_DETAILLEE', 'GO', 'PARCOURS',
                         'source_fictive.xlsx!Details Eval:2',
                         'DETAIL_CONSERVE_SANS_RESULTAT_GLOBAL_DEDUIT']
        enregistrement += [None] * len(C.COLONNES_TRACABILITE)
        for index, v in modifications.items():
            enregistrement[index] = v
        fichier = self.root / 'provenance_fictive.xlsx'
        C.ecrire_classeur(fichier, ['ID_DOSSIER_LOCAL', 'ORDRE_PERIODE'],
                          ['ID_DOSSIER_LOCAL', 'ORDRE_PERIODE'], {},
                          [{'ID_DOSSIER_LOCAL': 'D1', 'ORDRE_PERIODE': 1}],
                          [enregistrement], {}, {}, self.reef, {}, {}, None)
        return fichier, enregistrement

    def test_regression_provenance_nombres_precis_et_identifiants_longs(self):
        grand = 12345678901234567890
        precis = 0.12345678901234568
        fichier, attendu = self._provenance_fictive({3: precis, 15: grand, 16: grand + 1})
        C.verifier_tracabilite(fichier, [attendu], {})
        with open_read(fichier) as wb:
            r = rows_of(wb, 'PROVENANCE')[0]
        self.assertEqual(r['VALEUR'], repr(precis))
        self.assertEqual(r['EVALID_SOURCE'], str(grand))
        self.assertEqual(r['ROLEID_SOURCE'], str(grand + 1))

    def test_regression_provenance_date_et_duree_ordinaires(self):
        ts = dt.datetime(2026, 11, 13, 9, 20, 11, 123456)
        fichier, attendu = self._provenance_fictive({3: ts, 20: dt.date(2026, 11, 13)})
        C.verifier_tracabilite(fichier, [attendu], {})
        with open_read(fichier) as wb:
            r = rows_of(wb, 'PROVENANCE')[0]
        self.assertEqual(r['VALEUR'], ts.isoformat(sep=' '))
        self.assertEqual(r['LIGNE_HISTORIQUE'], '2026-11-13')
        simple, simple_attendu = self._provenance_fictive({3: 12.5})
        C.verifier_tracabilite(simple, [simple_attendu], {})
        with open_read(simple) as wb:
            self.assertEqual(rows_of(wb, 'PROVENANCE')[0]['VALEUR'], 12.5)

    def test_regression_provenance_modification_reelle_detectee(self):
        fichier, attendu = self._provenance_fictive({3: 'GO'})
        wb = load_workbook(fichier)
        wb['PROVENANCE']['D2'] = 'NOGO'
        wb.save(fichier); wb.close()
        with self.assertRaisesRegex(C.Erreur, r'ligne 2, colonne D \(VALEUR\)'):
            C.verifier_tracabilite(fichier, [attendu], {})

    def test_regression_provenance_texte_trop_long_non_tronque(self):
        with self.assertRaisesRegex(C.Erreur, 'Provenance trop longue'):
            self._provenance_fictive({3: 'x' * 32768})


def main():
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(LotC1C8)
    return 0 if unittest.TextTestRunner(verbosity=2).run(suite).wasSuccessful() else 1

if __name__ == '__main__':
    raise SystemExit(main())

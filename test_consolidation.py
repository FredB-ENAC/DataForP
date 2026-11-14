# -*- coding: utf-8 -*-
"""Tests synthetiques (donnees fictives) du fichier consolide."""
import datetime as dt, sys, tempfile, traceback
from pathlib import Path
from openpyxl import Workbook, load_workbook
import consolidation as C

D = dt.datetime


def colonnes_reef():
    out = ['ID_DOSSIER_LOCAL', 'NOM', 'PRENOM', 'NUMERO_PARCOURS', 'ORDRE_PERIODE', 'PROMOTION']
    for d in C.DEGREES:
        if d == 'APS2':
            out.append('APS2_MODALITE')
        out.append(d + '_CORRECTIVE')
        for role in C.ROLES[d]:
            stem = d if d in ('BTWR', 'ADI') else d + '_' + role
            out += [stem + '_RESULTAT_INITIAL', stem + '_CONFIRMATION', stem + '_RESULTAT_CONFIRMATION']
        out.append(d + '_RESULTAT_FINAL')
        for role in C.ROLES[d]:
            out.append(d + '_NOTE' if d in ('BTWR', 'ADI') else d + '_' + role + '_NOTE')
        if d in ('BSURV', 'ACS1'):
            out.append(d + '_NOTE')
    return out + ['MODE_ENTREE', 'CAUSE_REPRISE', 'DEGREE_REPRISE', 'ISSUE_PERIODE', 'CAUSE_ISSUE',
                  'DEGREE_REPERE_ISSUE', 'POSITION_ISSUE', 'DEGREE_REPERE_ISSUE_2', 'PROMOTION_SUITE',
                  'COMMENTAIRE', 'ETAT_RECONSTRUCTION', 'ORIGINE_ISSUE', 'COMMENTAIRE_ISSUE']


def reef(p):
    wb = Workbook(); ws = wb.active; ws.title = 'REEF_REBUILD'; h = colonnes_reef(); ws.append(h)
    def ligne(**kw):
        ws.append([kw.get(c) for c in h])
    ligne(ID_DOSSIER_LOCAL='D1', NOM='MARTIN', PRENOM='Paul', NUMERO_PARCOURS=1, ORDRE_PERIODE=1, PROMOTION='20A',
          BTWR_NOTE=14.0, ISSUE_PERIODE='REDOUBLEMENT', COMMENTAIRE='decision conservee')
    ligne(ID_DOSSIER_LOCAL='D1', NOM='MARTIN', PRENOM='Paul', NUMERO_PARCOURS=1, ORDRE_PERIODE=2, PROMOTION='20C',
          MODE_ENTREE='REPRISE')
    ligne(ID_DOSSIER_LOCAL='D2', NOM='DURAND', PRENOM='Léa', NUMERO_PARCOURS=1, ORDRE_PERIODE=1, PROMOTION='21B')
    ligne(ID_DOSSIER_LOCAL='D3', NOM='BERNARD', PRENOM='Luc', NUMERO_PARCOURS=1, ORDRE_PERIODE=1, PROMOTION='22A')
    ligne(ID_DOSSIER_LOCAL='D4', NOM='BERNARD', PRENOM='Luc', NUMERO_PARCOURS=1, ORDRE_PERIODE=1, PROMOTION='22A')
    ligne(ID_DOSSIER_LOCAL='D5', NOM='GARCIA LOPEZ', PRENOM='Ana', NUMERO_PARCOURS=1, ORDRE_PERIODE=1, PROMOTION='23A')
    ligne(ID_DOSSIER_LOCAL='D6', NOM='PETIT', PRENOM='Jean', NUMERO_PARCOURS=1, ORDRE_PERIODE=1, PROMOTION='19B')
    ligne(ID_DOSSIER_LOCAL='D6', NOM='PETIT', PRENOM='Jean', NUMERO_PARCOURS=1, ORDRE_PERIODE=2, PROMOTION='19B')
    wb.save(p)


def parcours(p):
    wb = Workbook(); ws = wb.active; ws.title = 'Parcours'
    ws.append(['Nom', 'Prénom', 'id_Aurion', 'EleveId', 'Promotion', 'Degré', 'Promotion fichier', 'Degre fichier',
               'Position 1', 'Test1', 'Note Position 1', 'Resultat 1', 'Rattrapage 1', 'Position 2', 'Test2',
               'Note Position 2', 'Resultat 2', 'Rattrapage 2', 'Note', 'Résultat'])
    ws.append(['Martin', 'paul', None, None, '20A', 'BTWR', None, None, 'TWR', None, 14.0, 'Go', None, None, None,
               None, None, None, 14.0, 'Go'])
    ws.append(['MARTIN', 'Paul', None, None, '20C', 'APS1', None, None, 'PL', None, 12.5, 'Go', None, 'AP', None, 11.0,
               'Go', None, 11.75, 'Go'])
    ws.append(['DURAND', 'Lea', None, None, '21B', 'ADI', None, None, 'LOC', None, 9.0, 'NoGo', None, None, None,
               None, None, None, 9.0, 'NoGo'])
    ws.append(['DURAND', 'Lea', None, None, '21B', 'ADI', None, None, 'LOC', None, 13.0, 'Go', None, None, None,
               None, None, None, 13.0, 'Go'])
    ws.append(['BERNARD', 'Luc', None, None, '22A', 'BTWR', None, None, 'TWR', None, 15, 'Go', None, None, None,
               None, None, None, 15, 'Go'])
    ws.append(['PETIT', 'Jean', None, None, '19B', 'BTWR', None, None, 'TWR', None, 10, 'Go', None, None, None,
               None, None, None, 10, 'Go'])
    ws.append(['INCONNU', 'Max', None, None, '20A', 'BTWR', None, None, 'TWR', None, 10, 'Go', None, None, None,
               None, None, None, 10, 'Go'])
    wb.create_sheet('Synthese').append(['Total', 7])
    wb.save(p)


def notes(p):
    wb = Workbook(); ws = wb.active; ws.title = 'APS1'
    ws.cell(1, 1, 'Paul MARTIN'); ws.cell(1, 2, 'Class:\u00a0MCTA20C'); ws.cell(1, 3, 'Closed'); ws.cell(1, 4, 2)
    ws.cell(1, 5, 'PL: 12.5'); ws.cell(1, 6, 'AP: 11.5'); ws.cell(1, 7, 12.0)
    ws.cell(2, 1, 'DOMAINE'); ws.cell(2, 2, 'Degree:\u00a0APS1_V2')
    ws.cell(4, 1, 'Ana GARCIA'); ws.cell(4, 2, 'Class: MCTA23A'); ws.cell(4, 5, 'PL: 14'); ws.cell(4, 6, 'AP: 15')
    ws.cell(4, 7, 14.5); ws.cell(5, 2, 'Degree: APS1_V2')
    wb.save(p)


def sessions(p):
    wb = Workbook(); ws = wb.active; ws.title = 'APS1'
    r = 1
    for d, res in ((D(2026, 4, 14, 9), 'No Go'), (D(2026, 4, 20, 9), 'Go')):
        ws.cell(r, 1, 'MARTIN Paul'); ws.cell(r, 2, 'Class: MCTA20C'); ws.cell(r, 3, 'Closed'); ws.cell(r, 4, d)
        ws.cell(r, 6, res); ws.cell(r + 1, 2, 'Degree: APS1_V2'); ws.cell(r + 2, 2, 'Sim: Test (PL)')
        r += 4
    wb.save(p)


def astral(p):
    wb = Workbook(); ws = wb.active; ws.title = 'MCTA20C'
    ws.append(['', 'Apprenant', 'BTWR', 'BSURV', 'ADI', 'APS1', 'APS2', 'ACS1', 'ACS2'])
    ws.append(['', 'MARTIN Paul', '100% done', '100% done', '100% done', '100% done', '', '', ''])
    ws.append(['', 'ENAC', '', '', '', '2 remedial(s)', '', '', ''])
    ws.append(['', 'TOTAL', '', '', '', '', '', '', ''])
    wb.save(p)


def klif(p):
    wb = Workbook(); ws = wb.active; ws.title = '11A'
    h = ['PROMO', "Date d'entrée à l'ENAC", 'Nom du stagiaire', 'Prénom du stagiaire', 'Affectation', 'Zone',
         "Date d'affectation", 'Date obtention de la MU ZZ', 'Arrêt Formation', 'Nouveau centre', "Date d'arrivée",
         'Date obtention de la MU ZZ', 'Arrêt Formation', "Nombre d'échecs"]
    ws.append(h); ws.append(['11A', D(2011, 9, 1), 'MARTIN', 'Paul', 'LFBB', 'W', D(2014, 1, 1), D(2015, 1, 1)])
    ws = wb.create_sheet('20A'); ws.append(h)
    ws.append(['20A', D(2020, 9, 1), 'MARTIN', 'Paul', 'LFRR', 'W', D(2023, 1, 1), None, 'OUI', 'LFBB',
               D(2024, 1, 1), D(2025, 1, 1), 'NON', 1])
    ws = wb.create_sheet('21B'); ws.append(h)
    ws.append(['21B', D(2021, 9, 1), 'DURAND', 'Léa', 'LFPG', 'E', D(2023, 6, 1), D(2024, 6, 1), 'NON'])
    ws = wb.create_sheet('23A'); ws.append(h)
    ws.append(['23A', D(2023, 2, 1), 'GARCIA', 'Ana', 'LFMM', 'E', D(2025, 6, 1), None, None])
    ws.append([None, None, None, None, None, None, None, None, None, None, None, None, None, None])
    ws.append(['Moyenne', None, None, None, None, None, None, None, None])
    wb.save(p)


def ligne_de(ws, did, ordre):
    h = [c.value for c in ws[1]]
    for r in ws.iter_rows(min_row=2, values_only=True):
        v = dict(zip(h, r))
        if v['ID_DOSSIER_LOCAL'] == did and v['ORDRE_PERIODE'] == ordre:
            return v
    raise AssertionError('ligne absente')


def main():
    ok = ko = 0
    resultats = []
    with tempfile.TemporaryDirectory() as t:
        t = Path(t)
        f = {k: t / (k + '.xlsx') for k in ('reef', 'parcours', 'notes', 'sessions', 'astral', 'klif')}
        reef(f['reef']); parcours(f['parcours']); notes(f['notes']); sessions(f['sessions'])
        astral(f['astral']); klif(f['klif'])
        avant = {k: C.sha(v) for k, v in f.items()}
        src = {'PARCOURS': f['parcours'], 'NOTES': f['notes'], 'SESSIONS': f['sessions'],
               'ASTRAL': f['astral'], 'KLIF': f['klif']}
        out = t / 'resultats'
        fichier, bilan = C.consolider(f['reef'], src, out, journal=lambda *_: None)
        wb = load_workbook(fichier)
        ws = wb['PARCOURS_CONSOLIDES']
        h = [c.value for c in ws[1]]
        d1a, d1b, d2 = ligne_de(ws, 'D1', 1), ligne_de(ws, 'D1', 2), ligne_de(ws, 'D2', 1)
        d3, d5, d6 = ligne_de(ws, 'D3', 1), ligne_de(ws, 'D5', 1), ligne_de(ws, 'D6', 1)
        ex = [g.value for g in wb['A_EXAMINER']['A']]
        reef_h = colonnes_reef()

        def test(nom, cond):
            nonlocal ok, ko
            resultats.append(('OK  ' if cond else 'ECHEC ') + nom)
            ok, ko = ok + bool(cond), ko + (not cond)

        test('sources inchangees (REEF_rebuild compris)', all(C.sha(v) == avant[k] for k, v in f.items()))
        test('nouveau fichier distinct de REEF_rebuild', fichier.resolve() != f['reef'].resolve())
        test('toutes les colonnes REEF_rebuild conservees dans le meme ordre',
             [c for c in h if c in reef_h] == reef_h)
        test('8 lignes REEF conservees, population ouverte maintenue', sum(1 for r in ws.iter_rows(min_row=2,values_only=True) if str(r[h.index('ID_DOSSIER_LOCAL')]).startswith('D')) == 8 and ws.max_row-1 > 8)
        verification = bilan['verification_fichier_enregistre']
        test('fichier enregistre relu et toutes les cellules comparees',
             verification['statut'] == 'CONFORME_APRES_RELECTURE' and verification['lignes_verifiees'] == ws.max_row-1)
        test('les quatre sources ont des cellules reellement enregistrees',
             all(n > 0 for n in verification['cellules_non_vides_par_source'].values()))
        test('premiere cellule remplie de chaque source verifiable par adresse',
             all(ws[x['cellule']].value is not None for x in verification['premieres_cellules_par_source'].values()))
        test('aucune decision annoncee comme lue ou appliquee au premier lancement',
             bilan['decisions_lues'] == 0 and bilan['decisions_appliquees'] == 0)
        test('decision et commentaire REEF conserves', d1a['ISSUE_PERIODE'] == 'REDOUBLEMENT'
             and d1a['COMMENTAIRE'] == 'decision conservee')
        test('colonnes complementaires placees apres le Degree',
             h.index('BTWR_PARCOURS_NOTE') > h.index('BTWR_NOTE') and h.index('BTWR_PARCOURS_NOTE') < h.index('BSURV_CORRECTIVE'))
        test('Parcours BTWR rattache a la periode 1 (20A)', d1a['BTWR_PARCOURS_NOTE'] == 14.0)
        test('concordance avec la note REEF', d1a['BTWR_ETAT_SOURCES'].startswith('SOURCES:'))
        test('Parcours APS1 rattache a la reprise (20C) et pas a 20A',
             d1b['APS1_PL_PARCOURS_NOTE'] == 12.5 and d1a['APS1_PL_PARCOURS_NOTE'] is None)
        test('note Parcours APS1 par role AP', d1b['APS1_AP_PARCOURS_NOTE'] == 11.0)
        test('Notes Socrates rattachees (MCTA20C -> 20C)', d1b['APS1_SOCRATES_NOTE'] == 12.0
             and d1b['APS1_AP_SOCRATES_NOTE'] == 11.5)
        test('Sessions : nombre, NOGO, dates', d1b['APS1_SESSIONS_NB'] == 2 and d1b['APS1_SESSIONS_NOGO'] == 1
             and d1b['APS1_SESSIONS_PREMIERE'].date() == dt.date(2026, 4, 14))
        test('ASTRAL : remedials APS1 de la periode', d1b['APS1_ASTRAL_REMEDIALS'] == 2)
        test('deux passages ADI differents non fusionnes', d2['ADI_PARCOURS_NOTE'] == 'PLUSIEURS_VALEURS'
             and 'PLUSIEURS_VALEURS_MEME_PERIODE' in ex)
        test('homonymes meme promotion : aucune valeur, a examiner', d3['BTWR_PARCOURS_NOTE'] is None
             and 'HOMONYMES_MEME_PROMOTION' in ex)
        test('periode ambigue (meme promotion repetee) non choisie', d6['BTWR_PARCOURS_NOTE'] is None
             and 'PERIODE_AMBIGUE' in ex)
        test('eleve hors REEF integre selon population ouverte', any(r[h.index('NOM')]=='INCONNU' for r in ws.iter_rows(min_row=2,values_only=True)))
        test('KLIF : feuille 11A ignoree', bilan['lecture']['KLIF']['feuilles_anterieures_12A_ignorees'] == 1)
        test('KLIF une seule fois, sur la derniere periode', d1b['KLIF_PORTEE'] == 'SUIVI_KLIF_DU_DOSSIER'
             and d1a['KLIF_PORTEE'] == 'VOIR_PERIODE_2' and d1a['KLIF_AFFECTATION_1'] is None)
        test('KLIF dossier D1 : premiere affectation en echec', d1b['KLIF_ISSUE_AFFECTATION_1'] == 'ECHEC_ARRET_DE_FORMATION'
             and d1b['KLIF_AFFECTATION_1'] == 'LFRR')
        test('KLIF second centre et duree', d1b['KLIF_NOUVEAU_CENTRE'] == 'LFBB' and d1b['KLIF_MOIS_ARRIVEE_MU_2'] == 12.0)
        test('KLIF entree ENAC -> MU finale', d1b['KLIF_MOIS_ENTREE_ENAC_MU'] == 52.0)
        test('KLIF D2 qualifiee en 12 mois', d2['KLIF_ISSUE_AFFECTATION_1'] == 'QUALIFIEE'
             and d2['KLIF_MOIS_AFFECTATION_MU_1'] == 12.0)
        test('variante de nom (GARCIA / GARCIA LOPEZ) proposee, non utilisee', d5['KLIF_PORTEE'] == 'AUCUNE_LIGNE_KLIF_RACCORDEE'
             and d5['APS1_SOCRATES_NOTE'] is None)
        dec = out / C.DECISIONS_NOM
        test('fichier de decisions cree', dec.exists())
        wd = load_workbook(dec); wsd = wd.active; hd = [c.value for c in wsd[1]]
        n_prop = 0
        for row in wsd.iter_rows(min_row=2):
            v = {hd[i]: c for i, c in enumerate(row)}
            if v['ID_DOSSIER_LOCAL'].value == 'D5':
                v['DECISION'].value = 'OUI'; n_prop += 1
        wd.save(dec)
        test('propositions pour KLIF et Notes', n_prop == 2)
        # source vivante : ajout d'une ligne Parcours, relance
        wp = load_workbook(f['parcours']); wp['Parcours'].append(['GARCIA LOPEZ', 'Ana', None, None, '23A', 'BTWR',
            None, None, 'TWR', None, 16, 'Go', None, None, None, None, None, None, 16, 'Go']); wp.save(f['parcours'])
        fichier2, bilan2 = C.consolider(f['reef'], src, out, journal=lambda *_: None)
        ws2 = load_workbook(fichier2)['PARCOURS_CONSOLIDES']
        d5b = ligne_de(ws2, 'D5', 1)
        test('decision OUI appliquee a la relance (KLIF)', d5b['KLIF_LIEN'] == 'DECISION_VALIDEE'
             and d5b['KLIF_ISSUE_AFFECTATION_1'] == 'NON_QUALIFIEE_A_LA_DATE_DU_FICHIER')
        test('decision OUI appliquee a la relance (Notes)', d5b['APS1_SOCRATES_NOTE'] == 14.5)
        test('decisions lues distinguees des decisions effectivement retrouvees',
             bilan2['decisions_lues'] == 2 and bilan2['decisions_appliquees'] == 2)
        test('promotion differente non affectee meme si le dossier a une seule periode',
             C.Moteur.periode({'periodes': [{'ordre': 1, 'promo': '18A'}]}, {'promo': '18B'})
             == (None, 'PERIODE_SANS_VIS_A_VIS'))
        # Simule une perte de valeur entre le calcul et l'ecriture du fichier.
        attendues = [dict(zip(h, row)) for row in ws.iter_rows(min_row=2, values_only=True)]
        attendues[0]['BTWR_PARCOURS_NOTE'] = 19.0
        erreur_detectee = False
        try:
            C.verifier_classeur(fichier, h, attendues, reef_h, len(list(wb['PROVENANCE'].rows)) - 1)
        except C.Erreur:
            erreur_detectee = True
        test('divergence entre calcul et fichier enregistre detectee', erreur_detectee)
        test('donnee ajoutee a la source integree', d5b['BTWR_PARCOURS_NOTE'] == 16)
        wd2 = load_workbook(dec).active
        test('decisions conservees apres relance', sum(1 for r in wd2.iter_rows(min_row=2, values_only=True)
                                                        if r[-1] == 'OUI') == 2)
        test('fichier precedent conserve', fichier.exists() and fichier2 != fichier)
        test('bilan partageable sans nom', all('MARTIN' not in p.read_text(encoding='utf-8')
                                               for p in (out / 'PARTAGEABLE').glob('*.json')))
        test('REEF_rebuild toujours inchange', C.sha(f['reef']) == avant['reef'])
    print('\n'.join(resultats))
    print('%d tests reussis, %d echecs' % (ok, ko))
    return 0 if ko == 0 else 1


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc(); sys.exit(1)

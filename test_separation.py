# -*- coding: utf-8 -*-
"""Tests fictifs integres au consolidateur : aucune source reelle n'est modifiee."""
import tempfile,unittest,json
from pathlib import Path
from openpyxl import Workbook,load_workbook
from collections import defaultdict,Counter
import consolidation as C
import test_consolidation as H

class Separation(unittest.TestCase):
 def setUp(self):
  self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)
  self.reef=self.root/'reef.xlsx';H.reef(self.reef);self.source=self.root/'parcours.xlsx'
  self.head=['Nom','Prénom','EleveId','Promotion','Degré','Position 1','Note Position 1','Resultat 1','Position 2','Note Position 2','Resultat 2','Note','Résultat']
  self.dhead=['Nom','Prénom','Promotion','Degré','Passage','EvalId','RoleId','Rôle','Résultat','Rattrapage','Simulation','Source','Feuille source','Ligne source']
 def tearDown(self):self.tmp.cleanup()
 def synth(self, res='GO',note=12,degree='BTWR',role='TWR',pres='GO',ident='A'):
  return ['MARTIN','Paul',ident,'20A',degree,role,note,pres,None,None,None,note,res]
 def detail(self,res='NOGO',passage='0',degree='BTWR',role='Tour'):
  return ['MARTIN','Paul','20A',degree,passage,'E1','R1',role,res,'0','TEST','archive.xlsx','Simple',2]
 def write(self,synths,details):
  w=Workbook();ws=w.active;ws.title='Parcours';ws.append(self.head)
  for r in synths:ws.append(r)
  ws=w.create_sheet('Détails Eval');ws.append(self.dhead)
  for r in details:ws.append(r)
  w.save(self.source)
 def execute(self,synths,details):
  self.write(synths,details);before=[C.sha(self.source),C.sha(self.reef)]
  f,b=C.consolider(self.reef,{'PARCOURS':self.source},self.root/'resultats',journal=lambda *_:None)
  self.assertEqual(before,[C.sha(self.source),C.sha(self.reef)])
  w=load_workbook(f,read_only=True,data_only=True)
  def rows(ws):
   it=ws.iter_rows(values_only=True);h=next(it);return [dict(zip(h,r)) for r in it]
  out=rows(w['PARCOURS_CONSOLIDES']);p=rows(w['PROVENANCE']);e=rows(w['A_EXAMINER']);w.close()
  return next(r for r in out if r['ID_DOSSIER_LOCAL']=='D1' and r['ORDRE_PERIODE']==1),p,e,b,f
 def test_01_go_details_mixtes(self):
  r,p,e,b,_=self.execute([self.synth()],[self.detail('GO'),self.detail('NOGO','1')])
  self.assertEqual(r['BTWR_PARCOURS_RESULTAT'],'GO');self.assertEqual(len([x for x in p if x['COLONNE']=='OBSERVATION_DETAILLEE']),2)
  self.assertNotIn('PLUSIEURS_VALEURS_MEME_PERIODE',b['points_a_examiner_par_cause'])
 def test_02_nogo_non_remplace(self):
  r,*_=self.execute([self.synth('NOGO')],[self.detail('GO')]);self.assertEqual(r['BTWR_PARCOURS_RESULTAT'],'NOGO')
 def test_03_syntheses_contradictoires(self):
  r,p,e,b,_=self.execute([self.synth(),self.synth('NOGO')],[])
  self.assertEqual(r['BTWR_PARCOURS_RESULTAT'],'PLUSIEURS_VALEURS')
  cas=[x for x in e if x['CHAMP']=='BTWR_PARCOURS_RESULTAT'];self.assertEqual({x['VALEUR_CANDIDATE'] for x in cas},{'GO','NOGO'})
  self.assertEqual(len({x['ID_CAS'] for x in cas}),1)
 def test_04_synthese_absente(self):
  r,p,e,b,_=self.execute([],[self.detail('GO')]);self.assertIsNone(r['BTWR_PARCOURS_RESULTAT']);self.assertTrue(any(x['COLONNE']=='OBSERVATION_DETAILLEE' for x in p))
 def test_05_resultat_global_vide(self):
  r,*_=self.execute([self.synth(None)],[self.detail('GO')]);self.assertIsNone(r['BTWR_PARCOURS_RESULTAT'])
 def test_06_resultat_autre_conserve(self):
  r,*_=self.execute([self.synth('NON_EVALUE')],[self.detail('GO')]);self.assertEqual(r['BTWR_PARCOURS_RESULTAT'],'NON_EVALUE')
 def test_07_role_pl_distinct_global_et_detail(self):
  r,p,e,b,_=self.execute([self.synth('GO',degree='APS1',role='Cordon',pres='NOGO')],[self.detail('GO',degree='APS1',role='Cordon')])
  self.assertEqual(r['APS1_PARCOURS_RESULTAT'],'GO');self.assertEqual(r['APS1_PL_PARCOURS_RESULTAT'],'NOGO')
  self.assertEqual(next(x for x in p if x['COLONNE']=='OBSERVATION_DETAILLEE')['ROLE'],'PL')
 def test_08_repetitions_provenance_complete(self):
  r,p,e,b,_=self.execute([self.synth(ident=str(i)) for i in range(12)],[])
  self.assertEqual(r['BTWR_PARCOURS_RESULTAT'],'GO')
  self.assertEqual(len([x for x in p if x['COLONNE']=='BTWR_PARCOURS_RESULTAT']),12)
 def test_09_passages_identiques_non_fusionnes(self):
  details=[self.detail('GO',str(i)) for i in range(12)]
  r,p,e,b,_=self.execute([self.synth()],details)
  obs=[x for x in p if x['COLONNE']=='OBSERVATION_DETAILLEE'];self.assertEqual(len(obs),12)
  self.assertEqual({x['PASSAGE_TECHNIQUE_SOURCE'] for x in obs},{str(i) for i in range(12)})
 def test_10_tous_candidats_et_reperes(self):
  r,p,e,b,_=self.execute([self.synth('GO' if i%2 else 'NOGO',ident=str(i)) for i in range(14)],[])
  obs=[x for x in e if x['CHAMP']=='BTWR_PARCOURS_RESULTAT'];self.assertEqual(len(obs),14)
  self.assertEqual(len({x['REPERES'] for x in obs}),14)
  self.assertTrue(all(x['AIDE'] and x['NIVEAU']=='SYNTHESE_PARCOURS' for x in obs))
 def test_11_rattachement_incertain(self):
  row=self.detail();row[2]='19Z'
  r,p,e,b,_=self.execute([self.synth()],[row])
  obs=next(x for x in p if x['COLONNE']=='OBSERVATION_DETAILLEE');self.assertIsNone(obs['ORDRE_PERIODE']);self.assertEqual(obs['STATUT_RATTACHEMENT'],'PERIODE_SANS_VIS_A_VIS')
 def test_12_pas_synthese_artificielle(self):
  a=self.synth(None);b=self.synth();b[6]=None;b[11]=None
  r,p,e,bilan,_=self.execute([a,b],[])
  self.assertEqual(r['BTWR_PARCOURS_RESULTAT'],'PLUSIEURS_VALEURS')
  self.assertIn('SYNTHESE_COMPOSITE_NON_JUSTIFIEE',bilan['points_a_examiner_par_cause'])
 def test_13_relance_enrichissement(self):
  r,p,e,b,f=self.execute([self.synth()],[self.detail()])
  r2,p2,e2,b2,f2=self.execute([self.synth()],[self.detail(),self.detail('GO','1')])
  self.assertEqual(b['dossiers_consolides'],b2['dossiers_consolides']);self.assertNotEqual(f,f2)
  self.assertEqual(sum(x['COLONNE']=='OBSERVATION_DETAILLEE' for x in p2),2)
  self.assertTrue(b2['verification_fichier_enregistre']['provenance_cellules_verifiees'])
 def test_14_detail_seul_pas_nouveau_dossier(self):
  row=self.detail();row[0]='AUTRE';row[1]='FICTIF'
  r,p,e,b,_=self.execute([],[row]);self.assertEqual(b['dossiers_consolides'],6)
  self.assertTrue(any(x['COLONNE']=='OBSERVATION_DETAILLEE' and not x['ID_DOSSIER_LOCAL'] for x in p))
 def test_15_alias_et_filtre_preserves(self):
  self.assertEqual(C.degree_de('BASIC_TOWER'),'BTWR');self.assertEqual(C.degree_de('BASIC_ACC'),'BSURV')
  for pr in ['90A','99Z','00A','11D']:self.assertTrue(C.promotion_avant_12a(pr))
  self.assertFalse(C.promotion_avant_12a('12A'))
 def test_16_sans_eleveid(self):
  r,p,e,b,_=self.execute([self.synth(ident='A'),self.synth(ident='B')],[])
  self.assertEqual(r['BTWR_PARCOURS_RESULTAT'],'GO');self.assertNotIn('HOMONYMES_SOURCE',b['points_a_examiner_par_cause'])
 def test_18_registre_ancien_mis_ecart_sans_renumerotation(self):
  reg={'version':1,'dossiers':{'OLD':{'nom':'A'},'MIX':{'nom':'B'}},'periodes':{'OLD':{'90A':{'ordre':1}},'MIX':{'99A':{'ordre':1},'24A':{'ordre':2}}}}
  b=C.ecarter_anciennes_periodes_registre(reg)
  self.assertEqual(b['dossiers_mis_a_ecart_avant_12A'],1)
  self.assertEqual(reg['periodes']['MIX']['24A']['ordre'],2)
  self.assertEqual(reg['hors_perimetre_avant_12A']['dossiers']['OLD'],{'nom':'A'})
  self.assertEqual(C.ecarter_anciennes_periodes_registre(reg)['periodes_mises_a_ecart_avant_12A'],0)
 def test_19_alias_restitues(self):
  r,p,e,b,_=self.execute([self.synth(degree='BASIC_TOWER'),self.synth(degree='BASIC_ACC',role='CO')],[])
  self.assertEqual(r['BTWR_PARCOURS_NOTE'],12);self.assertEqual(r['BSURV_CO_PARCOURS_NOTE'],12)
 def test_20_cas_stable_apres_enrichissement(self):
  r,p,e,b,_=self.execute([self.synth(),self.synth('NOGO')],[])
  ids={x['ID_CAS'] for x in e if x['ID_CAS']}
  r,p,e,b,_=self.execute([self.synth(),self.synth('NOGO'),self.synth(ident='C')],[])
  self.assertEqual(ids,{x['ID_CAS'] for x in e if x['ID_CAS']})
 def test_21_contradiction_meme_observation_detail(self):
  r,p,e,b,_=self.execute([self.synth()],[self.detail('GO'),self.detail('NOGO')])
  self.assertEqual(r['BTWR_PARCOURS_RESULTAT'],'GO')
  self.assertIn('DETAIL_EVALUATION_DIVERGENTE',b['points_a_examiner_par_cause'])
  self.assertEqual({x['VALEUR_CANDIDATE'] for x in e if x['CHAMP']=='BTWR_DETAIL_RESULTAT'},{'GO','NOGO'})
 def test_17_integrite_trace_detectee(self):
  r,p,e,b,f=self.execute([self.synth()],[self.detail()])
  from openpyxl import load_workbook
  ww=load_workbook(f);ww['PROVENANCE'].cell(2,4,'MODIFIE');ww.save(f);ww.close()
  exp=[[row.get(c) for c in ['ID_DOSSIER_LOCAL','ORDRE_PERIODE','COLONNE','VALEUR','SOURCE','REPERES (fichier!feuille:ligne)','REGLE']+C.COLONNES_TRACABILITE] for row in p]
  with self.assertRaises(C.Erreur):C.verifier_tracabilite(f,exp,{})

def main():
 result=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(Separation))
 return 0 if result.wasSuccessful() else 1
if __name__=='__main__':raise SystemExit(main())

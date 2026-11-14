# Livraison logicielle - consolidateur existant, lot C1 a C8

**Revision livree : `2026-11-13-lot-c1-c8-r1-provenance`**  
**Reference effectivement controlee : `2026-11-13-controle-couverture`**

## Nature de la livraison

Cette archive contient le **code Python executable complet du consolidateur existant**, son lanceur `LANCER_WINDOWS.bat`, ses dependances, ses tests, le mode d'emploi, le manifeste SHA256 et les preuves des essais synthetiques. Elle n'est pas un patch a installer sur une ancienne version ni une nouvelle application. Aucune source reelle, decision nominative ou registre reel n'est inclus.

Les 12 fichiers du manifeste d'origine ont ete lus completement et leurs empreintes SHA256 ont toutes concorde avec `MANIFESTE_BASE.json` avant modification. Le lanceur et `requirements.txt` de la livraison ont **exactement la meme empreinte** que les fichiers d'origine.

## C1-C8 - Modifications appliquees

- **C1** : exclure explicitement les feuilles auxiliaires REEF `REPRISES_A_PRECISER`, `LIRE_MOI`, `_META_LOCAL` et Notes `AUDIT_DEGREES`; contenu conserve en audit brut; feuille inconnue non exclue.
- **C2** : exclusions des promotions numeriques P dans KLIF, Notes et Sessions, avant rattachement/creation; exclusion avant 12A limitee a KLIF; familles explicitement hors cible uniquement; ancien registre conserve.
- **C3** : detection de `Moyenne` normalise, dans toute colonne KLIF, sans identite; identite presente : conflit a examiner, non retenu comme eleve; aucun nombre de moyennes impose.
- **C4** : regle explicite, strictement limitee aux deux dossiers 18A autorises et aux Degres ACS1/ACS2, avec conservation du libelle brut et absence de repli arbitraire.
- **C5** : presence remedial ASTRAL uniquement sur mention explicite au Degree utile; ABES_ et TOTAL hors champ; pourcentage ou vide n'est pas un NON ni un zero derive.
- **C6** : exposition de l'Origine KLIF brute et de la duree source `Nombre de mois de qualification ENAC exclus` en mois, distincte des durees recalculees; A/Matricule non utilises pour identite; prolongations non interpretees.
- **C7** : prise en charge de la note reconstituee fournie pour `Note=NOGO`, distincte du resultat, sans nouveau calcul; origine/statut traces; role 2 neutralise uniquement si le brut l'indique et reste conserve.
- **C8** : niveaux DETAIL_EVALUATION prioritaires sur des traces d'origine; omission et absence de rattachement detectees/conservees; pas de conversion de detail en synthese.

## Tests executes et limites de preuve

`python consolidation.py --autotest` : **105 controles reussis, 0 echec**, dont 40 historiques, 21 separation, 25 couverture (un test historique mis a jour selon C2), 15 C1-C8, et 4 controles de regression de provenance. Test C4 : 6 syntheses, 24 details fictifs sur les deux dossiers, repartition 20 en periode 1 et 10 en periode 2. Tests C7 : 27 **observations fictives** (ce n'est pas une verification des 27 observations reelles). Une relecture des classeurs generes et les comparaisons des empreintes des sources sont comprises dans les tests. Syntaxe Python 3.10 verifiee; execution de recette sous Linux avec Python `3.13.5`.

**Non executes et non certifies** : lancement sous Windows, affichage Tkinter, performance et justesse sur les classeurs nominatifs, compteurs des fichiers reels, validite de tous les rattachements des dossiers 18A reels. Les donnees reelles restent sur le poste utilisateur.

## Mise en service et retour arriere

1. **Conserver** le dossier de l'application de reference; extraire **tous les fichiers** de ce ZIP dans un **nouveau dossier**. Ne pas remplacer uniquement `consolidation.py`.
2. Sur Windows, utiliser le lanceur existant `LANCER_WINDOWS.bat --tests` pour controler la suite. Puis `LANCER_WINDOWS.bat` pour selectionner les sources habituelles et **le meme dossier de resultats** afin de reutiliser `DECISIONS_CORRESPONDANCES.xlsx` et `REGISTRE_POPULATION_LOCAL.json`. Avant la relance, effectuer une **copie de sauvegarde** locale de ces deux fichiers (et du dossier de resultats). Les sources sont lues, un nouveau classeur de consolidation est cree.
3. Controler les variations de dossiers, periodes, exclusions, moyennes KLIF, ASTRAL, notes reconstituees, les deux dossiers 18A, les traces et alertes; ne pas assimiler une baisse due a une exclusion autorisee a une perte de donnees.
4. En cas d'anomalie, ne pas valider la nouvelle production et reprendre l'ancien dossier applicatif avec la sauvegarde des decisions/registre.

Consulter `RESULTATS_TESTS.txt`, `BILAN_TESTS.json`, `MATRICE_RECEPTION_C1_C8.csv`, `MODIFICATIONS_C1_C8.diff`, `MANIFESTE.json` et `MANIFESTE_BASE.json`. Les derniers arbitrages sont documentes dans `SPECIFICATION_C1_C8.pdf` fourni par l'utilisateur.

## Correctif R1 de reception apres erreur « Provenance differente apres sauvegarde »

- Constat : un test local reproduit que des valeurs Python longues ou tres precises peuvent etre modifiees lors de l'ecriture puis relecture Excel. Le controle d'origine comparait strictement les types et valeurs attendus, ce qui peut generer un echec de verification. **Aucune preuve ne permet d'attribuer avec certitude l'erreur de la ligne reelle 32275 a ce cas** sans les donnees locales.
- Correction bornee : conversion en **texte exact** uniquement des valeurs de provenance dont la precision numerique ou temporelle risque d'etre alteree ; autres types de provenance conserves. La comparaison apres sauvegarde reste stricte sur la valeur effectivement preparee pour Excel.
- Confidentialite : en cas de nouvel echec, le message indique ligne, colonne, champ et types, jamais les valeurs source.
- Surete : test de divergence reelle toujours bloquant ; cellule texte trop longue refusee plutot que tronquee ; pas de suppression du controle.
- Tests : 4 nouveaux cas de regression dans `test_lot_c1_c8.py`, en plus des 101 controles C1-C8 precedents ; test de charge supplementaire hors suite sur 32 274 lignes, avec valeurs sensibles a la ligne Excel 32 275 (`ESSAI_PROVENANCE_32275.txt`).
- Limites : pas de rejeu Windows ni de verification sur les classeurs nominatifs ; un nouvel echec reste possible s'il a une autre cause. L'application, les decisions et le registre reels ne sont ni inclus ni modifies.

**Installation : extraire le ZIP complet dans un dossier neuf. Ne pas copier seulement le module Python.** Le lanceur et le fichier de dependances sont bit a bit identiques a la livraison C1-C8 precedente.

## Correctif R1 de reception apres erreur « Provenance differente apres sauvegarde »

- Constat : un test local reproduit que des valeurs Python longues ou tres precises peuvent etre modifiees lors de l'ecriture puis relecture Excel. Le controle d'origine comparait strictement les types et valeurs attendus, ce qui peut generer un echec de verification. **Aucune preuve ne permet d'attribuer avec certitude l'erreur de la ligne reelle 32275 a ce cas** sans les donnees locales.
- Correction bornee : conversion en **texte exact** uniquement des valeurs de provenance dont la precision numerique ou temporelle risque d'etre alteree ; autres types de provenance conserves. La comparaison apres sauvegarde reste stricte sur la valeur effectivement preparee pour Excel.
- Confidentialite : en cas de nouvel echec, le message indique ligne, colonne, champ et types, jamais les valeurs source.
- Surete : test de divergence reelle toujours bloquant ; cellule texte trop longue refusee plutot que tronquee ; pas de suppression du controle.
- Tests : 4 nouveaux cas de regression dans `test_lot_c1_c8.py`, en plus des 101 controles C1-C8 precedents ; test de charge supplementaire hors suite sur 32 274 lignes, avec valeurs sensibles a la ligne Excel 32 275 (`ESSAI_PROVENANCE_32275.txt`).
- Limites : pas de rejeu Windows ni de verification sur les classeurs nominatifs ; un nouvel echec reste possible s'il a une autre cause. L'application, les decisions et le registre reels ne sont ni inclus ni modifies.

**Installation : extraire le ZIP complet dans un dossier neuf. Ne pas copier seulement le module Python.** Le lanceur et le fichier de dependances sont bit a bit identiques a la livraison C1-C8 precedente.

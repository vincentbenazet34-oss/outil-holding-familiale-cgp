# Prototype Holding Familiale CGP – v16

Application Streamlit d'aide au diagnostic préventif dans le cadre d'une stratégie de holding familiale.

## Lancer l'application

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Nouveautés v16

- Renforcement de la prise en compte du **besoin de revenus réguliers** dans l'analyse.
- Le besoin de revenus influence désormais plus clairement :
  - le risque de manque de liquidité ;
  - le risque de fragilisation du conjoint ou des proches ;
  - le risque de blocage de gouvernance lié aux dividendes ;
  - le risque de dépendance au patrimoine professionnel.
- Les questions relatives à la politique de distribution et à la prévoyance apparaissent dans davantage de situations pertinentes.
- Les questions liées au conjoint sont prises en compte dès qu'un conjoint est présent, même si l'objectif de protection n'est pas sélectionné comme prioritaire.
- Les questions de dialogue familial sont posées plus largement lorsqu'il existe plusieurs enfants.
- Les informations manquantes signalent maintenant explicitement les points à vérifier lorsque le besoin de revenus est moyen ou élevé.

## Principe

L'outil ne fournit pas une recommandation juridique ou fiscale définitive. Il aide le CGP à structurer son raisonnement :

Objectifs → Questions clés → Signaux d'alerte → Risques → Outils à étudier → Actions préventives → Suivi.

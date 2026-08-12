# Vision Amah — Prospection

Outil de prospection pour [Vision Amah](https://vision-amah.vercel.app), une activité de création de sites web pour les commerces qui n'en ont pas.

Le script `prospect_finder.py` cherche sur OpenStreetMap les commerces de Conakry (Guinée) qui ont un numéro de téléphone mais aucun site web, et génère pour chacun un kit de prospection prêt à l'emploi : un message WhatsApp personnalisé et une page de démo qui montre concrètement à quoi pourrait ressembler leur site.

**Aucun envoi n'est automatisé.** L'envoi des messages reste manuel : automatiser WhatsApp expose le numéro à un bannissement.

## Ce que ça fait

- Interroge OpenStreetMap (API Overpass) pour trouver les commerces sans site web
- Nettoie et dédoublonne les résultats (numéros de téléphone, établissements saisis plusieurs fois)
- Classe chaque commerce par secteur (restauration, beauté, mode, santé, hébergement, auto, etc.) et adapte le message et la mise en page de la démo en conséquence
- Génère une page de démo personnalisée par prospect (`demos/`), avec une vraie photo/un vrai menu quand OpenStreetMap en référence
- Produit un CSV, une page interactive (`prospects.html` — recherche, filtres, statuts, relances, statistiques) et un PDF
- Garde en mémoire tous les prospects déjà trouvés d'une exécution à l'autre
- Journalise les messages envoyés dans un classeur Excel dédié, séparé du suivi manuel existant

## Installation

```
pip install requests fpdf2 openpyxl
```

## Utilisation

```
python prospect_finder.py            # Conakry par défaut
python prospect_finder.py "Kindia"   # ou toute autre ville reconnue par OpenStreetMap
```

Le script peut aussi être relancé automatiquement chaque semaine via `lancer_hebdo.bat` (voir la tâche planifiée Windows "Vision Amah - Prospection hebdomadaire").

## Documentation technique

Voir [`CLAUDE.md`](CLAUDE.md) pour l'architecture détaillée, la référence des fonctions, et l'historique des bugs corrigés.

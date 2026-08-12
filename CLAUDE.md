# CLAUDE.md

Ce fichier sert de guide à Claude Code (claude.ai/code) pour travailler dans ce dépôt.

## Description

Un outil Python en un seul fichier (`prospect_finder.py`) pour Vision Amah, une activité freelance de création de sites web : il trouve sur OpenStreetMap les commerces de Conakry (Guinée) qui ont un numéro de téléphone mais pas de site web, et transforme chacun en un kit de prospection WhatsApp prêt à l'emploi (message personnalisé + page de démo personnalisée). Il n'envoie jamais rien automatiquement — automatiser WhatsApp expose le numéro à un bannissement, donc l'envoi reste manuel.

## Commandes

```
pip install requests fpdf2 openpyxl
python prospect_finder.py            # Conakry par défaut
python prospect_finder.py "Kindia"   # ou tout autre nom de ville reconnu par OSM
```

Il n'y a pas de suite de tests. Pour vérifier un changement :
```
python -m py_compile prospect_finder.py   # vérification rapide de la syntaxe
python prospect_finder.py                 # exécution réelle contre l'API Overpass en direct — pas de mocks
```
Une exécution réelle est le seul moyen de tester la majeure partie du pipeline (requête Overpass, extraction, dédoublonnage, rendu). Pour forcer les prospects à être retraités comme "nouveaux" lors d'un test (par ex. après avoir modifié les gabarits de message/démo), réinitialise le fichier d'accumulation : `echo {} > %LOCALAPPDATA%\ProspectionVisionAmah\deja_contactes.json`.

`lancer_hebdo.bat` (`cd /d "%~dp0"` puis lance le script) est ce que la tâche planifiée Windows "Vision Amah - Prospection hebdomadaire" (tous les lundis 8h) invoque réellement — teste toujours en passant par le dossier de travail qu'il fixe, pas seulement depuis un shell déjà positionné avec `cd`, car des bugs liés au dossier de travail se sont déjà produits ici (voir les pièges plus bas).

## Architecture

Tout vit dans `prospect_finder.py` ; `main()` en bas du fichier est le pipeline dans l'ordre : fusionner un éventuel export du journal en attente → interroger Overpass → extraire → dédoublonner → fusionner avec la base des fiches connues → générer les sorties → sauvegarder la base → ouvrir automatiquement le PDF.

### Le modèle d'accumulation (ne pas régresser là-dessus)

`prospects.csv` / `prospects.html` / `prospects.pdf` sont toujours régénérés à partir de **l'ensemble accumulé complet** de tous les prospects jamais trouvés (`charger_fiches_connues()` / `sauvegarder_fiches_connues()`, stocké comme un dict `id_osm -> fiche` dans `%LOCALAPPDATA%\ProspectionVisionAmah\deja_contactes.json`), jamais seulement à partir des nouveautés de l'exécution en cours. C'était un vrai bug à l'origine : une version précédente générait les sorties à partir de `nouvelles` uniquement, donc l'exécution hebdomadaire planifiée effaçait silencieusement chaque lundi tous les prospects trouvés précédemment (avec leur statut/relance/stats côté navigateur). Si tu touches à `main()`, continue de régénérer à partir du dict fusionné `connues`, jamais à partir de `nouvelles` seul. Le texte du message est aussi reconstruit pour chaque fiche à chaque exécution (pas seulement les nouvelles) pour que les anciennes profitent des améliorations de gabarit — ce qui veut dire que le message affiché pour un prospect peut changer après avoir été marqué "contacté" si `MESSAGES` a été modifié depuis.

Un ancien format de ce fichier JSON (une simple liste d'identifiants, sans détail des fiches) est détecté et écarté plutôt que migré — `charger_fiches_connues()` repart simplement de zéro s'il tombe sur une liste au lieu d'un dict.

### Le système de personnalisation par famille

Tout ce qui concerne la façon dont un prospect est démarché est indexé par "famille" (une parmi `restauration, beaute, mode, informatique, alimentation, sante, hebergement, auto, services, loisirs, generique`), assignée par `famille()` à partir de la valeur du tag OSM via le dict d'appartenance `FAMILLES`. Une famille pilote, indépendamment, via des dicts séparés utilisant les mêmes clés :
- `MESSAGES` — le texte du message WhatsApp (doit décrire ce que la démo jointe montre réellement — ça a déjà été synchronisé une fois volontairement, ne pas laisser les deux diverger à nouveau)
- `MISE_EN_PAGE_GROUPE` — quelle variante de mise en page pour la démo : `"menu"` (liste avec prix, restauration), `"rdv"` (liste de prestations réservables, beaute/sante/hebergement/auto/services), `"catalogue"` (grille de produits, mode/alimentation/informatique/generique), ou `"agenda"` (liste d'événements, loisirs)
- `ELEMENTS_GROUPE` / `SECTIONS_GROUPE` — les 3 noms d'exemple et le titre de section utilisés dans cette mise en page
- `THEME_GROUPE` — thème clair par défaut ; loisirs/hebergement/auto reçoivent plutôt la palette sombre `VARS_THEME["sombre"]` (esthétique nightlife/hôtellerie haut de gamme/garage)
- `COULEUR_GROUPE`, `ICONES_GROUPE`, `TAGLINES`, `LABELS_GROUPE` — couleur d'accent, emoji, accroche du hero, libellé de catégorie

`corps_demo()` utilise `MISE_EN_PAGE_GROUPE` pour choisir quel bloc HTML construire. Ajouter une nouvelle famille implique de toucher tous les dicts ci-dessus plus `FILTRES_OSM` (les filtres de la requête Overpass) — en oublier un et cette famille retombe silencieusement sur `"generique"`.

### Pages de démo (`demos/*.html`)

`ecrire_demos()` génère un mockup HTML statique et autonome par prospect via `rendre_demo()`, classé sous `demos/<slug-du-nom>.html` (ou `%LOCALAPPDATA%\...\demos` si le dossier du projet n'est pas accessible en écriture — voir `dossier_demos_ecrivable()`). `image_osm()` récupère une **vraie** photo depuis le tag OSM `image=` ou une référence `wikimedia_commons=File:...` (via la redirection `Special:FilePath` de Commons) quand il en existe une, avec repli sur une galerie de type placeholder sinon — ça ne touche volontairement jamais aux photos Google Maps, puisque ça demanderait une clé API Google payante, ce qui va à l'encontre de la contrainte du projet "aucune clé API, aucune facturation". Même logique pour les tags `website:menu`/`contact:menu` (un lien de menu déjà existant) — attention, le tag OSM `menu=` seul est une description du type de cuisine, pas une URL, et ne doit pas être traité comme telle.

### Les sorties et où elles vivent

- `prospects.csv` / `.html` / `.pdf` : écrits dans le dossier du projet, ou là où atterrit le repli à 3 niveaux de `ouvrir_en_ecriture()` (voir plus bas) si c'est bloqué.
- `demos/*.html` : un par prospect, lié depuis les trois sorties ci-dessus.
- `Journal des envois - Vision Amah.xlsx` : un journal des messages envoyés **séparé, géré par le script** (openpyxl), construit par `fusionner_journal_envois()` à partir de `journal_envois.csv` (que le bouton "Exporter le journal" de `prospects.html` produit côté navigateur, puisqu'une page statique ne peut pas écrire de fichier toute seule — il réexporte *tout* l'historique des statuts à chaque fois, et la fusion dédoublonne sur `(telephone, date)` pour que les réimports ne dupliquent pas les lignes).
- `Vision Amah - Suivi Prospection.xlsx` : **le fichier de suivi manuel de l'utilisateur, préexistant** (prospects différents, contient des dessins/commentaires Excel). Le script ne doit jamais y écrire — un aller-retour openpyxl risquerait de corrompre cette mise en forme/ces commentaires, et l'utilisateur a explicitement choisi un fichier dédié séparé plutôt que d'intégrer à celui-ci.

### L'état côté navigateur de `prospects.html`

Pas de backend — le statut de chaque prospect (`a_contacter`/`contacte`/`repondu`/`signe`/`pas_interesse`, chacun avec un horodatage) vit dans le `localStorage` du navigateur, indexé par `${nom}|${telephone}`. Un prospect marqué `contacte` depuis 5 jours ou plus sans changement de statut affiche un badge "à relancer". Le panneau de stats et l'export du journal dérivent tous les deux de ce même état côté client — il n'existe aucune source de vérité côté serveur pour "est-ce que ça a vraiment été envoyé", seulement ce que l'utilisateur coche.

### Le mécanisme de repli à l'écriture

`ouvrir_en_ecriture()` / `chemin_de_repli()` implémentent un repli à 3 niveaux utilisé partout dans le script : essayer le chemin cible → retenter avec un nom de fichier horodaté dans le même dossier (gère un fichier verrouillé, ouvert dans Excel) → se replier sur `%LOCALAPPDATA%\ProspectionVisionAmah\` (gère le Contrôle d'accès aux dossiers de Windows Defender qui bloque complètement l'écriture dans Bureau/Documents). `ecrire_pdf()` et `fusionner_journal_envois()` réimplémentent chacun cette même logique à 3 niveaux indépendamment plutôt que d'appeler `ouvrir_en_ecriture()` — duplication connue, pas encore factorisée.

### Pièges spécifiques à Windows

- L'appel à `sys.stdout.reconfigure(encoding="utf-8", ...)` près du début du fichier est structurant : sans lui, l'encodage `cp1252` par défaut de la console Windows plante sur les emojis/tirets cadratins utilisés dans les propres `print()` du script — c'est déjà arrivé, ne pas le retirer.
- Le dossier de travail compte : lancer via Win+R ou certaines configurations de tâche planifiée place par défaut `C:\Windows\system32` comme dossier courant, ce qui casse les chemins de sortie relatifs du script. `lancer_hebdo.bat` s'en protège avec `cd /d "%~dp0"` ; si tu ajoutes un autre point d'entrée, fais pareil.
- Le `input()` tout à la fin du bloc `__main__` (garde ouverte la fenêtre de console en cas de double-clic) est protégé par `sys.stdin.isatty()` **et** encadré par un `try/except EOFError`, car `isatty()` seul n'est pas fiable dans tous les terminaux susceptibles de lancer ce script (y compris certains shells d'agents/CI) — garder les deux protections si tu touches à ce bloc.

### Gestion des numéros de téléphone

`normaliser_telephones()` cible spécifiquement les numéros guinéens (format de sortie `224XXXXXXXXX` pour les liens `wa.me`). Elle découpe d'abord sur `/;,`, et ne se replie sur un découpage par espaces que si le morceau entier échoue à être analysé directement — cet ordre compte, car un numéro unique correctement formaté avec des espaces internes (ex. `"+224 611 76 85 52"`, le format utilisé dans `SIGNATURE`) doit être analysé par le chemin direct et ne jamais atteindre le repli par espaces, sous peine d'être débité en fragments inexploitables.

## Bugs trouvés et corrigés (revue du 2026-08-11)

Une revue de code complète (`code-review`, effort élevé) a été menée sur tout le fichier avant la création de ce CLAUDE.md. Chaque correctif ci-dessous a été vérifié par un test unitaire ciblé ou une exécution réelle avant/après, pas seulement accepté sur la foi du rapport.

1. **🔴 Critique — les anciens prospects disparaissaient à chaque lancement.** `main()` appelait `ecrire_csv/html/pdf(nouvelles)` (seulement les prospects jamais vus), donc chaque exécution écrasait `prospects.csv/html/pdf` avec uniquement les nouveautés du jour — la tâche planifiée du lundi aurait silencieusement effacé tous les prospects précédents (et leur statut/relance/stats) chaque semaine. Corrigé en stockant les fiches complètes (pas seulement leurs identifiants) dans `deja_contactes.json` via `charger_fiches_connues()`/`sauvegarder_fiches_connues()`, et en régénérant toujours les sorties à partir de l'ensemble accumulé complet. Vérifié : deux exécutions consécutives affichent désormais toutes les deux "300 prospects au total" au lieu que la seconde tombe à 0. Voir la section "Le modèle d'accumulation" ci-dessus.
2. **🟠 Injection possible via un nom de commerce.** Les noms de prospects viennent d'OSM (base modifiable par n'importe qui) et étaient insérés tels quels dans un bloc `<script>` JSON dans `ecrire_html()` — un nom contenant littéralement `</script>` aurait cassé la page. Corrigé avec `.replace("</", "<\\/")` sur le JSON avant insertion.
3. **🟡 Guillemets courbes silencieusement supprimés au lieu d'être convertis (PDF).** `texte_pdf()` avait une table de conversion censée transformer guillemets/tirets courbes en équivalents ASCII avant l'encodage Latin-1 requis par fpdf2 — mais une erreur de copier-coller comparait des caractères déjà identiques (ne faisait rien). Un nom copié depuis Word avec une vraie apostrophe courbe se retrouvait donc amputé dans le PDF. Corrigé avec les bons points de code Unicode, vérifiés explicitement via `ord()` (une confusion guillemet droit/courbe est quasi invisible à l'œil dans certains rendus — c'est littéralement comme ce bug est apparu la première fois).
4. **🟡 Numéros séparés par un simple espace parfois perdus.** `normaliser_telephones()` ne découpait que sur `/;,`, pas sur l'espace seul — un champ OSM du type `"620585080 664243429"` (sans slash) était rejeté en bloc comme "sans numéro exploitable". Corrigé avec un repli qui ne se déclenche que si l'analyse directe du morceau entier échoue, pour ne jamais casser un numéro correctement formaté avec des espaces internes.
5. **🟡 Dédoublonnage téléphone incomplet.** `dedupliquer()` ne comparait que le premier numéro de chaque fiche — deux fiches partageant un numéro qui n'était pas en première position sur l'une d'elles n'étaient pas détectées comme doublons. Corrigé pour comparer sur l'ensemble des numéros de chaque fiche.

**Non corrigés (qualité de code, pas des bugs de comportement)** :
- Logique de repli à l'écriture dupliquée indépendamment dans `ouvrir_en_ecriture()`, `ecrire_pdf()` et `fusionner_journal_envois()` au lieu d'être factorisée.
- `dedupliquer()` a une passe en O(n²) (scan linéaire + recalcul de `cle_nom()` à chaque comparaison) — sans impact à l'échelle actuelle (centaines de fiches), mais coûterait cher sur un volume bien plus grand.
- `zone_la_plus_proche()` réimplémente la même formule de distance approximative que `distance_km()` au lieu de l'appeler.

## Référence des fonctions

| Fonction | Rôle |
|---|---|
| `famille(valeur_osm)` | Classe une valeur de tag OSM (shop/amenity/tourism/office/craft/leisure) dans une des 11 familles via `FAMILLES`, ou `"generique"` par défaut. |
| `zone_la_plus_proche(lat, lon)` | Trouve le repère de `REPERES` le plus proche de coordonnées données, pour assigner un nom de quartier lisible. |
| `distance_km(a_lat, a_lon, b_lat, b_lon)` | Distance approximative en km entre deux points (projection équirectangulaire simplifiée). |
| `_numero_depuis_chiffres(chiffres)` | Valide/complète une suite de chiffres en numéro guinéen `224XXXXXXXXX`, ou renvoie `None`. |
| `normaliser_telephones(brut)` | Extrait un ou plusieurs numéros valides d'un champ téléphone OSM brut (gère `/;,` et le repli par espace). |
| `image_osm(tags)` | Récupère l'URL d'une vraie photo depuis les tags OSM `image` ou `wikimedia_commons` (jamais Google Maps). |
| `cle_nom(nom)` | Normalise un nom (sans accents, sans casse, alphanumérique) pour le dédoublonnage par proximité. |
| `interroger_overpass(ville)` | Construit et envoie la requête Overpass QL, avec retries et bascule entre serveurs miroirs. |
| `extraire(elements)` | Transforme les éléments OSM bruts en fiches exploitables (nom, numéros, coordonnées, famille…), rejette celles sans nom/numéro valide. |
| `dedupliquer(fiches)` | Élimine les doublons : même nom + proximité géographique, puis numéro de téléphone partagé. |
| `score(f)` | Score de priorité — email et/ou réseau social déjà renseignés sur OSM = fiche "haute priorité". |
| `charger_fiches_connues()` | Charge la base accumulée (dict `id_osm -> fiche`) de tous les prospects déjà rencontrés lors d'exécutions précédentes. |
| `sauvegarder_fiches_connues(fiches_par_id)` | Sauvegarde cette base accumulée sur disque. |
| `chemin_de_repli(chemin)` | Calcule un nom de fichier de repli horodaté, pour le cas où le fichier cible est verrouillé. |
| `ouvrir_en_ecriture(chemin, **kwargs)` | Ouvre un fichier en écriture avec repli à 3 niveaux (cible → nom horodaté → dossier applicatif). |
| `fusionner_journal_envois()` | Absorbe l'export du journal (`journal_envois.csv`) dans le classeur Excel dédié, en dédoublonnant. |
| `ecrire_csv(fiches, chemin)` | Génère `prospects.csv` (classement/sauvegarde, colonnes incluant le message et le lien de démo). |
| `ecrire_html(fiches, chemin)` | Génère `prospects.html`, la page interactive (recherche, filtres, statuts, relances, stats, export du journal). |
| `corps_demo(groupe, sections, elements, wa)` | Construit le bloc HTML central de la page de démo selon la mise en page de la famille (menu/rdv/catalogue/agenda). |
| `photo_hero_html(image, icone)` | Construit le HTML de la photo héro de la démo — vraie photo OSM si disponible, sinon icône sur fond dégradé. |
| `galerie_html(image)` | Construit la galerie photo de la démo — vraie photo + emplacements vides, ou 3 emplacements vides. |
| `lien_menu_html(menu_url)` | Construit le bouton "Voir le menu déjà en ligne" si un lien de menu existe sur OSM. |
| `rendre_demo(p)` | Assemble la page de démo HTML complète pour un prospect (thème, couleurs, contenu, photo, menu). |
| `dossier_demos_ecrivable(base)` | Choisit un dossier accessible en écriture pour stocker les pages de démo (projet, ou dossier applicatif en repli). |
| `ecrire_demos(fiches)` | Génère toutes les pages de démo et attache le chemin du fichier à chaque fiche (`p["demo"]`). |
| `texte_pdf(s)` | Nettoie un texte pour l'encodage Latin-1 requis par fpdf2 (conversion guillemets/tirets courbes, suppression des emojis). |
| `ecrire_pdf(fiches, chemin)` | Génère `prospects.pdf`, le format ouvert automatiquement en fin d'exécution. |
| `main(ville)` | Orchestre tout le pipeline : journal → Overpass → extraction → dédoublonnage → fusion avec l'historique → génération des sorties → sauvegarde. |

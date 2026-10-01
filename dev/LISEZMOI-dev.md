# MapTools · Développement

Ce dossier sert à faire évoluer MapTools. Il n'est pas nécessaire pour l'utiliser.

## Sources

| Fichier | Contenu |
|---|---|
| `src/MapTools.src.html` | Structure HTML et JavaScript de l'outil, avec les marqueurs `__CSS__`, `__DEMO__`, `__PAPA__` et `__VIS__` |
| `src/styles.css` | Feuille de style complète (thèmes clair et sombre, impression) |
| `src/demo-data.js` | Données fictives de l'exemple et de l'avant / après |
| `build.py` | Assemble `MapTools.html` (version locale) et `out/MapTools-en-ligne.html` (bibliothèques sur CDN, exemple chargé d'office) |

Ne pas modifier `MapTools.html` directement : modifier les sources, puis, depuis le dossier `MapTools` :

```
python dev/build.py
```

Le script vérifie la syntaxe JavaScript si Node.js est installé. Après toute modification, régénérer `SHA256SUMS.txt` (commande dans `docs/REFERENCE.md`).

## Tests

**Interface** (Python, Playwright et Chromium) :

```
pip install playwright
python -m playwright install chromium
python dev/tests/test_interface.py
```

Charge l'exemple et vérifie la carte, les 13 vues rapides, la réinitialisation des filtres, les onglets de détails, le chemin d'héritage, le rapport et ses parties, l'acceptation d'un risque, le plan d'action CSV, les listes d'export, l'avant / après, la restauration après rafraîchissement et l'effacement. Code de sortie 1 en cas d'échec.

**Scripts PowerShell** (PowerShell 7, sans tenant) :

```
.\dev\tests\Test-Scripts.ps1
```

Exécute les quatre exports avec des cmdlets Graph, Active Directory, Exchange Online et SharePoint simulées (`tests/mocks`) et vérifie que les CSV attendus sont produits. Ces tests valident la logique, pas le comportement réel des API : un passage sur un vrai tenant reste nécessaire après toute modification des scripts.

## Conventions

- Scripts PowerShell enregistrés en **UTF-8 avec BOM** (sinon Windows PowerShell 5.1 lit mal les accents).
- Ne jamais nommer une variable locale `vis` dans le JavaScript : elle masque la bibliothèque vis-network.
- Une variable PowerShell ne doit pas porter le nom d'un paramètre du script (PowerShell ignore la casse : `$sansMfa` écrase `-SansMFA`).
- Version : constante `AUTHOR` du HTML, bannières et blocs `.NOTES` des scripts, en-têtes des documents, historique de `docs/REFERENCE.md`.

# Copie Diélectrie → PV de Test

Recopie les résultats des tests diélectriques (fichiers `.xls` du banc) dans les PV de test Excel (`.xlsm` / `.xlsx`).
Les fichiers source et cible sont associés par l'identifiant `xxxxxx-yyyyyyy-zzzzz` présent dans leur nom.

Nécessite Windows et Excel installé (pilotage par COM).

## Types de copie

Un type de copie = onglet cible, extension cible et correspondance des cellules source → cible.

Ils se modifient dans le logiciel avec le bouton **Configurer…** : créer, dupliquer, renommer, supprimer un type, et éditer ses cellules.
La configuration est enregistrée dans `modes.json`, **à côté de l'exe** (l'exe doit donc être dans un dossier où l'on peut écrire).
Pour partager une configuration avec un autre poste, copier ce fichier à côté de l'autre exe.
Sans `modes.json`, ou s'il est invalide, le logiciel utilise les valeurs d'origine (`DEFAULT_MODES` dans `copie_dielectrie_gui.py`) :

| Mode | Onglet cible | Extension |
|---|---|---|
| HubView-xT | Test final (1_4) | .xlsm |
| NVR | Test final (1_4) | .xlsm |
| HubView-MF19 | Test Final(1_3) | .xlsx |
| Poste de Locution | Test final (1_4) | .xlsm |

## Récupérer l'exe

Chaque push sur `main` compile l'exe sur un runner Windows (GitHub Actions) :
onglet **Actions** → dernier run « Build exe » → section **Artifacts** → `copie_dielectrie_gui`.

Pour une version téléchargeable dans **Releases** : sur GitHub, **Releases** → **Draft a new release**,
choisir un nouveau tag (ex. `v1.1.0`) sur `main`, puis **Publish release**.
L'exe est compilé et joint automatiquement à la release en 1 à 2 minutes.

## Compiler en local

Sous Windows, avec Python dans le PATH : double-cliquer sur `build.bat`. L'exe est créé dans `dist\`.

# Copie Diélectrie → PV de Test

Recopie les résultats des tests diélectriques (fichiers `.xls` du banc) dans les PV de test Excel (`.xlsm` / `.xlsx`).
Les fichiers source et cible sont associés par l'identifiant `xxxxxx-yyyyyyy-zzzzz` présent dans leur nom.

Nécessite Windows et Excel installé (pilotage par COM).

## Types de copie

Définis dans le dictionnaire `MODES` en tête de `copie_dielectrie_gui.py` : onglet cible, extension cible et correspondance des cellules source → cible. Ajouter un type = ajouter une entrée.

| Mode | Onglet cible | Extension |
|---|---|---|
| HubView-xT | Test final (1_4) | .xlsm |
| NVR | Test final (1_4) | .xlsm |
| HubView-MF19 | Test Final(1_3) | .xlsx |
| Poste de Locution | Test final (1_4) | .xlsx |

## Récupérer l'exe

Chaque push sur `main` compile l'exe sur un runner Windows (GitHub Actions) :
onglet **Actions** → dernier run « Build exe » → section **Artifacts** → `copie_dielectrie_gui`.

Pour une version téléchargeable dans **Releases**, pousser un tag :

```bash
git tag v1.0.0
git push origin v1.0.0
```

## Compiler en local

Sous Windows, avec Python dans le PATH : double-cliquer sur `build.bat`. L'exe est créé dans `dist\`.

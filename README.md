# Registre des Tirages — Analyseur 01–90

Site web statique (HTML/CSS/JS, sans backend) qui analyse des fichiers Excel de
tirages de loterie 01–90 : fréquences des numéros, paires, et lots (dizaines).

Tout le traitement se fait **dans le navigateur** de la personne qui l'utilise —
aucun fichier n'est envoyé vers un serveur.

## Contenu du dossier

```
analyseur-tirages/
├── index.html   → structure de la page
├── style.css    → design
├── app.js       → lecture Excel + calculs + affichage
└── README.md    → ce fichier
```

## Format de fichier attendu

Un fichier `.xlsx` avec une colonne "Résultats" contenant les numéros séparés
par des tirets, par exemple :

| Date | Heure | Résultats |
|---|---|---|
| 20/10/2025 | 0 | 34-17-76-88-30 |


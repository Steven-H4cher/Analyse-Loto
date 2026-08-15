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

Si aucune colonne ne s'appelle exactement "Résultats", le site essaie de
deviner automatiquement la bonne colonne.

## Comment mettre le site en ligne (gratuit)

### Option 1 — GitHub Pages (recommandé, gratuit et durable)
1. Crée un compte sur [github.com](https://github.com) si tu n'en as pas.
2. Crée un nouveau dépôt (repository), par exemple `analyseur-tirages`.
3. Mets les 3 fichiers (`index.html`, `style.css`, `app.js`) à la racine du dépôt
   — soit en les glissant-déposant sur la page GitHub ("Add file" → "Upload files"),
   soit avec Git en ligne de commande.
4. Va dans **Settings → Pages**, choisis la branche `main` et le dossier `/ (root)`,
   puis clique sur **Save**.
5. Après 1-2 minutes, le site est accessible à une adresse du type :
   `https://<ton-nom-utilisateur>.github.io/analyseur-tirages/`

### Option 2 — Netlify Drop (le plus rapide, aucun compte requis au départ)
1. Va sur [app.netlify.com/drop](https://app.netlify.com/drop)
2. Glisse-dépose le dossier `analyseur-tirages` entier dans la zone indiquée.
3. Le site est en ligne immédiatement avec une adresse `https://un-nom-genere.netlify.app`
   (tu peux créer un compte gratuit ensuite pour garder ce lien et le personnaliser).

### Option 3 — Vercel
Même principe que Netlify : dépose le dossier sur [vercel.com](https://vercel.com)
après avoir créé un compte gratuit.

## Aucune de ces options ne coûte rien pour ce type de site (statique, sans base de données).

## Prochaines évolutions possibles
- Mémoriser plusieurs fichiers uploadés pour cumuler l'historique
- Export du rapport en PDF ou image
- Ré-ajout de l'analyse "Absence par lot" et "Mariages de lots" (déjà codées
  dans une version précédente, juste masquées pour l'instant)

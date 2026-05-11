# 💅 BGlam · Générateur de Commandes

Outil interne de gestion et génération de bons de commande fournisseur.

## 🚀 Déploiement GitHub Pages (5 minutes)

### 1. Créer le repo GitHub

1. Va sur [github.com](https://github.com) → **New repository**
2. Nom : `bglam-commandes` (ou ce que tu veux)
3. Visibilité : **Private** (recommandé — c'est un outil interne)
4. Clique **Create repository**

### 2. Uploader les fichiers

Option A — via l'interface GitHub (le plus simple) :
1. Dans ton repo → **Add file** → **Upload files**
2. Glisse tout le dossier `bglam-commandes/`
3. Commit : "Initial commit"

Option B — via terminal :
```bash
cd bglam-commandes
git init
git remote add origin https://github.com/TON-USERNAME/bglam-commandes.git
git add .
git commit -m "Initial commit"
git push -u origin main
```

### 3. Activer GitHub Pages

1. Dans le repo → **Settings** → **Pages**
2. Source : **Deploy from a branch**
3. Branch : `main` / `/ (root)`
4. Clique **Save**

Ton outil sera accessible à :
`https://TON-USERNAME.github.io/bglam-commandes/`

---

## 📸 Ajouter les photos produits

### Méthode 1 — Glisser-déposer dans l'outil (le plus rapide)

Dans l'outil, chaque produit a une zone `+` pour glisser une photo directement.
La photo est stockée dans le navigateur (localStorage) — elle reste entre les sessions.

**Limite : 5-10 MB de photos max dans localStorage.**

### Méthode 2 — Dossier `images/` sur GitHub (recommandé pour la durabilité)

1. Crée un dossier `images/` dans le repo
2. Nomme chaque photo avec la référence exacte du produit :
   - `BA-398.jpg`
   - `BA-399.png`
   - etc.
3. Upload sur GitHub
4. L'outil détecte automatiquement les photos dans `images/` si la photo locale n'est pas définie

**Format accepté :** JPG, PNG, WEBP — recommandé : 400×400px minimum

---

## 📥 Importer un nouveau fichier commande Excel

1. Clique **📥 Importer Excel** dans l'outil
2. Glisse ton fichier `.xlsx` ou `.xls`
3. L'outil détecte les colonnes automatiquement (ITEM NO, PRICE, PCS/CTN, etc.)
4. Vérifie le mapping si besoin, puis clique **Importer et fusionner**

**Règles de fusion :**
- Nouvelle référence → ajoutée au catalogue
- Prix ou MOQ différent → mis à jour
- Photo déjà renseignée → jamais écrasée
- Quantités en cours → jamais touchées

---

## 🗂️ Structure du repo

```
bglam-commandes/
├── index.html          ← L'outil complet
├── images/             ← Photos produits (optionnel)
│   ├── BA-398.jpg
│   ├── BA-399.jpg
│   └── ...
└── README.md
```

---

## 💾 Sauvegarde du catalogue

- Le catalogue et les photos sont sauvegardés automatiquement dans le navigateur (localStorage)
- Pour une sauvegarde externe : clique **💾 Backup** → télécharge un fichier JSON
- Pour restaurer : importe le JSON comme un fichier Excel (le mapping se fait automatiquement)

---

## ✨ Fonctionnalités

| Fonction | Description |
|---|---|
| 📥 Importer Excel | Fusionne n'importe quel fichier commande fournisseur |
| 📸 Photo par produit | Glisser-déposer ou URL — stocké localement |
| 🏷️ Catégories BGlam | Tags par catégorie du site bglam-re.com |
| 🔍 Recherche/filtres | Par réf, description, catégorie, groupe |
| 🖨 Aperçu commande | Page imprimable avec photos |
| ⬇ Export Excel | Bon de commande au format fournisseur |
| 💾 Backup catalogue | Export JSON du catalogue complet |

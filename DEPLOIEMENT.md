# 🚀 Déploiement de PlanniPro sur Render.com (gratuit)

## En 5 étapes — environ 10 minutes

---

### Étape 1 — Créer un compte GitHub (si pas déjà fait)
→ https://github.com/signup  
Gratuit, juste une adresse mail.

---

### Étape 2 — Mettre les fichiers sur GitHub

1. Aller sur https://github.com/new
2. Nommer le dépôt : `plannipro`
3. Laisser en **Public**, cliquer **Create repository**
4. Cliquer **uploading an existing file**
5. Déposer TOUS les fichiers du dossier `plannipro/` :
   - `app.py`
   - `requirements.txt`
   - `render.yaml`
   - `templates/index.html` ← créer le dossier `templates` d'abord
6. Cliquer **Commit changes**

---

### Étape 3 — Créer un compte Render.com (gratuit)
→ https://render.com/  
Se connecter avec GitHub (bouton "Sign in with GitHub").

---

### Étape 4 — Déployer l'application

1. Sur Render.com, cliquer **+ New** → **Web Service**
2. Sélectionner votre dépôt `plannipro`
3. Render détecte automatiquement le `render.yaml`
4. Cliquer **Create Web Service**
5. ⏳ Attendre 3-5 minutes que le déploiement se termine

---

### Étape 5 — Votre site est en ligne !

Render vous donne une URL comme :  
`https://plannipro.onrender.com`

→ Donnez cette URL à votre mère, c'est tout !

---

## ⚠️ Notes importantes

- **Gratuit** : le plan Free de Render.com est suffisant
- **Mise en veille** : après 15 min d'inactivité, le site "dort".  
  Le premier chargement après une pause prend ~30 secondes, c'est normal.
- **Données** : les fichiers uploadés sont supprimés automatiquement après téléchargement
- **Mises à jour** : modifier les fichiers sur GitHub → Render redéploie automatiquement

---

## En cas de problème

Envoyer un message avec le texte d'erreur affiché sur Render.com  
(bouton "Logs" dans le dashboard).

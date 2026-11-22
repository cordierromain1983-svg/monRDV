# Rendez-vous avant le conseil de classe

Une page web toute simple : l'élève clique sur un créneau libre, tape son prénom, valide. C'est tout.

- **La page** (`index.html`) est hébergée gratuitement sur GitHub Pages.
- **Les réservations** sont enregistrées dans un Google Sheet, grâce à un petit script (`Code.gs`).
- Les élèves n'ont besoin d'**aucun compte**. Ils voient « libre » ou « pris », jamais les prénoms des autres.
- Un même prénom ne peut réserver qu'**un seul créneau**, et deux élèves ne peuvent pas prendre le même.

Installation : environ 10 minutes, une seule fois.

---

## Étape 1 — Le Google Sheet et son script (5 min)

1. Créez un Google Sheet vide (sheets.new), nommez-le par exemple « RDV conseil de classe ».
2. Menu **Extensions → Apps Script**.
3. Effacez le contenu de l'éditeur, collez tout le contenu de `Code.gs`, puis cliquez sur 💾 **Enregistrer**.
4. En haut, choisissez la fonction **`installer`** dans la liste déroulante, puis cliquez sur **▶ Exécuter**.
   - Google demande une autorisation : **Examiner les autorisations** → votre compte →
     **Paramètres avancés** → **Accéder à … (non sécurisé)** → **Autoriser**.
     (Ce message est normal : le script est le vôtre et n'a pas été « vérifié » par Google.)
   - Retournez sur le Google Sheet : un onglet **Créneaux** est apparu avec des créneaux d'exemple.
5. Cliquez sur **Déployer → Nouveau déploiement**, puis sur la roue ⚙ → **Application Web**.
   - Exécuter en tant que : **Moi**
   - Qui a accès : **Tout le monde**
   - **Déployer**, puis **copiez l'URL de l'application Web** (elle se termine par `/exec`).

## Étape 2 — Brancher la page (1 min)

Ouvrez `index.html` avec un éditeur de texte (Bloc-notes, TextEdit…) et collez l'URL ici :

```js
url: 'https://script.google.com/macros/s/XXXXXXXX/exec',
```

Vous pouvez aussi modifier `titre` et `sousTitre` juste en dessous.

## Étape 3 — Mettre en ligne sur GitHub Pages (4 min)

1. Sur github.com, cliquez sur **New repository**. Nom : par exemple `rdv-conseil`, visibilité **Public**, puis **Create repository**.
2. Cliquez sur **uploading an existing file**, déposez `index.html`, puis **Commit changes**.
3. **Settings → Pages** → Source : **Deploy from a branch** → Branch : **main** / **(root)** → **Save**.
4. Une à deux minutes plus tard, la page est en ligne à l'adresse :
   `https://VOTRE-PSEUDO.github.io/rdv-conseil/`

Partagez ce lien aux élèves (ENT, Pronote, QR code au tableau…).

---

## Au quotidien : tout se gère dans le Google Sheet

| Je veux… | Je fais… |
|---|---|
| Ajouter un créneau | J'ajoute une ligne : Jour + Heure, Prénom vide |
| Retirer un créneau | Je supprime la ligne |
| Libérer un créneau (annulation) | J'efface le prénom et la date |
| Fermer les réservations | Je supprime les lignes encore libres |
| Préparer un autre conseil de classe | Je vide les lignes et j'en écris de nouvelles |

Les jours et heures s'écrivent librement (« Jeudi 4 déc. », « 17h15 »…). La page les affiche tels quels, regroupés par jour, dans l'ordre du tableau.

**Pas besoin de redéployer** quand vous modifiez le tableau. Il faut redéployer seulement si vous modifiez le **code** du script :
Déployer → Gérer les déploiements → ✏ → Version : **Nouvelle version**. L'URL reste la même.

## En cas de souci

- **« Impossible de charger les créneaux »** : vérifiez que l'URL collée se termine bien par `/exec` et que l'accès est réglé sur **Tout le monde**.
- **La page affiche « Mode démonstration »** : l'URL n'a pas été collée dans `index.html`, ou la nouvelle version n'a pas encore été envoyée sur GitHub.
- **Deux élèves ont le même prénom** : le second reçoit un message lui proposant d'ajouter l'initiale de son nom (ex. « Léa M. »).

# Rendez-vous avant le conseil de classe

Page élève : https://cordierromain1983-svg.github.io/monRDV/

L'élève touche un créneau, puis son prénom, répond au QCM de préparation (avec un message libre facultatif), et c'est réservé.
Il reçoit un **code personnel à 4 caractères** qui lui permet de modifier ses réponses, de changer de créneau ou d'annuler, y compris depuis un autre appareil.
Il peut ajouter le rendez-vous à son agenda (Google, iPhone ou Outlook), avec un rappel une heure avant.

Tout se gère dans le Google Sheet **« RDV conseil de classe »** (Google Drive).

## Les onglets

| Onglet | Rôle |
|---|---|
| **Créneaux** | Une ligne = un créneau (Jour, Heure, Prénom vide). Le prénom s'ajoute quand un élève réserve. |
| **Élèves** | La liste de la classe. Une ligne **rouge** = élève pas encore inscrit. Les réponses au QCM et le message libre arrivent ici. |
| **Questions** | Le QCM : une question par ligne, les choix séparés par `/`. Ajouter, retirer ou reformuler une question est pris en compte immédiatement. |
| **Réglages** | Titre de la page, lieu, durée d'un rendez-vous, **date du conseil de classe**. |

## Au quotidien

| Je veux… | Je fais… |
|---|---|
| Ajouter / retirer un créneau | J'ajoute ou je supprime une ligne dans **Créneaux** |
| Libérer le créneau d'un élève | J'efface son prénom dans **Créneaux** |
| Ajouter un élève | J'écris son prénom en bas de la colonne A de **Élèves** |
| Voir qui ne s'est pas inscrit | Lignes en rouge dans **Élèves** |
| Effacer toutes les réservations et réponses | Menu **Rendez-vous → Purger** dans le tableau |

## Purge automatique

Indiquez la date du conseil dans **Réglages** (format `jj/mm/aaaa`).
Le lendemain, les rendez-vous, codes, réponses et messages sont effacés automatiquement et la page affiche « Les rendez-vous sont terminés ».
La liste de classe, les questions et les créneaux sont conservés : pour le conseil suivant, changez la date, ajustez les créneaux, et c'est reparti.

## Technique

- `index.html` : la page élève (GitHub Pages).
- `Code.gs` : le script Google relié au tableau (Extensions → Apps Script).
  Après toute modification du script : **Déployer → Gérer les déploiements → ✏ → Version : Nouvelle version → Déployer** (l'adresse ne change pas).
- Sans adresse de script dans `CONFIG.url`, la page s'ouvre en mode démonstration.

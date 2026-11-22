/**
 * Rendez-vous avant le conseil de classe — script Google Apps Script
 *
 * Ce script relie la page des élèves (hébergée sur GitHub Pages) à ce Google Sheet.
 * - La feuille « Créneaux » contient : Jour | Heure | Prénom | Réservé le
 * - Pour proposer un créneau : ajoutez une ligne avec Jour et Heure, laissez Prénom vide.
 * - Pour libérer un créneau : effacez le prénom (et la date de réservation).
 * - Pour retirer un créneau : supprimez la ligne.
 *
 * Installation : lancez une fois la fonction « installer », puis déployez en application Web
 * (voir le guide README.md).
 */

const NOM_FEUILLE = 'Créneaux';

/** À lancer une seule fois : crée la feuille avec des créneaux d'exemple. */
function installer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(NOM_FEUILLE) || ss.insertSheet(NOM_FEUILLE, 0);
  if (sh.getLastRow() > 0) {
    console.log('La feuille « ' + NOM_FEUILLE + ' » existe déjà : rien n\'a été modifié.');
    return;
  }
  const lignes = [['Jour', 'Heure', 'Prénom', 'Réservé le']];
  ['Lundi 30 novembre', 'Mardi 1er décembre'].forEach(function (jour) {
    for (let m = 16 * 60 + 30; m < 18 * 60; m += 10) {
      lignes.push([jour, Math.floor(m / 60) + 'h' + String(m % 60).padStart(2, '0'), '', '']);
    }
  });
  sh.getRange('A:C').setNumberFormat('@'); // tout en texte : pas de conversion automatique en date/heure
  sh.getRange(1, 1, lignes.length, 4).setValues(lignes);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8eefc');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 4, 170);
}

/** Liste des créneaux (sans les prénoms : les élèves voient seulement « pris » ou « libre »). */
function doGet() {
  try {
    const creneaux = lireCreneaux_(feuille_()).map(function (c) {
      return { id: c.id, jour: c.jour, heure: c.heure, libre: c.prenom === '' };
    });
    return json_({ ok: true, creneaux: creneaux });
  } catch (err) {
    return json_({ ok: false, erreur: 'Impossible de lire les créneaux.' });
  }
}

/** Réservation d'un créneau : { id, jour, heure, prenom } */
function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, erreur: 'Demande invalide.' });
  }

  const prenom = String(req.prenom || '').replace(/\s+/g, ' ').trim().slice(0, 40);
  if (!prenom) return json_({ ok: false, erreur: 'Indique ton prénom.' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return json_({ ok: false, erreur: 'Beaucoup de monde en même temps… Réessaie dans quelques secondes.' });
  }
  try {
    const sh = feuille_();
    const creneaux = lireCreneaux_(sh);

    // Un seul rendez-vous par prénom (insensible aux majuscules et aux accents)
    const cle = normaliser_(prenom);
    const doublon = creneaux.find(function (c) { return normaliser_(c.prenom) === cle; });
    if (doublon) {
      return json_({
        ok: false, code: 'doublon',
        erreur: 'Le prénom « ' + prenom + ' » a déjà un rendez-vous. ' +
          'Si ce n\'est pas toi, ajoute l\'initiale de ton nom (ex. : ' + prenom + ' M.).'
      });
    }

    // On vérifie que la ligne correspond toujours au créneau affiché (au cas où le tableau a été modifié)
    const c = creneaux.find(function (x) { return x.id === Number(req.id); });
    if (!c || c.jour !== String(req.jour || '').trim() || c.heure !== String(req.heure || '').trim()) {
      return json_({ ok: false, code: 'change', erreur: 'La liste des créneaux a changé. Choisis à nouveau.' });
    }
    if (c.prenom !== '') {
      return json_({ ok: false, code: 'pris', erreur: 'Oups, ce créneau vient d\'être pris. Choisis-en un autre.' });
    }

    // Le préfixe ' empêche qu'un prénom commençant par = + - @ soit lu comme une formule
    const valeur = /^[=+\-@]/.test(prenom) ? "'" + prenom : prenom;
    sh.getRange(c.id, 3, 1, 2).setValues([[valeur, new Date()]]);
    SpreadsheetApp.flush();
    return json_({ ok: true, jour: c.jour, heure: c.heure, prenom: prenom });
  } catch (err) {
    return json_({ ok: false, erreur: 'Une erreur est survenue. Réessaie.' });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- outils internes ---------- */

function feuille_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOM_FEUILLE);
  if (!sh) throw new Error('Feuille « ' + NOM_FEUILLE + ' » introuvable : lancez installer().');
  return sh;
}

function lireCreneaux_(sh) {
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 3).getDisplayValues()
    .map(function (r, i) {
      return { id: i + 2, jour: r[0].trim(), heure: r[1].trim(), prenom: r[2].trim() };
    })
    .filter(function (c) { return c.jour !== '' && c.heure !== ''; });
}

function normaliser_(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

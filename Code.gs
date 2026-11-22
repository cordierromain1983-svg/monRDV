/**
 * Rendez-vous avant le conseil de classe — script Google Apps Script (v2)
 *
 * Onglets du tableau :
 *  - Créneaux  : Jour | Heure | Prénom | Réservé le      → une ligne = un créneau (laisser Prénom vide)
 *  - Élèves    : Prénom | Rendez-vous | réponses…        → la liste de classe ; Rendez-vous vide = pas encore inscrit
 *  - Questions : Question | Choix (séparés par /)        → le QCM de préparation, modifiable
 *  - Réglages  : titre, lieu, durée, date du conseil     → après la date du conseil, tout est purgé automatiquement
 *
 * Installation : lancer une fois « installer », puis déployer en application Web (accès : Tout le monde).
 * Après une modification de ce code : Déployer → Gérer les déploiements → ✏ → Nouvelle version.
 */

const F_CRENEAUX = 'Créneaux';
const F_ELEVES = 'Élèves';
const F_QUESTIONS = 'Questions';
const F_REGLAGES = 'Réglages';

const R = {
  titre: 'Titre de la page',
  lieu: 'Lieu du rendez-vous',
  duree: 'Durée d\'un rendez-vous (minutes)',
  conseil: 'Date du conseil de classe (jj/mm/aaaa)',
  purge: 'Dernière purge (automatique)'
};

const COL = {
  prenom: 'Prénom',
  rdv: 'Rendez-vous',
  message: 'Message libre',
  maj: 'Mis à jour le',
  code: 'Code élève (ne pas modifier)'
};

const QUESTIONS_DEFAUT = [
  ['Comment juges-tu ton semestre ?', 'Très bien / Plutôt bien / Moyen / Difficile'],
  ['Ton travail personnel à la maison ?', 'Régulier / Irrégulier / Presque inexistant'],
  ['Ta participation en classe ?', 'Je participe souvent / De temps en temps / Rarement'],
  ['Ta principale difficulté ?', 'Comprendre les cours / M\'organiser / Rester motivé·e / Aucune en particulier'],
  ['Tu veux parler d\'orientation ?', 'Oui / Non / Je ne sais pas encore']
];

/* ======================= Installation & menu ======================= */

/** Crée les onglets manquants. Ne modifie jamais un onglet déjà rempli. */
function installer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let cr = ss.getSheetByName(F_CRENEAUX);
  if (!cr) cr = ss.insertSheet(F_CRENEAUX, 0);
  if (cr.getLastRow() === 0) {
    cr.getRange('A:C').setNumberFormat('@');
    cr.getRange(1, 1, 1, 4).setValues([['Jour', 'Heure', 'Prénom', 'Réservé le']]);
    entete_(cr, 4);
  }

  let el = ss.getSheetByName(F_ELEVES);
  if (!el) el = ss.insertSheet(F_ELEVES);
  if (el.getLastRow() === 0) {
    const entetes = [COL.prenom, COL.rdv]
      .concat(QUESTIONS_DEFAUT.map(function (q) { return q[0]; }))
      .concat([COL.message, COL.maj, COL.code]);
    el.getRange(1, 1, 1, entetes.length).setValues([entetes]);
    el.getRange('A:B').setNumberFormat('@');
    entete_(el, entetes.length);
    el.setFrozenColumns(1);
    // Rendez-vous vide = élève pas encore inscrit → en rouge clair
    const regle = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",$B2="")')
      .setBackground('#fde8e4')
      .setRanges([el.getRange('A2:B200')])
      .build();
    el.setConditionalFormatRules([regle]);
  }

  let qs = ss.getSheetByName(F_QUESTIONS);
  if (!qs) qs = ss.insertSheet(F_QUESTIONS);
  if (qs.getLastRow() === 0) {
    qs.getRange(1, 1, 1, 2).setValues([['Question', 'Choix (séparés par /)']]);
    qs.getRange(2, 1, QUESTIONS_DEFAUT.length, 2).setValues(QUESTIONS_DEFAUT);
    entete_(qs, 2);
    qs.setColumnWidth(1, 300);
    qs.setColumnWidth(2, 420);
  }

  let rg = ss.getSheetByName(F_REGLAGES);
  if (!rg) rg = ss.insertSheet(F_REGLAGES);
  if (rg.getLastRow() === 0) {
    rg.getRange('B:B').setNumberFormat('@');
    rg.getRange(1, 1, 6, 2).setValues([
      ['Réglage', 'Valeur'],
      [R.titre, 'Rendez-vous avant le conseil de classe'],
      [R.lieu, ''],
      [R.duree, '10'],
      [R.conseil, ''],
      [R.purge, '']
    ]);
    entete_(rg, 2);
    rg.setColumnWidth(1, 300);
    rg.setColumnWidth(2, 300);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Rendez-vous')
    .addItem('Purger les réservations et réponses maintenant', 'purgerMaintenant')
    .addToUi();
  try { purgerSiBesoin_(); } catch (e) {}
}

function purgerMaintenant() {
  const ui = SpreadsheetApp.getUi();
  const rep = ui.alert('Purger ?',
    'Tous les rendez-vous, codes et réponses des élèves seront effacés. La liste de classe et les créneaux sont conservés.',
    ui.ButtonSet.YES_NO);
  if (rep === ui.Button.YES) {
    purger_();
    ui.alert('Purge effectuée.');
  }
}

/* ======================= Web : lecture & actions ======================= */

function doGet() {
  return avecVerrou_(function () {
    purgerSiBesoin_();
    return etat_();
  });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, erreur: 'Demande invalide.' });
  }
  return avecVerrou_(function () {
    purgerSiBesoin_();
    if (estTermine_(reglages_())) {
      return { ok: false, code: 'ferme', erreur: 'Les rendez-vous sont terminés.' };
    }
    switch (req.action) {
      case 'reserver': return reserver_(req);
      case 'annuler': return annuler_(req);
      case 'moi': return moi_(req);
      default: return { ok: false, erreur: 'Action inconnue.' };
    }
  });
}

function etat_() {
  const reg = reglages_();
  const creneaux = lireCreneaux_();
  const eleves = lireEleves_();
  synchroniserRdv_(eleves, creneaux);
  return {
    ok: true,
    titre: reg[R.titre] || 'Rendez-vous avant le conseil de classe',
    lieu: reg[R.lieu] || '',
    duree: Number(reg[R.duree]) || 10,
    ferme: estTermine_(reg),
    questions: lireQuestions_(),
    eleves: eleves.lignes.map(function (e) {
      return { prenom: e.prenom, inscrit: !!creneauDe_(creneaux, e.prenom) };
    }),
    creneaux: creneaux.map(function (c) {
      return { id: c.id, jour: c.jour, heure: c.heure, libre: c.prenom === '' };
    })
  };
}

/** Réserver, déplacer un rendez-vous ou mettre à jour ses réponses. */
function reserver_(req) {
  const eleves = lireEleves_();
  const eleve = trouverEleve_(eleves, req.prenom);
  if (!eleve) return { ok: false, erreur: 'Prénom introuvable dans la liste de la classe.' };

  const creneaux = lireCreneaux_();
  const actuel = creneauDe_(creneaux, eleve.prenom);
  if (actuel && !codeValide_(eleve, req.code)) {
    return { ok: false, code: 'code', erreur: 'Tu as déjà un rendez-vous. Utilise « Retrouver mon rendez-vous » avec ton code pour le modifier.' };
  }

  const cible = creneaux.find(function (x) { return x.id === Number(req.id); });
  if (!cible || cible.jour !== String(req.jour || '').trim() || cible.heure !== String(req.heure || '').trim()) {
    return { ok: false, code: 'change', erreur: 'La liste des créneaux a changé. Choisis à nouveau.' };
  }
  if (cible.prenom !== '' && normaliser_(cible.prenom) !== normaliser_(eleve.prenom)) {
    return { ok: false, code: 'pris', erreur: 'Oups, ce créneau vient d\'être pris. Choisis-en un autre.' };
  }

  // Réponses : uniquement les choix proposés
  const questions = lireQuestions_();
  const reponses = req.reponses || {};
  for (let i = 0; i < questions.length; i++) {
    const v = reponses[questions[i].q];
    if (!v || questions[i].choix.indexOf(v) === -1) {
      return { ok: false, erreur: 'Réponds à toutes les questions avant de valider.' };
    }
  }
  const message = String(req.message || '').trim().slice(0, 1000);

  const sh = feuille_(F_CRENEAUX);
  const maintenant = new Date();
  creneaux.forEach(function (c) {
    if (c.id !== cible.id && normaliser_(c.prenom) === normaliser_(eleve.prenom)) {
      sh.getRange(c.id, 3, 1, 2).setValues([['', '']]); // libère l'ancien créneau
    }
  });
  if (!actuel || actuel.id !== cible.id) {
    sh.getRange(cible.id, 3, 1, 2).setValues([[texte_(eleve.prenom), maintenant]]);
  }

  const code = (actuel && eleve.code) ? eleve.code : nouveauCode_();
  const valeurs = {};
  valeurs[COL.rdv] = cible.jour + ' à ' + cible.heure;
  valeurs[COL.message] = texte_(message);
  valeurs[COL.maj] = maintenant;
  valeurs[COL.code] = code;
  questions.forEach(function (q) { valeurs[q.q] = reponses[q.q]; });
  ecrireEleve_(eleves, eleve, valeurs);
  SpreadsheetApp.flush();

  return { ok: true, prenom: eleve.prenom, code: code, id: cible.id, jour: cible.jour, heure: cible.heure };
}

function annuler_(req) {
  const eleves = lireEleves_();
  const eleve = trouverEleve_(eleves, req.prenom);
  if (!eleve || !codeValide_(eleve, req.code)) return { ok: false, code: 'code', erreur: 'Code incorrect.' };
  const sh = feuille_(F_CRENEAUX);
  lireCreneaux_().forEach(function (c) {
    if (normaliser_(c.prenom) === normaliser_(eleve.prenom)) sh.getRange(c.id, 3, 1, 2).setValues([['', '']]);
  });
  const vide = {};
  eleves.entetes.forEach(function (h) { if (h && h !== COL.prenom) vide[h] = ''; });
  ecrireEleve_(eleves, eleve, vide);
  SpreadsheetApp.flush();
  return { ok: true };
}

/** Retrouver son rendez-vous (autre appareil) grâce à son code. */
function moi_(req) {
  const eleves = lireEleves_();
  const eleve = trouverEleve_(eleves, req.prenom);
  if (!eleve || !codeValide_(eleve, req.code)) return { ok: false, code: 'code', erreur: 'Prénom ou code incorrect.' };
  const c = creneauDe_(lireCreneaux_(), eleve.prenom);
  if (!c) return { ok: false, code: 'aucun', erreur: 'Tu n\'as pas de rendez-vous en ce moment.' };
  const reponses = {};
  lireQuestions_().forEach(function (q) { reponses[q.q] = eleve.valeurs[q.q] || ''; });
  return {
    ok: true, prenom: eleve.prenom, code: eleve.code, id: c.id, jour: c.jour, heure: c.heure,
    reponses: reponses, message: eleve.valeurs[COL.message] || ''
  };
}

/* ======================= Purge ======================= */

function purgerSiBesoin_() {
  const reg = reglages_();
  const cle = cleConseil_(reg);
  if (!cle || aujourdhui_() <= cle) return;
  const derniere = String(reg[R.purge] || '').replace(/\D/g, '');
  if (derniere && derniere >= cle) return; // déjà purgé pour ce conseil
  purger_();
}

function purger_() {
  const cr = feuille_(F_CRENEAUX);
  if (cr.getLastRow() > 1) cr.getRange(2, 3, cr.getLastRow() - 1, 2).clearContent();
  const el = feuille_(F_ELEVES);
  if (el.getLastRow() > 1 && el.getLastColumn() > 1) {
    el.getRange(2, 2, el.getLastRow() - 1, el.getLastColumn() - 1).clearContent();
  }
  ecrireReglage_(R.purge, Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyy-MM-dd'));
}

function estTermine_(reg) {
  const cle = cleConseil_(reg);
  return !!cle && aujourdhui_() > cle;
}

/** Date du conseil sous la forme aaaammjj (texte), ou '' si absente/illisible. */
function cleConseil_(reg) {
  const v = reg[R.conseil];
  if (v instanceof Date) return Utilities.formatDate(v, 'Europe/Paris', 'yyyyMMdd');
  const m = String(v || '').match(/(\d{1,2})\s*[\/.\-]\s*(\d{1,2})\s*[\/.\-]\s*(\d{2,4})/);
  if (!m) return '';
  const an = m[3].length === 2 ? '20' + m[3] : m[3];
  return an + ('0' + m[2]).slice(-2) + ('0' + m[1]).slice(-2);
}

function aujourdhui_() {
  return Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyyMMdd');
}

/* ======================= Lecture / écriture ======================= */

function feuille_(nom) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nom);
  if (!sh) throw new Error('Onglet « ' + nom + ' » introuvable : lancez installer().');
  return sh;
}

function lireCreneaux_() {
  const sh = feuille_(F_CRENEAUX);
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 3).getDisplayValues()
    .map(function (r, i) { return { id: i + 2, jour: r[0].trim(), heure: r[1].trim(), prenom: r[2].trim() }; })
    .filter(function (c) { return c.jour !== '' && c.heure !== ''; });
}

function creneauDe_(creneaux, prenom) {
  const cle = normaliser_(prenom);
  return creneaux.find(function (c) { return c.prenom !== '' && normaliser_(c.prenom) === cle; }) || null;
}

function lireEleves_() {
  const sh = feuille_(F_ELEVES);
  const nl = sh.getLastRow(), nc = Math.max(sh.getLastColumn(), 1);
  const entetes = sh.getRange(1, 1, 1, nc).getDisplayValues()[0].map(function (h) { return h.trim(); });
  const lignes = [];
  if (nl >= 2) {
    sh.getRange(2, 1, nl - 1, nc).getDisplayValues().forEach(function (r, i) {
      const prenom = String(r[0]).trim();
      if (!prenom) return;
      const valeurs = {};
      entetes.forEach(function (h, j) { if (h) valeurs[h] = r[j]; });
      lignes.push({ ligne: i + 2, prenom: prenom, code: String(valeurs[COL.code] || '').trim(), valeurs: valeurs });
    });
  }
  return { sh: sh, entetes: entetes, lignes: lignes };
}

function trouverEleve_(eleves, prenom) {
  const cle = normaliser_(prenom || '');
  if (!cle) return null;
  return eleves.lignes.find(function (e) { return normaliser_(e.prenom) === cle; }) || null;
}

function ecrireEleve_(eleves, eleve, valeurs) {
  Object.keys(valeurs).forEach(function (h) {
    let j = eleves.entetes.indexOf(h);
    if (j === -1) { // nouvelle question ajoutée par le professeur → nouvelle colonne
      j = eleves.entetes.length;
      eleves.entetes.push(h);
      eleves.sh.getRange(1, j + 1).setValue(h).setFontWeight('bold').setBackground('#ece9ff');
    }
    eleves.sh.getRange(eleve.ligne, j + 1).setValue(valeurs[h]);
  });
}

/** Tient la colonne « Rendez-vous » à jour si le professeur modifie les créneaux à la main. */
function synchroniserRdv_(eleves, creneaux) {
  const j = eleves.entetes.indexOf(COL.rdv);
  if (j === -1) return;
  eleves.lignes.forEach(function (e) {
    const c = creneauDe_(creneaux, e.prenom);
    const attendu = c ? c.jour + ' à ' + c.heure : '';
    if ((e.valeurs[COL.rdv] || '') !== attendu) eleves.sh.getRange(e.ligne, j + 1).setValue(attendu);
  });
}

function lireQuestions_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(F_QUESTIONS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues()
    .map(function (r) {
      return {
        q: r[0].trim(),
        choix: r[1].split('/').map(function (s) { return s.trim(); }).filter(function (s) { return s; })
      };
    })
    .filter(function (x) { return x.q && x.choix.length; });
}

function reglages_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(F_REGLAGES);
  const reg = {};
  if (!sh || sh.getLastRow() < 2) return reg;
  const brut = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const aff = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues();
  brut.forEach(function (r, i) {
    const k = String(r[0]).trim();
    if (k) reg[k] = r[1] instanceof Date ? r[1] : aff[i][1].trim();
  });
  return reg;
}

function ecrireReglage_(cle, valeur) {
  const sh = feuille_(F_REGLAGES);
  const n = sh.getLastRow();
  const cles = n >= 1 ? sh.getRange(1, 1, n, 1).getDisplayValues().map(function (r) { return r[0].trim(); }) : [];
  const i = cles.indexOf(cle);
  if (i >= 0) sh.getRange(i + 1, 2).setValue(valeur);
  else sh.getRange(n + 1, 1, 1, 2).setValues([[cle, valeur]]);
}

/* ======================= Outils ======================= */

function avecVerrou_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return json_({ ok: false, erreur: 'Beaucoup de monde en même temps… Réessaie dans quelques secondes.' });
  }
  try {
    return json_(fn());
  } catch (err) {
    console.error(err);
    return json_({ ok: false, erreur: 'Une erreur est survenue. Réessaie.' });
  } finally {
    lock.releaseLock();
  }
}

function codeValide_(eleve, code) {
  return !!eleve.code && String(code || '').trim().toUpperCase() === eleve.code.toUpperCase();
}

function nouveauCode_() {
  const a = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let c = '';
  for (let i = 0; i < 4; i++) c += a.charAt(Math.floor(Math.random() * a.length));
  return c;
}

/** Empêche qu'un texte commençant par = + - @ soit interprété comme une formule. */
function texte_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function normaliser_(s) {
  return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function entete_(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#ece9ff');
  sh.setFrozenRows(1);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

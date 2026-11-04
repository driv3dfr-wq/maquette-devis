/**
 * Atelier Devis — Maquettes 3D : sauvegarde des projets clients sur Google Drive.
 *
 * Mise en route (une seule fois, avec le compte Google driv3dfr@gmail.com) :
 *  1. script.google.com → Nouveau projet → coller ce fichier à la place du code existant → Enregistrer.
 *  2. Choisir la fonction « installer » → Exécuter → accepter les autorisations.
 *     Le journal affiche la CLÉ SECRÈTE à recopier dans l'appli.
 *  3. Déployer → Nouveau déploiement → Application web
 *     Exécuter en tant que : Moi · Qui a accès : Tout le monde → copier l'URL …/exec.
 *  4. Dans l'appli : ☁ Drive → coller l'URL et la clé → Tester → Enregistrer.
 *
 * Sur le Drive : dossier « Atelier Devis »
 *   - un fichier JSON par projet (nom : « Client — Projet.json »)
 *   - tableur « Atelier Devis — Clients » (une ligne par projet, mise à jour à chaque sauvegarde)
 *   - sous-dossier « Sauvegardes » : copie complète datée, une par semaine (lancer installer() crée le déclencheur).
 *
 * Sans la clé, le script ne répond rien : l'URL seule ne suffit pas pour lire les fiches clients.
 */
const PROPS = PropertiesService.getScriptProperties();
const NOM_DOSSIER = 'Atelier Devis';
const NOM_TABLEUR = 'Atelier Devis — Clients';
const ENTETES = ['Identifiant', 'Client', 'Projet', 'N° devis', 'E-mail', 'Téléphone', 'Adresse', 'Échelle', 'Total (€)', 'Dernière sauvegarde', 'Fichier'];

/** À lancer une seule fois depuis l'éditeur. Relancer ne casse rien et réaffiche la clé. */
function installer() {
  let cle = PROPS.getProperty('CLE');
  if (!cle) { cle = Utilities.getUuid().replace(/-/g, '').slice(0, 12); PROPS.setProperty('CLE', cle); }
  dossier_();
  tableur_();
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'sauvegardeHebdo')) {
    ScriptApp.newTrigger('sauvegardeHebdo').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  }
  Logger.log('Clé secrète à recopier dans l\'appli : ' + cle);
}

function dossier_() {
  const id = PROPS.getProperty('DOSSIER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(NOM_DOSSIER);
  PROPS.setProperty('DOSSIER', f.getId());
  return f;
}

function sousDossier_(nom) {
  const d = dossier_(), it = d.getFoldersByName(nom);
  return it.hasNext() ? it.next() : d.createFolder(nom);
}

function tableur_() {
  const id = PROPS.getProperty('TABLEUR');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create(NOM_TABLEUR);
  DriveApp.getFileById(ss.getId()).moveTo(dossier_());
  const sh = ss.getSheets()[0];
  sh.setName('Clients');
  sh.getRange(1, 1, 1, ENTETES.length).setValues([ENTETES]).setFontWeight('bold');
  sh.setFrozenRows(1);
  PROPS.setProperty('TABLEUR', ss.getId());
  return ss;
}

function sortie_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return sortie_({ ok: false, err: 'Utiliser l\'appli Atelier Devis.' });
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return sortie_({ ok: false, err: 'Requête illisible' }); }
  return sortie_(api_(b));
}

function nomFichier_(b) {
  const propre = s => String(s || '').replace(/[\\/:*?"<>|]/g, ' ').trim();
  return (propre(b.client) ? propre(b.client) + ' — ' : '') + (propre(b.nom) || 'Projet') + '.json';
}

function api_(b) {
  const cle = PROPS.getProperty('CLE');
  if (!cle || !b || b.key !== cle) return { ok: false, err: 'Clé incorrecte' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (b.action) {
      case 'ping':
        return { ok: true };

      case 'list': {
        const it = dossier_().getFilesByType(MimeType.PLAIN_TEXT);
        const projets = [];
        while (it.hasNext()) {
          const f = it.next();
          let meta = {};
          try { meta = JSON.parse(f.getDescription() || '{}'); } catch (e) {}
          if (!meta.uid) continue;
          projets.push({ id: f.getId(), uid: meta.uid, nom: meta.nom || f.getName(), client: meta.client || '', echelle: meta.echelle || '', savedAt: f.getLastUpdated().toISOString() });
        }
        projets.sort((a, c) => c.savedAt.localeCompare(a.savedAt));
        return { ok: true, projets: projets };
      }

      case 'load': {
        const f = DriveApp.getFileById(b.id);
        if (!estDansDossier_(f)) return { ok: false, err: 'Fichier inconnu' };
        return { ok: true, data: f.getBlob().getDataAsString('UTF-8') };
      }

      case 'save': {
        if (!b.uid || typeof b.data !== 'string') return { ok: false, err: 'Données manquantes' };
        const meta = JSON.stringify({ uid: b.uid, nom: b.nom || '', client: b.client || '', echelle: b.echelle || '' });
        const ligne = trouverLigne_(b.uid);
        let f = null;
        if (ligne && ligne.fichierId) { try { f = DriveApp.getFileById(ligne.fichierId); if (f.isTrashed()) f = null; } catch (e) { f = null; } }
        if (f) {
          f.setContent(b.data);
          f.setName(nomFichier_(b));
        } else {
          f = dossier_().createFile(nomFichier_(b), b.data, MimeType.PLAIN_TEXT);
        }
        f.setDescription(meta);
        majTableur_(b, f, ligne);
        return { ok: true, id: f.getId() };
      }

      case 'delete': {
        const f = DriveApp.getFileById(b.id);
        if (!estDansDossier_(f)) return { ok: false, err: 'Fichier inconnu' };
        f.setTrashed(true); // va dans la corbeille Drive : récupérable 30 jours
        return { ok: true };
      }

      default:
        return { ok: false, err: 'Action inconnue' };
    }
  } catch (err) {
    return { ok: false, err: String(err && err.message || err) };
  } finally {
    lock.releaseLock();
  }
}

function estDansDossier_(f) {
  const id = dossier_().getId(), parents = f.getParents();
  while (parents.hasNext()) if (parents.next().getId() === id) return true;
  return false;
}

function trouverLigne_(uid) {
  const sh = tableur_().getSheets()[0];
  const n = sh.getLastRow();
  if (n < 2) return null;
  const vals = sh.getRange(2, 1, n - 1, ENTETES.length).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0] === uid) {
      const lien = String(vals[i][10] || '');
      const m = lien.match(/[-\w]{25,}/);
      return { row: i + 2, fichierId: m ? m[0] : '' };
    }
  }
  return null;
}

function majTableur_(b, f, ligne) {
  const sh = tableur_().getSheets()[0];
  const valeurs = [[b.uid, b.client || '', b.nom || '', b.numero || '', b.email || '', b.telephone || '', b.adresse || '',
    b.echelle ? '1:' + b.echelle : '', typeof b.total === 'number' ? Math.round(b.total * 100) / 100 : '', new Date(), f.getUrl()]];
  const row = ligne ? ligne.row : sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, ENTETES.length).setValues(valeurs);
}

/** Copie complète datée du dossier (déclencheur hebdomadaire créé par installer). */
function sauvegardeHebdo() {
  const cible = sousDossier_('Sauvegardes').createFolder(Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyy-MM-dd'));
  const it = dossier_().getFilesByType(MimeType.PLAIN_TEXT);
  while (it.hasNext()) { const f = it.next(); f.makeCopy(f.getName(), cible); }
}

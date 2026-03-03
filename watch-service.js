const { google } = require('googleapis');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

// ─── CONFIGURATION ──────────────────────────────────────────────────────
const API_URL = process.env.API_URL || 'https://backend-rappels-whatsapp-production.up.railway.app/api';
const ADMIN_TOKEN = process.env.ADMIN_TOKEN;
const FOLDER_NAME = process.env.DRIVE_FOLDER_NAME || 'Rappels_RDV_WhatsApp';
const CHECK_INTERVAL = parseInt(process.env.CHECK_INTERVAL) || 300000; // 5 minutes

// ─── GOOGLE AUTH ────────────────────────────────────────────────────────
function getAuthClient() {
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  const token = JSON.parse(process.env.GOOGLE_TOKEN);
  
  const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, (redirect_uris && redirect_uris[0]) || '');
  oAuth2Client.setCredentials(token);
  
  return oAuth2Client;
}

// ─── NETTOYAGE TELEPHONE ────────────────────────────────────────────────
function cleanPhone(raw) {
  if (!raw) return null;
  let phone = String(raw).replace(/[\s.\-\/\(\)]/g, '').trim();
  
  if (phone.startsWith('+')) {
    return phone;
  }
  if (phone.startsWith('00')) {
    return '+' + phone.substring(2);
  }
  if (phone.length === 10 && phone.startsWith('0')) {
    return '+32' + phone.substring(1);
  }
  if (phone.length === 9 && !phone.startsWith('0')) {
    return '+32' + phone;
  }
  return '+' + phone;
}

// ─── NETTOYAGE HEURE ───────────────────────────────────────────────────
function cleanTime(raw) {
  if (!raw) return '';
  const str = String(raw).trim();
  
  // Format HH:MM:SS → HH:MM
  const match = str.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (match) {
    return `${match[1].padStart(2, '0')}:${match[2]}`;
  }
  
  // Format décimal Excel (ex: 0.354166... = 08:30)
  const num = parseFloat(str);
  if (!isNaN(num) && num >= 0 && num < 1) {
    const totalMinutes = Math.round(num * 24 * 60);
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  }
  
  return str;
}

// ─── NETTOYAGE DATE ────────────────────────────────────────────────────
function cleanDate(raw) {
  if (!raw) return null;
  const str = String(raw).trim();
  
  // Format DD-MM-YYYY ou DD/MM/YYYY
  const match = str.match(/(\d{1,2})[\-\/](\d{1,2})[\-\/](\d{4})/);
  if (match) {
    return `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`;
  }
  
  // Format YYYY-MM-DD (déjà bon)
  const match2 = str.match(/(\d{4})[\-\/](\d{1,2})[\-\/](\d{1,2})/);
  if (match2) {
    return `${match2[1]}-${match2[2].padStart(2, '0')}-${match2[3].padStart(2, '0')}`;
  }
  
  // Excel serial date number
  const num = parseFloat(str);
  if (!isNaN(num) && num > 40000 && num < 60000) {
    const date = new Date((num - 25569) * 86400 * 1000);
    return date.toISOString().split('T')[0];
  }
  
  return str;
}

// ─── PARSE EXCEL AVEC XLSX (NODE.JS PUR) ────────────────────────────────
function parseExcelBuffer(buffer, fileName) {
  const XLSX = require('xlsx');
  
  let workbook;
  
  try {
    workbook = XLSX.read(buffer, { type: 'buffer' });
    console.log('   ✅ Lecture avec type buffer réussie');
  } catch (e1) {
    console.log('   ⚠️ Échec type buffer, essai avec array...');
    try {
      const uint8 = new Uint8Array(buffer);
      workbook = XLSX.read(uint8, { type: 'array' });
      console.log('   ✅ Lecture avec type array réussie');
    } catch (e2) {
      console.log('   ⚠️ Échec type array, essai avec base64...');
      try {
        const base64 = buffer.toString('base64');
        workbook = XLSX.read(base64, { type: 'base64' });
        console.log('   ✅ Lecture avec type base64 réussie');
      } catch (e3) {
        console.log('   ⚠️ Échec type base64, essai en écrivant sur disque...');
        const tmpPath = '/tmp/temp_import' + path.extname(fileName);
        fs.writeFileSync(tmpPath, buffer);
        try {
          workbook = XLSX.readFile(tmpPath);
          console.log('   ✅ Lecture via fichier disque réussie');
        } catch (e4) {
          try {
            workbook = XLSX.readFile(tmpPath, { type: 'binary' });
            console.log('   ✅ Lecture via fichier binaire réussie');
          } catch (e5) {
            try {
              const binary = fs.readFileSync(tmpPath, 'binary');
              workbook = XLSX.read(binary, { type: 'binary' });
              console.log('   ✅ Lecture binaire manuelle réussie');
            } catch (e6) {
              throw new Error('Impossible de lire le fichier Excel: ' + e6.message);
            }
          }
        } finally {
          try { fs.unlinkSync(tmpPath); } catch(e) {}
        }
      }
    }
  }
  
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  
  console.log(`   📊 ${rows.length} lignes trouvées (dont 1 en-tête)`);
  
  if (rows.length < 2) {
    throw new Error('Fichier vide ou sans données');
  }
  
  const patients = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 6) continue;
    
    const dateRaw = row[0];
    const prenom = String(row[3] || '').trim();
    const nom = String(row[4] || '').trim();
    const heureRaw = row[5];
    const gsmRaw = row[9];
    
    if (!prenom && !nom) continue;
    
    const date = cleanDate(dateRaw);
    const heure = cleanTime(heureRaw);
    const telephone = cleanPhone(gsmRaw);
    
    if (!telephone) {
      console.log(`   ⚠️ Pas de téléphone pour ${prenom} ${nom}, ignoré`);
      continue;
    }
    
    patients.push({ nom, prenom, telephone, date_rdv: date, heure_rdv: heure });
    console.log(`   ✅ ${prenom} ${nom} - ${telephone} - ${date} ${heure}`);
  }
  
  return patients;
}

// ─── VÉRIFIER SI PATIENT EXISTE DÉJÀ ───────────────────────────────────
async function patientExists(patient, existingPatients) {
  // Normaliser l'heure pour comparaison (enlever les secondes si présentes)
  const normalizeHeure = (h) => {
    if (!h) return '';
    const match = h.match(/(\d{2}):(\d{2})/);
    return match ? `${match[1]}:${match[2]}` : h;
  };
  
  const patientHeure = normalizeHeure(patient.heure_rdv);
  
  return existingPatients.some(p => 
    p.nom.toLowerCase() === patient.nom.toLowerCase() &&
    p.prenom.toLowerCase() === patient.prenom.toLowerCase() &&
    p.date_rdv === patient.date_rdv &&
    normalizeHeure(p.heure_rdv) === patientHeure
  );
}

// ─── RÉCUPÉRER TOUS LES PATIENTS EXISTANTS ─────────────────────────────
async function getExistingPatients() {
  try {
    const response = await axios.get(`${API_URL}/patients`, {
      headers: { Authorization: `Bearer ${ADMIN_TOKEN}` }
    });
    return response.data.patients || response.data || [];
  } catch (err) {
    console.error('   ⚠️ Erreur récupération patients existants:', err.message);
    return [];
  }
}

// ─── IMPORT PATIENTS DANS LA BASE (AVEC ANTI-DOUBLON) ──────────────────
async function importPatients(patients) {
  let success = 0;
  let skipped = 0;
  let errors = 0;
  
  // Récupérer les patients existants pour vérifier les doublons
  console.log('   🔍 Vérification des doublons...');
  const existingPatients = await getExistingPatients();
  console.log(`   📋 ${existingPatients.length} patients déjà en base`);
  
  for (const patient of patients) {
    try {
      // Vérifier si le patient existe déjà (même nom + date + heure)
      if (await patientExists(patient, existingPatients)) {
        console.log(`   ⏭️ Doublon ignoré: ${patient.prenom} ${patient.nom} (${patient.date_rdv} ${patient.heure_rdv})`);
        skipped++;
        continue;
      }
      
      // Créer le patient
      const res = await axios.post(`${API_URL}/patients`, {
        ...patient,
        statut_envoi: 'en_attente',
        reponse: 'en_attente',
        nouveau: true
      }, {
        headers: { Authorization: `Bearer ${ADMIN_TOKEN}` }
      });
      
      // Récupérer l'ID du patient créé et mettre nouveau=true via PATCH
      const patientId = res.data?.patient?.id;
      if (patientId) {
        await axios.patch(`${API_URL}/patients/${patientId}`, { nouveau: true }, {
          headers: { Authorization: `Bearer ${ADMIN_TOKEN}` }
        });
        
        // Ajouter à la liste des patients existants pour éviter les doublons dans le même fichier
        existingPatients.push({
          nom: patient.nom,
          prenom: patient.prenom,
          date_rdv: patient.date_rdv,
          heure_rdv: patient.heure_rdv
        });
      }
      
      success++;
    } catch (err) {
      const status = err.response?.status;
      const msg = err.response?.data?.error || err.message;
      console.error(`   ❌ Erreur import ${patient.prenom} ${patient.nom}: ${status} - ${msg}`);
      errors++;
    }
  }
  
  return { success, skipped, errors };
}

// ─── SURVEILLANCE GOOGLE DRIVE ─────────────────────────────────────────
async function checkGoogleDrive() {
  try {
    const auth = getAuthClient();
    const drive = google.drive({ version: 'v3', auth });
    
    const folderRes = await drive.files.list({
      q: `name='${FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
      fields: 'files(id, name)',
    });
    
    if (!folderRes.data.files || folderRes.data.files.length === 0) {
      console.log(`   📁 Dossier "${FOLDER_NAME}" non trouvé`);
      return;
    }
    
    const folderId = folderRes.data.files[0].id;
    
    const filesRes = await drive.files.list({
      q: `'${folderId}' in parents and trashed=false and (mimeType='application/vnd.ms-excel' or mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or name contains '.xls')`,
      fields: 'files(id, name, createdTime, mimeType)',
      orderBy: 'createdTime desc',
    });
    
    const files = filesRes.data.files || [];
    
    if (files.length === 0) {
      console.log(`   📭 Aucun fichier Excel dans le dossier`);
      return;
    }
    
    console.log(`   📄 ${files.length} fichier(s) trouvé(s)`);
    
    for (const file of files) {
      console.log('');
      console.log('   ════════════════════════════════════════════════════');
      console.log(`   📄 Traitement: ${file.name}`);
      console.log('   ════════════════════════════════════════════════════');
      
      try {
        console.log('   📦 Téléchargement...');
        const response = await drive.files.get(
          { fileId: file.id, alt: 'media' },
          { responseType: 'arraybuffer' }
        );
        
        const buffer = Buffer.from(response.data);
        console.log(`   📦 Taille: ${buffer.length} octets`);
        
        console.log('   📊 Lecture des données...');
        const patients = parseExcelBuffer(buffer, file.name);
        
        if (patients.length === 0) {
          console.log('   ⚠️ Aucun patient trouvé dans le fichier');
        } else {
          console.log(`   📤 Import de ${patients.length} patient(s)...`);
          const result = await importPatients(patients);
          console.log(`   ✅ ${result.success} importé(s), ${result.skipped} doublon(s) ignoré(s), ${result.errors} erreur(s)`);
        }
        
        console.log('   🗑️ Suppression du fichier (RGPD)...');
        await drive.files.delete({ fileId: file.id });
        console.log('   ✅ Fichier supprimé du Drive');
        
      } catch (err) {
        console.error(`   ❌ ERREUR: ${err.message}`);
      }
    }
    
  } catch (err) {
    console.error(`   ❌ Erreur surveillance: ${err.message}`);
  }
}

// ─── DÉMARRAGE ─────────────────────────────────────────────────────────
console.log('');
console.log('═══════════════════════════════════════════════════════════');
console.log('       SERVICE DE SURVEILLANCE AUTOMATIQUE');
console.log('       Google Drive → Base de données');
console.log('       (avec détection anti-doublon)');
console.log('═══════════════════════════════════════════════════════════');
console.log('');
console.log(`👀 Surveillance de Google Drive démarrée...`);
console.log(`📁 Dossier: ${FOLDER_NAME}`);
console.log(`⏰ Vérification toutes les ${CHECK_INTERVAL / 1000 / 60} minutes`);
console.log('');

checkGoogleDrive();
setInterval(checkGoogleDrive, CHECK_INTERVAL);

process.on('SIGTERM', () => {
  console.log('🛑 Arrêt du service de surveillance');
  process.exit(0);
});

// ═══════════════════════════════════════════════════════════════════════
// SERVICE DE SURVEILLANCE AUTOMATIQUE - GOOGLE DRIVE → BASE DE DONNÉES
// Vérifie Google Drive toutes les 5 minutes et importe automatiquement
// ═══════════════════════════════════════════════════════════════════════

const XLSX = require('xlsx');
const xlsx = require('node-xlsx');
const { google } = require('googleapis');
const fs = require('fs');
const axios = require('axios');
const path = require('path');
require('dotenv').config();

// Configuration
const FOLDER_NAME = 'Rappels_RDV_WhatsApp';
const CHECK_INTERVAL = 5 * 60 * 1000; // 5 minutes
const API_URL = process.env.API_URL || 'http://localhost:5000/api';
const PROCESSED_FILES_LOG = './processed_files.json';

// ─── GARDER LA TRACE DES FICHIERS DÉJÀ TRAITÉS ────────────────────────
function loadProcessedFiles() {
  if (fs.existsSync(PROCESSED_FILES_LOG)) {
    return JSON.parse(fs.readFileSync(PROCESSED_FILES_LOG, 'utf8'));
  }
  return [];
}

function saveProcessedFile(fileId, fileName) {
  const processed = loadProcessedFiles();
  processed.push({ fileId, fileName, processedAt: new Date().toISOString() });
  fs.writeFileSync(PROCESSED_FILES_LOG, JSON.stringify(processed, null, 2));
}

function isFileProcessed(fileId) {
  const processed = loadProcessedFiles();
  return processed.some(f => f.fileId === fileId);
}

// ─── AUTHENTIFICATION GOOGLE DRIVE ────────────────────────────────────
async function authorize() {
  try {
    const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
    const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

    const token = JSON.parse(process.env.GOOGLE_TOKEN);
    oAuth2Client.setCredentials(token);
    return oAuth2Client;
  } catch (error) {
    console.error('❌ Erreur authentification Google:', error.message);
    throw error;
  }
}

// ─── RÉCUPÉRER LES NOUVEAUX FICHIERS EXCEL ────────────────────────────
async function getNewExcelFiles(auth) {
  const drive = google.drive({ version: 'v3', auth });

  // 1. Trouver le dossier
  const folderResponse = await drive.files.list({
    q: `name='${FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder'`,
    fields: 'files(id, name)',
    spaces: 'drive'
  });

  if (folderResponse.data.files.length === 0) {
    console.log(`⚠️  Dossier "${FOLDER_NAME}" introuvable`);
    return [];
  }

  const folderId = folderResponse.data.files[0].id;

  // 2. Récupérer tous les fichiers Excel non traités
  const filesResponse = await drive.files.list({
    q: `'${folderId}' in parents and (name contains '.xls' or name contains '.xlsx')`,
    orderBy: 'modifiedTime desc',
    fields: 'files(id, name, modifiedTime)',
    pageSize: 10
  });

  // Filtrer les fichiers déjà traités
  const newFiles = filesResponse.data.files.filter(file => !isFileProcessed(file.id));
  
  return { drive, folderId, newFiles };
}

// ─── TÉLÉCHARGER UN FICHIER ───────────────────────────────────────────
async function downloadFile(drive, fileId, fileName) {
  const tempPath = path.join('/tmp', fileName);
  const dest = fs.createWriteStream(tempPath);
  
  const response = await drive.files.get(
    { fileId: fileId, alt: 'media' },
    { responseType: 'stream' }
  );

  return new Promise((resolve, reject) => {
    response.data
      .on('end', () => resolve(tempPath))
      .on('error', reject)
      .pipe(dest);
  });
}

// ─── CONVERTIR XLS EN XLSX ────────────────────────────────────────────
async function convertXlsToXlsx(xlsPath) {
  const xlsxPath = xlsPath.replace('.xls', '.xlsx');
  
  return new Promise((resolve, reject) => {
    const { exec } = require('child_process');
    const pythonScript = `
import sys
import pandas as pd
try:
    df = pd.read_excel('${xlsPath}', engine='xlrd', header=None)
    df.to_excel('${xlsxPath}', index=False, header=False, engine='openpyxl')
    print('OK')
except Exception as e:
    print(f'ERROR: {e}')
    sys.exit(1)
`;
    
    exec(`python3 -c "${pythonScript.replace(/\n/g, ';')}"`, (error, stdout, stderr) => {
      if (error || stdout.includes('ERROR')) {
        reject(new Error('Conversion échouée'));
      } else {
        resolve(xlsxPath);
      }
    });
  });
}

// ─── NETTOYER LE NUMÉRO DE TÉLÉPHONE ───────────────────────────────────
function cleanPhoneNumber(phone) {
  if (!phone) return null;
  let cleaned = String(phone).replace(/[\s.\-/]/g, '');
  
  if (cleaned.startsWith('00')) {
    cleaned = '+' + cleaned.substring(2);
  } else if (cleaned.startsWith('+')) {
    // OK
  } else if (cleaned.startsWith('0') && cleaned.length === 10) {
    cleaned = '+32' + cleaned.substring(1);
  } else if (!cleaned.startsWith('0') && cleaned.length === 9) {
    cleaned = '+32' + cleaned;
  } else if (!cleaned.startsWith('+')) {
    cleaned = '+' + cleaned;
  }
  
  return cleaned;
}

// ─── NETTOYER L'HEURE ──────────────────────────────────────────────────
function cleanTime(time) {
  if (!time) return null;
  
  if (typeof time === 'number') {
    const hours = Math.floor(time * 24);
    const minutes = Math.round((time * 24 - hours) * 60);
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  }
  
  const str = String(time);
  const match = str.match(/(\d{1,2}):(\d{2})/);
  if (match) {
    return `${match[1].padStart(2, '0')}:${match[2]}`;
  }
  
  return null;
}

// ─── PARSER LE FICHIER EXCEL ──────────────────────────────────────────
async function parseExcelFile(filePath) {
  try {
    let fileToRead = filePath;
    if (filePath.endsWith('.xls') && !filePath.endsWith('.xlsx')) {
      console.log('🔄 Conversion .xls → .xlsx...');
      fileToRead = await convertXlsToXlsx(filePath);
    }
    
    const workSheetsFromFile = xlsx.parse(fileToRead);
    const sheet = workSheetsFromFile[0];
    const data = sheet.data;
    
    const patients = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[3] || !row[4] || !row[9]) continue;
      
      const dateRdv = row[0];
      const prenom = row[3];
      const nom = row[4];
      const heureRaw = row[5];
      const telephone = row[9];
      
      const telephoneClean = cleanPhoneNumber(telephone);
      const heureClean = cleanTime(heureRaw);
      
      let dateFormatted;
      if (typeof dateRdv === 'number') {
        const excelDate = new Date((dateRdv - 25569) * 86400 * 1000);
        dateFormatted = excelDate.toISOString().split('T')[0];
      } else if (typeof dateRdv === 'string') {
        const parts = dateRdv.split('-');
        if (parts.length === 3) {
          dateFormatted = `${parts[2]}-${parts[1]}-${parts[0]}`;
        } else {
          dateFormatted = dateRdv;
        }
      }
      
      if (!telephoneClean || !heureClean || !dateFormatted) continue;
      
      patients.push({
        nom,
        prenom,
        telephone: telephoneClean,
        date_rdv: dateFormatted,
        heure_rdv: heureClean,
        statut_envoi: 'en_attente',
        reponse: 'en_attente',
        traite: false,
        nouveau: true // MARQUER COMME NOUVEAU
      });
    }
    
    return patients;
  } catch (error) {
    console.error('❌ Erreur parsing:', error.message);
    throw error;
  }
}

// ─── IMPORTER DANS LA BASE DE DONNÉES ─────────────────────────────────
async function importToDatabase(patients) {
  const token = process.env.ADMIN_TOKEN;
  let imported = 0;
  let skipped = 0;
  
  for (const patient of patients) {
    try {
      await axios.post(
        `${API_URL}/patients`,
        patient,
        {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
          }
        }
      );
      imported++;
    } catch (error) {
      if (error.response?.status === 409) {
        skipped++;
      } else {
        console.error(`❌ Erreur import ${patient.prenom} ${patient.nom}:`, error.message);
      }
    }
  }
  
  return { imported, skipped };
}

// ─── SUPPRIMER LE FICHIER DE GOOGLE DRIVE ─────────────────────────────
async function deleteFileFromDrive(drive, fileId) {
  try {
    await drive.files.delete({ fileId });
    console.log('✓ Fichier supprimé de Google Drive (RGPD)');
  } catch (error) {
    console.error('⚠️  Impossible de supprimer le fichier:', error.message);
  }
}

// ─── TRAITER UN FICHIER ───────────────────────────────────────────────
async function processFile(drive, file) {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`📄 Traitement: ${file.name}`);
  console.log(`${'='.repeat(60)}`);
  
  try {
    // 1. Télécharger
    console.log('📥 Téléchargement...');
    const filePath = await downloadFile(drive, file.id, file.name);
    
    // 2. Parser
    console.log('📊 Lecture des données...');
    const patients = await parseExcelFile(filePath);
    console.log(`✓ ${patients.length} patients trouvés`);
    
    if (patients.length === 0) {
      console.log('⚠️  Aucun patient à importer');
      return;
    }
    
    // 3. Importer
    console.log('💾 Import dans la base de données...');
    const { imported, skipped } = await importToDatabase(patients);
    console.log(`✓ ${imported} patients importés, ${skipped} déjà existants`);
    
    // 4. Marquer comme traité
    saveProcessedFile(file.id, file.name);
    
    // 5. Supprimer de Google Drive
    console.log('🗑️  Suppression du fichier (RGPD)...');
    await deleteFileFromDrive(drive, file.id);
    
    // 6. Nettoyer les fichiers temporaires
    try {
      fs.unlinkSync(filePath);
      if (filePath.endsWith('.xls')) {
        fs.unlinkSync(filePath.replace('.xls', '.xlsx'));
      }
    } catch (e) {}
    
    console.log('✅ Traitement terminé avec succès\n');
    
  } catch (error) {
    console.error('❌ ERREUR:', error.message);
  }
}

// ─── BOUCLE DE SURVEILLANCE ───────────────────────────────────────────
async function watchDrive() {
  console.log('👁️  Surveillance de Google Drive démarrée...');
  console.log(`📁 Dossier: ${FOLDER_NAME}`);
  console.log(`⏰ Vérification toutes les ${CHECK_INTERVAL / 60000} minutes\n`);
  
  const check = async () => {
    try {
      const auth = await authorize();
      const { drive, newFiles } = await getNewExcelFiles(auth);
      
      if (newFiles.length === 0) {
        console.log(`[${new Date().toLocaleTimeString()}] ✓ Aucun nouveau fichier`);
        return;
      }
      
      console.log(`[${new Date().toLocaleTimeString()}] 🆕 ${newFiles.length} nouveau(x) fichier(s) détecté(s) !`);
      
      for (const file of newFiles) {
        await processFile(drive, file);
      }
      
    } catch (error) {
      console.error('❌ Erreur surveillance:', error.message);
    }
  };
  
  // Première vérification immédiate
  await check();
  
  // Puis toutes les 5 minutes
  setInterval(check, CHECK_INTERVAL);
}

// ─── DÉMARRAGE ────────────────────────────────────────────────────────
if (require.main === module) {
  console.log('═══════════════════════════════════════════════════════════');
  console.log('     SERVICE DE SURVEILLANCE AUTOMATIQUE');
  console.log('     Google Drive → Base de données');
  console.log('═══════════════════════════════════════════════════════════\n');
  
  watchDrive().catch(error => {
    console.error('❌ ERREUR FATALE:', error);
    process.exit(1);
  });
}

module.exports = { watchDrive };


import express from 'express';
import path from 'path';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI } from "@google/genai";
import axios from 'axios';
import dotenv from 'dotenv';
import fs from 'fs';
import os from 'os';
import FormData from 'form-data';
import * as XLSX from 'xlsx';
import puppeteer from 'puppeteer-core';
import { execFileSync } from 'child_process';

dotenv.config();

const app = express();
const PORT = 5000;

app.use(express.json());

// Disable cache only for HTML navigation requests so Replit preview loads fresh,
// but allow JS/CSS/assets to be cached normally (prevents slow white-screen on load)
if (process.env.NODE_ENV !== 'production') {
  app.use((req, res, next) => {
    const accept = req.headers['accept'] || '';
    if (accept.includes('text/html')) {
      res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
      res.setHeader('Pragma', 'no-cache');
    }
    next();
  });
}

// Serve saved print images statically
const PRINTS_DIR = path.join(process.cwd(), 'prints');
if (!fs.existsSync(PRINTS_DIR)) fs.mkdirSync(PRINTS_DIR, { recursive: true });
app.use('/prints', express.static(PRINTS_DIR));

// ── Auto-upload helper: upload a local file to Telegram archive channel ──────
// Returns `tg:{fileId}:{encodedFilename}` on success, or the original local URL on failure.
async function autoUploadToTelegram(localFilePath: string, clientCaption: string, chatIdOverride?: string): Promise<string | null> {
  try {
    const config = readTelegramConfig();
    const archiveChatId = chatIdOverride || config.printArchiveChatId || config.archiveChatId || config.alertChatId || '';
    if (!config.token || !archiveChatId) return null;
    const filename = path.basename(localFilePath);
    if (!fs.existsSync(localFilePath)) return null;
    const form = new FormData();
    form.append('chat_id', archiveChatId);
    form.append('document', fs.createReadStream(localFilePath), { filename });
    form.append('caption', `📁 ${clientCaption}\n🗂 ${filename}`);
    const response = await axios.post(
      `https://api.telegram.org/bot${config.token}/sendDocument`,
      form,
      { headers: form.getHeaders(), maxContentLength: Infinity, maxBodyLength: Infinity, timeout: 60000 }
    );
    const fileId = response.data?.result?.document?.file_id;
    if (fileId) {
      try { fs.unlinkSync(localFilePath); } catch {}
      console.log(`[AUTO-TG] Uploaded & deleted: ${filename}`);
      return `tg:${fileId}:${encodeURIComponent(filename)}`;
    }
  } catch (err: any) {
    const tgErr = err.response?.data?.description || err.message;
    console.warn(`[AUTO-TG] Upload failed (non-fatal): ${tgErr}`);
  }
  return null;
}
// ─────────────────────────────────────────────────────────────────────────────

// Initialize Gemini
const ai = new GoogleGenAI({ 
  apiKey: process.env.GEMINI_API_KEY as string,
  httpOptions: {
    headers: {
      'User-Agent': 'aistudio-build',
    }
  }
});

// JSON Database Storage
const ENTRIES_FILE = path.join(process.cwd(), 'data_entries.json');
const OFFICES_FILE = path.join(process.cwd(), 'offices.json');
const TELEGRAM_FILE = path.join(process.cwd(), 'telegram_config.json');
const CREDENTIALS_FILE = path.join(process.cwd(), 'admin_credentials.json');
const AUTHENTICATED_CHATS_FILE = path.join(process.cwd(), 'authenticated_chats.json');
const USERS_FILE = path.join(process.cwd(), 'users.json');
const SITE_IDENTITY_FILE = path.join(process.cwd(), 'site_identity.json');

function readUsers(): any[] {
  try {
    if (fs.existsSync(USERS_FILE)) {
      const fileContent = fs.readFileSync(USERS_FILE, 'utf-8');
      return JSON.parse(fileContent);
    }
  } catch (e) {
    console.error('Error reading users file:', e);
  }
  return [
    {
      id: 'admin-default',
      username: 'admin',
      password: '1234',
      name: 'مدير النظام (الرئيسي)',
      permissions: ['admin', 'edit_entries', 'excel_sync', 'bot_settings', 'update_checking', 'add_clients', 'edit_clients', 'download_data']
    }
  ];
}

function writeUsers(users: any[]) {
  try {
    fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2), 'utf-8');
  } catch (e) {
    console.error('Error writing users file:', e);
  }
}

function readOffices(): any[] {
  try {
    if (fs.existsSync(OFFICES_FILE)) {
      return JSON.parse(fs.readFileSync(OFFICES_FILE, 'utf-8'));
    }
  } catch (e) { console.error('Error reading offices file:', e); }
  return [];
}

function writeOffices(offices: any[]) {
  try { fs.writeFileSync(OFFICES_FILE, JSON.stringify(offices, null, 2), 'utf-8'); }
  catch (e) { console.error('Error writing offices file:', e); }
}

function readEntries(): any[] {
  try {
    if (fs.existsSync(ENTRIES_FILE)) {
      const fileContent = fs.readFileSync(ENTRIES_FILE, 'utf-8');
      return JSON.parse(fileContent);
    }
  } catch (e) {
    console.error('Error reading entries file:', e);
  }
  return [];
}

function writeEntries(entries: any[]) {
  try {
    fs.writeFileSync(ENTRIES_FILE, JSON.stringify(entries, null, 2), 'utf-8');
  } catch (e) {
    console.error('Error writing entries file:', e);
  }
}

function readTelegramConfig() {
  try {
    if (fs.existsSync(TELEGRAM_FILE)) {
      const fileContent = fs.readFileSync(TELEGRAM_FILE, 'utf-8');
      return JSON.parse(fileContent);
    }
  } catch (e) {
    console.error('Error reading telegram config file:', e);
  }
  return { token: '', enabled: false, alertChatId: '' };
}

function writeTelegramConfig(config: any) {
  try {
    fs.writeFileSync(TELEGRAM_FILE, JSON.stringify(config, null, 2), 'utf-8');
  } catch (e) {
    console.error('Error writing telegram config file:', e);
  }
}

function readCredentials() {
  try {
    if (fs.existsSync(CREDENTIALS_FILE)) {
      return JSON.parse(fs.readFileSync(CREDENTIALS_FILE, 'utf-8'));
    }
  } catch (e) {
    console.error('Error reading credentials file:', e);
  }
  return { username: 'admin', password: '1234' };
}

function writeCredentials(creds: any) {
  try {
    fs.writeFileSync(CREDENTIALS_FILE, JSON.stringify(creds, null, 2), 'utf-8');
  } catch (e) {
    console.error('Error writing credentials file:', e);
  }
}

// ─── Developer Panel Auth ──────────────────────────────────────────────────
const DEV_CREDENTIALS_FILE = path.join(process.cwd(), 'dev_credentials.json');
const DEV_TOKEN = 'dev_session_' + Math.random().toString(36).slice(2);

function readDevCredentials() {
  try {
    if (fs.existsSync(DEV_CREDENTIALS_FILE)) {
      return JSON.parse(fs.readFileSync(DEV_CREDENTIALS_FILE, 'utf-8'));
    }
  } catch {}
  return { username: 'devmaster', password: 'Dev@9999' };
}

function writeDevCredentials(creds: any) {
  try {
    fs.writeFileSync(DEV_CREDENTIALS_FILE, JSON.stringify(creds, null, 2), 'utf-8');
  } catch (e) {
    console.error('Error writing dev credentials:', e);
  }
}

function checkDevToken(req: any, res: any): boolean {
  const token = req.headers['x-dev-token'];
  if (!token || token !== DEV_TOKEN) {
    res.status(401).json({ error: 'Unauthorized' });
    return false;
  }
  return true;
}

app.post('/api/dev/login', (req: any, res: any) => {
  const { username, password } = req.body;
  if (!username || !password) {
    return res.status(400).json({ error: 'Missing credentials' });
  }
  const creds = readDevCredentials();
  if (username.trim() === creds.username && password.trim() === creds.password) {
    return res.json({ success: true, token: DEV_TOKEN });
  }
  res.status(401).json({ error: 'بيانات الدخول خاطئة' });
});

app.get('/api/dev/stats', (req: any, res: any) => {
  if (!checkDevToken(req, res)) return;
  try {
    const entries = readEntries();
    const offices = readOffices();
    const users = readUsers();
    const entriesSize = fs.existsSync(ENTRIES_FILE) ? fs.statSync(ENTRIES_FILE).size : 0;
    const memUsed = process.memoryUsage();
    const formatBytes = (b: number) => b > 1048576 ? (b / 1048576).toFixed(1) + ' MB' : (b / 1024).toFixed(1) + ' KB';
    res.json({
      stats: {
        totalEntries: entries.length,
        totalOffices: offices.length,
        totalUsers: users.length,
        dataSize: formatBytes(entriesSize),
      },
      sys: {
        uptime: process.uptime(),
        nodeVersion: process.version,
        platform: process.platform,
        memUsed: formatBytes(memUsed.heapUsed),
        memTotal: formatBytes(memUsed.heapTotal),
        cpuLoad: os.loadavg()[0].toFixed(2),
      },
      logs: []
    });
  } catch (e: any) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/dev/credentials', (req: any, res: any) => {
  if (!checkDevToken(req, res)) return;
  const { username, password } = req.body;
  if (!username || !password) return res.status(400).json({ error: 'Missing fields' });
  writeDevCredentials({ username: username.trim(), password: password.trim() });
  res.json({ success: true });
});

app.post('/api/dev/clear-logs', (req: any, res: any) => {
  if (!checkDevToken(req, res)) return;
  res.json({ success: true });
});

app.post('/api/dev/reset-entries', (req: any, res: any) => {
  if (!checkDevToken(req, res)) return;
  try {
    fs.writeFileSync(ENTRIES_FILE, JSON.stringify([], null, 2), 'utf-8');
    res.json({ success: true });
  } catch (e: any) {
    res.status(500).json({ error: e.message });
  }
});
// ──────────────────────────────────────────────────────────────────────────

function readAuthenticatedChats(): number[] {
  try {
    if (fs.existsSync(AUTHENTICATED_CHATS_FILE)) {
      return JSON.parse(fs.readFileSync(AUTHENTICATED_CHATS_FILE, 'utf-8'));
    }
  } catch (e) {
    console.error('Error reading authenticated chats:', e);
  }
  return [];
}

function writeAuthenticatedChats(chats: number[]) {
  try {
    fs.writeFileSync(AUTHENTICATED_CHATS_FILE, JSON.stringify(chats, null, 2), 'utf-8');
  } catch (e) {
    console.error('Error writing authenticated chats:', e);
  }
}

// REST Api for Two-Way synchronization with React state
app.get('/api/entries', (req, res) => {
  res.json(readEntries());
});

app.post('/api/entries', (req, res) => {
  const { entries } = req.body;
  if (Array.isArray(entries)) {
    writeEntries(entries);
    res.json({ success: true });
  } else {
    res.status(400).json({ error: 'Entries must be an array' });
  }
});

// GET /api/prints — list only print PDF files (برنت_) in the prints directory
app.get('/api/prints', (req, res) => {
  try {
    const files = fs.readdirSync(PRINTS_DIR)
      .filter((f: string) => f.toLowerCase().endsWith('.pdf') && f.startsWith('برنت_'))
      .map((f: string) => {
        const filePath = path.join(PRINTS_DIR, f);
        const stats = fs.statSync(filePath);
        return { filename: f, url: `/prints/${f}`, size: stats.size, createdAt: stats.mtimeMs };
      })
      .sort((a: any, b: any) => b.createdAt - a.createdAt);
    const totalSize = files.reduce((sum: number, f: any) => sum + f.size, 0);
    res.json({ files, totalSize });
  } catch {
    res.json({ files: [], totalSize: 0 });
  }
});

// GET /api/visas — list only visa PDF files (تاشيره_) in the prints directory
app.get('/api/visas', (req, res) => {
  try {
    const files = fs.readdirSync(PRINTS_DIR)
      .filter((f: string) => f.toLowerCase().endsWith('.pdf') && f.startsWith('تاشيره_'))
      .map((f: string) => {
        const filePath = path.join(PRINTS_DIR, f);
        const stats = fs.statSync(filePath);
        return { filename: f, url: `/prints/${f}`, size: stats.size, createdAt: stats.mtimeMs };
      })
      .sort((a: any, b: any) => b.createdAt - a.createdAt);
    const totalSize = files.reduce((sum: number, f: any) => sum + f.size, 0);
    res.json({ files, totalSize });
  } catch {
    res.json({ files: [], totalSize: 0 });
  }
});

// DELETE /api/entries/pdf — delete the PDF file and clear its URL from the matching entry
app.delete('/api/entries/pdf', (req, res) => {
  const { printImageUrl } = req.body as { printImageUrl?: string };
  if (!printImageUrl) return res.json({ success: false, error: 'no url' });
  try {
    const filename = path.basename(printImageUrl);
    const filePath = path.join(PRINTS_DIR, filename);
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
      console.log(`[PDF] Deleted: ${filePath}`);
    }
    // Clear the URL from the matching entry
    const isVisaFile = filename.startsWith('تاشيره_');
    const isPrintFile = filename.startsWith('برنت_');
    if (isVisaFile || isPrintFile) {
      const entries = readEntries();
      const urlPath = `/prints/${filename}`;
      let changed = false;
      const updated = entries.map((e: any) => {
        if (isVisaFile && (e.visaImageUrl === urlPath || e.visaImageUrl === printImageUrl)) {
          changed = true;
          return { ...e, visaImageUrl: undefined, visaImageBase64: undefined };
        }
        if (isPrintFile && (e.printImageUrl === urlPath || e.printImageUrl === printImageUrl)) {
          changed = true;
          return { ...e, printImageUrl: undefined, printImageBase64: undefined };
        }
        return e;
      });
      if (changed) {
        writeEntries(updated);
        console.log(`[PDF] Cleared URL from entry for file: ${filename}`);
      }
    }
    res.json({ success: true });
  } catch (err: any) {
    console.warn('[PDF] Delete failed:', err.message);
    res.json({ success: false, error: err.message });
  }
});

app.get('/api/offices', (req, res) => { res.json(readOffices()); });
app.post('/api/offices', (req, res) => {
  const { offices } = req.body;
  if (Array.isArray(offices)) { writeOffices(offices); res.json({ success: true }); }
  else res.status(400).json({ error: 'Offices must be an array' });
});

// ── Site Identity (persistent on server) ──────────────────────────────────────
function readSiteIdentity() {
  try {
    if (fs.existsSync(SITE_IDENTITY_FILE)) {
      return JSON.parse(fs.readFileSync(SITE_IDENTITY_FILE, 'utf-8'));
    }
  } catch {}
  return null;
}
function writeSiteIdentity(identity: any) {
  fs.writeFileSync(SITE_IDENTITY_FILE, JSON.stringify(identity, null, 2), 'utf-8');
}
app.get('/api/site-identity', (req, res) => {
  const identity = readSiteIdentity();
  if (identity) res.json(identity);
  else res.json(null);
});
// Use a higher body limit (2 MB) to accommodate base64-encoded logo images
app.post('/api/site-identity', express.json({ limit: '2mb' }), (req, res) => {
  const identity = req.body;
  if (!identity || typeof identity !== 'object') {
    return res.status(400).json({ error: 'Invalid identity object' });
  }
  writeSiteIdentity(identity);
  res.json({ success: true });
});
// ─────────────────────────────────────────────────────────────────────────────

// ── Auto-Check Configuration (persistent on server) ───────────────────────────
const AUTOCHECK_CONFIG_FILE = path.join(process.cwd(), 'autocheck_config.json');
function readAutoCheckConfig() {
  try {
    if (fs.existsSync(AUTOCHECK_CONFIG_FILE)) {
      return JSON.parse(fs.readFileSync(AUTOCHECK_CONFIG_FILE, 'utf-8'));
    }
  } catch {}
  return null;
}
function writeAutoCheckConfig(config: any) {
  fs.writeFileSync(AUTOCHECK_CONFIG_FILE, JSON.stringify(config, null, 2), 'utf-8');
}
app.get('/api/autocheck-config', (req, res) => {
  const config = readAutoCheckConfig();
  res.json(config || {});
});
app.post('/api/autocheck-config', (req, res) => {
  const config = req.body;
  if (!config || typeof config !== 'object') return res.status(400).json({ error: 'Invalid config' });
  writeAutoCheckConfig(config);
  res.json({ success: true });
});
// ─────────────────────────────────────────────────────────────────────────────

// ── Google Sheets Configuration (persistent on server) ────────────────────────
const GOOGLE_SHEETS_CONFIG_FILE = path.join(process.cwd(), 'google_sheets_config.json');
function readGoogleSheetsConfig() {
  try {
    if (fs.existsSync(GOOGLE_SHEETS_CONFIG_FILE)) {
      return JSON.parse(fs.readFileSync(GOOGLE_SHEETS_CONFIG_FILE, 'utf-8'));
    }
  } catch {}
  return null;
}
function writeGoogleSheetsConfig(config: any) {
  fs.writeFileSync(GOOGLE_SHEETS_CONFIG_FILE, JSON.stringify(config, null, 2), 'utf-8');
}
app.get('/api/google-sheets-config', (req, res) => {
  const config = readGoogleSheetsConfig();
  res.json(config || {});
});
app.post('/api/google-sheets-config', (req, res) => {
  const config = req.body;
  if (!config || typeof config !== 'object') return res.status(400).json({ error: 'Invalid config' });
  writeGoogleSheetsConfig(config);
  res.json({ success: true });
});
// ─────────────────────────────────────────────────────────────────────────────

// ── Transaction Status Options (persistent on server) ─────────────────────────
const TX_STATUS_OPTIONS_FILE = path.join(process.cwd(), 'tx_status_options.json');
const DEFAULT_TX_STATUS_OPTIONS = [
  { id: 'ts1', label: 'مستمرة', type: 'مستمرة' },
  { id: 'ts2', label: 'موقفة', type: 'موقفة' },
  { id: 'ts3', label: 'مرتجع تم الترحيل من جديد', type: 'مستمرة' },
  { id: 'ts4', label: 'جاهزة للاستلام', type: 'مستمرة' },
];
function readTxStatusOptions() {
  try {
    if (fs.existsSync(TX_STATUS_OPTIONS_FILE)) {
      return JSON.parse(fs.readFileSync(TX_STATUS_OPTIONS_FILE, 'utf-8'));
    }
  } catch {}
  return DEFAULT_TX_STATUS_OPTIONS;
}
function writeTxStatusOptions(opts: any[]) {
  fs.writeFileSync(TX_STATUS_OPTIONS_FILE, JSON.stringify(opts, null, 2), 'utf-8');
}
app.get('/api/tx-status-options', (req, res) => { res.json(readTxStatusOptions()); });
app.post('/api/tx-status-options', (req, res) => {
  const { options } = req.body;
  if (Array.isArray(options)) { writeTxStatusOptions(options); res.json({ success: true }); }
  else res.status(400).json({ error: 'Options must be an array' });
});

app.get('/api/telegram-config', (req, res) => {
  res.json(readTelegramConfig());
});

// GET /api/tg-download — proxy download from Telegram using stored file_id
app.get('/api/tg-download', async (req, res) => {
  const fileId = req.query.fileId as string;
  const filename = (req.query.filename as string) || 'file.pdf';
  if (!fileId) return res.status(400).json({ error: 'fileId required' });
  try {
    const config = readTelegramConfig();
    if (!config.token) return res.status(503).json({ error: 'Telegram not configured' });
    const fileInfo = await axios.get(`https://api.telegram.org/bot${config.token}/getFile?file_id=${fileId}`);
    const filePath = fileInfo.data.result.file_path;
    const downloadUrl = `https://api.telegram.org/file/bot${config.token}/${filePath}`;
    const fileRes = await axios.get(downloadUrl, { responseType: 'stream' });
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.setHeader('Content-Type', 'application/pdf');
    fileRes.data.pipe(res);
  } catch (err: any) {
    console.error('[TG-DOWNLOAD]', err.message);
    res.status(500).json({ error: 'فشل تحميل الملف من تيليغرام' });
  }
});

// POST /api/archive-upload — archive entry and upload its files to Telegram channel
// POST /api/bulk-upload-to-telegram — upload ALL remaining local files to Telegram
app.post('/api/bulk-upload-to-telegram', async (req, res) => {
  const config = readTelegramConfig();
  const printArchiveChatId = config.printArchiveChatId || config.archiveChatId || config.alertChatId || '';
  const visaArchiveChatId = config.visaArchiveChatId || config.archiveChatId || config.alertChatId || '';
  if (!config.token) return res.status(400).json({ error: 'no_token', message: 'لم يتم ضبط توكن البوت' });
  if (!printArchiveChatId && !visaArchiveChatId) return res.status(400).json({ error: 'no_chat_id', message: 'لم يتم ضبط قنوات الأرشيف' });

  const entries = readEntries();
  let uploaded = 0, skipped = 0, failed = 0;

  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    let changed = false;

    for (const field of ['printImageUrl', 'visaImageUrl'] as const) {
      const url: string | undefined = entry[field];
      if (!url || url.startsWith('tg:') || url.startsWith('data:')) { skipped++; continue; }
      const filename = path.basename(url);
      const filePath = path.join(PRINTS_DIR, filename);
      if (!fs.existsSync(filePath)) { skipped++; continue; }
      const caption = `📁 ${entry.applicantName || entry.passportNumber || ''}\n🗂 ${filename}`;
      const chatId = field === 'printImageUrl' ? printArchiveChatId : visaArchiveChatId;
      const tgUrl = await autoUploadToTelegram(filePath, caption, chatId);
      if (tgUrl) {
        entries[i][field] = tgUrl;
        changed = true;
        uploaded++;
      } else {
        failed++;
      }
    }

    if (changed) writeEntries(entries);
  }

  res.json({ success: true, uploaded, skipped, failed });
});

app.post('/api/archive-upload', async (req, res) => {
  const { entryId } = req.body as { entryId: string };
  if (!entryId) return res.status(400).json({ error: 'entryId required' });

  const config = readTelegramConfig();
  const entries = readEntries();
  const entryIndex = entries.findIndex((e: any) => e.id === entryId);
  if (entryIndex === -1) return res.status(404).json({ error: 'Entry not found' });

  const printArchiveChatId = config.printArchiveChatId || config.archiveChatId || config.alertChatId || '';
  const visaArchiveChatId = config.visaArchiveChatId || config.archiveChatId || config.alertChatId || '';

  if (!config.token) return res.status(400).json({ error: 'no_token', message: 'لم يتم ضبط توكن البوت في إعدادات التيليغرام' });
  if (!printArchiveChatId && !visaArchiveChatId) return res.status(400).json({ error: 'no_chat_id', message: 'لم يتم ضبط قنوات الأرشيف في إعدادات التيليغرام' });

  const entry = { ...entries[entryIndex] };
  const uploadResults: { printFileId?: string; visaFileId?: string; printFilename?: string; visaFilename?: string } = {};
  let lastUploadError = '';

  const uploadFile = async (url: string, chatId: string): Promise<{ fileId: string; filename: string } | null> => {
    if (!url || !chatId) return null;
    try {
      const filename = path.basename(url);
      const filePath = path.join(PRINTS_DIR, filename);
      if (!fs.existsSync(filePath)) return null;
      const form = new FormData();
      form.append('chat_id', chatId);
      form.append('document', fs.createReadStream(filePath), { filename });
      const clientName = entry.applicantName || entry.passportNumber || '';
      form.append('caption', `📁 أرشيف: ${clientName}\n🗂 ${filename}`);
      const response = await axios.post(
        `https://api.telegram.org/bot${config.token}/sendDocument`,
        form,
        { headers: form.getHeaders(), maxContentLength: Infinity, maxBodyLength: Infinity, timeout: 60000 }
      );
      const fileId = response.data?.result?.document?.file_id;
      if (fileId) {
        fs.unlinkSync(filePath);
        console.log(`[ARCHIVE] Uploaded to Telegram & deleted local: ${filename}`);
        return { fileId, filename };
      }
    } catch (err: any) {
      const tgErr = err.response?.data;
      console.error('[ARCHIVE] Upload failed:', err.message, tgErr ? JSON.stringify(tgErr) : '');
      lastUploadError = tgErr?.description || err.message || 'Unknown error';
    }
    return null;
  };

  if (entry.printImageUrl && !entry.printImageUrl.startsWith('tg:')) {
    console.log('[ARCHIVE] Uploading print file:', entry.printImageUrl, '→ chat:', printArchiveChatId);
    const result = await uploadFile(entry.printImageUrl, printArchiveChatId);
    if (result) {
      uploadResults.printFileId = result.fileId;
      uploadResults.printFilename = result.filename;
    }
  }

  if (entry.visaImageUrl && !entry.visaImageUrl.startsWith('tg:')) {
    console.log('[ARCHIVE] Uploading visa file:', entry.visaImageUrl, '→ chat:', visaArchiveChatId);
    const result = await uploadFile(entry.visaImageUrl, visaArchiveChatId);
    if (result) {
      uploadResults.visaFileId = result.fileId;
      uploadResults.visaFilename = result.filename;
    }
  }

  const hasFiles = !!(entry.printImageUrl || entry.visaImageUrl);
  const anyUploaded = !!(uploadResults.printFileId || uploadResults.visaFileId);

  if (hasFiles && !anyUploaded && lastUploadError) {
    return res.status(500).json({ error: 'upload_failed', message: lastUploadError });
  }

  entry.archived = true;
  entry.archivedAt = Date.now();
  if (uploadResults.printFileId) {
    entry.printImageUrl = `tg:${uploadResults.printFileId}:${encodeURIComponent(uploadResults.printFilename || 'print.pdf')}`;
    entry.printImageBase64 = undefined;
  }
  if (uploadResults.visaFileId) {
    entry.visaImageUrl = `tg:${uploadResults.visaFileId}:${encodeURIComponent(uploadResults.visaFilename || 'visa.pdf')}`;
    entry.visaImageBase64 = undefined;
  }

  entries[entryIndex] = entry;
  writeEntries(entries);

  res.json({ success: true, entry, uploaded: { print: !!uploadResults.printFileId, visa: !!uploadResults.visaFileId } });
});

app.post('/api/telegram-config', (req, res) => {
  const { token, enabled, alertChatId, printArchiveChatId, visaArchiveChatId } = req.body;
  const config = { token: token || '', enabled: !!enabled, alertChatId: alertChatId || '', printArchiveChatId: printArchiveChatId || '', visaArchiveChatId: visaArchiveChatId || '' };
  writeTelegramConfig(config);
  
  // Re-start or apply Telegram Bot polling runner
  startTelegramBotRunner();
  
  res.json({ success: true, config });
});

app.get('/api/credentials', (req, res) => {
  res.json(readCredentials());
});

app.post('/api/credentials', (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) {
    return res.status(400).json({ error: 'Username and password are required' });
  }
  const creds = { username: username.trim(), password: password.trim() };
  writeCredentials(creds);
  writeAuthenticatedChats([]); // securely reset bot sessions
  res.json({ success: true, credentials: creds });
});

// Users Management REST endpoints
app.get('/api/users', (req, res) => {
  res.json(readUsers());
});

app.post('/api/users', (req, res) => {
  const { users } = req.body;
  if (!users || !Array.isArray(users)) {
    return res.status(400).json({ error: 'Users must be an array' });
  }
  writeUsers(users);
  res.json({ success: true, users });
});

app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) {
    return res.status(400).json({ error: 'Username and password are required' });
  }

  const normalizedUser = username.trim().toLowerCase();
  const rawPass = password.trim();

  // 1. Check Master Admin Credentials
  const masterCreds = readCredentials();
  if (normalizedUser === masterCreds.username.toLowerCase() && rawPass === masterCreds.password) {
    return res.json({
      success: true,
      user: {
        id: 'master-admin',
        username: masterCreds.username,
        name: 'مدير النظام (الرئيسي)',
        permissions: ['admin', 'edit_entries', 'excel_sync', 'bot_settings', 'update_checking', 'add_clients', 'edit_clients', 'download_data']
      }
    });
  }

  // 2. Check Custom Users
  const users = readUsers();
  const matchedUser = users.find(u => u.username.toLowerCase() === normalizedUser && u.password === rawPass);
  if (matchedUser) {
    return res.json({
      success: true,
      user: {
        id: matchedUser.id,
        username: matchedUser.username,
        name: matchedUser.name || matchedUser.username,
        permissions: matchedUser.permissions || []
      }
    });
  }

  res.status(401).json({ error: 'اسم المستخدم أو رمز الدخول غير صحيح' });
});

// Telegram Polling System Variables
let telegramInterval: NodeJS.Timeout | null = null;
let lastUpdateId = 0;
let isPollingActive = false;
let currentBotToken = '';

// User state mappings for advanced bot navigation
let userActiveDept: { [chatId: number]: string } = {}; // 'all', 'today', 'with_print', 'issued'
let userWaitingForSearch: { [chatId: number]: boolean } = {};
let userLastMatches: { [chatId: number]: any[] } = {};

// Helper to check if entry was updated today
function isTodayEntry(e: any): boolean {
  if (e.updatedAtUnix) {
    const d = new Date(e.updatedAtUnix);
    const today = new Date();
    return d.getDate() === today.getDate() &&
           d.getMonth() === today.getMonth() &&
           d.getFullYear() === today.getFullYear();
  }
  return false;
}

async function startTelegramBotRunner() {
  if (telegramInterval) {
    clearTimeout(telegramInterval);
    telegramInterval = null;
  }
  
  const config = readTelegramConfig();
  if (!config.token || !config.enabled) {
    console.log('Telegram Bot system is paused or has no valid token.');
    return;
  }
  
  currentBotToken = config.token;
  console.log(`Telegram Bot actively started. Token prefix: ${config.token.substring(0, 8)}...`);
  lastUpdateId = 0; // reset offsets to check for live messages on start
  isPollingActive = false;

  // Proactively attempt to delete any existing webhook to prevent 409 Conflict errors
  try {
    await axios.get(`https://api.telegram.org/bot${config.token}/deleteWebhook`, { timeout: 3000 });
    console.log('Cleared webhook for Telegram bot successfully.');
  } catch (err: any) {
    console.log('Non-blocking webhook clear attempt:', err.message);
  }
  
  runPollingLoop();
}

function runPollingLoop() {
  const config = readTelegramConfig();
  if (!config.token || !config.enabled || config.token !== currentBotToken) {
    return; // Stop loop if state changed or deactivated
  }

  pollTelegramUpdates(config.token)
    .finally(() => {
      // Schedule the next poll only AFTER the current one is finished
      telegramInterval = setTimeout(runPollingLoop, 2000);
    });
}

async function pollTelegramUpdates(token: string) {
  if (isPollingActive) return;
  isPollingActive = true;
  try {
    const response = await axios.get(`https://api.telegram.org/bot${token}/getUpdates`, {
      params: {
        offset: lastUpdateId + 1,
        timeout: 5, // 5 seconds of long polling for telegram efficiency
      },
      timeout: 10000, // Higher client-side timeout to avoid aborting valid requests
    });
    
    if (response.data && response.data.ok) {
      const updates = response.data.result;
      for (const update of updates) {
        lastUpdateId = update.update_id;
        if (update.message) {
          await handleTelegramMessage(token, update.message);
        }
      }
    }
  } catch (err: any) {
    // Avoid spamming the logs with normal timeout / abort messages from long poll cycles
    if (err.code === 'ECONNABORTED' || err.message?.includes('timeout') || err.message?.includes('timeout of')) {
      // Expected long-polling timeout, fail silently to keep logs pristine
    } else {
      console.error('Telegram polling cycle error:', err.message);
    }
  } finally {
    isPollingActive = false;
  }
}

async function handleTelegramMessage(token: string, message: any) {
  const chatId = message.chat.id;
  const username = message.from?.username || message.from?.first_name || 'مستخدم';
  const textRaw = (message.text || message.caption || '').trim();
  const text = textRaw.toLowerCase();
  
  // 1. القائمة الرئيسية
  const mainKeyboardMarkup = {
    keyboard: [
      [{ text: '🏠 الرئيسية' }, { text: '📁 المعاملات' }],
      [{ text: '⚙️ إدارة البيانات' }]
    ],
    resize_keyboard: true,
    one_time_keyboard: false
  };

  // 2. قائمة الرئيسية الفروع
  const dashboardKeyboardMarkup = {
    keyboard: [
      [{ text: '🔄 تحديث الكل' }, { text: '📊 الإحصائيات' }],
      [{ text: '🔙 العودة للقائمة الرئيسية' }]
    ],
    resize_keyboard: true,
    one_time_keyboard: false
  };

  // 3. قائمة المعاملات - بها زر التصدير وأزرار الأقسام المعينة
  const transactionsKeyboardMarkup = {
    keyboard: [
      [{ text: '📥 تصدير ككشف Excel' }],
      [{ text: '📄 كافة المعاملات' }, { text: '🔄 تحديثات اليوم' }],
      [{ text: '📝 المعاملات لها طلب ( برنت )' }, { text: '🟢 المعاملات الجاهزة ( المؤشرة )' }],
      [{ text: '🔙 العودة للقائمة الرئيسية' }]
    ],
    resize_keyboard: true,
    one_time_keyboard: false
  };

  // أزرار اختيار القسم المطلوب تصديره للتنزيل المباشر
  const exportSectionsKeyboardMarkup = {
    keyboard: [
      [{ text: '📥 تصدير: كافة المعاملات' }, { text: '📥 تصدير: تحديثات اليوم' }],
      [{ text: '📥 تصدير: المعاملات لها طلب ( برنت )' }, { text: '📥 تصدير: المعاملات الجاهزة ( المؤشرة )' }],
      [{ text: '🔙 العودة لقسم المعاملات' }]
    ],
    resize_keyboard: true,
    one_time_keyboard: false
  };

  // أزرار البحث التخصيصية وكشف عملاء الأقسام
  const departmentOptionsKeyboardMarkup = {
    keyboard: [
      [{ text: '🔍 بحث واستعلام في قسمي' }, { text: '📋 كشف العملاء' }],
      [{ text: '🔙 العودة لقسم المعاملات' }]
    ],
    resize_keyboard: true,
    one_time_keyboard: false
  };

  // 4. قائمة إدارة البيانات
  const dataManagementKeyboardMarkup = {
    keyboard: [
      [{ text: '➕ إضافة معاملة جديدة' }, { text: '📥 استيراد عملاء (Excel)' }],
      [{ text: '❌ حذف معاملة' }, { text: '❓ مساعدة والتعليمات' }],
      [{ text: '🔙 العودة للقائمة الرئيسية' }]
    ],
    resize_keyboard: true,
    one_time_keyboard: false
  };

  // توافقية الاسم مع بقية الأكواد الخلفية
  const syncKeyboardMarkup = dashboardKeyboardMarkup;
  const reportsKeyboardMarkup = transactionsKeyboardMarkup;
  const helpKeyboardMarkup = dataManagementKeyboardMarkup;

  const sendText = async (textMsg: string, replyMarkup: any = mainKeyboardMarkup) => {
    try {
      await axios.post(`https://api.telegram.org/bot${token}/sendMessage`, {
        chat_id: chatId,
        text: textMsg,
        parse_mode: 'Markdown',
        reply_markup: replyMarkup
      });
    } catch (err: any) {
      console.error('Error sending telegram response:', err.message);
    }
  };

  // =============== SECURE PASSWORD ENFORCEMENT ===============
  const creds = readCredentials();
  const authedChats = readAuthenticatedChats();
  const isAuthed = authedChats.includes(chatId);

  if (!isAuthed) {
    const parts = textRaw.split(/\s+/).map((p: string) => p.trim());
    let providedUser = '';
    let providedPass = '';

    if (parts.length === 2) {
      providedUser = parts[0];
      providedPass = parts[1];
    } else if (parts.length === 3 && (parts[0] === 'دخول' || parts[0] === 'login')) {
      providedUser = parts[1];
      providedPass = parts[2];
    }

    // Check custom users who have admin or bot_settings permissions
    const customUsers = readUsers();
    const isCustomAuthed = customUsers.some(
      (u: any) =>
        u.username.toLowerCase() === providedUser.toLowerCase() &&
        u.password === providedPass &&
        (u.permissions.includes('admin') || u.permissions.includes('bot_settings'))
    );

    if (
      providedUser &&
      ((providedUser.toLowerCase() === creds.username.toLowerCase() &&
        providedPass === creds.password) ||
        isCustomAuthed)
    ) {
      authedChats.push(chatId);
      writeAuthenticatedChats(authedChats);

      const authSuccessMsg = `🟢 *تم التحقق والولوج إلى النظام بنجاح!*

مرحباً بك مجدداً يا *${username}* كمسؤول معتمد.
لقد تم إلغاء حظر الأوامر وتفعيل لوحة الأزرار والخيارات التفاعلية كاملة لحسابك بأمان.

👇 *الآن يمكنك استخدام الخيارات والتحكم بالقائمة التفاعلية أدناه:*`;

      await sendText(authSuccessMsg, mainKeyboardMarkup);
      return;
    }

    const accessDeniedMsg = `🔐 *تنبيه أمان: الوصول مقيد ومحمي بكلمة مرور!*

مرحباً بك يا *${username}*. عذراً، الوصول لخدمات وتحديثات النظام مقفل ويتطلب إدخل مفاتيح الاعتماد الخاصة بمسؤول النظام.

يرجى إرسال *اسم المستخدم* متبوعاً بـ *كلمة المرور* في رسالة نصية واحدة لتجاوز حماية البوت والتحقق من حسابك ومباشرة العمل.

📝 *طريقة كتابة الرسالة والتحقق:*
اكتب اليوزر مسافة الباسورد كالتالي:
\`اسم_المستخدم كلمة_المرور\`

_مثال:_ \`admin 1234\``;

    try {
      await axios.post(`https://api.telegram.org/bot${token}/sendMessage`, {
        chat_id: chatId,
        text: accessDeniedMsg,
        parse_mode: 'Markdown',
        reply_markup: {
          remove_keyboard: true
        }
      });
    } catch (err: any) {
      console.error('Error sending access denied telegram response:', err.message);
    }
    return;
  }

  // ==================== SUBMENU ROUTING AND COMMAND HANDLING ====================

  // A. Navigation Event: Main Menu / Return Back
  if (
    text === '/start' || 
    text === '🔙' || 
    text === 'العودة' || 
    text === 'رجوع' || 
    text.includes('العودة للقائمة الرئيسية') || 
    text.includes('القائمة الرئيسية') || 
    text === 'menu' || 
    text === 'القائمة'
  ) {
    const welcomeMsg = `🤖 *مرحباً بك يا ${username}!*

لقد تم فرز وتصنيف *لوحة الأزرار والخيارات التفاعلية* في قوائم مخصصة لتسهيل إدارتك السريعة لجميع العملاء والمعاملات.

👇 *اختر القسم المطلوب من الأزرار الرئيسية أدناه لتظهر لك بقية الخيارات والخرائط الفرعية:*`;
    await sendText(welcomeMsg, mainKeyboardMarkup);
    return;
  }

  // B. Navigation Event: الرئيسية
  if (text === '🏠 الرئيسية' || text === 'الرئيسية' || text === 'الرئيسيه') {
    const entries = readEntries();
    const total = entries.length;
    let issued = 0;
    let today = 0;
    let withPrint = 0;

    entries.forEach((e: any) => {
      const st = String(e.statusText || '').toLowerCase();
      if (st.includes('issued') || st.includes('اصدار') || st.includes('جاهز')) {
        issued++;
      }
      if (isTodayEntry(e)) {
        today++;
      }
      if (e.applicationNumber !== '---' || (e.applicantData && e.applicantData.trim() !== '')) {
        withPrint++;
      }
    });

    const msg = `🏠 *لوحة معلومات ومؤشرات إنجاز المعاملات:*

👥 *كافة المعاملات بالسيستم:* \`${total}\`
🟢 *المعاملات الجاهزة ( المؤشرة ):* \`${issued}\`
🔄 *تحديثات اليوم:* \`${today}\`
📝 *المعاملات لها طلب ( برنت ):* \`${withPrint}\`

👇 *اختر الفئة أو التبويب المطلوب أدناه لتصفح المعاملات أو إجراء عمليات البحث المتقدم والتصدير عليها:*`;
    await sendText(msg, dashboardKeyboardMarkup);
    return;
  }

  // C. Navigation Event: المعاملات
  if (text === '📁 المعاملات' || text === 'المعاملات' || text === 'معاملات' || text === '🔙 العودة لقسم المعاملات') {
    const msg = `📁 *قسم كشوفات وعمليات المعاملات الحالية:*

يمكنك تصفح المعاملات حسب كل أقسام وصفحة الأدمن، والبحث السريع، والتحميل المباشر لكشوفات الـ Excel.

👇 *اختر الفئة للتصفح والبحث الذكي، أو انقر التصدير لتحميل كشف Excel من الأزرار أدناه كما في صفحة الأدمن:*`;
    await sendText(msg, transactionsKeyboardMarkup);
    return;
  }

  // --- EXPORT SELECTOR VIEW ---
  if (text === '📥 تصدير ككشف excel' || text === 'تصدير' || text === 'تصدير ككشف إكسل') {
    const msg = `📥 *تصدير تقارير وكشوفات Excel:*

الرجاء اختيار الفئة/القسم المطلوب لتصدير كافة معاملاته في ملف Excel منسق جاهز للتحميل والتصدير للمحمول فوراً:`;
    await sendText(msg, exportSectionsKeyboardMarkup);
    return;
  }

  // --- SUB-DEPARTMENTS SELECTION ---
  if (
    text === '📄 كافة المعاملات' || 
    text === 'كافة المعاملات' ||
    text === '🔄 تحديثات اليوم' || 
    text === 'تحديثات اليوم' ||
    text === '📝 معاملات لها طلب ( برنت )' ||
    text === '📝 المعاملات لها طلب ( برنت )' ||
    text === 'المعاملات لها طلب ( برنت )' ||
    text === '🟢 المعاملات الجاهزة ( المؤشرة )' || 
    text === 'المعاملات الجاهزة ( المؤشرة )'
  ) {
    let dept = 'all';
    let label = 'كافة المعاملات';
    
    if (text === '🔄 تحديثات اليوم' || text === 'تحديثات اليوم') {
      dept = 'today';
      label = 'تحديثات اليوم';
    } else if (text === '📝 معاملات لها طلب ( برنت )' || text === '📝 المعاملات لها طلب ( برنت )' || text === 'المعاملات لها طلب ( برنت )') {
      dept = 'with_print';
      label = 'المعاملات لها طلب ( برنت )';
    } else if (text === '🟢 المعاملات الجاهزة ( المؤشرة )' || text === 'المعاملات الجاهزة ( المؤشرة )') {
      dept = 'issued';
      label = 'المعاملات الجاهزة ( المؤشرة )';
    }
    
    userActiveDept[chatId] = dept;
    userWaitingForSearch[chatId] = false;
    
    const entries = readEntries();
    let filteredCount = 0;
    if (dept === 'all') {
      filteredCount = entries.length;
    } else if (dept === 'today') {
      filteredCount = entries.filter(isTodayEntry).length;
    } else if (dept === 'with_print') {
      filteredCount = entries.filter((e: any) => e.applicationNumber !== '---' || (e.applicantData && e.applicantData.trim() !== '')).length;
    } else if (dept === 'issued') {
      filteredCount = entries.filter((e: any) => {
        const st = String(e.statusText || '').toLowerCase();
        return st.includes('issued') || st.includes('اصدار') || st.includes('جاهز');
      }).length;
    }

    const msg = `📋 القسم المختار حاليا: *${label}* (${filteredCount} معاملة)
    
الرجاء اختيار أحد الخيارين أدناه لتصفح أو البحث عن العملاء داخل هذا القسم كما بالأدمن:`;
    await sendText(msg, departmentOptionsKeyboardMarkup);
    return;
  }

  // --- SEARCH PER SECTION ACTION ---
  if (text === '🔍 بحث واستعلام في قسمي' || text === '🔍 بحث واستعلام في هذا القسم') {
    const dept = userActiveDept[chatId] || 'all';
    const deptName = dept === 'all' ? 'كافة المعاملات' : dept === 'today' ? 'تحديثات اليوم' : dept === 'with_print' ? 'المعاملات لها طلب ( برنت )' : 'المعاملات الجاهزة ( المؤشرة )';
    
    userWaitingForSearch[chatId] = true;
    const msg = `🔍 *جاري الاستماع للبحث الفوري في قسم (${deptName}):*
    
الآن، يرجى إرسال أي معلومة للعميل المراد تتبعه (الاسم بالكامل، أو رقم الجواز، أو رقم الهاتف) وسيعرض البوت النتائج الماثلة فوراً.`;
    
    await sendText(msg, departmentOptionsKeyboardMarkup);
    return;
  }

  // --- SHOW DEPARTMENT CUSTOMER LIST ---
  if (text === '📋 كشف العملاء') {
    const dept = userActiveDept[chatId] || 'all';
    const deptName = dept === 'all' ? 'كافة المعاملات' : dept === 'today' ? 'تحديثات اليوم' : dept === 'with_print' ? 'المعاملات لها طلب ( برنت )' : 'المعاملات الجاهزة ( المؤشرة )';
    
    const entries = readEntries();
    let filtered: any[] = [];
    
    if (dept === 'all') {
      filtered = entries;
    } else if (dept === 'today') {
      filtered = entries.filter(isTodayEntry);
    } else if (dept === 'with_print') {
      filtered = entries.filter((e: any) => e.applicationNumber !== '---' || (e.applicantData && e.applicantData.trim() !== ''));
    } else if (dept === 'issued') {
      filtered = entries.filter((e: any) => {
        const st = String(e.statusText || '').toLowerCase();
        return st.includes('issued') || st.includes('اصدار') || st.includes('جاهز');
      });
    }

    if (filtered.length === 0) {
      await sendText(`⚠️ عذراً، لا يوجد عملاء أو معاملات مسجلة حالياً في قسم (${deptName}).`, departmentOptionsKeyboardMarkup);
      return;
    }

    userLastMatches[chatId] = filtered; // Save the reference list to enable details by index/number

    let reply = `📋 *كشف العملاء في قسم (${deptName}) [إجمالي: ${filtered.length}]:*\n`;
    reply += `_أرسل رقم المعاملة المتسلسل (مثال: 1) لعرض كافة التفاصيل والبيانات للمسافر تلقائياً._\n\n`;

    const itemsToDisplay = filtered.slice(0, 40); // safe limit to avoid telegram size constraints
    itemsToDisplay.forEach((m: any, idx: number) => {
      reply += `*${idx + 1}.* 👤 *${m.applicantName || 'غير مسجل'}*\n🛂 جواز: \`${m.passportNumber}\`\n\n`;
    });

    if (filtered.length > 40) {
      reply += `_... تم إظهار أول 40 عميل من الكشف الكلي لتسهيل العرض._`;
    }

    await sendText(reply, departmentOptionsKeyboardMarkup);
    return;
  }
  // Cleaned remaining duplicate handlers block perfectly

  // D. Navigation Event: إدارة البيانات
  if (text === '⚙️ إدارة البيانات' || text === 'إدارة البيانات' || text === 'ادارة البيانات') {
    const msg = `⚙️ *تجهيز وضبط وإلغاء بيانات السجل المزدوج:*

هنا يمكنك إدراج عملاء جدد للجدول فوراً يدوياً، أو رفع ملف إكسل XLSX للمزامنة الترحيلية تلقائياً، والتحكم بحذف السجلات.

👇 *اختر المهمة المطلوبة من الخيارات كصفحة التحكم بالأدمن:*`;
    await sendText(msg, dataManagementKeyboardMarkup);
    return;
  }

  // 1. Check if a document (Excel file) has been uploaded
  if (message.document) {
    const doc = message.document;
    const fileName = doc.file_name || '';
    const fileId = doc.file_id;
    
    // Check extension
    if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls') && !fileName.endsWith('.csv')) {
      await sendText('⚠️ *عذراً، نوع الملف المرفق غير مدعوم!* يرجى إرسال ملف Excel صالح بصيغة (`.xlsx` أو `.xls` أو `.csv`) يحتوي على بيانات واضحة للعملاء.');
      return;
    }
    
    // Notify parsing started
    await sendText(`📥 *جاري استلام وتحليل ملف Excel الحامل لاسم:*\n\`${fileName}\`...\nيرجى الانتظار لحين مطابقة السجلات وتنزيلها.`);
    
    try {
      // Get file path from Telegram API
      const fileRes = await axios.get(`https://api.telegram.org/bot${token}/getFile`, {
        params: { file_id: fileId }
      });
      
      if (!fileRes.data || !fileRes.data.ok) {
        throw new Error('فشل جلب مسار الملف من سيرفر تليجرام');
      }
      
      const filePathOnTelegram = fileRes.data.result.file_path;
      
      // Download the actual file buffer
      const downloadRes = await axios.get(`https://api.telegram.org/file/bot${token}/${filePathOnTelegram}`, {
        responseType: 'arraybuffer'
      });
      
      const buffer = Buffer.from(downloadRes.data);
      
      // Read sheet workbook using XLSX
      const workbook = XLSX.read(buffer, { type: 'buffer' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      if (!jsonData || jsonData.length === 0) {
        await sendText('❌ الملف المرفوع فارغ تماماً ولا يحتوي على أية صفوف للبيانات.');
        return;
      }
      
      let headerRowIndex = 0;
      let passportColIdx = -1;
      let applicationColIdx = -1;
      let nameColIdx = -1;
      let phoneColIdx = -1;
      let professionColIdx = -1;

      // Look at the first 10 rows for matching headers
      for (let r = 0; r < Math.min(jsonData.length, 10); r++) {
        const row = jsonData[r] || [];
        for (let c = 0; c < row.length; c++) {
          const val = String(row[c] || '').trim().toLowerCase();
          if (val.includes('جواز') || val.includes('passport') || val.includes('pass') || val.includes('رقم الجواز')) {
            passportColIdx = c;
            headerRowIndex = r;
          }
          if (val.includes('طلب') || val.includes('application') || val.includes('app') || val.includes('رقم الطلب')) {
            applicationColIdx = c;
            headerRowIndex = r;
          }
          if (val.includes('اسم') || val.includes('name') || val.includes('مستفيد') || val.includes('الاسم') || val.includes('العميل')) {
            nameColIdx = c;
            headerRowIndex = r;
          }
          if (val.includes('هاتف') || val.includes('تلفون') || val.includes('جوال') || val.includes('phone') || val.includes('mobile') || val.includes('التلفون')) {
            phoneColIdx = c;
            headerRowIndex = r;
          }
          if (val.includes('مهنة') || val.includes('المهنة') || val.includes('profession') || val.includes('job') || val.includes('الوظيفة')) {
            professionColIdx = c;
            headerRowIndex = r;
          }
        }
        if (passportColIdx !== -1 && nameColIdx !== -1) {
          break; // Found perfect headers row
        }
      }

      // Fallback defaults if headers could not be found
      if (passportColIdx === -1 || nameColIdx === -1) {
        passportColIdx = 1;
        nameColIdx = 0;
        applicationColIdx = 2;
        professionColIdx = 3;
        phoneColIdx = 4;
        headerRowIndex = 0; // Assume first row as data
      }

      const entries = readEntries();
      let addedCount = 0;
      let skippedCount = 0;

      for (let r = headerRowIndex + 1; r < jsonData.length; r++) {
        const row = jsonData[r];
        if (!row || row.length === 0) continue;

        const rawName = String(row[nameColIdx] || '').trim();
        const rawPassport = String(row[passportColIdx] || '').trim().toUpperCase();
        const rawAppNo = String(row[applicationColIdx] || '').trim();
        const rawProfession = String(row[professionColIdx] || 'عامل').trim();
        const rawPhone = String(row[phoneColIdx] || '').trim();

        // Check if essential fields are valid
        if (!rawName || !rawPassport || rawPassport.length < 5 || !rawAppNo) {
          continue; // Incomplete row
        }

        const doubleCheck = entries.find((e: any) => e.passportNumber.toUpperCase() === rawPassport);
        if (doubleCheck) {
          skippedCount++;
          continue;
        }

        const newRow = {
          id: `excel-import-${Date.now()}-${r}-${Math.floor(Math.random() * 1000)}`,
          passportNumber: rawPassport,
          applicationNumber: rawAppNo,
          applicantName: rawName,
          nationality: 'اليمن',
          arrivalPoint: 'عدن',
          visaType: 'تأشيرة عمل',
          status: 'Idle',
          statusText: 'قيد الانتظار (Idle)',
          visaNumber: '---',
          phoneNumber: rawPhone,
          profession: rawProfession,
          lastUpdate: new Date().toLocaleTimeString('ar-YE', { hour: '2-digit', minute: '2-digit' }),
          applicantData: 'تم استيراده بنجاح كجزء من دفعة الـ Excel تلقائياً عبر البوت.'
        };

        entries.unshift(newRow);
        addedCount++;
      }

      if (addedCount > 0) {
        writeEntries(entries);
        await sendText(`✅ *تمت معالجة الملف واستيراد العملاء بنجاح!*
        
👥 *عدد العملاء الذين تمت إضافتهم:* \`${addedCount}\` عميل جديد.
⚠️ *عدد العملاء المكررين (تخطيهم):* \`${skippedCount}\` عميل.

_الملف الذي طُبّق:_ \`${fileName}\``);
      } else {
        await sendText(`ℹ️ *اكتملت المعالجة ولكن لم تتم إضافة عملاء جدد:*
        
⚠️ تم تخطي كافة البيانات أو لم يتم العثور على جوازات غير مسجلة من قبل (مكررة).
👥 *الصفحات المكررة / المتخطاة:* \`${skippedCount}\``);
      }

    } catch (err: any) {
      console.error('Failed parsing in-chat Excel upload:', err);
      await sendText(`❌ واجه البوت خطأ غير متوقع أثناء تفكيك وقراءة ملف إكسل: ${err.message}`);
    }
    return;
  }

  // COMMAND: START / HELP / MENU
  if (
    text === '❓ مساعدة والتعليمات' ||
    text === '/help' || 
    text === 'help' || 
    text.includes('مساعدة') || 
    text.includes('البداية') || 
    text.includes('تعليمات')
  ) {
    const helpMsg = `🤖 *دليل الأوامر والمساعده*

يمكنك استخدام الأزرار التفاعلية، أو كتابة أي من الأوامر التالية مباشرة بالدردشة:

🔎 *أوامر البحث والفحص والمضافة السريعة:*
• \`🔍 فحص [رقم الجواز]\` : لمعرفة البيانات وجاهزية التأشيرة لعميل معين.
  _مثال:_ \`فحص A1234567\`
• \`🔎 بحث [الاسم أو الجواز]\` : للبحث الكامل عن عميل بالاسم أو الجواز.
  _مثال:_ \`بحث خالد\`
• \`إضافة [الاسم | رقم الجواز | رقم الطلب]\` : لإضافة عميل جديد للسيستم والجدول فوراً بالخط العمودي.
  _مثال:_ \`إضافة معاذ | A1234567 | 10048392\`
• \`حذف [رقم الجواز]\` : لحذف وإزالة بيانات العميل نهائياً وبشكل آمن من السجلات.
  _مثال:_ \`حذف A1234567\``;
    await sendText(helpMsg, helpKeyboardMarkup);
    return;
  }

  // COMMAND: QUICK SEARCH GUIDE
  if (text.includes('استعلام وبحث') || text === '🔍 استعلام وبحث سريع') {
    const searchHelpMsg = `🔍 *دليل الاستعلام والبحث السريع:*
    
• للبحث الكامل بالاسم أو رقم الجواز، اكتب الأمر *بحث* متبوعاً بالاسم.
  _مثال:_ \`بحث أحمد\`
  
• لفحص جواز عميل محدد واستعراض حالته بالتفصيل، اكتب الأمر *فحص* متبوعاً برقم الجواز.
  _مثال:_ \`فحص A1234567\``;
    await sendText(searchHelpMsg, transactionsKeyboardMarkup);
    return;
  }

  // COMMAND: ADD CUSTOMER GUIDE (from button)
  if (text.includes('إضافة معاملة جديدة') || text === '➕ إضافة معاملة جديدة') {
    const addHelpMsg = `➕ *طريقة إضافة عميل أو مسافر جديد فوراً للجدول:*
    
لتسجيل وإدراج عميل جديد للجدول مباشرة، يرجى كتابة كلمة *إضافة* متبوعة بالبيانات مقسمة بالخط العمودي (\`|\`) كالتالي:

\`إضافة الاسم | رقم الجواز | رقم الطلب | المهنة | الهاتف\`

_مثال:_
\`إضافة معاذ الواحد | A1234567 | 10048392 | مهندس | 770123456\``;
    await sendText(addHelpMsg, dataManagementKeyboardMarkup);
    return;
  }

  // COMMAND: DELETE CUSTOMER GUIDE (from button)
  if (text.includes('حذف معاملة') || text === '❌ حذف معاملة') {
    const deleteHelpMsg = `❌ *دليل إيقاف وحذف المعاملات والعملاء:*
    
لإزالة أي عميل من قاعدة السجلات نهائياً، يرجى كتابة الأمر *حذف* متبوعاً برقم الجواز المطلوب إزالته من جداول البوت والموقع.

_مثال:_ \`حذف A1234567\``;
    await sendText(deleteHelpMsg, dataManagementKeyboardMarkup);
    return;
  }

  // COMMAND: EXCEL EXPLAIN / GUIDE
  if (
    text.includes('استيراد') || 
    text.includes('اكسل') || 
    text.includes('excel') || 
    text.includes('ملف')
  ) {
    const guideMsg = `📥 *دليل استيراد العملاء عبر ملف Excel بالتليجرام:*
    
أهلاً بك! لتنزيل واستيراد قاعدة بيانات عملاء جديدة دفعة واحدة، يرجى القيام بـ:

1️⃣ قم بإرسال أو سحب ملف الإكسل (*.xlsx* أو *.xls*) مباشرة داخل هذه المحادثة كملف (Document).
2️⃣ تأكد أن الملف يحتوي على أعمدة في الصف الأول بالأسماء التالية (أو ما يقابلها):
   • *الاسم الكامل* (أو الاسم)
   • *رقم الجواز* (أو الجواز)
   • *رقم الطلب* (أو الطلب)
   • *المهنة* (اختياري)
   • *الهاتف* (اختياري)

⚡ وسيقوم البوت بالتقاط الملف ومطابقة الأعمدة آلياً لتلقي المعاملات وحفظها في ثوانٍ معدودة!`;
    await sendText(guideMsg, syncKeyboardMarkup);
    return;
  }

  // COMMAND: STATS
  if (
    text === '/stats' || 
    text.includes('إحصائيات') || 
    text.includes('الاحصائيات') || 
    text === '📊'
  ) {
    const entries = readEntries();
    const total = entries.length;
    let issued = 0;
    let underProcess = 0;
    let pendingApproval = 0;
    let others = 0;

    entries.forEach((e: any) => {
      const st = String(e.statusText || '').toLowerCase();
      if (st.includes('issued') || st.includes('اصدار') || st.includes('جاهز')) {
        issued++;
      } else if (st.includes('process') || st.includes('اجراء')) {
        underProcess++;
      } else if (st.includes('pending') || st.includes('موافقة') || st.includes('انتظار')) {
        pendingApproval++;
      } else {
        others++;
      }
    });

    const statsMsg = `📊 *تقرير إحصائيات المعاملات الحالية:*

👥 *إجمالي الجوازات والعملاء:* \`${total}\`
🟢 *جاهزة ومطبوعة (Issued):* \`${issued}\`
🟡 *تحت الإجراء (Under Process):* \`${underProcess}\`
⏳ *بانتظار الموافقة والرصد:* \`${pendingApproval}\`
` + (others > 0 ? `⚪ *معاملات أخرى/جديدة:* \`${others}\`\n` : '') + `
_تاريخ ووقت التقرير:_ ${new Date().toLocaleString('ar-YE')}`;
    await sendText(statsMsg, helpKeyboardMarkup);
    return;
  }

  // COMMAND: SEARCH
  if (text.startsWith('/search ') || text.startsWith('بحث ') || text.startsWith('البحث ')) {
    const query = textRaw.substring(textRaw.indexOf(' ') + 1).trim();
    if (!query) {
      await sendText('⚠️ يرجى كتابة الاسم أو الجواز غرض البحث بعد الأمر. مثال: `بحث أحمد`');
      return;
    }

    const entries = readEntries();
    const matches = entries.filter((e: any) => {
      return (e.applicantName || '').includes(query) || 
             (e.passportNumber || '').toUpperCase().includes(query.toUpperCase()) ||
             (e.applicationNumber || '').includes(query);
    });

    if (matches.length === 0) {
      await sendText(`❌ لم نعثر على أي مطابقة للبحث: *${query}*`);
      return;
    }

    let reply = `🔎 *نتائج البحث في السجلات لـ (${query}) [${matches.length} نتيجة]:*\n\n`;
    matches.slice(0, 15).forEach((m: any, idx: number) => {
      reply += `*${idx + 1}.* 👤 *${m.applicantName || 'غير مسمى'}*\n`;
      reply += `   🛂 الجواز: \`${m.passportNumber}\` | 📁 الطلب: \`${m.applicationNumber}\`\n`;
      reply += `   ⚙️ المهنة: \`${m.profession || 'غير مسمى'}\` | 🟢 الحالة: *${m.statusText || 'انتظار'}*\n\n`;
    });

    if (matches.length > 15) {
      reply += `_... تم إظهار أول 15 نتيجة تطابق._`;
    }
    await sendText(reply);
    return;
  }

  // COMMAND: CHECK
  if (text.startsWith('/check ') || text.startsWith('فحص ') || text.startsWith('تحقق ') || text.startsWith('استعلام ')) {
    const query = textRaw.substring(textRaw.indexOf(' ') + 1).trim().toUpperCase();
    if (!query) {
      await sendText('⚠️ يرجى إدخال رقم الجواز المستهلك للتفتيش. مثال: `فحص A1234567`');
      return;
    }

    const entries = readEntries();
    const found = entries.find((e: any) => 
      (e.passportNumber || '').toUpperCase() === query || 
      (e.applicationNumber || '') === query
    );

    if (!found) {
      await sendText(`❌ المعذرة، لا يوجد عميل مسجل برقم الجواز/الطلب: *${query}*`);
      return;
    }

    const chMsg = `ℹ️ *بيانات جواز العميل المفصلة:*

👤 *الاسم الكامل:* ${found.applicantName || 'غير مسجل'}
🛂 *رقم الجواز:* \`${found.passportNumber}\`
📁 *رقم الطلب:* \`${found.applicationNumber}\`
⚙️ *المهنة الأساسية:* ${found.profession || 'غير مسجلة'}
📞 *رقم الهاتف:* ${found.phoneNumber || 'لا يوجد'}
📍 *وصول السفارة:* ${found.arrivalPoint || 'عدن'}
🏷️ *فئة التأشيرة:* ${found.visaType || 'عمل'}

📊 *الرمز والوضعية التفاعلية:*
👉 *${found.statusText || 'قيد الانتظار'}*
🛂 *رقم التأشيرة المنجزة:* \`${found.visaNumber || '---'}\`

📝 *معلومات وفصل إضافي:*
_${found.applicantData || 'لا توجد بيانات ملحقة.'}_`;
    await sendText(chMsg);
    return;
  }

  // COMMAND: DELETE
  if (text.startsWith('/delete ') || text.startsWith('حذف ') || text.startsWith('إزالة ') || text.startsWith('ازالة ')) {
    const query = textRaw.substring(textRaw.indexOf(' ') + 1).trim().toUpperCase();
    if (!query) {
      await sendText('⚠️ يرجى إدخال رقم الجواز للحذف والمسح التام. مثال: `حذف A1234567`');
      return;
    }

    const entries = readEntries();
    const index = entries.findIndex((e: any) => 
      (e.passportNumber || '').toUpperCase() === query || 
      (e.applicationNumber || '') === query
    );

    if (index === -1) {
      await sendText(`❌ لم يتم الكشف على عميل مسجل للمسح بمعرف الجواز/الطلب: *${query}*`);
      return;
    }

    const removed = entries[index];
    entries.splice(index, 1);
    writeEntries(entries);

    // Delete associated PDF file if exists
    if (removed.printImageUrl) {
      try {
        const pdfFilename = path.basename(removed.printImageUrl);
        const pdfFilePath = path.join(PRINTS_DIR, pdfFilename);
        if (fs.existsSync(pdfFilePath)) fs.unlinkSync(pdfFilePath);
      } catch {}
    }

    await sendText(`✅ *تم حذف وإلغاء العميل بنجاح من قاعدة البيانات!*
👤 *الاسم:* ${removed.applicantName}
🛂 *الجواز:* \`${removed.passportNumber}\``);
    return;
  }

  // COMMAND: ADD CUSTOMER
  if (text.startsWith('/add ') || text.startsWith('إضافة ') || text.startsWith('اضافة ')) {
    const rawParams = textRaw.substring(textRaw.indexOf(' ') + 1).trim();
    if (!rawParams) {
      await sendText(`⚠️ *يرجى إرسال بيانات العميل مقسمة بصيغة الخط العمودي | كالتالي:*
\`إضافة الاسم | رقم الجواز | رقم الطلب | المهنة | الهاتف\`

_مثال:_ \`إضافة خالد الواحد | A9182312 | 10093822 | مهندس | 770123456\``);
      return;
    }

    const parts = rawParams.split('|').map(p => p.trim());
    if (parts.length < 3) {
      await sendText('⚠️ يرجى إدخال (الاسم الكامل، ورقم الجواز، ورقم الطلب) على الأقل لإكمال الإدخال.');
      return;
    }

    const name = parts[0];
    const passport = parts[1].toUpperCase();
    const appNo = parts[2];
    const profession = parts[3] || 'عامل';
    const phone = parts[4] || '';

    const entries = readEntries();
    const doubleCheck = entries.find((e: any) => e.passportNumber.toUpperCase() === passport);
    if (doubleCheck) {
      await sendText(`⚠️ تنبيه: العميل مسجل في النظام مسبقاً باسم: *${doubleCheck.applicantName}*`);
      return;
    }

    const newRow = {
      id: `bot-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      passportNumber: passport,
      applicationNumber: appNo,
      applicantName: name,
      nationality: 'اليمن',
      arrivalPoint: 'عدن',
      visaType: 'تأشيرة عمل',
      status: 'Idle',
      statusText: 'قيد الانتظار (Idle)',
      visaNumber: '---',
      phoneNumber: phone,
      profession: profession,
      lastUpdate: new Date().toLocaleTimeString('ar-YE', { hour: '2-digit', minute: '2-digit' }),
      applicantData: 'تم إدخال المعاملة بنجاح ومدرج للاستعلام الآلي.'
    };

    entries.unshift(newRow);
    writeEntries(entries);

    await sendText(`✅ *تمت إضافة وتسجيل العميل الجديد بنجاح في قاعدة البيانات!*
👤 *الاسم:* ${name}
🛂 *الجواز:* \`${passport}\`
📁 *الطلب:* \`${appNo}\`
⚙️ *المهنة:* ${profession}
📞 *الهاتف:* ${phone || '---'}`);
    return;
  }

  // COMMAND: SYNC (AUTO CHECK)
  if (text === '/sync' || text === 'تحديث الكل' || text === 'تحديث' || text === 'تشييك' || text === 'المزامنة') {
    const warnMsg = `⚠️ *تنبيه الأمان والبيانات الحقيقية (خادم إنجاز الموحد):*

بوابة التأشيرات الحكومية الرسمية (MOFA) محمية برموز أمان وصور تحقق (CAPTCHA) تمنع الاستعلام التلقائي الصامت في الخلفية؛ وتفادياً لإنتاج أي بيانات افتراضية أو خاطئة غير معبرة عن واقع العميل، *تم إيقاف المزامنة التلقائية الافتراضية للبوت*.

للحصول على بيانات معتمدة وحقيقية 100%، يرجى الاستعلام وتحديث كروت المسافرين بشكل فردي حقيقي ومطابقة منسوخ البوابة من لوحة التحكم الرئيسية للنظام بالويب بالمدخل الفردي المعتمد بالمنسوج الفعلي لجواز السفر.`;
    
    await sendText(warnMsg, syncKeyboardMarkup);
    return;
  }

  // COMMAND: DOWNLOAD EXCEL REPORT & DYNAMIC EXPORT BUTTONS
  if (
    text.startsWith('📥 تصدير:') || 
    text.includes('تصدير:') ||
    text.startsWith('/report') || 
    text.startsWith('تقرير') || 
    text.startsWith('كشف') ||
    (text.includes('كشف') && !text.includes('كشف العملاء')) ||
    text.includes('تقرير')
  ) {
    const entries = readEntries();
    if (entries.length === 0) {
      await sendText('⚠️ قاعدة البيانات فارغة تماماً، لا توجد سجلات لتصديرها.', transactionsKeyboardMarkup);
      return;
    }

    const cleanText = textRaw.replace(/[📄🟢🟡📥]/g, '').trim();
    const arg = cleanText.replace('/report', '').replace('تقرير', '').replace('كشف', '').replace('تصدير:', '').replace('تصدير', '').trim().toLowerCase();
    
    let targetDept = 'all';
    let titleStr = 'كافة المعاملات';

    if (arg.includes('جاهز') || arg.includes('issued') || arg.includes('الجاهزة') || arg.includes('المؤشرة')) {
      targetDept = 'issued';
      titleStr = 'المعاملات الجاهزة ( المؤشرة )';
    } else if (arg.includes('يوم') || arg.includes('today') || arg.includes('اليوم')) {
      targetDept = 'today';
      titleStr = 'تحديثات اليوم';
    } else if (arg.includes('طلب') || arg.includes('print') || arg.includes('برنت')) {
      targetDept = 'with_print';
      titleStr = 'المعاملات لها طلب ( برنت )';
    }

    await sendText(`📊 جاري صياغة وجمع ورقة الـ Excel لـ (${titleStr}). يرجى الانتظار ثوانٍ معدودة...`, transactionsKeyboardMarkup);

    try {
      let filtered = entries;
      if (targetDept === 'issued') {
        filtered = entries.filter((e: any) => {
          const st = String(e.statusText || '').toLowerCase();
          return st.includes('issued') || st.includes('اصدار') || st.includes('جاهز');
        });
      } else if (targetDept === 'today') {
        filtered = entries.filter(isTodayEntry);
      } else if (targetDept === 'with_print') {
        filtered = entries.filter((e: any) => e.applicationNumber !== '---' || (e.applicantData && e.applicantData.trim() !== ''));
      }

      const sData = filtered.map((e: any) => ({
        'رقم جواز العميل': e.passportNumber,
        'رقم الهاتف': e.phoneNumber || '',
        'اسم العميل': e.applicantName || '',
        'المهنة': e.profession || '',
        'حالة المعاملة من منصة التأشيرات': e.statusText || 'قيد الانتظار (Idle)',
        'رقم الطلب': e.applicationNumber,
        'رقم التأشيرة': e.visaNumber || '---',
        'الجنسية': e.nationality,
        'جهة الوصول': e.arrivalPoint,
        'نوع التأشيرة': e.visaType,
        'تاريخ التحديث': e.lastUpdate || ''
      }));

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(sData);

      ws['!cols'] = [
        { wch: 18 }, { wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 35 },
        { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 22 }
      ];

      XLSX.utils.book_append_sheet(wb, ws, "العملاء");
      
      const fileName = `excel_report_${Date.now()}.xlsx`;
      const filePath = path.join(process.cwd(), fileName);
      XLSX.writeFile(wb, filePath);

      const form = new FormData();
      form.append('chat_id', chatId);
      form.append('document', fs.createReadStream(filePath), {
        filename: `${titleStr}.xlsx`,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      form.append('caption', `📦 تم استخراج وطباعة كشف إلكتروني لـ *${titleStr}* بنجاح.\n\n👤 عدد العملاء: \`${filtered.length} معاملة\``);
      form.append('reply_markup', JSON.stringify(transactionsKeyboardMarkup));

      await axios.post(`https://api.telegram.org/bot${token}/sendDocument`, form, {
        headers: form.getHeaders(),
        maxContentLength: Infinity,
        maxBodyLength: Infinity
      });

      fs.unlinkSync(filePath);
    } catch (e: any) {
      console.error('Error compiling/sending Excel file over telegram chat:', e);
      await sendText(`❌ واجهنا عطل فني في تصدير كشف الإكسل ومشاركته: ${e.message}`, transactionsKeyboardMarkup);
    }
    return;
  }

  // --- DETAILED CLIENT VIEW VIA SEQUENTIAL INDEX NUMBER (1, 2, 3...) ---
  const isNumeric = /^\d+$/.test(textRaw.trim());
  if (isNumeric) {
    const idx = parseInt(textRaw.trim(), 10) - 1;
    const lastMatches = userLastMatches[chatId];
    if (lastMatches && idx >= 0 && idx < lastMatches.length) {
      const found = lastMatches[idx];
      const detailMsg = `ℹ️ *تفاصيل المعاملة للعميل الرقم (${idx + 1}):*
      
👤 *الاسم الكامل:* ${found.applicantName || 'غير مسجل'}
🛂 *رقم الجواز:* \`${found.passportNumber}\`
📁 *رقم الطلب:* \`${found.applicationNumber}\`
⚙️ *المهنة الأساسية:* ${found.profession || 'غير مسجلة'}
📞 *رقم الهاتف:* ${found.phoneNumber || 'لا يوجد'}
📍 *وصول السفارة:* ${found.arrivalPoint || 'عدن'}
🏷️ *فئة التأشيرة:* ${found.visaType || 'عمل'}
📊 *وضعية الحالة:* *${found.statusText || 'Idle'}*
🛂 *رقم التأشيرة المنجزة:* \`${found.visaNumber || '---'}\`
📅 *آخر تحديث:* ${found.lastUpdate || '---'}
📝 *ملاحظات إضافية:* _${found.applicantData || 'لا توجد بيانات ملحقة.'}_`;

      await sendText(detailMsg, departmentOptionsKeyboardMarkup);
      return;
    }
  }

  // --- ARBITRARY IN-DEPARTMENT DIRECT SEARCH FOR STRINGS ---
  const isMenuButton = [
    '🏠 الرئيسية', 'الرئيسية', 'الرئيسيه',
    '📁 المعاملات', 'المعاملات', 'معاملات',
    '⚙️ إدارة البيانات', 'إدارة البيانات', 'ادارة البيانات',
    '🔙 العودة للقائمة الرئيسية', '🔙 العودة لقسم المعاملات',
    '📥 تصدير ككشف excel', 'تصدير ككشف excel', 'تصدير ككشف إكسل', 'تصدير',
    '📄 كافة المعاملات', '🟢 المعاملات الجاهزة', '🟡 تحت الإجراء', '⏳ بانتظار الموافقة',
    '📥 تصدير: كافة المعاملات', '📥 تصدير: المعاملات الجاهزة', '📥 تصدير: تحت الإجراء', '📥 تصدير: بانتظار الموافقة',
    '🔍 بحث واستعلام في قسمي', '🔍 بحث واستعلام في هذا القسم', '📋 كشف العملاء'
  ].map(s => s.toLowerCase()).includes(textRaw.trim().toLowerCase());

  if (!isMenuButton && !isNumeric && !textRaw.startsWith('/') && !textRaw.startsWith('إضافة ') && !textRaw.startsWith('اضافة ') && !textRaw.startsWith('حذف ') && !textRaw.startsWith('فحص ') && !textRaw.startsWith('بحث ')) {
    const query = textRaw.trim();
    const dept = userActiveDept[chatId] || 'all';
    const deptName = dept === 'all' ? 'كافة المعاملات' : dept === 'issued' ? 'المعاملات الجاهزة' : dept === 'process' ? 'تحت الإجراء' : 'بانتظار الموافقة';
    
    const entries = readEntries();
    let filteredByDept = entries;
    if (dept === 'issued') {
      filteredByDept = entries.filter((e: any) => {
        const st = String(e.statusText || '').toLowerCase();
        return st.includes('issued') || st.includes('اصدار') || st.includes('جاهز');
      });
    } else if (dept === 'process') {
      filteredByDept = entries.filter((e: any) => {
        const st = String(e.statusText || '').toLowerCase();
        return st.includes('process') || st.includes('اجراء');
      });
    } else if (dept === 'pending_approval') {
      filteredByDept = entries.filter((e: any) => {
        const st = String(e.statusText || '').toLowerCase();
        return st.includes('pending') || st.includes('موافقة') || st.includes('انتظار');
      });
    }

    const matches = filteredByDept.filter((e: any) => {
      const nameMatch = (e.applicantName || '').includes(query);
      const passportMatch = (e.passportNumber || '').toUpperCase().includes(query.toUpperCase());
      const phoneMatch = (e.phoneNumber || '').includes(query);
      const appNoMatch = (e.applicationNumber || '').includes(query);
      return nameMatch || passportMatch || phoneMatch || appNoMatch;
    });

    if (matches.length > 0) {
      userLastMatches[chatId] = matches;
      let reply = `🔎 *نتائج البحث في قسم (${deptName}) لـ [${query}] (${matches.length} مطابقة):*\n`;
      reply += `_أرسل رقم المعاملة المتسلسل (مثال: 1) لعرض كافة التفاصيل والبيانات فوراً._\n\n`;

      matches.slice(0, 30).forEach((m: any, idx: number) => {
        reply += `*${idx + 1}.* 👤 *${m.applicantName || 'غير مسجل'}*\n🛂 جواز: \`${m.passportNumber}\` | 📞 هاتف: \`${m.phoneNumber || '---'}\`\n\n`;
      });

      if (matches.length > 30) {
        reply += `_... تم إظهار أول 30 نتيجة مطابقة._`;
      }
      await sendText(reply, departmentOptionsKeyboardMarkup);
      return;
    }
  }

  // DEFAULT REPLY if the message didn't match any commands
  await sendText(`⚠️ الأمر المكتوب غير مفهوم لمُعالج إنجاز الذكي أو لا يطابق أي عميل.
اكتب *مساعدة* أو \`/help\` لعرض قائمة الأوامر المعرّفة في السيستم بشكل مفصّل.`);
}

// REST API core routes definition for main app
app.post('/api/parse-data', async (req, res) => {
  const { rawText } = req.body;
  
  try {
    // Professional Regex Algorithm for Passport, Visa Application Numbers, and Names
    // Passport: Usually 1-2 letters followed by 6-9 digits, or 7-12 digits
    const passportRegex = /\b[A-Z]{1,2}[0-9]{6,12}\b|\b[0-9]{7,12}\b/gi;
    // Application Number: Typically 8-12 digits
    const appNoRegex = /\b[0-9]{8,15}\b/g;
    // Name attempt: Look for Arabic names or Uppercase names
    const nameRegex = /([أ-ي]{2,}\s?){2,4}|([A-Z]{2,}\s?){2,4}/g;

    const passports = (rawText.match(passportRegex) || []).map(p => p.toUpperCase());
    const appNos = rawText.match(appNoRegex) || [];
    const names = rawText.match(nameRegex) || [];

    const results = [];
    const count = Math.max(passports.length, appNos.length);
    
    for (let i = 0; i < count; i++) {
      results.push({
        passportNumber: passports[i] || '',
        applicationNumber: appNos[i] || '',
        applicantName: names[i] || ''
      });
    }

    res.json(results);
  } catch (error) {
    console.error('Parsing error:', error);
    res.status(500).json({ error: 'Failed to process data' });
  }
});

// ─── MOFA Live Proxy ────────────────────────────────────────────────────────

const MOFA_BASE = 'https://visa.mofa.gov.sa';
const MOFA_FORM_URL = `${MOFA_BASE}/visaperson/getapplicantdata`;
const CHROMIUM_PATH = '/nix/store/qa9cnw4v5xkxyip6mb9kxqfq1z4x2dx1-chromium-138.0.7204.100/bin/chromium';

// Browser-based MOFA session store
interface MofaBrowserSession {
  browser: Awaited<ReturnType<typeof puppeteer.launch>>;
  page: any;
  captchaBuffer: Buffer | null;
  autoSolvedCaptcha: string;
  nationalities: { value: string; label: string }[];
  visaTypes: { value: string; label: string }[];
  embassies: { value: string; label: string }[];
  createdAt: number;
}

// Solve captcha image using ddddocr Python script
function solveCaptchaWithOcr(imageBuffer: Buffer): string {
  const tmpPath = path.join(os.tmpdir(), `captcha_${Date.now()}.png`);
  try {
    fs.writeFileSync(tmpPath, imageBuffer);
    const result = execFileSync('python3', [path.join(process.cwd(), 'solve_captcha.py'), tmpPath], { timeout: 15000 });
    return result.toString().trim();
  } catch (e: any) {
    console.warn('[MOFA] ddddocr solve failed:', e.message);
    return '';
  } finally {
    try { fs.unlinkSync(tmpPath); } catch {}
  }
}
const mofaBrowserSessions = new Map<string, MofaBrowserSession>();

// Auto-cleanup stale sessions after 8 minutes
setInterval(() => {
  const now = Date.now();
  for (const [id, s] of mofaBrowserSessions.entries()) {
    if (now - s.createdAt > 8 * 60 * 1000) {
      try { s.browser.close(); } catch {}
      mofaBrowserSessions.delete(id);
    }
  }
}, 60 * 1000);

// ── KSA Visa Tracking Sessions ──────────────────────────────────────────────
interface KsaVisaSession {
  browser: Awaited<ReturnType<typeof puppeteer.launch>>;
  page: any;
  captchaBuffer: Buffer | null;
  autoSolvedCaptcha: string;
  createdAt: number;
}
const ksaVisaSessions = new Map<string, KsaVisaSession>();

setInterval(() => {
  const now = Date.now();
  for (const [id, s] of ksaVisaSessions.entries()) {
    if (now - s.createdAt > 8 * 60 * 1000) {
      try { s.browser.close(); } catch {}
      ksaVisaSessions.delete(id);
    }
  }
}, 60 * 1000);

// GET /api/ksavisa/init — open ksavisa.sa track-application, screenshot captcha
app.get('/api/ksavisa/init', async (req, res) => {
  if (ksaVisaSessions.size >= 5) {
    const oldest = [...ksaVisaSessions.entries()].sort((a, b) => a[1].createdAt - b[1].createdAt)[0];
    if (oldest) { try { oldest[1].browser.close(); } catch {} ksaVisaSessions.delete(oldest[0]); }
  }
  let browser: any = null;
  try {
    browser = await puppeteer.launch({
      executablePath: CHROMIUM_PATH,
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu', '--no-zygote'],
    });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 900 });
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
    await page.setExtraHTTPHeaders({ 'Accept-Language': 'en-US,en;q=0.9,ar;q=0.8' });

    console.log('[KSAVISA] Navigating to track-application...');
    await page.goto('https://ksavisa.sa/track-application', { waitUntil: 'networkidle2', timeout: 35000 });
    // Wait for Angular to fully render
    await new Promise(r => setTimeout(r, 3500));

    // Accept cookies dialog if present
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const acceptBtn = btns.find((b: any) => (b.textContent || '').trim() === 'Accept');
      if (acceptBtn) (acceptBtn as HTMLElement).click();
    }).catch(() => {});
    await new Promise(r => setTimeout(r, 600));

    // Wait for form to be ready (Angular hydration check)
    await page.waitForSelector('#applicationNumber', { timeout: 10000 }).catch(() => {});

    // ksavisa.sa captcha is a base64 data URI in img.captcha-img
    const captchaSrc: string = await page.$eval(
      'img.captcha-img, img[alt="Captcha"], img[alt*="captcha" i]',
      (el: any) => el.src || ''
    ).catch(() => '');

    let captchaBuffer: Buffer | null = null;
    let autoSolvedCaptcha = '';

    if (captchaSrc.startsWith('data:image')) {
      // Extract base64 directly — most reliable
      const base64Data = captchaSrc.split(',')[1] || '';
      captchaBuffer = Buffer.from(base64Data, 'base64');
      console.log('[KSAVISA] Captcha extracted from base64 src, length:', captchaBuffer.length);
    } else if (captchaSrc) {
      // URL-based captcha — screenshot element
      const captchaEl = await page.$('img.captcha-img, img[alt="Captcha"]').catch(() => null);
      if (captchaEl) {
        captchaBuffer = Buffer.from(await captchaEl.screenshot({ type: 'png' }));
        console.log('[KSAVISA] Captcha screenshot from URL src');
      }
    }

    if (!captchaBuffer) {
      // Fallback: full page screenshot
      captchaBuffer = Buffer.from(await page.screenshot({ type: 'png', fullPage: false }));
      console.log('[KSAVISA] No captcha found — using full page screenshot');
    } else {
      autoSolvedCaptcha = solveCaptchaWithOcr(captchaBuffer);
      console.log(`[KSAVISA] Captcha auto-solved: "${autoSolvedCaptcha}"`);
    }

    const sessionId = Math.random().toString(36).slice(2) + Date.now().toString(36);
    ksaVisaSessions.set(sessionId, { browser, page, captchaBuffer, autoSolvedCaptcha, createdAt: Date.now() });
    res.json({ sessionId, autoSolvedCaptcha, hasCaptchaEl: !!captchaSrc });
  } catch (err: any) {
    console.error('[KSAVISA] init error:', err.message);
    try { browser?.close(); } catch {}
    res.status(502).json({ error: 'فشل الاتصال بموقع KSA Visa. تأكد من الاتصال بالإنترنت.' });
  }
});

// GET /api/ksavisa/captcha — return captcha image buffer
app.get('/api/ksavisa/captcha', (req, res) => {
  const sessionId = req.query.sessionId as string;
  const session = ksaVisaSessions.get(sessionId);
  if (!session || !session.captchaBuffer) return res.status(404).send('Session not found');
  res.set('Content-Type', 'image/png');
  res.set('Cache-Control', 'no-store');
  res.send(session.captchaBuffer);
});

// POST /api/ksavisa/submit — fill form, submit, capture result page as PDF
app.post('/api/ksavisa/submit', async (req, res) => {
  const { sessionId, visaNumber, passportNumber, captcha, applicantName } = req.body as {
    sessionId: string; visaNumber: string; passportNumber: string; captcha: string; applicantName?: string;
  };
  const session = ksaVisaSessions.get(sessionId);
  if (!session) return res.status(400).json({ error: 'انتهت الجلسة، يرجى إعادة تحميل النموذج.' });

  const { page, browser } = session;
  try {
    // ── Step 1: Fill form using Angular-compatible native value setter ──────
    // NOTE: NO named const functions inside evaluate — esbuild wraps them with
    // __name() which is undefined in browser context and causes crashes.
    console.log(`[KSAVISA] Filling form: visa=${visaNumber} passport=${passportNumber} captcha=${captcha}`);
    await page.evaluate((vals: string[]) => {
      const [visa, pass, cap] = vals;
      const ns = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
      // Inline fill — no named inner functions
      const a = document.getElementById('applicationNumber') as HTMLInputElement | null;
      if (a) { if (ns) ns.call(a, visa); else a.value = visa; a.dispatchEvent(new Event('input', { bubbles: true })); a.dispatchEvent(new Event('change', { bubbles: true })); }
      const b = document.getElementById('passportNumber') as HTMLInputElement | null;
      if (b) { if (ns) ns.call(b, pass); else b.value = pass; b.dispatchEvent(new Event('input', { bubbles: true })); b.dispatchEvent(new Event('change', { bubbles: true })); }
      const c = document.getElementById('captchaInput') as HTMLInputElement | null;
      if (c) { if (ns) ns.call(c, cap); else c.value = cap; c.dispatchEvent(new Event('input', { bubbles: true })); c.dispatchEvent(new Event('change', { bubbles: true })); }
    }, [String(visaNumber || ''), String(passportNumber || ''), String(captcha || '')]);

    await new Promise(r => setTimeout(r, 600));

    // ── Step 2: Submit form ─────────────────────────────────────────────────
    // Remove disabled attribute (Angular sets it when captcha invalid) then click
    await page.evaluate(() => {
      const btn = document.querySelector('button[type="submit"]') as HTMLButtonElement | null;
      if (btn) { btn.disabled = false; btn.removeAttribute('disabled'); btn.click(); }
    });
    // Also hit Enter on captcha field as secondary trigger
    await page.focus('#captchaInput').catch(() => {});
    await page.keyboard.press('Enter');
    console.log('[KSAVISA] Form submitted — waiting for result…');

    // ── Step 3: Wait for result section (VIEW button or passport in text) ───
    // Use waitForFunction to detect success: "View Results" heading appears
    const gotResult = await Promise.race<boolean>([
      page.waitForFunction(() => {
        const txt = (document.body?.innerText || '').replace(/\s+/g, ' ');
        return txt.includes('View Results') || txt.includes('نتائج');
      }, { timeout: 14000 }).then(() => true).catch(() => false),
      new Promise<boolean>(r => setTimeout(() => r(false), 13000)),
    ]);

    await new Promise(r => setTimeout(r, 800));
    const pageText: string = await page.evaluate(() => (document.body?.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 3000));
    console.log(`[KSAVISA] gotResult=${gotResult}  pageText[0:300]: ${pageText.slice(0, 300)}`);

    // ── Error detection ─────────────────────────────────────────────────────
    const hasDataErr =
      pageText.toLowerCase().includes('no results') ||
      pageText.toLowerCase().includes('not found') ||
      pageText.toLowerCase().includes('no application found') ||
      pageText.includes('غير موجود') || pageText.includes('لم يتم العثور') ||
      pageText.includes('لا توجد نتائج');

    if (!gotResult) {
      ksaVisaSessions.delete(sessionId);
      try { browser.close(); } catch {}
      if (hasDataErr) {
        return res.json({ success: false, notFound: true,  error: 'البيانات غير صحيحة أو غير موجودة، يرجى التحقق من رقم التأشيرة ورقم الجواز.' });
      }
      return res.json({ success: false, notFound: false, error: 'رمز الصورة غير صحيح، يرجى المحاولة مرة أخرى.' });
    }

    // ── Step 4: Prepare PDF output path ────────────────────────────────────
    const safeName = (applicantName || '').replace(/\s+/g, '_');
    const safeFilename = `تاشيره_${passportNumber}_${safeName}`.replace(/[^\u0600-\u06FFa-zA-Z0-9_\-]/g, '_').replace(/_+/g, '_').slice(0, 120);
    const pdfPath = path.join(PRINTS_DIR, `${safeFilename}.pdf`);
    try {
      const oldFiles = fs.readdirSync(PRINTS_DIR).filter((f: string) => f.startsWith('تاشيره_') && f.includes(passportNumber));
      for (const f of oldFiles) { try { fs.unlinkSync(path.join(PRINTS_DIR, f)); } catch {} }
    } catch {}

    // ── Step 5: Click "VIEW" button → intercept new tab ────────────────────
    // The VIEW button (class: btn btn-primary ng-star-inserted) appears in results.
    // Set up targetcreated listener BEFORE clicking so we don't miss the event.
    console.log('[KSAVISA] Setting up new-tab listener and clicking VIEW button…');
    let visaPage: any = null;

    const newPagePromise = new Promise<any>((resolve) => {
      let settled = false;
      const listener = async (target: any) => {
        if (settled) return;
        // Only intercept page targets (not service workers / workers)
        if (target.type() !== 'page') return;
        const p = await target.page().catch(() => null);
        if (p) {
          settled = true;
          browser.off('targetcreated', listener);
          resolve(p);
        }
      };
      browser.on('targetcreated', listener);
      setTimeout(() => {
        if (!settled) { settled = true; browser.off('targetcreated', listener); resolve(null); }
      }, 10000);
    });

    // Click the VIEW / عرض button
    const btnClicked: boolean = await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button, a'));
      // Priority: exact "VIEW" button (the result-section btn), then Arabic عرض
      const priorities = [
        (t: string) => t === 'VIEW',
        (t: string) => t.includes('عرض البيانات'),
        (t: string) => t === 'عرض',
        (t: string) => t.toLowerCase() === 'view',
        (t: string) => t.toLowerCase().includes('view visa'),
        (t: string) => t.toLowerCase().includes('view data'),
      ];
      for (const match of priorities) {
        const btn = btns.find((el: any) => match((el.textContent || '').trim()));
        if (btn) { (btn as HTMLElement).click(); return true; }
      }
      return false;
    });
    console.log(`[KSAVISA] VIEW button clicked: ${btnClicked}`);

    visaPage = await newPagePromise;

    if (visaPage) {
      console.log('[KSAVISA] ✅ New tab opened for visa image');
      // Wait for the visa image page to fully load
      await Promise.race([
        visaPage.waitForNavigation({ timeout: 18000, waitUntil: 'networkidle2' }).catch(() => {}),
        new Promise(r => setTimeout(r, 15000)),
      ]);
      await new Promise(r => setTimeout(r, 2000));
      const visaUrl: string = visaPage.url();
      console.log(`[KSAVISA] Visa page URL: ${visaUrl}`);
    } else {
      console.log('[KSAVISA] No new tab — capturing current page instead');
      await new Promise(r => setTimeout(r, 2000));
    }

    // ── Step 6: Capture PDF from the visa page (new tab) or current page ───
    const targetPage: any = visaPage || page;

    // Wait for document to be fully loaded
    await targetPage.evaluate(() => {
      return new Promise<void>((resolve) => {
        if ((document as any).readyState === 'complete') { resolve(); return; }
        window.addEventListener('load', () => resolve(), { once: true });
        setTimeout(resolve, 8000);
      });
    }).catch(() => {});
    await new Promise(r => setTimeout(r, 2000));

    const contentWidth: number = await targetPage.evaluate(() =>
      Math.max((document.body || {}).scrollWidth || 0, (document.documentElement || {}).scrollWidth || 0, 1280)
    ).catch(() => 1280);
    await targetPage.setViewport({ width: Math.max(contentWidth, 1280) + 40, height: 1200 });

    // Scroll through page to trigger lazy-loaded images
    await targetPage.evaluate(() => {
      window.scrollTo(0, document.body.scrollHeight);
    }).catch(() => {});
    await new Promise(r => setTimeout(r, 1000));
    await targetPage.evaluate(() => {
      window.scrollTo(0, 0);
    }).catch(() => {});

    // Wait for ALL images to finish loading
    await targetPage.evaluate(() => {
      const imgs = Array.from(document.images);
      if (imgs.length === 0) return Promise.resolve();
      const pending = imgs.filter((img: any) => !img.complete);
      if (pending.length === 0) return Promise.resolve();
      return Promise.race([
        Promise.all(pending.map((img: any) => new Promise<void>(resolve => {
          img.onload = () => resolve();
          img.onerror = () => resolve();
        }))),
        new Promise<void>(resolve => setTimeout(resolve, 8000)),
      ]);
    }).catch(() => {});
    await new Promise(r => setTimeout(r, 1500));

    // ── Hide top site-navbar only, keep visa document header, fit single page ─
    await targetPage.evaluate(() => {
      const style = document.createElement('style');
      style.textContent = `
        /* Hide ONLY the site navigation bar at the very top — not the visa document content */
        body > app-root > * > header,
        body > app-root > header,
        body > header,
        body > nav,
        app-header, app-navbar, app-nav,
        app-root > app-layout > header,
        app-root > app-layout > nav,
        .site-header, .site-nav, .site-navbar,
        .top-header, .top-nav,
        .main-header:first-of-type,
        [class*="site-header"], [class*="site-nav"],
        [id*="site-header"], [id*="site-nav"] {
          display: none !important;
        }
        /* Remove body/html spacing */
        body, html { margin: 0 !important; padding: 0 !important; }
        /* Remove any top padding from the first main container */
        body > * { padding-top: 0 !important; margin-top: 0 !important; }
        app-root { padding-top: 0 !important; margin-top: 0 !important; }
      `;
      document.head.appendChild(style);

      // Also try: hide the first child of body/app-root IF it contains no visa-like content
      // (the navbar typically has text like "موقع حكومي" or "KSA" navigation links)
      const root = document.querySelector('app-root') || document.body;
      const children = Array.from(root.children) as HTMLElement[];
      for (const child of children) {
        const tag = child.tagName.toLowerCase();
        const txt = child.innerText || '';
        // If element is a nav/header tag or contains the site-nav text → hide it
        if (
          ['header','nav','app-header','app-navbar','app-nav'].includes(tag) ||
          (txt.includes('موقع حكومي') && !txt.includes('Visa No')) ||
          (txt.includes('كيفية التحقق') && !txt.includes('Visa No'))
        ) {
          child.style.display = 'none';
        }
      }

      // Find the machine-readable zone (MRZ: lines like 1<YEM... / passport number<)
      // Hide everything that comes after it
      const allLeaves = Array.from(document.querySelectorAll('*')).filter(
        (el: any) => el.children.length === 0
      );
      let mrzEl: Element | null = null;
      for (const el of allLeaves) {
        const txt = ((el as HTMLElement).innerText || '').trim();
        if (/^[A-Z0-9<]{20,}$/.test(txt)) mrzEl = el;
      }
      if (mrzEl) {
        let node: Element | null = mrzEl;
        while (node && node.parentElement && node.parentElement !== document.body) {
          node = node.parentElement;
        }
        if (node) {
          let sib = node.nextElementSibling;
          while (sib) { (sib as HTMLElement).style.display = 'none'; sib = sib.nextElementSibling; }
        }
      }
    }).catch(() => {});
    // ──────────────────────────────────────────────────────────────────────

    // ── Inject visa info stamp into the page before PDF capture ──────────
    await targetPage.evaluate((info: string[]) => {
      const [vNum, pNum, aName] = info;
      const banner = document.createElement('div');
      banner.id = '__visa_stamp__';
      banner.style.cssText = [
        'width:100%',
        'box-sizing:border-box',
        'background:#1a3d6e',
        'color:#ffffff',
        'font-family:Arial,"Segoe UI",Tahoma,sans-serif',
        'padding:10px 18px',
        'display:flex',
        'flex-direction:row',
        'justify-content:space-between',
        'align-items:center',
        'gap:12px',
        'margin:0',
        'border-bottom:3px solid #c8a227',
        'page-break-inside:avoid',
        'print-color-adjust:exact',
        '-webkit-print-color-adjust:exact',
      ].join(';');

      const cell = (label: string, value: string) => {
        const d = document.createElement('div');
        d.style.cssText = 'display:flex;flex-direction:column;align-items:center;min-width:0;';
        const l = document.createElement('span');
        l.textContent = label;
        l.style.cssText = 'font-size:9px;color:#a8c4e0;letter-spacing:0.5px;white-space:nowrap;text-transform:uppercase;';
        const v = document.createElement('span');
        v.textContent = value || '—';
        v.style.cssText = 'font-size:13px;font-weight:bold;color:#ffffff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:200px;direction:rtl;';
        d.appendChild(l);
        d.appendChild(v);
        return d;
      };

      // Title: "تأشيرة" label
      const titleEl = document.createElement('div');
      titleEl.style.cssText = 'display:flex;flex-direction:column;align-items:center;min-width:0;';
      const titleVal = document.createElement('span');
      titleVal.textContent = 'تأشيرة';
      titleVal.style.cssText = 'font-size:22px;font-weight:900;color:#c8a227;letter-spacing:2px;white-space:nowrap;';
      titleEl.appendChild(titleVal);
      banner.appendChild(titleEl);

      banner.appendChild(cell('رقم الجواز / Passport', pNum));
      banner.appendChild(cell('اسم العميل / Name', aName));

      document.body.insertBefore(banner, document.body.firstChild);
    }, [String(visaNumber || ''), String(passportNumber || ''), String(applicantName || '')]).catch(() => {});
    // ─────────────────────────────────────────────────────────────────────

    // Measure exact content bottom after hiding nav — find last visible element
    const contentHeightPx: number = await targetPage.evaluate(() => {
      // Find the bottom of the last visible element (MRZ or last child)
      const allEls = Array.from(document.querySelectorAll('*')) as HTMLElement[];
      let maxBottom = document.body.scrollHeight;
      for (const el of allEls) {
        const style = window.getComputedStyle(el);
        if (style.display === 'none' || style.visibility === 'hidden') continue;
        const rect = el.getBoundingClientRect();
        if (rect.bottom > 0 && rect.bottom + window.scrollY > maxBottom) {
          maxBottom = rect.bottom + window.scrollY;
        }
      }
      return Math.max(maxBottom, document.body.scrollHeight);
    }).catch(() => 1400);

    // 1 CSS px = 25.4/96 mm at default 96dpi
    const heightMm = Math.ceil(contentHeightPx * (25.4 / 96)) + 6;

    const pdfBuffer = await targetPage.pdf({
      width: '210mm',
      height: `${heightMm}mm`,
      printBackground: true,
      margin: { top: '0mm', bottom: '0mm', left: '5mm', right: '5mm' },
    });
    fs.writeFileSync(pdfPath, pdfBuffer);
    console.log(`[KSAVISA] ✅ Saved visa PDF → ${pdfPath}`);
    const clientCaption = `${applicantName || ''} — ${passportNumber || ''}`.trim();
    const tgVisaUrl = await autoUploadToTelegram(pdfPath, clientCaption);
    const visaPdfUrl = tgVisaUrl || `/prints/${safeFilename}.pdf`;

    ksaVisaSessions.delete(sessionId);
    try { browser.close(); } catch {}
    res.json({ success: true, visaPdfUrl, pageText });
  } catch (err: any) {
    console.error('[KSAVISA] submit error:', err.message);
    ksaVisaSessions.delete(sessionId);
    try { browser.close(); } catch {}
    res.status(500).json({ error: 'فشل جلب بيانات التأشيرة: ' + err.message });
  }
});

// GET /api/mofa/init — launch real Chrome, load MOFA page, screenshot captcha
app.get('/api/mofa/init', async (req, res) => {
  // Limit concurrent browser sessions to 5
  if (mofaBrowserSessions.size >= 5) {
    const oldest = [...mofaBrowserSessions.entries()].sort((a, b) => a[1].createdAt - b[1].createdAt)[0];
    if (oldest) { try { oldest[1].browser.close(); } catch {} mofaBrowserSessions.delete(oldest[0]); }
  }
  let browser: any = null;
  try {
    browser = await puppeteer.launch({
      executablePath: CHROMIUM_PATH,
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu', '--no-zygote'],
    });
    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 900 });
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
    await page.setExtraHTTPHeaders({ 'Accept-Language': 'ar,en-US;q=0.7,en;q=0.3' });

    await page.goto(MOFA_FORM_URL, { waitUntil: 'networkidle2', timeout: 30000 });

    // Wait for captcha image and JS to finish refreshing it
    await page.waitForSelector('#imgCaptcha', { timeout: 15000 });
    await new Promise(r => setTimeout(r, 2500));

    // Screenshot just the captcha image element
    const captchaEl = await page.$('#imgCaptcha');
    let captchaBuffer: Buffer | null = null;
    let autoSolvedCaptcha = '';
    if (captchaEl) {
      captchaBuffer = Buffer.from(await captchaEl.screenshot({ type: 'png' }));
      autoSolvedCaptcha = solveCaptchaWithOcr(captchaBuffer);
      console.log(`[MOFA] ddddocr captcha result: "${autoSolvedCaptcha}"`);
    }

    // Extract form select options from live DOM
    const nationalities: {value:string;label:string}[] = await page.evaluate(() =>
      Array.from((document.getElementById('Nationality') as HTMLSelectElement)?.options ?? [])
        .filter((o:any) => o.value).map((o:any) => ({ value: o.value, label: o.text.trim() }))
    );
    const visaTypes: {value:string;label:string}[] = await page.evaluate(() =>
      Array.from((document.getElementById('VisaType') as HTMLSelectElement)?.options ?? [])
        .filter((o:any) => o.value).map((o:any) => ({ value: o.value, label: o.text.trim() }))
    );
    const embassies: {value:string;label:string}[] = await page.evaluate(() =>
      Array.from((document.getElementById('Embassy') as HTMLSelectElement)?.options ?? [])
        .filter((o:any) => o.value).map((o:any) => ({ value: o.value, label: o.text.trim() }))
    );

    const sessionId = Math.random().toString(36).slice(2) + Date.now().toString(36);
    mofaBrowserSessions.set(sessionId, { browser, page, captchaBuffer, autoSolvedCaptcha, nationalities, visaTypes, embassies, createdAt: Date.now() });

    console.log(`[MOFA] Browser session created: ${sessionId}, nats: ${nationalities.length}, vts: ${visaTypes.length}, captcha: "${autoSolvedCaptcha}"`);
    res.json({ sessionId, nationalities, visaTypes, embassies, autoSolvedCaptcha });
  } catch (err: any) {
    console.error('MOFA init error:', err.message);
    try { browser?.close(); } catch {}
    res.status(502).json({ error: 'فشل الاتصال بمنصة التأشيرات. تأكد من الاتصال بالإنترنت.' });
  }
});

// GET /api/mofa/captcha — serve pre-screenshotted captcha from browser session
app.get('/api/mofa/captcha', async (req, res) => {
  const sessionId = req.query.sessionId as string;
  const session = mofaBrowserSessions.get(sessionId);
  if (!session || !session.captchaBuffer) return res.status(404).send('Session not found');
  res.set('Content-Type', 'image/png');
  res.set('Cache-Control', 'no-store');
  res.send(session.captchaBuffer);
});

// POST /api/mofa/submit — use live browser page to fill & submit the MOFA form
app.post('/api/mofa/submit', async (req, res) => {
  const { sessionId, passportNumber, nationality, visaType, embassy, captcha, existingPrintUrl, applicantName: frontendApplicantName } = req.body;
  const session = mofaBrowserSessions.get(sessionId);
  if (!session) return res.status(400).json({ error: 'انتهت الجلسة، يرجى إعادة تحميل النموذج.' });

  const { page, browser } = session;
  try {
    console.log(`[MOFA] Browser submit: passport=${passportNumber} nat=${nationality} vt=${visaType} emb=${embassy}`);

    // Fill form fields via DOM and fire change/input events so JS validation triggers
    // NOTE: no named const functions inside evaluate — esbuild wraps them with __name which breaks in browser context
    await page.evaluate((pp: string, nat: string, vt: string, emb: string) => {
      const ppEl = document.getElementById('PassportNumber') as HTMLInputElement;
      if (ppEl) { ppEl.value = pp; ppEl.dispatchEvent(new Event('input', { bubbles: true })); ppEl.dispatchEvent(new Event('change', { bubbles: true })); }
      const natEl = document.getElementById('Nationality') as HTMLSelectElement;
      if (natEl) { natEl.value = nat; natEl.dispatchEvent(new Event('input', { bubbles: true })); natEl.dispatchEvent(new Event('change', { bubbles: true })); }
      const vtEl = document.getElementById('VisaType') as HTMLSelectElement;
      if (vtEl) { vtEl.value = vt; vtEl.dispatchEvent(new Event('input', { bubbles: true })); vtEl.dispatchEvent(new Event('change', { bubbles: true })); }
      const embEl = document.getElementById('Embassy') as HTMLSelectElement;
      if (embEl) { embEl.value = emb; embEl.dispatchEvent(new Event('input', { bubbles: true })); embEl.dispatchEvent(new Event('change', { bubbles: true })); }
    }, passportNumber, nationality, visaType, embassy);

    // Log what was actually set in the form for debugging
    const formState = await page.evaluate(() => {
      const pp  = (document.getElementById('PassportNumber') as HTMLInputElement)?.value;
      const nat = (document.getElementById('Nationality')    as HTMLSelectElement)?.value;
      const vt  = (document.getElementById('VisaType')       as HTMLSelectElement)?.value;
      const emb = (document.getElementById('Embassy')        as HTMLSelectElement)?.value;
      return { pp, nat, vt, emb };
    });
    console.log(`[MOFA] Form state after fill:`, JSON.stringify(formState));

    // Clear captcha field and type the user's code
    await page.evaluate(() => { (document.getElementById('Captcha') as HTMLInputElement).value = ''; });
    await page.focus('#Captcha');
    await page.type('#Captcha', String(captcha), { delay: 30 });

    // Click submit and wait for navigation
    const submitBtn = await page.$('input[type="submit"], button[type="submit"], input[name="Submit"]');
    if (!submitBtn) throw new Error('زر الإرسال غير موجود في الصفحة');

    await Promise.all([
      page.waitForNavigation({ timeout: 30000, waitUntil: 'networkidle2' }),
      submitBtn.click(),
    ]);

    const urlAfterNav: string = page.url();
    console.log(`[MOFA] After submit URL: ${urlAfterNav}`);

    // If still on form page → captcha wrong or data mismatch
    const quickHtml: string = await page.content();
    if (quickHtml.includes('id="PassportNumber"') || quickHtml.includes('id="Captcha"')) {
      mofaBrowserSessions.delete(sessionId);
      try { browser.close(); } catch {}
      // Extract plain text from page to detect the actual error message
      const quickBodyText = quickHtml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
      const errMatch = quickHtml.match(/class="[^"]*(?:alert-danger|validation-summary-errors|text-danger)[^"]*"[^>]*>([\s\S]{0,300}?)<\/(?:div|ul|span)>/i);
      const errText = errMatch ? errMatch[1].replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() : '';
      // Check if page body contains "no data" message (data mismatch, not captcha error)
      const isNoData = quickBodyText.includes('حدث خطأ يرجى التأكد من البيانات') ||
                       quickBodyText.includes('يرجى التأكد من البيانات المدخلة') ||
                       errText.includes('حدث خطأ') ||
                       errText.includes('يرجى التأكد من البيانات');
      const finalErr = isNoData
        ? 'حدث خطأ يرجى التأكد من البيانات المدخلة والمحاولة مرة أخرى'
        : (errText || 'رمز الصورة المدخل غير صحيح');
      console.log(`[MOFA] Form page error: isNoData=${isNoData}, errText="${errText}", finalErr="${finalErr}"`);
      return res.json({ success: false, notFound: isNoData, error: finalErr });
    }

    // ✅ We are on the results page — wait extra time for AJAX/jTable to fully render data
    console.log(`[MOFA] On results page, waiting 4s for AJAX data to load...`);
    await new Promise(r => setTimeout(r, 4000));

    // Re-read page after AJAX settles
    const html: string = await page.content();
    console.log(`[MOFA] HTML length after AJAX wait: ${html.length}`);

    // Log body text to understand page structure
    const bodyTextForLog: string = await page.evaluate(() => (document.body?.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 3000));
    console.log(`[MOFA] Body text after AJAX: ${bodyTextForLog}`);

    // Check for "not found" ONLY after AJAX wait
    const notFoundInBody = bodyTextForLog.includes('لم يتم العثور') || bodyTextForLog.includes('لاتوجد نتائج') || bodyTextForLog.includes('No Results');
    const notFoundInHtml = html.includes('ResultNotFound');
    console.log(`[MOFA] notFoundInBody=${notFoundInBody}, notFoundInHtml=${notFoundInHtml}`);

    if (notFoundInHtml && !bodyTextForLog.includes('اسم') && !bodyTextForLog.includes('جواز')) {
      mofaBrowserSessions.delete(sessionId);
      try { browser.close(); } catch {}
      return res.json({ success: false, notFound: true, error: 'لم يتم العثور على أي نتائج، يرجى التأكد من البيانات المدخلة.' });
    }

    // ✅ Extract data using live DOM — NO named inner functions (esbuild __name issue)
    const data: any = await page.evaluate(() => {
      const bodyText = (document.body?.innerText || '').replace(/\s+/g, ' ').trim();

      // Build label→value maps once using anonymous callbacks (no named functions)
      const lm: Record<string, string> = {};
      Array.from(document.querySelectorAll('td,th,label,span,div,p,li,dt,dd')).forEach(el => {
        const lbl = (el.textContent || '').trim();
        if (!lbl || lbl.length > 80) return;
        const key = lbl.replace(/:?\s*$/, '').trim();
        if (lm[key]) return;
        let sib = el.nextElementSibling;
        while (sib) {
          const v = (sib.textContent || '').trim();
          if (v) { lm[key] = v; return; }
          sib = sib.nextElementSibling;
        }
        if (el.parentElement) {
          const sibs = Array.from(el.parentElement.children);
          const i = sibs.indexOf(el as Element);
          for (let j = i + 1; j < sibs.length; j++) {
            const v = (sibs[j].textContent || '').trim();
            if (v) { lm[key] = v; return; }
          }
        }
      });

      const rm: Record<string, string> = {};
      Array.from(document.querySelectorAll('tr')).forEach(row => {
        const cells = Array.from(row.querySelectorAll('td,th'));
        for (let i = 0; i < cells.length - 1; i++) {
          const lbl = (cells[i].textContent || '').trim();
          if (!lbl) continue;
          for (let j = i + 1; j < cells.length; j++) {
            const v = (cells[j].textContent || '').trim();
            if (v && !rm[lbl]) { rm[lbl] = v; break; }
          }
        }
      });

      const dm: Record<string, string> = {};
      Array.from(document.querySelectorAll('dt')).forEach(dt => {
        const lbl = (dt.textContent || '').trim();
        const dd = dt.nextElementSibling;
        if (dd && dd.tagName === 'DD') dm[lbl] = (dd.textContent || '').trim();
      });

      // Inlined lookups only — no named functions (avoids esbuild __name)
      return {
        applicantName:      lm['اسم صاحب الطلب']||rm['اسم صاحب الطلب']||dm['اسم صاحب الطلب']||lm['الاسم']||rm['الاسم']||lm['Applicant Name']||rm['Applicant Name']||'',
        nameEnglish:        lm['Name']||rm['Name']||lm['الاسم بالانجليزي']||rm['الاسم بالانجليزي']||'',
        applicationNumber:  lm['رقم الطلب']||rm['رقم الطلب']||dm['رقم الطلب']||lm['Application Number']||rm['Application Number']||(bodyText.match(/\b(70\d{8,11}|E\d{8,12})\b/)||[])[1]||'',
        applicationDate:    lm['تاريخ الطلب']||rm['تاريخ الطلب']||dm['تاريخ الطلب']||lm['Application Date']||rm['Application Date']||(bodyText.match(/تاريخ الطلب[:\s]+(\d{4}\/\d{2}\/\d{2})/)||[])[1]||(bodyText.match(/تاريخ الطلب[:\s]+(\d{2}\/\d{2}\/\d{4})/)||[])[1]||'',
        embassy:            lm['الممثلية في']||rm['الممثلية في']||dm['الممثلية في']||lm['Embassy']||rm['Embassy']||'',
        visaNumber:         lm['رقم التأشيرة']||rm['رقم التأشيرة']||dm['رقم التأشيرة']||lm['Visa Number']||rm['Visa Number']||(bodyText.match(/(?:تم إصدار التأشيرة|التأشيرة|Visa)[^0-9]{0,20}برقم\s*(\d{6,15})/)||[])[1]||(bodyText.match(/برقم\s+(\d{6,15})/)||[])[1]||(bodyText.match(/\b(V\d{7,12})\b/)||[])[1]||(bodyText.match(/\b(40\d{8,11})\b/)||[])[1]||'',
        documentNumber:     lm['رقم المستند']||rm['رقم المستند']||dm['رقم المستند']||lm['رقم الوثيقة']||rm['رقم الوثيقة']||dm['رقم الوثيقة']||lm['Document Number']||rm['Document Number']||(bodyText.match(/(?:رقم المستند|رقم الوثيقة|Document Number|Doc No)\s*[:：\-]?\s*([A-Za-z0-9]{4,25})/)||[])[1]||'',
        entriesCount:       lm['عدد مرات الدخول']||rm['عدد مرات الدخول']||dm['عدد مرات الدخول']||lm['Number of Entries']||rm['Number of Entries']||'',
        requesterName:      lm['اسم الشخص/الجهة الطالبة']||rm['اسم الشخص/الجهة الطالبة']||lm['اسم الجهة الطالبة']||rm['اسم الجهة الطالبة']||lm['Requester']||rm['Requester']||'',
        passportType:       lm['نوع الجواز']||rm['نوع الجواز']||dm['نوع الجواز']||lm['Passport Type']||rm['Passport Type']||'',
        passportExpiry:     lm['تاريخ الانتهاء']||rm['تاريخ الانتهاء']||dm['تاريخ الانتهاء']||lm['Expiry Date']||rm['Expiry Date']||lm['تاريخ انتهاء الجواز']||rm['تاريخ انتهاء الجواز']||'',
        birthDate:          lm['تاريخ الميلاد']||rm['تاريخ الميلاد']||dm['تاريخ الميلاد']||lm['Date of Birth']||rm['Date of Birth']||'',
        birthPlace:         lm['مكان الميلاد']||rm['مكان الميلاد']||dm['مكان الميلاد']||lm['Place of Birth']||rm['Place of Birth']||'',
        currentNationality: lm['الجنسية الحالية']||rm['الجنسية الحالية']||dm['الجنسية الحالية']||lm['Current Nationality']||rm['Current Nationality']||lm['الجنسية']||rm['الجنسية']||dm['الجنسية']||'',
        gender:             lm['الجنس']||rm['الجنس']||dm['الجنس']||lm['Gender']||rm['Gender']||'',
        profession:         lm['المهنة']||rm['المهنة']||dm['المهنة']||lm['Profession']||rm['Profession']||lm['المسمى الوظيفي']||rm['المسمى الوظيفي']||'',
        purpose:            lm['الغرض']||rm['الغرض']||dm['الغرض']||lm['Purpose']||rm['Purpose']||lm['الغرض من الزيارة']||rm['الغرض من الزيارة']||'',
        nationality:        lm['الجنسية']||rm['الجنسية']||dm['الجنسية']||lm['Nationality']||rm['Nationality']||'',
        arrivalPoint:       lm['جهة القدوم']||rm['جهة القدوم']||dm['جهة القدوم']||lm['Arrival Point']||rm['Arrival Point']||'',
        visaType:           lm['نوع التأشيرة']||rm['نوع التأشيرة']||dm['نوع التأشيرة']||lm['Visa Type']||rm['Visa Type']||'',
        rawText:            bodyText.slice(0, 3000),
      };
    });

    console.log(`[MOFA] Extracted data: name=${data.applicantName}, appNo=${data.applicationNumber}, visa=${data.visaNumber}, profession=${data.profession}`);
    console.log(`[MOFA] rawText snippet: ${data.rawText.slice(0, 500)}`);

    // ── Auto-capture PDF from print page BEFORE closing browser ──────────────
    let printPdfUrl: string | null = null;

    // If the entry already has a print file, skip capture and reuse it
    const hasPrintAlready = !!(existingPrintUrl && existingPrintUrl.trim() !== '' && existingPrintUrl.trim() !== '---');
    if (hasPrintAlready) {
      printPdfUrl = existingPrintUrl.trim();
      console.log(`[MOFA] pdf-capture: skipped — existing print file reused: ${printPdfUrl}`);
    }

    if (!hasPrintAlready) try {
      const clientName = (data.applicantName || '').replace(/\s+/g, '_');
      const safeFilename = `برنت_${passportNumber}_${clientName}`.replace(/[^\u0600-\u06FFa-zA-Z0-9_\-]/g, '_').replace(/_+/g, '_').slice(0, 120);
      const pdfPath = path.join(PRINTS_DIR, `${safeFilename}.pdf`);
      // Delete old print file if exists (to ensure fresh update)
      try {
        const oldFiles = fs.readdirSync(PRINTS_DIR).filter((f: string) => f.includes(passportNumber) && (f.endsWith('.pdf') || f.endsWith('.jpg')));
        for (const oldFile of oldFiles) {
          const oldPath = path.join(PRINTS_DIR, oldFile);
          if (oldPath !== pdfPath) fs.unlinkSync(oldPath);
        }
      } catch {}  

      // Try to find and click the print button
      const printBtnHandle = await page.evaluateHandle(() => {
        const candidates = Array.from(document.querySelectorAll('a, button, input[type="button"], input[type="submit"], input[type="image"]'));
        return candidates.find((el: any) => {
          const txt = ((el.textContent || '') + (el.value || '') + (el.title || '') + (el.alt || '')).trim();
          const hr  = (el.href || el.onclick?.toString() || '').toString();
          return txt.includes('طباعة') || txt.toLowerCase().includes('print') || hr.toLowerCase().includes('print');
        }) || null;
      });

      const printElement = printBtnHandle.asElement();
      let printPage = page;

      if (printElement) {
        console.log('[MOFA] pdf-capture: found print button, clicking…');
        const newPagePromise = new Promise<any>(resolve => {
          const listener = async (target: any) => {
            const p = await target.page().catch(() => null);
            if (p) { browser.off('targetcreated', listener); resolve(p); }
          };
          browser.on('targetcreated', listener);
          setTimeout(() => { browser.off('targetcreated', listener); resolve(null); }, 6000);
        });
        await printElement.click().catch(() => {});
        const newPage = await newPagePromise;
        if (newPage) {
          await newPage.waitForNavigation({ timeout: 15000, waitUntil: 'networkidle2' }).catch(() => {});
          await new Promise(r => setTimeout(r, 1500));
          printPage = newPage;
        } else {
          await new Promise(r => setTimeout(r, 2000));
        }
      }

      // Get actual content width so nothing gets clipped
      const contentWidth: number = await (printPage as any).evaluate(() =>
        Math.max(
          (document.body || {}).scrollWidth || 0,
          (document.documentElement || {}).scrollWidth || 0,
          1587
        )
      ).catch(() => 1587);
      await (printPage as any).setViewport({ width: contentWidth + 40, height: 1122 }).catch(() => {});
      await new Promise(r => setTimeout(r, 600));

      // Generate PDF using A4 portrait format
      const pdfBuffer = await (printPage as any).pdf({
        format: 'A4',
        landscape: false,
        printBackground: true,
        pageRanges: '1',
        margin: { top: '10mm', bottom: '10mm', left: '8mm', right: '8mm' },
      });
      fs.writeFileSync(pdfPath, pdfBuffer);
      console.log(`[MOFA] pdf-capture: saved PDF → ${pdfPath}`);
      const clientCaption = `${(data.applicantName || frontendApplicantName || '')} — ${passportNumber || ''}`.trim();
      const tgPrintUrl = await autoUploadToTelegram(pdfPath, clientCaption);
      printPdfUrl = tgPrintUrl || `/prints/${safeFilename}.pdf`;
    } catch (pdfErr: any) {
      console.warn('[MOFA] pdf-capture failed (non-fatal):', pdfErr.message);
    }
    // ─────────────────────────────────────────────────────────────────────────

    // Now close the browser AFTER data extraction and PDF capture
    mofaBrowserSessions.delete(sessionId);
    try { browser.close(); } catch {}

    // Determine status from known keywords
    const bt = data.rawText;
    let statusText = '';
    if (data.visaNumber || html.includes('تم اصدار التأشيرة') || html.includes('Visa Issued') || bt.includes('تم اصدار التأشيرة')) {
      statusText = 'تم اصدار التأشيرة (Issued)';
    } else if (html.includes('تحت الاجراء') || html.includes('Under Process') || html.includes('تحت الإجراء') || bt.includes('تحت الاجراء') || bt.includes('تحت الإجراء')) {
      statusText = 'تحت الاجراء ومطابقة الجوازات (Under Process)';
    } else if (html.includes('لاتوجد تأشيرة صادرة') || html.includes('VisaNotExported') || bt.includes('لاتوجد تأشيرة صادرة')) {
      statusText = 'لا توجد تأشيرة صادرة';
    } else if (bt.includes('مدخل') || bt.includes('قيد المعالجة') || bt.includes('تحت المعالجة')) {
      statusText = 'مدخل - قيد المعالجة';
    } else if (data.applicantName) {
      statusText = 'تم الاستعلام - البيانات متاحة';
    } else {
      statusText = 'تم الاستعلام';
    }

    res.json({
      success: true,
      data: {
        applicantName: data.applicantName,
        nameEnglish: data.nameEnglish,
        applicationNumber: data.applicationNumber,
        applicationDate: data.applicationDate,
        embassy: data.embassy,
        visaNumber: data.visaNumber,
        documentNumber: data.documentNumber,
        entriesCount: data.entriesCount,
        requesterName: data.requesterName,
        passportType: data.passportType,
        passportExpiry: data.passportExpiry,
        birthDate: data.birthDate,
        birthPlace: data.birthPlace,
        currentNationality: data.currentNationality,
        gender: data.gender,
        profession: data.profession,
        purpose: data.purpose,
        nationality: data.nationality,
        arrivalPoint: data.arrivalPoint,
        visaType: data.visaType,
        statusText,
        printPdfUrl,
        rawHtml: html.slice(0, 8000),
      }
    });
  } catch (err: any) {
    console.error('MOFA submit error:', err.message);
    mofaBrowserSessions.delete(sessionId);
    try { browser.close(); } catch {}
    res.status(502).json({ error: 'فشل الاتصال بمنصة التأشيرات. يرجى المحاولة مرة أخرى.' });
  }
});

// GET /api/mofa/print-screenshot — legacy: full-page PNG as base64 (kept for fallback)
app.get('/api/mofa/print-screenshot', async (req, res) => {
  const sessionId = req.query.sessionId as string;
  const session = mofaBrowserSessions.get(sessionId);
  if (!session) return res.status(404).json({ error: 'الجلسة غير موجودة أو انتهت. يرجى إعادة الاستعلام.' });
  try {
    const { page } = session;
    const screenshotBuffer = await page.screenshot({ type: 'png', fullPage: true });
    const base64 = `data:image/png;base64,${Buffer.from(screenshotBuffer).toString('base64')}`;
    res.json({ base64 });
  } catch (err: any) {
    res.status(500).json({ error: 'فشل تصوير صفحة البرنت.' });
  }
});

// POST /api/mofa/capture-print — click the print button on MOFA results page,
// capture the print-friendly page and save it to ./prints/ on the server.
app.post('/api/mofa/capture-print', async (req, res) => {
  const { sessionId, filename } = req.body as { sessionId: string; filename: string };
  const session = mofaBrowserSessions.get(sessionId);
  if (!session) return res.status(404).json({ error: 'الجلسة غير موجودة أو انتهت. يرجى إعادة الاستعلام.' });

  try {
    const { page, browser } = session;

    // Ensure prints directory exists
    const printsDir = path.join(process.cwd(), 'prints');
    if (!fs.existsSync(printsDir)) fs.mkdirSync(printsDir, { recursive: true });

    // Safe filename (keep Arabic chars, latin, digits, dash, underscore) — prefix with برنت
    const rawFilename = (filename || 'print').startsWith('برنت') ? (filename || 'print') : `برنت_${filename || 'print'}`;
    const safeFilename = rawFilename.replace(/[^\u0600-\u06FF\w\-]/g, '_').replace(/_+/g, '_').slice(0, 120);
    const filePath = path.join(printsDir, `${safeFilename}.png`);

    console.log(`[MOFA] capture-print: looking for print button, file="${safeFilename}.png"`);

    // Find the print button by Arabic text or common print-related attributes
    const printBtnHandle = await page.evaluateHandle(() => {
      const candidates = Array.from(document.querySelectorAll('a, button, input[type="button"], input[type="submit"], input[type="image"]'));
      return candidates.find((el: any) => {
        const txt = ((el.textContent || '') + (el.value || '') + (el.title || '') + (el.alt || '')).trim();
        const hr  = (el.href || el.onclick?.toString() || '').toString();
        return txt.includes('طباعة') || txt.toLowerCase().includes('print') || hr.toLowerCase().includes('print');
      }) || null;
    });

    const printElement = printBtnHandle.asElement();
    let printPage = page;

    if (printElement) {
      console.log(`[MOFA] capture-print: found print button, clicking…`);
      // Listen for new tab/page before clicking
      const newPagePromise = new Promise<any>(resolve => {
        const listener = async (target: any) => {
          const p = await target.page().catch(() => null);
          if (p) { browser.off('targetcreated', listener); resolve(p); }
        };
        browser.on('targetcreated', listener);
        setTimeout(() => { browser.off('targetcreated', listener); resolve(null); }, 6000);
      });

      await printElement.click().catch(() => {});

      const newPage = await newPagePromise;
      if (newPage) {
        console.log(`[MOFA] capture-print: new tab opened`);
        await newPage.waitForNavigation({ timeout: 15000, waitUntil: 'networkidle2' }).catch(() => {});
        await new Promise(r => setTimeout(r, 1500));
        printPage = newPage;
      } else {
        // No new tab — results may have changed on current page; wait briefly
        await new Promise(r => setTimeout(r, 2000));
        printPage = page;
      }
    } else {
      console.log(`[MOFA] capture-print: no print button found, capturing current page`);
    }

    // Capture full-page screenshot in memory only (not saved to disk)
    await printPage.setViewport({ width: 1280, height: 900 }).catch(() => {});
    const buffer = await printPage.screenshot({ type: 'jpeg', quality: 75, fullPage: true });
    console.log(`[MOFA] capture-print: captured ${buffer.length} bytes (in-memory only, not saved)`);

    const base64 = `data:image/jpeg;base64,${buffer.toString('base64')}`;
    res.json({ url: null, filename: `${safeFilename}.jpg`, base64 });
  } catch (err: any) {
    console.error('[MOFA] capture-print error:', err.message);
    res.status(500).json({ error: 'فشل التقاط صورة البرنت: ' + err.message });
  }
});

// ─── End MOFA Proxy ──────────────────────────────────────────────────────────

async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true, allowedHosts: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(process.cwd(), 'dist')));
    app.get('*', (req, res) => {
      res.sendFile(path.join(process.cwd(), 'dist', 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
  
  // Start the passive Telegram polling thread on server boot
  startTelegramBotRunner();
}

// ── Health Certificate Check Sessions ────────────────────────────────────────
// POST /api/healthcert/check — launch browser, navigate to MOFA checkmedicalresult,
// auto-solve captcha, fill applicationNumber + passportNumber, submit, return result.
// Retries automatically up to 5 times on captcha failures.
app.post('/api/healthcert/check', async (req, res) => {
  const { applicationNumber, passportNumber } = req.body as { applicationNumber: string; passportNumber: string };
  if (!applicationNumber || !passportNumber) {
    return res.status(400).json({ error: 'رقم الطلب ورقم الجواز مطلوبان' });
  }

  const HEALTH_URL      = 'https://visa.mofa.gov.sa/visaperson/checkmedicalresult';
  const MAX_BROWSER_ATTEMPTS = 3;   // reopen browser only on hard connection errors
  const MAX_CAPTCHA_TRIES    = 10;  // retries within one browser session for wrong captcha

  // ── helper: read page signals ──────────────────────────────────────────────
  const readPageSignals = async (page: any): Promise<{ successMsg: string; alertMsg: string; bodyText: string; stillForm: boolean }> =>
    page.evaluate(() => {
      const successEl = document.querySelector('.success-msg, .alert-success') as HTMLElement | null;
      const alertEl   = document.querySelector('.alert-danger, .alert-warning')  as HTMLElement | null;
      return {
        successMsg: (successEl?.innerText || '').replace(/\s+/g, ' ').trim(),
        alertMsg:   (alertEl?.innerText   || '').replace(/\s+/g, ' ').trim(),
        bodyText:   (document.body?.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 3000),
        stillForm:  !!(document.getElementById('Captcha') || document.getElementById('imgCaptcha')),
      };
    });

  // ── helper: classify text into a result message ────────────────────────────
  const classifyResult = (successMsg: string, bodyText: string): string | null => {
    const all = successMsg + ' ' + bodyText;
    const notIssued =
      all.includes('لم يتم') || all.includes('لم تصدر') ||
      all.includes('غير صادرة') || all.includes('لم يتم اصدار') || all.includes('لم يتم إصدار');
    const noData =
      bodyText.includes('لا توجد') || bodyText.includes('غير موجود') ||
      bodyText.includes('no results') || bodyText.includes('not found');
    const issued = !notIssued && (
      all.includes('تم اصدار') || all.includes('تم إصدار') ||
      all.includes('الشهادة صادرة') || all.includes('صادرة وجاهزة') ||
      (successMsg.includes('صادرة') && !successMsg.includes('غير'))
    );
    if (issued)    return 'تم اصدار الشهادة الصحية ✓';
    if (notIssued) return 'لم يتم اصدار الشهادة الصحية بعد';
    if (noData)    return 'لا توجد بيانات لهذا الطلب';
    return null;
  };

  for (let browserAttempt = 0; browserAttempt < MAX_BROWSER_ATTEMPTS; browserAttempt++) {
    let browser: any = null;
    try {
      console.log(`[HEALTHCERT] Browser session ${browserAttempt + 1}/${MAX_BROWSER_ATTEMPTS}`);
      browser = await puppeteer.launch({
        executablePath: CHROMIUM_PATH,
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu', '--no-zygote'],
      });
      const page = await browser.newPage();
      await page.setViewport({ width: 1280, height: 900 });
      await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
      await page.setExtraHTTPHeaders({ 'Accept-Language': 'ar,en-US;q=0.7,en;q=0.3' });

      // ── inner captcha retry loop (reuses the same browser) ─────────────────
      for (let captchaTry = 0; captchaTry < MAX_CAPTCHA_TRIES; captchaTry++) {
        console.log(`[HEALTHCERT] Captcha try ${captchaTry + 1}/${MAX_CAPTCHA_TRIES}`);

        // Navigate fresh to get a clean captcha on every retry
        await page.goto(HEALTH_URL, { waitUntil: 'networkidle2', timeout: 30000 });
        await page.waitForSelector('#imgCaptcha', { timeout: 12000 }).catch(() => {});
        await new Promise(r => setTimeout(r, 1500));

        // Read & solve captcha
        const captchaEl = await page.$('#imgCaptcha').catch(() => null);
        let captchaSolved = '';
        if (captchaEl) {
          const buf = Buffer.from(await captchaEl.screenshot({ type: 'png' }));
          captchaSolved = solveCaptchaWithOcr(buf);
          console.log(`[HEALTHCERT] OCR result: "${captchaSolved}"`);
        }
        if (!captchaSolved) {
          console.log('[HEALTHCERT] Empty OCR — retrying captcha');
          await new Promise(r => setTimeout(r, 500));
          continue;
        }

        // Fill all fields
        const typeInto = async (selector: string, value: string) => {
          try {
            await page.click(selector, { clickCount: 3 });
            await page.type(selector, value, { delay: 40 });
            return true;
          } catch { return false; }
        };
        const appOk     = await typeInto('#AppNo',      applicationNumber);
        const passOk    = await typeInto('#PassportNo',  passportNumber);
        const captchaOk = await typeInto('#Captcha',     captchaSolved);
        console.log(`[HEALTHCERT] Fields: app=${appOk} pass=${passOk} captcha=${captchaOk}`);

        // Submit
        const submitBtn = await page.$('input[type="submit"], button[type="submit"], input[name="Submit"]').catch(() => null);
        await Promise.all([
          page.waitForNavigation({ timeout: 25000, waitUntil: 'networkidle2' }).catch(() => {}),
          submitBtn
            ? submitBtn.click()
            : page.evaluate(() => {
                const btns = Array.from(document.querySelectorAll('input[type="submit"], button[type="submit"], button, input[type="button"]')) as HTMLElement[];
                const b = btns.find(el => (el instanceof HTMLInputElement ? el.value : el.textContent || '').trim().includes('بحث'));
                if (b) b.click();
              }),
        ]);
        await new Promise(r => setTimeout(r, 2500));

        // Read response
        const pd = await readPageSignals(page);
        console.log(`[HEALTHCERT] stillForm=${pd.stillForm} success="${pd.successMsg}" alert="${pd.alertMsg}"`);
        console.log(`[HEALTHCERT] Body[0:200]: ${pd.bodyText.slice(0, 200)}`);

        // Detect captcha-wrong explicitly
        const captchaWrong =
          pd.alertMsg.includes('رمز الصورة') || pd.alertMsg.includes('رمز التحقق') ||
          pd.bodyText.includes('رمز الصورة المدخل غير صحيح') || pd.bodyText.includes('رمز التحقق غير صحيح') ||
          pd.bodyText.includes('الرمز غير صحيح') || pd.bodyText.includes('كود التحقق خاطئ');

        if (captchaWrong) {
          console.log(`[HEALTHCERT] Captcha wrong on try ${captchaTry + 1} — retrying with fresh captcha`);
          await new Promise(r => setTimeout(r, 400));
          continue; // go back to top of captcha loop → page.goto loads fresh captcha
        }

        // Still on form with no message at all → treat as failed captcha too
        if (pd.stillForm && !pd.successMsg && !pd.alertMsg) {
          console.log(`[HEALTHCERT] Still on form, no message — retrying captcha`);
          await new Promise(r => setTimeout(r, 400));
          continue;
        }

        // Try to classify a known result
        const knownResult = classifyResult(pd.successMsg, pd.bodyText);
        if (knownResult) {
          try { browser.close(); } catch {}
          console.log(`[HEALTHCERT] Result: "${knownResult}"`);
          return res.json({ success: true, result: knownResult });
        }

        // Left the form but unknown text — use raw div text as fallback
        const rawResult = (pd.successMsg || pd.alertMsg || '').slice(0, 120);
        if (!pd.stillForm && rawResult) {
          try { browser.close(); } catch {}
          console.log(`[HEALTHCERT] Raw fallback result: "${rawResult}"`);
          return res.json({ success: true, result: rawResult });
        }

        // Unrecognized page state — retry
        console.log(`[HEALTHCERT] Unrecognized page on try ${captchaTry + 1} — retrying`);
        await new Promise(r => setTimeout(r, 500));
      }

      // Exhausted all captcha tries in this browser session
      try { browser.close(); } catch {}
      console.log('[HEALTHCERT] Exhausted captcha tries — reopening browser if possible');

    } catch (err: any) {
      console.error(`[HEALTHCERT] Browser session ${browserAttempt + 1} error: ${err.message}`);
      if (browser) { try { browser.close(); } catch {} }
      if (browserAttempt === MAX_BROWSER_ATTEMPTS - 1) {
        return res.status(502).json({ error: 'فشل الاتصال بمنصة الشهادة الصحية. يرجى المحاولة مرة أخرى.' });
      }
      await new Promise(r => setTimeout(r, 1500));
    }
  }

  return res.status(500).json({ error: 'فشل الحصول على نتيجة بعد عدة محاولات. يرجى المحاولة مرة أخرى.' });
});

startServer();

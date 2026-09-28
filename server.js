const express = require('express');
const cors = require('cors');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Imap = require('imap');
const nodemailer = require('nodemailer');
const WebSocket = require('ws');
const http = require('http');
const admin = require('firebase-admin');
const cron = require('node-cron');
const multer = require('multer');
const cloudinary = require('cloudinary').v2;

// ------------------------------------------------------------
//  CLOUDINARY (opcional)
// ------------------------------------------------------------
const CLOUDINARY_ENABLED =
  !!process.env.CLOUDINARY_CLOUD_NAME &&
  !!process.env.CLOUDINARY_API_KEY &&
  !!process.env.CLOUDINARY_API_SECRET;

if (CLOUDINARY_ENABLED) {
  cloudinary.config({
    cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
    api_key: process.env.CLOUDINARY_API_KEY,
    api_secret: process.env.CLOUDINARY_API_SECRET,
    secure: true,
  });
  console.log('✅ Cloudinary configurado');
} else {
  console.log('⚠️ Cloudinary sin configurar → fallback a Firebase Storage / local');
}

// ------------------------------------------------------------
//  FIREBASE ADMIN (FCM)
// ------------------------------------------------------------
let db = null;
let storageBucket = null;
try {
  let adminVersion = 'desconocida';
  try { adminVersion = require('firebase-admin/package.json').version; } catch (_) {}
  console.log('🔧 firebase-admin versión:', adminVersion);

  const hasApps = admin && Array.isArray(admin.apps) && admin.apps.length > 0;

  if (!hasApps) {
    if (process.env.FIREBASE_PRIVATE_KEY) {
      try {
        let formattedKey = process.env.FIREBASE_PRIVATE_KEY;
        formattedKey = formattedKey.replace(/^"|"$/g, '').replace(/\\n/g, '\n');
        admin.initializeApp({
          credential: admin.credential.cert({
            projectId: process.env.FIREBASE_PROJECT_ID,
            clientEmail: process.env.FIREBASE_CLIENT_EMAIL,
            privateKey: formattedKey,
          }),
          storageBucket: process.env.FIREBASE_STORAGE_BUCKET ||
            `${process.env.FIREBASE_PROJECT_ID}.appspot.com`,
        });
        console.log('✅ Firebase Admin inicializado con variables de entorno');
      } catch (e) {
        console.error('❌ Error al inicializar Firebase con env vars:', e.message);
      }
    } else {
      try {
        const serviceAccount = require('./serviceAccountKey.json');
        admin.initializeApp({
          credential: admin.credential.cert(serviceAccount),
          storageBucket: process.env.FIREBASE_STORAGE_BUCKET ||
            `${serviceAccount.project_id}.appspot.com`,
        });
        console.log('✅ Firebase Admin inicializado con serviceAccountKey.json');
      } catch (e) {
        console.error('⚠️ Sin credenciales de Firebase.');
      }
    }
  }

  if (admin && Array.isArray(admin.apps) && admin.apps.length > 0) {
    db = admin.firestore();
    console.log('✅ Firestore listo');
    try {
      storageBucket = admin.storage().bucket();
    } catch (e) { storageBucket = null; }
  }
} catch (e) {
  console.error('❌ Error crítico inicializando Firebase:', e.message);
  db = null;
}

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const server = http.createServer(app);
const wss = new WebSocket.Server({ server });

// ------------------------------------------------------------
//  Microsoft OAuth2
// ------------------------------------------------------------
const MICROSOFT_CLIENT_ID = process.env.MICROSOFT_CLIENT_ID || '';
const MICROSOFT_SCOPES =
  'https://outlook.office.com/IMAP.AccessAsUser.All https://outlook.office.com/SMTP.Send offline_access openid profile';
const msAccessTokenCache = new Map();

function isMicrosoftOAuthAccount(email, password) {
  if (!password) return false;
  return password.length > 200 && !password.includes(' ') && !password.includes('\n');
}

async function getMicrosoftAccessToken(refreshToken) {
  if (!MICROSOFT_CLIENT_ID) throw new Error('MICROSOFT_CLIENT_ID no configurado');
  const cacheKey = refreshToken.substring(0, 40);
  const cached = msAccessTokenCache.get(cacheKey);
  if (cached && cached.expiresAt > Date.now()) return cached.accessToken;

  const params = new URLSearchParams();
  params.append('client_id', MICROSOFT_CLIENT_ID);
  params.append('grant_type', 'refresh_token');
  params.append('refresh_token', refreshToken);
  params.append('scope', MICROSOFT_SCOPES);

  const res = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Microsoft token error ${res.status}: ${text.substring(0, 200)}`);
  }
  const data = await res.json();
  if (!data.access_token) throw new Error('Microsoft no devolvió access_token');
  const expiresIn = (data.expires_in || 3600) * 1000;
  msAccessTokenCache.set(cacheKey, {
    accessToken: data.access_token,
    expiresAt: Date.now() + expiresIn - 60000,
  });
  return data.access_token;
}

async function resolveAccessToken(email, password) {
  if (!isMicrosoftOAuthAccount(email, password)) return null;
  try { return await getMicrosoftAccessToken(password); }
  catch (e) {
    console.error(`❌ Error obteniendo access_token Microsoft para ${email}:`, e.message);
    return null;
  }
}

// ------------------------------------------------------------
//  WORKERS IMAP + CIRCUIT BREAKER
// ------------------------------------------------------------
const activeWorkers = new Map();
const failedAccounts = new Map();
const MAX_FAILURES = 5;
const FAILURE_COOLDOWN_MS = 60 * 60 * 1000;

function recordWorkerFailure(email) {
  const entry = failedAccounts.get(email) || { count: 0, until: 0 };
  entry.count++;
  if (entry.count >= MAX_FAILURES) {
    entry.until = Date.now() + FAILURE_COOLDOWN_MS;
    console.log(`🚫 Cuenta ${email} bloqueada tras ${MAX_FAILURES} intentos. Cooldown 1h.`);
  }
  failedAccounts.set(email, entry);
}
function resetWorkerFailure(email) { failedAccounts.delete(email); }
function isAccountBlocked(email) {
  const entry = failedAccounts.get(email);
  if (!entry) return false;
  if (entry.until && Date.now() < entry.until) return true;
  if (entry.until && Date.now() >= entry.until) {
    failedAccounts.delete(email);
    return false;
  }
  return false;
}

// ------------------------------------------------------------
//  AUTO-CONFIG UNIVERSAL
// ------------------------------------------------------------
function getAutoConfig(email) {
  if (!email) return null;
  const domain = email.split('@')[1]?.toLowerCase();
  if (!domain) return null;

  if (domain === 'gmail.com' || domain === 'googlemail.com') {
    return { imapHost: 'imap.gmail.com', imapPort: 993, smtpHost: 'smtp.gmail.com', smtpPort: 465, smtpSecure: true, provider: 'gmail' };
  }
  if (['outlook.com', 'hotmail.com', 'live.com', 'msn.com', 'outlook.es'].some(d => domain === d)) {
    return { imapHost: 'outlook.office365.com', imapPort: 993, smtpHost: 'smtp.office365.com', smtpPort: 587, smtpSecure: false, provider: 'outlook' };
  }
  if (['office365.com', 'office.com'].some(d => domain.endsWith(d))) {
    return { imapHost: 'outlook.office365.com', imapPort: 993, smtpHost: 'smtp.office365.com', smtpPort: 587, smtpSecure: false, provider: 'office365' };
  }
  if (domain.endsWith('.onmicrosoft.com')) {
    return { imapHost: 'outlook.office365.com', imapPort: 993, smtpHost: 'smtp.office365.com', smtpPort: 587, smtpSecure: false, provider: 'office365' };
  }
  if (domain.includes('yahoo.')) {
    return { imapHost: 'imap.mail.yahoo.com', imapPort: 993, smtpHost: 'smtp.mail.yahoo.com', smtpPort: 465, smtpSecure: true, provider: 'yahoo' };
  }
  if (domain.includes('zoho.')) {
    return { imapHost: 'imap.zoho.com', imapPort: 993, smtpHost: 'smtp.zoho.com', smtpPort: 465, smtpSecure: true, provider: 'zoho' };
  }
  if (domain === 'icloud.com' || domain === 'me.com' || domain === 'mac.com') {
    return { imapHost: 'imap.mail.me.com', imapPort: 993, smtpHost: 'smtp.mail.me.com', smtpPort: 587, smtpSecure: false, provider: 'icloud' };
  }
  if (domain.includes('gmx.')) {
    return { imapHost: 'imap.gmx.com', imapPort: 993, smtpHost: 'mail.gmx.com', smtpPort: 587, smtpSecure: false, provider: 'gmx' };
  }
  if (domain.includes('ionos.') || domain.includes('1and1.')) {
    return { imapHost: 'imap.ionos.es', imapPort: 993, smtpHost: 'smtp.ionos.es', smtpPort: 587, smtpSecure: false, provider: 'ionos' };
  }
  return { imapHost: 'mail.' + domain, imapPort: 993, smtpHost: 'mail.' + domain, smtpPort: 465, smtpSecure: true, provider: 'custom', domain };
}

function getSmtpCandidates(email) {
  const auto = getAutoConfig(email);
  if (!auto) return [];
  const candidates = [], seen = new Set();
  const add = (host, port, secure) => {
    const key = `${host}:${port}`;
    if (seen.has(key)) return;
    seen.add(key);
    candidates.push({ host, port, secure });
  };
  add(auto.smtpHost, auto.smtpPort, auto.smtpSecure);
  add(auto.smtpHost, 465, true);
  add(auto.smtpHost, 587, false);
  add(auto.smtpHost, 25, false);
  if (auto.provider === 'custom' && auto.domain) {
    const d = auto.domain;
    add('smtp.' + d, 465, true); add('smtp.' + d, 587, false);
    add('mail.' + d, 465, true); add('mail.' + d, 587, false);
    add('mail.' + d, 25, false);
    add(d, 465, true); add(d, 587, false);
  }
  return candidates;
}

function getImapCandidates(email) {
  const auto = getAutoConfig(email);
  if (!auto) return [];
  const candidates = [], seen = new Set();
  const add = (host, port, secure) => {
    const key = `${host}:${port}`;
    if (seen.has(key)) return;
    seen.add(key);
    candidates.push({ host, port, secure });
  };
  add(auto.imapHost, auto.imapPort, true);
  add(auto.imapHost, 143, false);
  if (auto.provider === 'custom' && auto.domain) {
    const d = auto.domain;
    add('imap.' + d, 993, true); add('imap.' + d, 143, false);
    add('mail.' + d, 993, true); add('mail.' + d, 143, false);
  }
  return candidates;
}

// ------------------------------------------------------------
//  PERSISTENCIA DE CUENTAS
// ------------------------------------------------------------
async function saveAccount(email, password, imapHost) {
  if (!db) return;
  try {
    await db.collection('user_accounts').doc(email).set({
      email, password, imapHost,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    }, { merge: true });
  } catch (e) { console.error(`⚠️ Error guardando cuenta ${email}:`, e.message); }
}

async function restoreWorkers() {
  if (!db) return;
  try {
    const snapshot = await db.collection('user_accounts').get();
    console.log(`🔄 Restaurando ${snapshot.size} worker(s) desde Firestore...`);
    for (const doc of snapshot.docs) {
      const data = doc.data();
      if (data.email && data.password) {
        startImapWorker(data.email, data.password, data.imapHost);
      }
    }
  } catch (e) { console.error('⚠️ Error restaurando workers:', e.message); }
}

async function getSavedLastUid(email) {
  if (!db) return null;
  try {
    const doc = await db.collection('user_states').doc(email).get();
    if (doc.exists) return doc.data().lastUid || null;
  } catch (_) {}
  return null;
}

async function saveLastUid(email, uid) {
  if (!db) return;
  try {
    await db.collection('user_states').doc(email).set({
      lastUid: uid,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    }, { merge: true });
  } catch (_) {}
}

async function connectImap(email, password, host, port, secure, accessToken = null) {
  const auth = accessToken
    ? { user: email, accessToken }
    : { user: email, pass: password };

  const config = {
    host, port, secure, auth,
    logger: false,
    tls: { rejectUnauthorized: false, minVersion: 'TLSv1.2', servername: host },
    connectionTimeout: 30000,
    greetingTimeout: 20000,
    socketTimeout: 35000,
  };
  const client = new ImapFlow(config);
  client.on('error', (err) => {
    console.error(`⚠️ Error ImapFlow (${email}):`, err.message);
  });
  await client.connect();
  return client;
}

async function connectImapAuto(email, password, preferredHost = null) {
  const accessToken = await resolveAccessToken(email, password);
  const candidates = getImapCandidates(email);

  if (preferredHost) {
    candidates.unshift({ host: preferredHost, port: 993, secure: true });
    candidates.unshift({ host: preferredHost, port: 143, secure: false });
  }

  let lastError = null;
  for (const c of candidates) {
    try {
      const client = await connectImap(email, password, c.host, c.port, c.secure, accessToken);
      return { client, host: c.host, port: c.port };
    } catch (e) { lastError = e; }
  }
  throw lastError || new Error('Todos los intentos IMAP fallaron');
}

async function createSmtpTransporter({ email, password, host, port, secure }) {
  const accessToken = await resolveAccessToken(email, password);
  const auth = accessToken
    ? { type: 'OAuth2', user: email, accessToken }
    : { user: email, pass: password };

  return nodemailer.createTransport({
    host, port, secure, auth,
    tls: { rejectUnauthorized: false },
    connectionTimeout: 10000,
    greetingTimeout: 10000,
    socketTimeout: 15000,
  });
}

// ------------------------------------------------------------
//  BREVO
// ------------------------------------------------------------
async function sendViaBrevo({ fromEmail, fromName, to, subject, html }) {
  const brevoKey = process.env.BREVO_API_KEY;
  if (!brevoKey) throw new Error('BREVO_API_KEY no configurada');
  const payload = {
    sender: { name: fromName || fromEmail.split('@')[0], email: fromEmail },
    to: [{ email: to }],
    subject,
    htmlContent: html,
  };
  const res = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: { 'accept': 'application/json', 'api-key': brevoKey, 'content-type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!res.ok) { const body = await res.text(); throw new Error(`Brevo ${res.status}: ${body}`); }
  return true;
}

// ------------------------------------------------------------
//  HELPERS DE TIMEZONE
// ------------------------------------------------------------
async function getUserTimezoneOffset(email) {
  if (!db || !email) return 0;
  try {
    const snap = await db.collection('fcm_tokens').where('email', '==', email).limit(1).get();
    if (snap.empty) return 0;
    const off = snap.docs[0].data().timezoneOffset;
    if (typeof off === 'number' && !isNaN(off)) return off;
    return 0;
  } catch (_) { return 0; }
}

function formatEventLocal(eventDate, offsetMinutes, referenceNow = new Date()) {
  const off = Number(offsetMinutes) || 0;
  const evLocal = new Date(eventDate.getTime() + off * 60 * 1000);
  const nowLocal = new Date(referenceNow.getTime() + off * 60 * 1000);

  const evY = evLocal.getUTCFullYear();
  const evM = evLocal.getUTCMonth();
  const evD = evLocal.getUTCDate();
  const evH = evLocal.getUTCHours();
  const evMin = evLocal.getUTCMinutes();

  const todayUtc = Date.UTC(
    nowLocal.getUTCFullYear(),
    nowLocal.getUTCMonth(),
    nowLocal.getUTCDate()
  );
  const evDayUtc = Date.UTC(evY, evM, evD);
  const diffDays = Math.round((evDayUtc - todayUtc) / (1000 * 60 * 60 * 24));

  const timeStr = `${String(evH).padStart(2, '0')}:${String(evMin).padStart(2, '0')}`;
  const dd = String(evD).padStart(2, '0');
  const mo = String(evM + 1).padStart(2, '0');

  let dayStr;
  if (diffDays === 0) dayStr = 'Hoy';
  else if (diffDays === 1) dayStr = 'Mañana';
  else if (diffDays === -1) dayStr = 'Ayer';
  else if (diffDays > 1 && diffDays < 7) {
    const days = ['dom', 'lun', 'mar', 'mié', 'jue', 'vie', 'sáb'];
    const idx = new Date(evDayUtc).getUTCDay();
    dayStr = days[idx];
  } else {
    dayStr = `${dd}/${mo}`;
  }

  const fullStr = `${dayStr} ${dd}/${mo} a las ${timeStr}`;
  return { dayStr, timeStr, fullStr };
}

function formatNowLocal(offsetMinutes, referenceNow = new Date()) {
  const off = Number(offsetMinutes) || 0;
  const local = new Date(referenceNow.getTime() + off * 60 * 1000);
  const h = String(local.getUTCHours()).padStart(2, '0');
  const m = String(local.getUTCMinutes()).padStart(2, '0');
  const d = String(local.getUTCDate()).padStart(2, '0');
  const mo = String(local.getUTCMonth() + 1).padStart(2, '0');
  return `${d}/${mo} ${h}:${m}`;
}

// ------------------------------------------------------------
//  AUSENCIAS
// ------------------------------------------------------------
async function getActiveAbsencePeriod(accountEmail) {
  if (!db) return null;
  try {
    const doc = await db.collection('absences_configs').doc(accountEmail).get();
    if (!doc.exists) return null;
    const data = doc.data() || {};
    const periods = Array.isArray(data.periods) ? data.periods : [];
    const now = Date.now();
    for (const p of periods) {
      if (!p.enabled) continue;
      const start = p.startDate ? Date.parse(p.startDate) : null;
      const end = p.endDate ? Date.parse(p.endDate) : null;
      if (start && now < start) continue;
      if (end && now > end) continue;
      return p;
    }
    return null;
  } catch (e) { return null; }
}

async function checkAndSendAutoReply({
  accountEmail, accountPassword, incomingFrom, incomingSubject, incomingUid,
}) {
  if (!db) return;
  try {
    const period = await getActiveAbsencePeriod(accountEmail);
    if (!period) return;
    const fromLower = (incomingFrom || '').toLowerCase();
    if (!fromLower.includes('@')) return;
    if (fromLower === accountEmail.toLowerCase()) return;
    const ignorePatterns = ['noreply@', 'no-reply@', 'no_reply@', 'mailer-daemon@', 'postmaster@', 'notifications@', 'notification@', 'bounce@', 'bounces@'];
    if (ignorePatterns.some((p) => fromLower.includes(p))) return;
    if (period.onlyContacts) {
      try {
        const cs = await db.collection('users').doc(accountEmail).collection('contacts').where('email', '==', fromLower).limit(1).get();
        if (cs.empty) return;
      } catch (_) {}
    }

    const replyId = `${accountEmail}__${period.id}__${fromLower}`.replace(/\//g, '_');
    const sentRef = db.collection('vacation_sent_replies').doc(replyId);
    const sentDoc = await sentRef.get();
    const intervalMs = (period.replyIntervalDays || 4) * 24 * 60 * 60 * 1000;
    if (sentDoc.exists) {
      const prevTs = sentDoc.data().sentAt?.toDate?.()?.getTime?.() || 0;
      if (prevTs && Date.now() - prevTs < intervalMs) return;
    }

    const escapedBody = (period.body || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
    const escapedTitle = (period.title || 'Ausencia').replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const html = `
<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#333;line-height:1.6;">
  <div style="background:#E3F2FD;border-left:4px solid #1A73E8;padding:10px 14px;margin-bottom:16px;border-radius:6px;">
    <strong style="color:#0D47A1;">📤 Respuesta automática — ${escapedTitle}</strong>
  </div>
  <div>${escapedBody}</div>
  <hr style="border:none;border-top:1px solid #ddd;margin:20px 0;">
  <div style="font-size:11px;color:#888;font-style:italic;">Este es un mensaje automático.</div>
</div>`;

    const replySubject = `Re: ${incomingSubject || '(Sin asunto)'}`;
    const autoReplySubject = period.subject ? `${period.subject} — ${replySubject}` : replySubject;

    let sent = false;
    const smtpCandidates = getSmtpCandidates(accountEmail);
    for (const c of smtpCandidates) {
      try {
        const transporter = await createSmtpTransporter({
          email: accountEmail, password: accountPassword,
          host: c.host, port: c.port, secure: c.secure,
        });
        await transporter.sendMail({
          from: accountEmail, to: incomingFrom,
          subject: autoReplySubject, html,
          headers: { 'Auto-Submitted': 'auto-replied', 'X-Auto-Response-Suppress': 'All' },
        });
        sent = true;
        break;
      } catch (_) {}
    }
    if (!sent && process.env.BREVO_API_KEY) {
      try {
        await sendViaBrevo({ fromEmail: accountEmail, fromName: '', to: incomingFrom, subject: autoReplySubject, html });
        sent = true;
      } catch (_) {}
    }
    if (sent) {
      await sentRef.set({
        accountEmail, periodId: period.id, periodTitle: period.title,
        senderEmail: fromLower, originalSubject: incomingSubject,
        originalUid: incomingUid,
        sentAt: admin.firestore.FieldValue.serverTimestamp(),
      });
    }
  } catch (e) { console.error('❌ Error en checkAndSendAutoReply:', e.message); }
}

// ------------------------------------------------------------
//  MOTOR DE REGLAS / FILTROS
// ------------------------------------------------------------
const _rulesCache = new Map();
const RULES_CACHE_TTL_MS = 5 * 60 * 1000;

// ------------------------------------------------------------
//  🔥 IA — Preferencias del usuario
// ------------------------------------------------------------
const _aiPrefsCache = new Map();
const AI_PREFS_CACHE_TTL_MS = 5 * 60 * 1000;

const AI_PREFS_DEFAULTS = {
  chat: true,
  naturalLanguageRules: false,
  detectUrgency: false,
  urgencyLevel: 'medium',
  autoClassify: false,
  categories: ['personal', 'trabajo', 'facturas', 'publicidad',
               'notificaciones', 'social', 'otro'],
};

async function getAiPreferences(email) {
  if (!db || !email) return { ...AI_PREFS_DEFAULTS };
  const now = Date.now();
  const cached = _aiPrefsCache.get(email);
  if (cached && cached.expiresAt > now) return cached.prefs;
  try {
    const doc = await db.collection('ai_preferences').doc(email).get();
    const data = doc.exists ? (doc.data() || {}) : {};
    const prefs = {
      chat: typeof data.chat === 'boolean' ? data.chat : AI_PREFS_DEFAULTS.chat,
      naturalLanguageRules: typeof data.naturalLanguageRules === 'boolean'
        ? data.naturalLanguageRules : AI_PREFS_DEFAULTS.naturalLanguageRules,
      detectUrgency: typeof data.detectUrgency === 'boolean'
        ? data.detectUrgency : AI_PREFS_DEFAULTS.detectUrgency,
      urgencyLevel: ['low', 'medium', 'high'].includes(data.urgencyLevel)
        ? data.urgencyLevel : AI_PREFS_DEFAULTS.urgencyLevel,
      autoClassify: typeof data.autoClassify === 'boolean'
        ? data.autoClassify : AI_PREFS_DEFAULTS.autoClassify,
      categories: Array.isArray(data.categories) && data.categories.length > 0
        ? data.categories.map((c) => String(c).toLowerCase())
        : AI_PREFS_DEFAULTS.categories,
    };
    _aiPrefsCache.set(email, { prefs, expiresAt: now + AI_PREFS_CACHE_TTL_MS });
    return prefs;
  } catch (e) {
    return { ...AI_PREFS_DEFAULTS };
  }
}

function invalidateAiPreferencesCache(email) {
  if (email) _aiPrefsCache.delete(email);
}

async function getRulesForAccount(email) {
  if (!db) return [];
  const now = Date.now();
  const cached = _rulesCache.get(email);
  if (cached && cached.expiresAt > now) return cached.rules;
  try {
    const doc = await db.collection('rules_configs').doc(email).get();
    const rules = doc.exists ? (doc.data()?.rules || []).filter((r) => r && r.enabled) : [];
    _rulesCache.set(email, { rules, expiresAt: now + RULES_CACHE_TTL_MS });
    return rules;
  } catch (e) { return []; }
}

function rulesNeedBody(rules) {
  return rules.some((r) => (r.conditions || []).some((c) => c.field === 'body'));
}

function evalCondition(c, ctx) {
  const field = c.field, op = c.operator, raw = c.value ?? '';
  let target = '';
  if (field === 'from') target = ctx.from || '';
  else if (field === 'to') target = ctx.to || '';
  else if (field === 'cc') target = ctx.cc || '';
  else if (field === 'subject') target = ctx.subject || '';
  else if (field === 'body') target = ctx.body || '';
  else if (field === 'hasAttachment') {
    if (op === 'isTrue') return !!ctx.hasAttachment;
    if (op === 'isFalse') return !ctx.hasAttachment;
    return false;
  } else if (field === 'sizeKb') {
    const n = Number(raw) || 0;
    const sz = Number(ctx.sizeKb) || 0;
    if (op === 'greaterThan') return sz > n;
    if (op === 'lessThan') return sz < n;
    return false;
  }
  const t = target.toLowerCase(), v = raw.toLowerCase();
  switch (op) {
    case 'contains': return t.includes(v);
    case 'notContains': return !t.includes(v);
    case 'equals': return t === v;
    case 'notEquals': return t !== v;
    case 'startsWith': return t.startsWith(v);
    case 'endsWith': return t.endsWith(v);
    case 'regex':
      try { return new RegExp(raw, 'i').test(target); } catch (_) { return false; }
    default: return false;
  }
}

function ruleMatches(rule, ctx) {
  const conds = rule.conditions || [];
  if (conds.length === 0) return false;
  const results = conds.map((c) => evalCondition(c, ctx));
  return rule.matchAll ? results.every(Boolean) : results.some(Boolean);
}

async function executeRuleActions({ accountEmail, accountPassword, accountHost, folder, uid, actions }) {
  let client;
  const forwards = [];
  try {
    const conn = await connectImapAuto(accountEmail, accountPassword, accountHost);
    client = conn.client;
    const lock = await client.getMailboxLock(folder);
    try {
      for (const a of actions) {
        const t = a.type;
        if (t === 'markRead') await client.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true }).catch(() => {});
        else if (t === 'markUnread') await client.messageFlagsRemove(String(uid), ['\\Seen'], { uid: true }).catch(() => {});
        else if (t === 'star') await client.messageFlagsAdd(String(uid), ['\\Flagged'], { uid: true }).catch(() => {});
        else if (t === 'unstar') await client.messageFlagsRemove(String(uid), ['\\Flagged'], { uid: true }).catch(() => {});
      }
      for (const a of actions) {
        const t = a.type, v = a.value || '';
        if (t === 'moveTo') {
          if (!v) continue;
          const list = await client.list();
          const found = list.find(f => f.path.toLowerCase() === v.toLowerCase() || f.name.toLowerCase() === v.toLowerCase());
          const target = found ? found.path : v;
          try { await client.messageMove(String(uid), target, { uid: true }); } catch (_) {}
          try { lock.release(); } catch (_) {}
          await client.logout(); client = null;
          return true;
        }
        if (t === 'deleteMessage') {
          const list = await client.list();
          const trash = list.find(f => f.specialUse === '\\Trash' || /papelera/i.test(f.name) || /^trash$/i.test(f.name))?.path || 'Trash';
          try { await client.messageMove(String(uid), trash, { uid: true }); }
          catch (_) { await client.messageDelete(String(uid), { uid: true }).catch(() => {}); }
          try { lock.release(); } catch (_) {}
          await client.logout(); client = null;
          return true;
        }
        if (t === 'markSpam') {
          const list = await client.list();
          const junk = list.find(f => f.specialUse === '\\Junk' || /junk/i.test(f.name) || /spam/i.test(f.name))?.path;
          if (junk) {
            try {
              await client.messageMove(String(uid), junk, { uid: true });
              try { lock.release(); } catch (_) {}
              await client.logout(); client = null;
              return true;
            } catch (_) {}
          }
        }
        if (t === 'forward' && v) forwards.push(v);
      }
    } finally { try { lock.release(); } catch (_) {} }
    await client.logout(); client = null;
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    return false;
  }
  for (const to of forwards) {
    try {
      await sendViaBrevo({ fromEmail: accountEmail, fromName: '', to, subject: '[Reenviado por regla]', html: '<p>Este mensaje ha sido reenviado por una regla de RSMail.</p>' });
    } catch (_) {}
  }
  return true;
}

async function applyRulesToMessage({ accountEmail, accountPassword, accountHost, uid, folder, parsed, envelope, size }) {
  if (!db) return;
  const rules = await getRulesForAccount(accountEmail);
  if (rules.length === 0) return;

  const fromAddr = envelope?.from?.[0]?.address || '';
  const fromName = envelope?.from?.[0]?.name || '';
  const toAddr = (envelope?.to || []).map((t) => t.address).join(', ');
  const ccAddr = (envelope?.cc || []).map((t) => t.address).join(', ');
  const subject = envelope?.subject || '';
  const hasAttachment = (parsed?.attachments || []).length > 0;
  const sizeKb = Math.round((size || 0) / 1024);

  const ctx = { from: `${fromName} <${fromAddr}>`, to: toAddr, cc: ccAddr, subject, body: parsed?.text || '', hasAttachment, sizeKb };

  for (const rule of rules) {
    if (!ruleMatches(rule, ctx)) continue;
    console.log(`✅ Regla "${rule.name}" coincide (UID ${uid})`);
    await executeRuleActions({ accountEmail, accountPassword, accountHost, folder, uid, actions: rule.actions || [] });
    if (rule.stopProcessing) break;
  }
}

// ------------------------------------------------------------
//  🔥 IA — Detección de urgencia (FASE 2)
// ------------------------------------------------------------

/** Umbrales según nivel de sensibilidad configurado por el usuario. */
function urgencyThresholdForLevel(level) {
  switch (level) {
    case 'low': return 85;
    case 'high': return 55;
    case 'medium':
    default: return 70;
  }
}

/**
 * Llama a Groq para saber si un correo es urgente.
 * Devuelve { score: 0-100, reason: '...' } o lanza.
 */
async function detectUrgencyWithGroq({ from, subject, preview, level }) {
  const groqKey = process.env.GROQ_API_KEY;
  if (!groqKey) throw new Error('GROQ_API_KEY no configurada');

  const prompt = `Evalúa si este correo es URGENTE para el destinatario.

Nivel de sensibilidad del usuario: ${level || 'medium'}
- low: solo correos muy obvios (plazos vencidos, emergencias reales)
- medium: correos importantes que requieren acción hoy
- high: también correos que requieren acción esta semana

Remitente: ${from || '(desconocido)'}
Asunto: ${subject || '(sin asunto)'}
Vista previa: ${(preview || '').substring(0, 300)}

Devuelve SOLO un JSON con esta forma exacta:
{ "score": <número entero 0-100>, "reason": "<frase muy corta en español, máx 60 caracteres>" }

Interpretación del score:
- 85-100: urgencia extrema (plazo hoy, emergencia, cancelación importante)
- 70-84: urgente (requiere acción hoy o mañana)
- 50-69: importante pero no urgente
- 0-49: no urgente (newsletters, notificaciones rutinarias, publicidad)

Analiza con criterio:
- Los correos de personas reales > correos automáticos
- Palabras como "urgente", "hoy", "mañana", "plazo", "vence", "importante" suben el score
- Palabras como "newsletter", "promoción", "no-reply", "suscripción" bajan el score
- Correos de remitentes no-reply o marketing raramente son urgentes
- Correos de jefes, clientes o con asunto personal pueden ser urgentes`;

  const messages = [
    { role: 'system', content: 'Eres un clasificador de urgencia de correos. Devuelves SOLO JSON válido, sin markdown ni texto extra.' },
    { role: 'user', content: prompt },
  ];

  let lastError = null;
  for (const model of GROQ_MODELS_FALLBACK) {
    try {
      const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${groqKey}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          model,
          messages,
          temperature: 0.1,
          max_tokens: 200,
          response_format: { type: 'json_object' },
        }),
      });

      if (!res.ok) {
        lastError = new Error(`Groq ${res.status} (${model})`);
        if (res.status === 404 || res.status === 400) continue;
        if (res.status === 429) throw lastError;
        continue;
      }

      const data = await res.json();
      const raw = data?.choices?.[0]?.message?.content || '{}';
      const parsed = JSON.parse(raw);
      const score = Math.max(0, Math.min(100, parseInt(parsed.score) || 0));
      const reason = String(parsed.reason || '').substring(0, 100);
      return { score, reason, model };
    } catch (e) {
      lastError = e;
      continue;
    }
  }
  throw lastError || new Error('Todos los modelos fallaron');
}

/**
 * Evalúa urgencia, guarda en Firestore y (si supera umbral) marca con \Flagged
 * y envía push prioritario.
 */
async function checkUrgencyAndNotify({
  email, password, host, uid, from, subject, preview, folder = 'INBOX',
}) {
  if (!db) return;

  let prefs;
  try { prefs = await getAiPreferences(email); }
  catch (_) { return; }
  if (!prefs.detectUrgency) return;

  let result;
  try {
    result = await detectUrgencyWithGroq({
      from, subject, preview,
      level: prefs.urgencyLevel,
    });
  } catch (e) {
    console.log(`⚠️ [AI-Urgency] Fallo detectando: ${e.message}`);
    return;
  }

  const threshold = urgencyThresholdForLevel(prefs.urgencyLevel);
  const isUrgent = result.score >= threshold;

  console.log(`🔥 [AI-Urgency] UID ${uid}: score=${result.score} (${prefs.urgencyLevel}, umbral ${threshold}) → ${isUrgent ? 'URGENTE' : 'normal'}`);

  // Guardar en Firestore (incluso si no es urgente, guardamos score bajo para debug)
  try {
    await db
      .collection('email_urgency')
      .doc(email)
      .collection('messages')
      .doc(String(uid))
      .set({
        score: result.score,
        reason: result.reason,
        level: prefs.urgencyLevel,
        isUrgent,
        folder,
        subject: subject || '',
        from: from || '',
        createdAt: admin.firestore.FieldValue.serverTimestamp(),
      });
  } catch (e) {
    console.log(`⚠️ No se pudo guardar urgencia en Firestore: ${e.message}`);
  }

  if (!isUrgent) return;

  // Marcar \Flagged
  try {
    const conn = await connectImapAuto(email, password, host);
    const client = conn.client;
    const lock = await client.getMailboxLock(folder);
    try {
      await client.messageFlagsAdd(String(uid), ['\\Flagged'], { uid: true });
    } finally { lock.release(); }
    await client.logout();
    console.log(`🔥 UID ${uid} marcado como \\Flagged`);
  } catch (e) {
    console.log(`⚠️ No se pudo marcar \\Flagged: ${e.message}`);
  }

  // Push prioritario (título distinto)
  try {
    await sendPushNotification(email, {
      title: `🔥 URGENTE · ${subject || '(Sin asunto)'}`,
      body: result.reason || `Correo de ${from}`,
      data: {
        type: 'new_email_urgent',
        sender: from || '',
        subject: subject || '',
        uid: String(uid),
        folder,
        urgentScore: String(result.score),
        urgentReason: result.reason || '',
        text: result.reason || subject || '',
      },
    }, { dataOnly: true });
    console.log(`📤 Push urgente enviado para UID ${uid}`);
  } catch (e) {
    console.log(`⚠️ No se pudo enviar push urgente: ${e.message}`);
  }
}

app.post('/api/ai/detect-urgency', async (req, res) => {
  if (!groqEnabled) {
    return res.status(503).json({ success: false, error: 'ai_unavailable' });
  }
  try {
    const { email, from, subject, preview, level } = req.body || {};
    if (!email) return res.status(400).json({ success: false, error: 'email_required' });

    const prefs = await getAiPreferences(email);
    if (!prefs.detectUrgency && !req.body.force) {
      return res.status(403).json({
        success: false,
        error: 'urgency_disabled',
        message: 'Activa "Detección de urgencia" en Mi Perfil → Asistente IA.',
      });
    }

    const result = await detectUrgencyWithGroq({
      from: from || '',
      subject: subject || '',
      preview: preview || '',
      level: level || prefs.urgencyLevel,
    });
    const threshold = urgencyThresholdForLevel(level || prefs.urgencyLevel);
    res.json({
      success: true,
      score: result.score,
      reason: result.reason,
      isUrgent: result.score >= threshold,
      threshold,
      level: level || prefs.urgencyLevel,
    });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ------------------------------------------------------------
//  TRADUCCIÓN AUTOMÁTICA
// ------------------------------------------------------------
const _translateCache = new Map();
const TRANSLATE_CACHE_TTL_MS = 6 * 60 * 60 * 1000;
const GOOGLE_HARD_LIMIT = 4500;

function _simpleHash(s) {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = ((h << 5) - h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(16);
}

async function googleTranslate(text, target, source = 'auto', format = 'text') {
  const url = new URL('https://translate.googleapis.com/translate_a/single');
  url.searchParams.set('client', 'gtx');
  url.searchParams.set('sl', source);
  url.searchParams.set('tl', target);
  url.searchParams.set('dt', 't');
  url.searchParams.set('q', text);
  if (format === 'html') url.searchParams.set('format', 'html');
  const res = await fetch(url.toString(), { headers: { 'User-Agent': 'RSMail/3.0' } });
  if (!res.ok) throw new Error(`google ${res.status}`);
  const data = await res.json();
  const translated = Array.isArray(data?.[0]) ? data[0].map((c) => c?.[0] || '').join('') : '';
  return { translated, detectedSource: data?.[2] || source };
}

async function libreTranslate(text, target, source = 'auto', format = 'text') {
  const res = await fetch('https://libretranslate.com/translate', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'User-Agent': 'RSMail/3.0' },
    body: JSON.stringify({ q: text, source: source === 'auto' ? 'auto' : source, target, format }),
  });
  if (!res.ok) throw new Error(`libre ${res.status}`);
  const data = await res.json();
  return { translated: data?.translatedText || text, detectedSource: data?.detectedLanguage?.language || source };
}

function _chunkText(text, maxLen = GOOGLE_HARD_LIMIT) {
  if (text.length <= maxLen) return [text];
  const chunks = []; let i = 0;
  while (i < text.length) {
    if (text.length - i <= maxLen) { chunks.push(text.substring(i)); break; }
    let cutAt = i + maxLen;
    const lowerBound = i + Math.floor(maxLen * 0.6);
    for (let k = cutAt; k > lowerBound; k--) {
      const c = text[k];
      if (c === ' ' || c === '\n' || c === '>' || c === '.') { cutAt = k; break; }
    }
    if (cutAt <= i) cutAt = i + maxLen;
    chunks.push(text.substring(i, cutAt));
    i = cutAt;
  }
  return chunks;
}

async function translateInChunks(text, target, source, format) {
  if (text.length <= GOOGLE_HARD_LIMIT) {
    let r;
    try { r = await googleTranslate(text, target, source, format); }
    catch (e1) { r = await libreTranslate(text, target, source, format); }
    return { translated: r.translated, detectedSource: r.detectedSource, chunked: false };
  }
  const chunks = _chunkText(text);
  const parts = []; let detected = source;
  for (const chunk of chunks) {
    let r;
    try { r = await googleTranslate(chunk, target, source, format); }
    catch (_) {
      try { r = await libreTranslate(chunk, target, source, format); }
      catch (_) { parts.push(chunk); continue; }
    }
    if (r.detectedSource && r.detectedSource !== 'auto') detected = r.detectedSource;
    parts.push(r.translated);
  }
  return { translated: parts.join(' '), detectedSource: detected, chunked: true };
}

app.post('/api/translate', async (req, res) => {
  try {
    const { text, target = 'es', source = 'auto', format = 'text' } = req.body || {};
    if (!text || typeof text !== 'string') return res.status(400).json({ error: 'text_required' });
    if (text.length > 200000) return res.status(413).json({ error: 'text_too_long' });

    const cacheKey = `${target}_${source}_${format}_${_simpleHash(text)}`;
    const now = Date.now();
    const cached = _translateCache.get(cacheKey);
    if (cached && now - cached.ts < TRANSLATE_CACHE_TTL_MS) {
      return res.json({ translated: cached.translated, detectedSource: cached.detectedSource, cached: true });
    }
    if (db) {
      try {
        const doc = await db.collection('translation_cache').doc(cacheKey).get();
        if (doc.exists && doc.data()?.translated) {
          const d = doc.data();
          _translateCache.set(cacheKey, { translated: d.translated, detectedSource: d.detectedSource || source, ts: now });
          return res.json({ translated: d.translated, detectedSource: d.detectedSource || source, cached: true });
        }
      } catch (_) {}
    }
    let result;
    try { result = await translateInChunks(text, target, source, format); }
    catch (e) { return res.status(502).json({ error: 'translate_failed', message: e.message }); }

    _translateCache.set(cacheKey, { translated: result.translated, detectedSource: result.detectedSource, ts: now });
    if (db) {
      db.collection('translation_cache').doc(cacheKey).set({
        translated: result.translated, detectedSource: result.detectedSource,
        target, source, format,
        createdAt: admin.firestore.FieldValue.serverTimestamp(),
      }).catch(() => {});
    }
    res.json({ translated: result.translated, detectedSource: result.detectedSource, cached: false, chunked: !!result.chunked });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

cron.schedule('0 */12 * * *', async () => {
  if (!db) return;
  try {
    const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 60);
    const snap = await db.collection('translation_cache').where('createdAt', '<', admin.firestore.Timestamp.fromDate(cutoff)).limit(500).get();
    for (const doc of snap.docs) await doc.ref.delete();
  } catch (_) {}
});

// ------------------------------------------------------------
//  CHAT — SUBIDA DE ARCHIVOS
// ------------------------------------------------------------
const multerMemory = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 },
});

app.post('/api/chat/upload', multerMemory.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, error: 'No file' });
    const originalName = req.file.originalname || `file_${Date.now()}`;
    const size = req.file.size || 0;
    const mimetype = req.file.mimetype || 'application/octet-stream';

    if (CLOUDINARY_ENABLED) {
      try {
        const isImage = mimetype.startsWith('image/');
        const isVideo = mimetype.startsWith('video/');
        const uploadResult = await new Promise((resolve, reject) => {
          const stream = cloudinary.uploader.upload_stream(
            {
              folder: 'rsmail_chat',
              resource_type: isImage ? 'image' : isVideo ? 'video' : 'raw',
              public_id: `${Date.now()}_${originalName.replace(/\.[^/.]+$/, '').replace(/[^\w\-]/g, '_')}`,
              unique_filename: true,
              overwrite: false,
              type: 'upload',
              access_mode: 'public',
            },
            (error, result) => { if (error) reject(error); else resolve(result); }
          );
          stream.end(req.file.buffer);
        });
        return res.json({
          success: true,
          url: uploadResult.secure_url,
          filename: originalName,
          size,
          provider: 'cloudinary',
          publicId: uploadResult.public_id,
        });
      } catch (cloudErr) {
        console.error('❌ Cloudinary falló:', cloudErr.message);
      }
    }

    if (storageBucket) {
      try {
        const safeName = originalName.replace(/[^\w\.\-]/g, '_');
        const fileName = `chat/${Date.now()}_${safeName}`;
        const fileRef = storageBucket.file(fileName);
        await fileRef.save(req.file.buffer, {
          contentType: mimetype,
          metadata: { cacheControl: 'public, max-age=31536000' },
          public: true,
          resumable: false,
        });
        const publicUrl = `https://storage.googleapis.com/${storageBucket.name}/${fileName}`;
        return res.json({
          success: true, url: publicUrl, filename: originalName, size,
          provider: 'firebase-storage',
        });
      } catch (fbErr) { console.error('❌ Firebase Storage falló:', fbErr.message); }
    }

    const fs = require('fs'), path = require('path');
    const dir = path.join(__dirname, 'uploads', 'chat');
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    const uniqueName = `${Date.now()}_${originalName.replace(/[^\w\.\-]/g, '_')}`;
    const filePath = path.join(dir, uniqueName);
    fs.writeFileSync(filePath, req.file.buffer);
    const baseUrl = `${req.protocol}://${req.get('host')}`;
    res.json({
      success: true,
      url: `${baseUrl}/api/chat/file/${encodeURIComponent(uniqueName)}`,
      filename: originalName, size, provider: 'local',
    });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.get('/api/chat/file/:filename', (req, res) => {
  try {
    const fs = require('fs'), path = require('path');
    const safeName = path.basename(req.params.filename);
    const filePath = path.join(__dirname, 'uploads', 'chat', safeName);
    if (!fs.existsSync(filePath)) return res.status(404).send('Not found');
    res.sendFile(filePath);
  } catch (e) { res.status(500).send(e.message); }
});

// ------------------------------------------------------------
//  WEBSOCKETS
// ------------------------------------------------------------
wss.on('connection', (ws) => {
  let userEmail = null;
  ws.on('message', async (message) => {
    try {
      const data = JSON.parse(message);
      if (data.type === 'login') {
        userEmail = data.email;
        if (activeWorkers.has(userEmail)) {
          activeWorkers.get(userEmail).ws = ws;
        } else {
          startImapWorker(data.email, data.password, data.imapHost, ws);
        }
      }
    } catch (e) { console.error('❌ Error WebSocket:', e.message); }
  });
  ws.on('close', () => {
    if (userEmail && activeWorkers.has(userEmail)) {
      activeWorkers.get(userEmail).ws = null;
    }
  });
});

// ------------------------------------------------------------
//  MOTOR IMAP PERSISTENTE
// ------------------------------------------------------------
function startImapWorker(email, password, customHost, ws = null) {
  if (activeWorkers.has(email)) {
    const existing = activeWorkers.get(email);
    if (ws) existing.ws = ws;
    return;
  }
  if (!password) return;
  if (isAccountBlocked(email)) return;

  const auto = getAutoConfig(email);
  const host = customHost || (auto ? auto.imapHost : 'mail.' + email.split('@')[1]);

  const workerState = {
    ws, email, password, host,
    lastUidNext: 0, client: null, active: true, lastAlive: Date.now(),
  };

  activeWorkers.set(email, workerState);
  console.log(`✅ Worker IMAP iniciado para ${email} en ${host}`);
  runImapLoop(workerState);
}

async function processEmailsInRange(state, startUid, endUidNext) {
  const safeStartUid = Math.max(1, startUid || 1);
  const safeEndUid = endUidNext - 1;
  if (safeStartUid > safeEndUid) {
    state.lastUidNext = endUidNext;
    await saveLastUid(state.email, endUidNext);
    return;
  }
  try {
    const fetchRange = `${safeStartUid}:${safeEndUid}`;
    console.log(`📥 Procesando correos en rango ${fetchRange} para ${state.email}`);
    const newIter = state.client.fetch(fetchRange, { uid: true, envelope: true }, { uid: true });
    for await (const msg of newIter) {
      if (msg.uid && msg.uid >= safeStartUid && msg.uid < endUidNext) {
        const from = msg.envelope?.from?.[0]?.address || 'Remitente desconocido';
        const subject = msg.envelope?.subject || 'Nuevo correo';
        console.log(`🔔 Correo entrante UID:${msg.uid} | De: ${from} | Asunto: ${subject}`);

        // Reglas de usuario
        try {
          const rules = await getRulesForAccount(state.email);
          if (rules.length > 0) {
            const needBody = rulesNeedBody(rules);
            let parsed = null, source = null;
            try {
              const fetchOpts = needBody ? { source: true } : { size: true };
              const full = await state.client.fetchOne(String(msg.uid), fetchOpts, { uid: true });
              if (full?.source) parsed = await simpleParser(full.source);
              source = full;
            } catch (e) { console.log(`⚠️ No se pudo bajar source: ${e.message}`); }
            await applyRulesToMessage({
              accountEmail: state.email, accountPassword: state.password,
              accountHost: state.host, uid: msg.uid, folder: 'INBOX',
              parsed, envelope: msg.envelope, size: source?.size || 0,
            });
          }
        } catch (e) { console.log(`⚠️ Error aplicando reglas: ${e.message}`); }

        // WebSocket
        if (state.ws && state.ws.readyState === WebSocket.OPEN) {
          state.ws.send(JSON.stringify({
            type: 'new_email', email: state.email,
            timestamp: new Date().toISOString(),
            from, subject, uid: msg.uid,
          }));
        }

        // Push normal (data-only)
        await sendPushNotification(state.email, {
          title: `📧 Nuevo correo de ${from}`,
          body: subject,
          data: {
            type: 'new_email',
            sender: from,
            subject,
            uid: String(msg.uid),
            folder: 'INBOX',
            text: subject,
          },
        }, { dataOnly: true });

        // 🔥 NUEVO: detección de urgencia (si el usuario la tiene activada)
        try {
          await checkUrgencyAndNotify({
            email: state.email,
            password: state.password,
            host: state.host,
            uid: msg.uid,
            from,
            subject,
            preview: subject, // sin preview real; la IA usará solo from+subject
            folder: 'INBOX',
          });
        } catch (e) {
          console.log(`⚠️ Error en detección de urgencia: ${e.message}`);
        }

        // Auto-reply de ausencia
        await checkAndSendAutoReply({
          accountEmail: state.email, accountPassword: state.password,
          incomingFrom: from, incomingSubject: subject, incomingUid: msg.uid,
        });
      }
    }
  } catch (err) {
    console.error(`❌ Error al procesar correos para ${state.email}:`, err.message);
  } finally {
    state.lastUidNext = endUidNext;
    await saveLastUid(state.email, endUidNext);
  }
}

async function runImapLoop(state) {
  while (state.active) {
    try {
      const conn = await connectImapAuto(state.email, state.password, state.host);
      state.client = conn.client;
      state.host = conn.host;
      resetWorkerFailure(state.email);

      await state.client.mailboxOpen('INBOX', { readOnly: true });

      const handleExists = async () => {
        try {
          state.lastAlive = Date.now();
          const status = await state.client.status('INBOX', { uidNext: true });
          const currentUidNext = status.uidNext || 1;
          if (currentUidNext > state.lastUidNext) {
            await processEmailsInRange(state, state.lastUidNext, currentUidNext);
          }
        } catch (e) { console.error(`❌ Error en handleExists:`, e.message); }
      };

      state.client.on('exists', handleExists);

      const statusInit = await state.client.status('INBOX', { uidNext: true });
      const currentUidNext = statusInit.uidNext || 1;
      const savedUid = await getSavedLastUid(state.email);

      if (savedUid && savedUid > 0 && savedUid < currentUidNext) {
        await processEmailsInRange(state, savedUid, currentUidNext);
      } else {
        state.lastUidNext = currentUidNext;
        await saveLastUid(state.email, currentUidNext);
      }

      let aliveLogCounter = 0;
      while (state.active && state.client.usable) {
        await new Promise(resolve => setTimeout(resolve, 5000));
        aliveLogCounter++;
        if (aliveLogCounter % 6 === 0) {
          state.lastAlive = Date.now();
        }
        if (aliveLogCounter % 12 === 0) {
          try {
            const s = await state.client.status('INBOX', { uidNext: true });
            if (s.uidNext && s.uidNext > state.lastUidNext) {
              await processEmailsInRange(state, state.lastUidNext, s.uidNext);
            }
          } catch (e) { break; }
        }
      }
    } catch (e) {
      if (state.active) {
        if (/Login is disabled|invalid credentials|auth/i.test(e.message)) {
          recordWorkerFailure(state.email);
        }
      }
    } finally {
      if (state.client) {
        try { state.client.removeAllListeners('exists'); } catch (_) {}
        await state.client.logout().catch(() => {});
        state.client = null;
      }
    }

    if (state.active) {
      if (isAccountBlocked(state.email)) {
        await new Promise(r => setTimeout(r, 30 * 60 * 1000));
      } else {
        await new Promise(r => setTimeout(r, 10000));
      }
    }
  }
}

// ------------------------------------------------------------
//  CRON: CALENDARIO
// ------------------------------------------------------------
cron.schedule('* * * * *', async () => {
  if (!db) return;
  try {
    const now = new Date();
    const snapshot = await db.collection('calendar_events').get();
    const offsetCache = new Map();
    const getOffset = async (email) => {
      const key = (email || '').toLowerCase();
      if (!key) return 0;
      if (offsetCache.has(key)) return offsetCache.get(key);
      const off = await getUserTimezoneOffset(key);
      offsetCache.set(key, off);
      return off;
    };

    for (const doc of snapshot.docs) {
      const event = doc.data();
      const rawTime = event.eventTime;
      const eventDate = rawTime?.toDate ? rawTime.toDate() : new Date(rawTime);
      if (isNaN(eventDate.getTime())) continue;

      const diffMs = eventDate.getTime() - now.getTime();
      const diffHours = diffMs / (1000 * 60 * 60);

      const recipientEmail = event.email
        || (Array.isArray(event.sharedEmails) && event.sharedEmails.length > 0
          ? event.sharedEmails[0] : null);
      if (!recipientEmail) continue;

      const offsetMin = await getOffset(recipientEmail);
      const { dayStr, timeStr, fullStr } = formatEventLocal(eventDate, offsetMin, now);

      const type = (event.type || 'cita').toLowerCase();
      let typeLabel, typeEmoji;
      if (type === 'cita') { typeLabel = 'Cita'; typeEmoji = '📅'; }
      else if (type === 'tarea') { typeLabel = 'Tarea'; typeEmoji = '✅'; }
      else if (type === 'alarma') { typeLabel = 'Alarma'; typeEmoji = '⏰'; }
      else { typeLabel = 'Evento'; typeEmoji = '📌'; }

      const nowLocalStr = formatNowLocal(offsetMin, now);
      const infoLine = `${typeEmoji} ${typeLabel} · ${fullStr}\n🔔 Aviso enviado: ${nowLocalStr}`;

      if (!event.notified1Day && diffHours <= 25 && diffHours > 23) {
        await sendPushNotification(recipientEmail, {
          title: `📅 Mañana: ${event.title}`,
          body: infoLine,
          data: {
            type: 'calendar_event', eventId: doc.id,
            eventTitle: event.title, eventTime: eventDate.toISOString(),
            eventDate: fullStr, eventDayShort: dayStr, eventTimeShort: timeStr,
            eventType: type, eventTypeLabel: typeLabel,
            notice: '1day', text: infoLine,
          },
        }, { dataOnly: true });
        await doc.ref.update({ notified1Day: true });
      }

      if (!event.notified1Hour && diffMs <= 65 * 60 * 1000 && diffMs > 55 * 60 * 1000) {
        await sendPushNotification(recipientEmail, {
          title: `⏰ En 1 hora: ${event.title}`,
          body: infoLine,
          data: {
            type: 'calendar_event', eventId: doc.id,
            eventTitle: event.title, eventTime: eventDate.toISOString(),
            eventDate: fullStr, eventDayShort: dayStr, eventTimeShort: timeStr,
            eventType: type, eventTypeLabel: typeLabel,
            notice: '1hour', text: infoLine,
          },
        }, { dataOnly: true });
        await doc.ref.update({ notified1Hour: true });
      }

      if (!event.notifiedEvent && diffMs <= 15 * 60 * 1000 && diffMs > 0) {
        await sendPushNotification(recipientEmail, {
          title: `⏰ Comienza pronto: ${event.title}`,
          body: infoLine,
          data: {
            type: 'calendar_event', eventId: doc.id,
            eventTitle: event.title, eventTime: eventDate.toISOString(),
            eventDate: fullStr, eventDayShort: dayStr, eventTimeShort: timeStr,
            eventType: type, eventTypeLabel: typeLabel,
            notice: '15min', text: infoLine,
          },
        }, { dataOnly: true });
        await doc.ref.update({ notifiedEvent: true });
      }
    }
  } catch (e) { console.error('❌ Error en Cron Job:', e.message); }
});

// ------------------------------------------------------------
//  CRON: REPORTE SEMANAL
// ------------------------------------------------------------
cron.schedule('0 8 * * 0', async () => {
  if (!db) return;
  try {
    const accountsSnap = await db.collection('user_accounts').get();
    if (accountsSnap.empty) return;

    const now = new Date();
    const nextMonday = new Date(now);
    const daysUntilMonday = (8 - now.getDay()) % 7 || 7;
    nextMonday.setDate(now.getDate() + daysUntilMonday);
    nextMonday.setHours(0, 0, 0, 0);

    const nextSundayEnd = new Date(nextMonday);
    nextSundayEnd.setDate(nextMonday.getDate() + 6);
    nextSundayEnd.setHours(23, 59, 59, 999);

    for (const accDoc of accountsSnap.docs) {
      const acc = accDoc.data();
      const email = acc.email;
      if (!email) continue;
      try {
        const emailLower = email.toLowerCase();
        const ownerSnap = await db.collection('calendar_events')
          .where('ownerEmail', '==', emailLower)
          .where('eventTime', '>=', admin.firestore.Timestamp.fromDate(nextMonday))
          .where('eventTime', '<=', admin.firestore.Timestamp.fromDate(nextSundayEnd))
          .get();
        const sharedSnap = await db.collection('calendar_events')
          .where('sharedEmails', 'arrayContains', email)
          .where('eventTime', '>=', admin.firestore.Timestamp.fromDate(nextMonday))
          .where('eventTime', '<=', admin.firestore.Timestamp.fromDate(nextSundayEnd))
          .get();

        const seenIds = new Set();
        const events = [];
        for (const d of ownerSnap.docs) {
          if (!seenIds.has(d.id)) { seenIds.add(d.id); events.push({ id: d.id, ...d.data() }); }
        }
        for (const d of sharedSnap.docs) {
          if (!seenIds.has(d.id)) { seenIds.add(d.id); events.push({ id: d.id, ...d.data() }); }
        }
        if (events.length === 0) continue;

        events.sort((a, b) => {
          const ta = a.eventTime?.toDate ? a.eventTime.toDate() : new Date(a.eventTime);
          const tb = b.eventTime?.toDate ? b.eventTime.toDate() : new Date(b.eventTime);
          return ta - tb;
        });

        const userOffset = await getUserTimezoneOffset(emailLower);
        const html = buildWeeklyReportHtml({
          recipientEmail: email,
          recipientName: (acc.email || '').split('@')[0],
          weekStart: nextMonday, weekEnd: nextSundayEnd,
          events, offsetMinutes: userOffset,
        });

        try {
          await sendViaBrevo({
            fromEmail: 'hola@rsmail.app', fromName: 'RSMail · Calendario',
            to: email,
            subject: `📅 Tu semana en RSMail: ${events.length} evento${events.length === 1 ? '' : 's'}`,
            html,
          });
        } catch (_) {}
      } catch (_) {}
    }
  } catch (e) { console.error('❌ [Weekly] Error general:', e.message); }
}, { timezone: 'Europe/Madrid' });

function buildWeeklyReportHtml({ recipientEmail, recipientName, weekStart, weekEnd, events, offsetMinutes }) {
  const off = Number(offsetMinutes) || 0;
  const fmtDate = (d) => {
    const local = new Date(d.getTime() + off * 60 * 1000);
    const days = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
    const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
      'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    return `${days[local.getUTCDay()]} ${local.getUTCDate()} de ${months[local.getUTCMonth()]}`;
  };
  const fmtTime = (d) => {
    const local = new Date(d.getTime() + off * 60 * 1000);
    return `${String(local.getUTCHours()).padStart(2, '0')}:${String(local.getUTCMinutes()).padStart(2, '0')}`;
  };

  const rangeStr = `${fmtDate(weekStart)} → ${fmtDate(weekEnd)}`;
  const cards = events.map((ev) => {
    const raw = ev.eventTime || ev.startTime;
    const dt = raw?.toDate ? raw.toDate() : new Date(raw);
    const dayStr = fmtDate(dt);
    const timeStr = fmtTime(dt);
    const type = (ev.type || 'cita').toLowerCase();
    let typeLabel = 'Evento', typeColor = '#1A73E8', typeEmoji = '📌';
    if (type === 'cita') { typeLabel = 'Cita'; typeColor = '#1A73E8'; typeEmoji = '📅'; }
    else if (type === 'tarea') { typeLabel = 'Tarea'; typeColor = '#FB8C00'; typeEmoji = '✅'; }
    else if (type === 'alarma') { typeLabel = 'Alarma'; typeColor = '#E53935'; typeEmoji = '⏰'; }

    const esc = (s) => String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const escNl = (s) => esc(s).replace(/\n/g, '<br>');

    return `<table width="100%" style="margin-bottom:14px;background:#fff;border-radius:14px;border:1px solid #E0E7EF;overflow:hidden;">
      <tr><td style="width:6px;background:${typeColor};"></td>
      <td style="padding:16px 18px;">
        <div style="font-size:17px;font-weight:700;color:#111;margin-bottom:6px;">${esc(ev.title || 'Evento')}</div>
        <div style="font-size:13px;color:#5A6B7B;">🗓 <b style="color:#1A73E8;">${esc(dayStr)}</b> · <b style="color:#1A73E8;">${esc(timeStr)}</b></div>
        <div style="font-size:13px;color:#5A6B7B;">${typeEmoji} ${typeLabel}</div>
      </td></tr></table>`;
  }).join('');

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#EEF3F8;font-family:sans-serif;">
<table width="100%" style="padding:24px 12px;"><tr><td align="center">
<table width="600" style="max-width:600px;">
<tr><td style="background:linear-gradient(135deg,#1A73E8,#0D47A1);padding:28px;border-radius:18px 18px 0 0;text-align:center;color:#fff;">
<div style="font-size:40px;">📅</div>
<div style="font-size:22px;font-weight:700;">Tu semana en RSMail</div>
<div style="font-size:13px;color:#B3D4FC;">${rangeStr}</div>
</td></tr>
<tr><td style="background:#fff;padding:22px 24px 8px;">
<div style="font-size:15px;color:#1A2A3A;">Hola <b>${recipientName || 'usuario'}</b>,</div>
<div style="font-size:14px;color:#5A6B7B;margin-top:6px;">Tienes <b style="color:#1A73E8;">${events.length} evento${events.length === 1 ? '' : 's'}</b> esta semana.</div>
</td></tr>
<tr><td style="background:#fff;padding:16px 24px;">${cards}</td></tr>
<tr><td style="background:#F5F8FB;padding:16px;text-align:center;border-radius:0 0 18px 18px;">
<div style="font-size:11px;color:#8896A5;">RSMail · Cada domingo a las 8:00</div>
</td></tr>
</table></td></tr></table></body></html>`;
}

// ------------------------------------------------------------
//  CRON: REANIMACIÓN DE WORKERS
// ------------------------------------------------------------
cron.schedule('*/3 * * * *', async () => {
  if (!db) return;
  try {
    const snapshot = await db.collection('user_accounts').get();
    for (const doc of snapshot.docs) {
      const data = doc.data();
      if (!data.email || !data.password) continue;
      if (isAccountBlocked(data.email)) continue;
      const state = activeWorkers.get(data.email);
      const isAlive = state && state.active && state.client && state.client.usable;
      if (!isAlive) {
        if (state) state.active = false;
        activeWorkers.delete(data.email);
        startImapWorker(data.email, data.password, data.imapHost);
      }
    }
  } catch (e) { console.error('⚠️ Error en cron:', e.message); }
});

// ------------------------------------------------------------
//  CRON: RECORDATORIOS
// ------------------------------------------------------------
cron.schedule('* * * * *', async () => {
  if (!db) return;
  try {
    const now = new Date();
    const snapshot = await db.collection('email_reminders').where('status', '==', 'pending').limit(100).get();
    if (snapshot.empty) return;
    const toSend = [];
    snapshot.forEach((doc) => {
      const data = doc.data();
      const ra = data.remindAt;
      if (!ra) return;
      const remindDate = ra.toDate ? ra.toDate() : new Date(ra);
      if (remindDate.getTime() <= now.getTime()) toSend.push(doc);
    });
    for (const doc of toSend) {
      const data = doc.data();
      const docRef = doc.ref;
      try {
        const accountEmail = data.accountEmail;
        if (!accountEmail) { await docRef.update({ status: 'failed', error: 'Sin cuenta' }); continue; }
        await docRef.update({ status: 'processing' });
        const from = data.emailFrom || 'remitente';
        const subject = data.emailSubject || '(Sin asunto)';
        const note = data.note || '';
        const title = note.length > 0 ? `🔔 ${note}` : `🔔 Recordatorio: ${subject}`;
        const body = note.length > 0 ? `Correo de ${from}: ${subject}` : `Correo de ${from}`;
        await sendPushNotification(accountEmail, {
          title, body,
          data: {
            type: 'email_reminder',
            uid: String(data.emailUid || ''),
            folder: data.emailFolder || 'INBOX',
            subject, sender: from, text: body,
          },
        }, { dataOnly: true });
        await docRef.update({ status: 'sent', sentAt: admin.firestore.FieldValue.serverTimestamp() });
      } catch (e) {
        await docRef.update({ status: 'failed', error: e.message });
      }
    }
  } catch (e) { console.error('❌ Error en cron recordatorios:', e.message); }
});

// ------------------------------------------------------------
//  CRON: LIMPIEZA AUTO-REPLIES
// ------------------------------------------------------------
cron.schedule('0 */6 * * *', async () => {
  if (!db) return;
  try {
    const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 30);
    const snap = await db.collection('vacation_sent_replies').where('sentAt', '<', admin.firestore.Timestamp.fromDate(cutoff)).get();
    for (const doc of snap.docs) await doc.ref.delete();
  } catch (_) {}
});

// ------------------------------------------------------------
//  FCM PUSH
// ------------------------------------------------------------
async function sendPushNotification(email, payload, options = {}) {
  if (!db) return;
  const dataOnly = options.dataOnly === true;
  try {
    const tokensSnapshot = await db.collection('fcm_tokens').where('email', '==', email).get();
    if (tokensSnapshot.empty) { console.log(`📴 Sin token FCM para ${email}`); return; }
    const tokens = [];
    tokensSnapshot.forEach(doc => tokens.push(doc.data().token));

    const rawData = payload.data || { type: 'general' };
    const safeData = {};
    Object.keys(rawData).forEach((k) => {
      safeData[k] = rawData[k] == null ? '' : String(rawData[k]);
    });

    const message = {
      data: safeData,
      android: { priority: 'high', ttl: 60 * 60 * 1000 },
      apns: {
        headers: {
          'apns-priority': '10',
          'apns-push-type': dataOnly ? 'background' : 'alert',
        },
        payload: {
          aps: dataOnly
            ? { 'content-available': 1 }
            : { sound: 'default', badge: 1, 'content-available': 1 },
        },
      },
      tokens,
    };

    if (!dataOnly) {
      message.notification = {
        title: payload.title || 'RSMAIL',
        body: payload.body || 'Nueva notificación',
      };
      message.android.notification = {
        channelId: 'rsmail_high_importance_channel',
        sound: 'default',
        priority: 'max',
        visibility: 'public',
        defaultVibrateTimings: true,
        defaultSound: true,
      };
    }

    const response = await admin.messaging().sendEachForMulticast(message);
    if (response.failureCount > 0) {
      const failedTokens = [];
      response.responses.forEach((resp, idx) => { if (!resp.success) failedTokens.push(tokens[idx]); });
      for (const token of failedTokens) {
        const snapshots = await db.collection('fcm_tokens').where('token', '==', token).get();
        snapshots.forEach(doc => doc.ref.delete());
      }
    }
  } catch (e) { console.error('❌ Error enviando Push:', e.message); }
}

// ------------------------------------------------------------
//  MODO CONFIDENCIAL
// ------------------------------------------------------------
app.post('/api/confidential/create', async (req, res) => {
  if (!db) return res.status(500).json({ success: false, error: 'Firestore no configurado' });
  try {
    const { ownerId, accountEmail, to, subject, body, password, expiresInDays, note } = req.body;
    if (!ownerId || !body || !expiresInDays) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
    let days = Number(expiresInDays);
    if (isNaN(days) || days < 1) days = 1;
    if (days > 15) days = 15;
    const expiresAt = new Date(); expiresAt.setDate(expiresAt.getDate() + days);
    const ref = await db.collection('confidential_emails').add({
      ownerId, accountEmail, to, subject, body,
      password: password || null,
      expiresAt: admin.firestore.Timestamp.fromDate(expiresAt),
      note: note || '', views: [],
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
      status: 'active',
    });
    res.json({
      success: true, id: ref.id,
      url: `https://rsmail-backend.onrender.com/confidential/${ref.id}`,
      expiresAt: expiresAt.toISOString(),
    });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.post('/api/confidential/open/:id', async (req, res) => {
  if (!db) return res.status(500).json({ success: false, error: 'Firestore no configurado' });
  try {
    const { id } = req.params;
    const { password } = req.body;
    const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress || 'unknown';
    const doc = await db.collection('confidential_emails').doc(id).get();
    if (!doc.exists) return res.status(404).json({ success: false, error: 'Enlace no encontrado' });
    const data = doc.data();
    const expiresAt = data.expiresAt?.toDate ? data.expiresAt.toDate() : new Date(data.expiresAt);
    if (expiresAt < new Date()) return res.status(410).json({ success: false, error: 'Enlace caducado' });
    if (data.password) {
      if (!password) return res.status(401).json({ success: false, error: 'Contraseña requerida' });
      if (password !== data.password) return res.status(401).json({ success: false, error: 'Contraseña incorrecta' });
    }
    const views = data.views || [];
    views.push({ at: new Date().toISOString(), ip });
    await doc.ref.update({ views });
    if (data.accountEmail) {
      sendPushNotification(data.accountEmail, {
        title: '🔓 Tu correo confidencial ha sido abierto',
        body: `"${data.subject || '(Sin asunto)'}" se ha abierto`,
        data: { type: 'confidential_opened', id },
      }).catch(() => {});
    }
    res.json({
      success: true,
      subject: data.subject || '(Sin asunto)',
      body: data.body || '',
      from: data.accountEmail, to: data.to,
      expiresAt: expiresAt.toISOString(),
      views: views.length,
    });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.get('/api/confidential/status/:id', async (req, res) => {
  if (!db) return res.status(500).json({ success: false, error: 'Firestore no configurado' });
  try {
    const doc = await db.collection('confidential_emails').doc(req.params.id).get();
    if (!doc.exists) return res.status(404).json({ success: false, error: 'No encontrado' });
    const data = doc.data();
    const expiresAt = data.expiresAt?.toDate ? data.expiresAt.toDate() : new Date(data.expiresAt);
    res.json({
      success: true, id: doc.id,
      subject: data.subject, to: data.to, note: data.note,
      status: expiresAt < new Date() ? 'expired' : 'active',
      expiresAt: expiresAt.toISOString(),
      views: (data.views || []).length,
      viewsList: data.views || [],
      hasPassword: !!data.password,
    });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.delete('/api/confidential/:id', async (req, res) => {
  if (!db) return res.status(500).json({ success: false, error: 'Firestore no configurado' });
  try {
    await db.collection('confidential_emails').doc(req.params.id).delete();
    res.json({ success: true });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.get('/confidential/:id', (req, res) => {
  const id = req.params.id;
  res.send(`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>RSMail · Confidencial</title>
<style>*{box-sizing:border-box;margin:0;padding:0;-webkit-user-select:none;user-select:none;}
body{font-family:-apple-system,sans-serif;background:linear-gradient(135deg,#0D47A1 0%,#1A73E8 100%);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:16px;color:#333;}
.container{background:#fff;border-radius:16px;max-width:720px;width:100%;box-shadow:0 10px 40px rgba(0,0,0,0.25);overflow:hidden;}
.header{background:#1A73E8;color:#fff;padding:20px;display:flex;align-items:center;gap:12px;}
.header h1{font-size:18px;font-weight:600;}.header .lock{font-size:24px;}
.banner{background:#FFF3CD;border-left:4px solid #FFC107;padding:12px 20px;font-size:13px;color:#856404;}
.content{padding:24px;}.subject{font-size:20px;font-weight:700;margin-bottom:8px;color:#111;}
.meta{font-size:12px;color:#888;margin-bottom:20px;}.body{font-size:15px;line-height:1.6;color:#222;white-space:pre-wrap;}
.footer{padding:16px 24px;background:#f5f5f5;font-size:11px;color:#999;text-align:center;}
.password-box{padding:40px 24px;text-align:center;}.password-box h2{font-size:18px;margin-bottom:16px;}
.password-box input{width:100%;max-width:300px;padding:12px 16px;border-radius:8px;border:2px solid #1A73E8;font-size:15px;outline:none;margin-bottom:12px;-webkit-user-select:text;user-select:text;}
.password-box button{background:#1A73E8;color:#fff;border:none;padding:12px 32px;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;}
.error{padding:40px 24px;text-align:center;color:#c62828;}
.loading{padding:40px;text-align:center;color:#1A73E8;font-size:14px;}
.spinner{width:40px;height:40px;border:4px solid #e0e0e0;border-top-color:#1A73E8;border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 16px;}
@keyframes spin{to{transform:rotate(360deg);}}</style></head>
<body><div class="container" id="app"><div class="loading"><div class="spinner"></div>Cargando...</div></div>
<script>
const CONF_ID='${id}';
document.addEventListener('contextmenu',e=>e.preventDefault());
document.addEventListener('copy',e=>e.preventDefault());
document.addEventListener('selectstart',e=>e.preventDefault());
document.addEventListener('keydown',e=>{if(e.ctrlKey&&['c','x','s','p','u','a'].includes(e.key.toLowerCase()))e.preventDefault();if(e.key==='F12')e.preventDefault();});
async function loadContent(password){const body={password:password||null};const res=await fetch('/api/confidential/open/'+CONF_ID,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});return{status:res.status,data:await res.json()};}
async function init(password){const app=document.getElementById('app');try{const{status,data}=await loadContent(password);
if(status===401){app.innerHTML='<div class="header"><span class="lock">🔒</span><h1>Mensaje confidencial</h1></div><div class="banner">Introduce la contraseña</div><div class="password-box"><h2>Contraseña</h2><input type="password" id="pwd" placeholder="Contraseña" autofocus><div><button onclick="submitPwd()">Abrir</button></div></div>';document.getElementById('pwd').addEventListener('keydown',e=>{if(e.key==='Enter')submitPwd();});window.submitPwd=()=>{const p=document.getElementById('pwd').value;if(!p)return;init(p);};return;}
if(!data.success){app.innerHTML='<div class="header"><span class="lock">🔒</span><h1>Confidencial</h1></div><div class="error"><h2>'+(status===410?'Enlace caducado':'No disponible')+'</h2><p>'+(data.error||'')+'</p></div>';return;}
const exp=new Date(data.expiresAt).toLocaleDateString('es-ES');
const safeBody=(data.body||'').replace(/</g,'&lt;').replace(/>/g,'&gt;');
app.innerHTML='<div class="header"><span class="lock">🔒</span><h1>Mensaje confidencial</h1></div><div class="banner">Caduca el '+exp+'</div><div class="content"><div class="subject">'+(data.subject||'(Sin asunto)').replace(/</g,'&lt;')+'</div><div class="meta">De: '+(data.from||'').replace(/</g,'&lt;')+' · Para: '+(data.to||'').replace(/</g,'&lt;')+'</div><div class="body">'+safeBody+'</div></div><div class="footer">RSMail · Caduca el '+exp+'</div>';}catch(e){app.innerHTML='<div class="error"><h2>Error</h2><p>'+e.message+'</p></div>';}}
init();
</script></body></html>`);
});

cron.schedule('*/30 * * * *', async () => {
  if (!db) return;
  try {
    const snap = await db.collection('confidential_emails').get();
    const now = Date.now();
    for (const doc of snap.docs) {
      const data = doc.data();
      const exp = data.expiresAt?.toDate ? data.expiresAt.toDate() : new Date(data.expiresAt);
      if (now - exp.getTime() > 30 * 24 * 60 * 60 * 1000) { await doc.ref.delete(); }
    }
  } catch (_) {}
});

// ------------------------------------------------------------
//  AUTH
// ------------------------------------------------------------
const handleAuth = async (req, res) => {
  const { email, password, host, port } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, error: 'Email y contraseña requeridos' });
  try {
    const conn = await connectImapAuto(email, password, host);
    const client = conn.client;
    await client.logout().catch(() => {});
    const auto = getAutoConfig(email);
    saveAccount(email, password, conn.host);
    startImapWorker(email, password, conn.host);
    return res.json({
      success: true, message: 'Autenticación exitosa',
      account: {
        email, password,
        imapHost: conn.host, imapPort: conn.port,
        imapSecurity: conn.port === 993 ? 'ssl' : 'starttls',
        smtpHost: auto.smtpHost, smtpPort: auto.smtpPort,
        smtpSecurity: auto.smtpSecure ? 'ssl' : 'starttls',
      },
    });
  } catch (e) {
    return res.status(401).json({ success: false, error: 'Credenciales inválidas: ' + e.message });
  }
};

// ------------------------------------------------------------
//  MICROSOFT LOGIN
// ------------------------------------------------------------
app.post('/api/microsoft/login', async (req, res) => {
  const { email, refreshToken } = req.body;
  if (!email || !refreshToken) return res.status(400).json({ success: false, error: 'Email y refreshToken requeridos' });
  if (!MICROSOFT_CLIENT_ID) return res.status(500).json({ success: false, error: 'Servidor sin MICROSOFT_CLIENT_ID' });
  try {
    const accessToken = await getMicrosoftAccessToken(refreshToken);
    const testClient = new ImapFlow({
      host: 'outlook.office365.com', port: 993, secure: true,
      auth: { user: email, accessToken },
      logger: false,
      tls: { rejectUnauthorized: false, minVersion: 'TLSv1.2' },
      connectionTimeout: 20000, greetingTimeout: 15000,
    });
    await testClient.connect();
    await testClient.logout().catch(() => {});
    saveAccount(email, refreshToken, 'outlook.office365.com');
    resetWorkerFailure(email);
    startImapWorker(email, refreshToken, 'outlook.office365.com');
    return res.json({
      success: true, message: 'Microsoft OAuth2 correcto',
      account: {
        email,
        imapHost: 'outlook.office365.com', imapPort: 993, imapSecurity: 'ssl',
        smtpHost: 'smtp.office365.com', smtpPort: 587, smtpSecurity: 'starttls',
      },
    });
  } catch (e) {
    return res.status(401).json({ success: false, error: 'Microsoft OAuth2 falló: ' + e.message });
  }
});

// ------------------------------------------------------------
//  API REST
// ------------------------------------------------------------
app.post('/api/login', handleAuth);
app.post('/api/verify', handleAuth);

app.post('/api/rules/invalidate', (req, res) => {
  const { email } = req.body || {};
  if (email) _rulesCache.delete(email);
  res.json({ ok: true });
});

app.get('/ping', async (req, res) => {
  if (db) {
    try {
      for (const [email, state] of activeWorkers.entries()) {
        const isAlive = state.active && state.client && state.client.usable;
        if (!isAlive) {
          state.active = false;
          activeWorkers.delete(email);
          if (isAccountBlocked(email)) continue;
          const doc = await db.collection('user_accounts').doc(email).get();
          if (doc.exists) {
            const data = doc.data();
            if (data.email && data.password) startImapWorker(data.email, data.password, data.imapHost);
          }
        }
      }
      if (activeWorkers.size === 0) {
        const snapshot = await db.collection('user_accounts').get();
        for (const d of snapshot.docs) {
          const data = d.data();
          if (data.email && data.password && !isAccountBlocked(data.email)) {
            startImapWorker(data.email, data.password, data.imapHost);
          }
        }
      }
    } catch (e) {}
  }
  res.json({ alive: true, ts: new Date().toISOString(), workers: activeWorkers.size, firestore: !!db, storage: !!storageBucket });
});

app.get('/api/debug/workers', (req, res) => {
  const workers = [];
  for (const [email, state] of activeWorkers.entries()) {
    workers.push({
      email, host: state.host, active: state.active,
      lastUidNext: state.lastUidNext,
      clientUsable: state.client?.usable ?? false,
      wsConnected: state.ws?.readyState === 1,
      lastAliveAgo: Math.floor((Date.now() - (state.lastAlive || 0)) / 1000) + 's',
      isMicrosoftOAuth: isMicrosoftOAuthAccount(email, state.password),
    });
  }
  res.json({
    count: workers.length, workers,
    uptime: process.uptime(), timestamp: new Date().toISOString(),
    microsoftClientIdConfigured: !!MICROSOFT_CLIENT_ID,
    firestoreAvailable: !!db,
    storageAvailable: !!storageBucket,
    groqConfigured: !!process.env.GROQ_API_KEY,
  });
});

app.get('/api/debug/auto-config/:email', (req, res) => {
  const email = req.params.email;
  const auto = getAutoConfig(email);
  res.json({
    email, provider: auto?.provider || 'unknown',
    imap: { host: auto?.imapHost, port: auto?.imapPort },
    smtp: { host: auto?.smtpHost, port: auto?.smtpPort, secure: auto?.smtpSecure },
  });
});

app.get('/api/debug/rules/:email', async (req, res) => {
  if (!db) return res.status(500).json({ error: 'Firestore no configurado' });
  try {
    const doc = await db.collection('rules_configs').doc(req.params.email).get();
    if (!doc.exists) return res.json({ count: 0, rules: [] });
    const rules = doc.data()?.rules || [];
    res.json({ count: rules.length, rules });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/debug/flags', async (req, res) => {
  const { email, password, host, folder = 'INBOX', uid } = req.body;
  if (!uid) return res.status(400).json({ success: false, error: 'uid requerido' });
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const lock = await client.getMailboxLock(folder, { readOnly: true });
    let flags = [];
    try {
      const msg = await client.fetchOne(String(uid), { flags: true }, { uid: true });
      let f = msg?.flags;
      if (f instanceof Set) f = Array.from(f);
      else if (!Array.isArray(f)) f = [];
      flags = f;
    } finally { lock.release(); }
    await client.logout();
    res.json({
      success: true, email, folder, uid, flags,
      isSeen: flags.includes('\\Seen'),
      isFlagged: flags.includes('\\Flagged'),
    });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/fcm-token', async (req, res) => {
  const { email, token, password, imapHost, timezoneOffset } = req.body;
  if (!email || !token) return res.status(400).json({ success: false, error: 'Email y token requeridos' });
  if (!db) return res.status(500).json({ success: false, error: 'Firestore no configurado' });
  try {
    const existing = await db.collection('fcm_tokens').where('email', '==', email).get();
    existing.forEach(doc => doc.ref.delete());

    const offset =
      (typeof timezoneOffset === 'number' && !isNaN(timezoneOffset))
        ? Math.max(-840, Math.min(840, Math.round(timezoneOffset)))
        : 0;

    await db.collection('fcm_tokens').add({
      email, token, timezoneOffset: offset,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    if (password) {
      saveAccount(email, password, imapHost);
      startImapWorker(email, password, imapHost);
    }
    res.json({ success: true });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.post('/api/send-notification', async (req, res) => {
  const { email, title, body, data } = req.body;
  if (!email) return res.status(400).json({ success: false, error: 'Email requerido' });
  try {
    await sendPushNotification(email, {
      title: title || '📅 Recordatorio',
      body: body || 'Evento programado',
      data: data || { type: 'calendar_event' },
    });
    res.json({ success: true });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.post('/api/chat/notify', async (req, res) => {
  if (!db) return res.status(500).json({ success: false, error: 'Firestore no configurado' });
  const { conversationId, senderEmail, senderName, recipientEmails, text, isGroup, groupName } = req.body;
  if (!conversationId || !senderEmail || !Array.isArray(recipientEmails)) {
    return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  }
  const senderDisplay = senderName || senderEmail.split('@')[0];
  const preview = (text || '').length > 120 ? text.substring(0, 120) + '…' : text || 'Nuevo mensaje';
  const title = isGroup && groupName ? `${groupName} · ${senderDisplay}` : `💬 ${senderDisplay}`;
  const notified = [];
  for (const recipient of recipientEmails) {
    const email = String(recipient).toLowerCase();
    if (email === senderEmail.toLowerCase()) continue;
    try {
      await sendPushNotification(email, {
        title, body: preview,
        data: {
          type: 'chat_message', conversationId, senderEmail,
          senderName: senderDisplay,
          isGroup: isGroup ? 'true' : 'false',
          groupName: groupName || '', text: preview,
        },
      }, { dataOnly: true });
      notified.push(email);
    } catch (_) {}
  }
  res.json({ success: true, notified });
});

app.post('/api/send-email', async (req, res) => {
  const { email, password, host, port, to, subject, body, attachments } = req.body;
  if (!email || !password || !to) return res.status(400).json({ success: false, error: 'Faltan campos' });
  const auto = getAutoConfig(email);
  const smtpHost = host || auto.smtpHost;
  const smtpPort = Number(port) || auto.smtpPort || 587;
  try {
    const transporter = await createSmtpTransporter({
      email, password, host: smtpHost, port: smtpPort, secure: smtpPort === 465,
    });
    await transporter.sendMail({
      from: email, to,
      subject: subject || '(Sin asunto)',
      html: body || '',
      attachments: attachments ? attachments.map(att => ({
        filename: att.filename,
        content: Buffer.from(att.content, 'base64'),
      })) : [],
    });
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/save-to-sent', async (req, res) => {
  const { email, password, host, port, to, cc, bcc, subject, body } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, error: 'Parámetros insuficientes' });

  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;

    const list = await client.list();
    const candidates = list.filter(f => {
      const name = (f.name || '').toLowerCase();
      const path = (f.path || '').toLowerCase();
      const su = (f.specialUse || '').toLowerCase();
      return (
        su === '\\sent' || name === 'sent' || name === 'sent items' ||
        name === 'sent messages' || name === 'enviados' ||
        name === 'elementos enviados' || name.includes('sent') ||
        name.includes('enviad') || path.includes('sent') || path.includes('enviad')
      );
    });

    let sentFolder = candidates[0]?.path || 'Sent';
    const messageId = `<${Date.now()}.${Math.random().toString(36).substring(2, 10)}@${email.split('@')[1]}>`;
    const date = new Date().toUTCString();
    const fromName = email.split('@')[0];

    const rawEmail = [
      `From: ${fromName} <${email}>`,
      `To: ${to || ''}`,
      ...(cc ? [`Cc: ${cc}`] : []),
      ...(bcc ? [`Bcc: ${bcc}`] : []),
      `Subject: ${subject || '(Sin asunto)'}`,
      `Date: ${date}`,
      `Message-ID: ${messageId}`,
      `MIME-Version: 1.0`,
      `Content-Type: text/html; charset=UTF-8`,
      `Content-Transfer-Encoding: 8bit`,
      '', body || '',
    ].join('\r\n');

    let saved = false;
    try {
      await client.append(sentFolder, Buffer.from(rawEmail, 'utf8'), ['\\Seen']);
      saved = true;
    } catch (_) {}

    if (!saved) {
      const alternatives = ['Sent', 'Sent Items', 'Sent Messages', 'Enviados', 'INBOX.Sent', '[Gmail]/Sent Mail'];
      for (const alt of alternatives) {
        if (alt === sentFolder) continue;
        try {
          await client.append(alt, Buffer.from(rawEmail, 'utf8'), ['\\Seen']);
          saved = true;
          sentFolder = alt;
          break;
        } catch (_) {}
      }
    }

    await client.logout();
    client = null;

    if (!saved) {
      return res.status(500).json({ success: false, error: 'No se pudo guardar en Enviados' });
    }
    res.json({ success: true, message: 'Guardado', folder: sentFolder });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/folders', async (req, res) => {
  const { email, password, host, port } = req.body;
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const list = await client.list();
    await client.logout();
    const folders = list.map(f => ({ name: f.name, path: f.path, specialUse: f.specialUse || '' }));
    res.json({ success: true, folders });
  } catch (err) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: err.message });
  }
});

// ------------------------------------------------------------
//  MENSAJES
// ------------------------------------------------------------
function detectAttachmentsFromStructure(structure) {
  if (!structure) return false;
  const stack = [structure];
  while (stack.length) {
    const part = stack.pop();
    if (!part) continue;
    const disp = (part.disposition || '').toString().toLowerCase();
    if (disp === 'attachment') return true;
    if (part.dispositionParameters && part.dispositionParameters.filename) return true;
    if (part.parameters && part.parameters.name) return true;
    if (Array.isArray(part.childNodes)) for (const child of part.childNodes) stack.push(child);
  }
  return false;
}

app.post('/api/messages', async (req, res) => {
  const { email, password, host, port, folder = 'INBOX', limit = 20 } = req.body;
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const lock = await client.getMailboxLock(folder, { readOnly: true });
    const messages = [];
    try {
      const iter = client.fetch('1:*',
        { envelope: true, flags: true, bodyStructure: true },
        { max: limit, reverse: true });
      for await (const msg of iter) {
        let flags = msg.flags;
        if (flags instanceof Set) flags = Array.from(flags);
        else if (!Array.isArray(flags)) flags = [];

        const hasAttachments = detectAttachmentsFromStructure(msg.bodyStructure);
        const fromAddr = msg.envelope?.from?.[0]?.address || '';
        const fromName = msg.envelope?.from?.[0]?.name || '';
        const toAddr = msg.envelope?.to?.[0]?.address || '';
        const toName = msg.envelope?.to?.[0]?.name || '';

        messages.push({
          uid: msg.uid, id: msg.uid.toString(),
          subject: msg.envelope?.subject || '(Sin asunto)',
          from: fromAddr || fromName,
          fromName,
          to: toAddr || toName,
          toName,
          preview: '',
          date: msg.envelope?.date ? new Date(msg.envelope.date).toISOString() : new Date().toISOString(),
          hasAttachments, flags,
          isRead: flags.includes('\\Seen'),
          isFlagged: flags.includes('\\Flagged'),
        });
      }
    } finally { lock.release(); }
    await client.logout();

    // 🔥 NUEVO: enriquecer con urgencia desde Firestore (si la hay)
    if (db) {
      try {
        const urgencyCol = db.collection('email_urgency').doc(email).collection('messages');
        for (const m of messages) {
          try {
            const doc = await urgencyCol.doc(String(m.uid)).get();
            if (doc.exists) {
              const d = doc.data();
              m['isUrgent'] = d.isUrgent === true;
              m['urgentScore'] = d.score || 0;
              m['urgentReason'] = d.reason || '';
            }
          } catch (_) {}
        }
      } catch (_) {}
    }

    messages.sort((a, b) => new Date(b.date) - new Date(a.date));
    res.json({ success: true, messages, total: messages.length });
  } catch (err) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: err.message, messages: [] });
  }
});

app.post('/api/message-detail', async (req, res) => {
  const { email, password, host, port, folder = 'INBOX', uid } = req.body;
  if (!uid) return res.status(400).json({ success: false, error: 'UID requerido' });
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const lock = await client.getMailboxLock(folder, { readOnly: true });
    let parsed;
    try {
      const msg = await client.fetchOne(String(uid), { source: true }, { uid: true });
      if (msg?.source) parsed = await simpleParser(msg.source);
    } finally { lock.release(); }
    await client.logout();
    if (!parsed) return res.status(404).json({ success: false, error: 'Correo no encontrado' });
    const attachments = (parsed.attachments || []).map(att => ({
      filename: att.filename || 'adjunto',
      contentType: att.contentType,
      size: att.size,
      content: att.content ? att.content.toString('base64') : '',
    }));

    // 🔥 NUEVO: leer urgencia
    let urgency = null;
    if (db) {
      try {
        const doc = await db.collection('email_urgency').doc(email).collection('messages').doc(String(uid)).get();
        if (doc.exists) {
          const d = doc.data();
          urgency = {
            isUrgent: d.isUrgent === true,
            score: d.score || 0,
            reason: d.reason || '',
            level: d.level || 'medium',
          };
        }
      } catch (_) {}
    }

    res.json({
      success: true,
      message: {
        uid: Number(uid),
        subject: parsed.subject || '(Sin asunto)',
        from: parsed.from?.text || '',
        to: parsed.to?.text || '',
        date: parsed.date ? parsed.date.toISOString() : new Date().toISOString(),
        text: parsed.text || '',
        html: parsed.html || parsed.textAsHtml || parsed.text || '',
        attachments,
        urgency, // 🔥 nuevo
      },
    });
  } catch (err) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: err.message });
  }
});

// ------------------------------------------------------------
//  MOVER / CREAR / BORRAR CARPETA, DELETE MESSAGE, READ, FLAG
// ------------------------------------------------------------
app.post('/api/move-message', async (req, res) => {
  const { email, password, host, port, uid, fromFolder, toFolder } = req.body;
  if (!email || !password || !uid || !fromFolder || !toFolder) {
    return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  }
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const lock = await client.getMailboxLock(fromFolder);
    try {
      let targetFolder = toFolder;
      try {
        const list = await client.list();
        const found = list.find((f) =>
          f.path.toLowerCase() === toFolder.toLowerCase() ||
          f.name.toLowerCase() === toFolder.toLowerCase());
        if (found) targetFolder = found.path;
      } catch (_) {}
      await client.messageMove(String(uid), targetFolder, { uid: true });
      try { lock.release(); } catch (_) {}
      await client.logout();
      client = null;
      return res.json({ success: true, movedTo: targetFolder });
    } catch (e) {
      try { lock.release(); } catch (_) {}
      throw e;
    }
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    return res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/create-folder', async (req, res) => {
  const { email, password, host, folderName } = req.body;
  if (!email || !password || !folderName) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    await client.mailboxCreate(folderName);
    await client.logout();
    return res.json({ success: true, folder: folderName });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    return res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/delete-folder', async (req, res) => {
  const { email, password, host, folderName } = req.body;
  if (!email || !password || !folderName) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    let targetFolder = folderName;
    try {
      const list = await client.list();
      const found = list.find((f) =>
        f.path.toLowerCase() === folderName.toLowerCase() ||
        f.name.toLowerCase() === folderName.toLowerCase());
      if (found) targetFolder = found.path;
    } catch (_) {}
    await client.mailboxDelete(targetFolder);
    await client.logout();
    return res.json({ success: true, folder: targetFolder });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    return res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/delete-message', async (req, res) => {
  const { email, password, host, uid, folder = 'INBOX' } = req.body;
  if (!uid) return res.status(400).json({ success: false, error: 'Falta UID' });
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const lock = await client.getMailboxLock(folder);
    try {
      const isAlreadyTrash = folder.toLowerCase().includes('trash') || folder.toLowerCase().includes('papelera');
      if (isAlreadyTrash) {
        await client.messageDelete(String(uid), { uid: true });
      } else {
        const list = await client.list();
        const trashFolder = list.find(f =>
          f.specialUse === '\\Trash' ||
          /^trash$/i.test(f.name) ||
          /papelera/i.test(f.name) ||
          /inbox\.trash/i.test(f.path)
        )?.path || 'Trash';
        await client.messageMove(String(uid), trashFolder, { uid: true });
      }
    } finally { lock.release(); }
    await client.logout();
    res.json({ success: true });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/toggle-read', async (req, res) => {
  const { email, password, host, uid, folder = 'INBOX', read } = req.body;
  if (uid == null) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const lock = await client.getMailboxLock(folder);
    try {
      if (read) await client.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true });
      else await client.messageFlagsRemove(String(uid), ['\\Seen'], { uid: true });
    } finally { lock.release(); }
    await client.logout();
    res.json({ success: true });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/mark-all-read', async (req, res) => {
  const { email, password, host, folder = 'INBOX' } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const lock = await client.getMailboxLock(folder);
    let total = 0;
    try {
      const uids = [];
      const iter = client.fetch('1:*', { uid: true }, { uid: true });
      for await (const msg of iter) { if (msg.uid) uids.push(msg.uid); }
      total = uids.length;
      if (total > 0) await client.messageFlagsAdd(uids, ['\\Seen'], { uid: true });
    } finally { lock.release(); }
    await client.logout();
    res.json({ success: true, marked: total });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/toggle-flagged', async (req, res) => {
  const { email, password, host, uid, folder = 'INBOX', flagged } = req.body;
  if (uid == null) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const lock = await client.getMailboxLock(folder);
    try {
      if (flagged) await client.messageFlagsAdd(String(uid), ['\\Flagged'], { uid: true });
      else await client.messageFlagsRemove(String(uid), ['\\Flagged'], { uid: true });
    } finally { lock.release(); }
    await client.logout();
    res.json({ success: true });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

// ------------------------------------------------------------
//  CLOUD / ADJUNTOS / SUSCRIPCIONES
// ------------------------------------------------------------
function collectAttachmentParts(structure, prefix = '') {
  const out = [];
  const walk = (node, currentPart) => {
    if (!node) return;
    const part = currentPart || node.part || '';
    const disposition = (node.disposition || '').toString().toLowerCase();
    const filename = node.dispositionParameters?.filename || node.parameters?.name || '';
    const isAttachment = disposition === 'attachment' || filename.length > 0;
    if (isAttachment && part) {
      out.push({
        partId: String(part),
        filename: filename || `adjunto_${part}`,
        contentType: node.type || 'application/octet-stream',
        size: node.size || 0,
      });
    }
    if (Array.isArray(node.childNodes)) {
      node.childNodes.forEach((child) => walk(child, child.part || part));
    }
  };
  walk(structure, prefix);
  return out;
}

app.post('/api/scan-attachments', async (req, res) => {
  const { email, password, host, folder = 'INBOX', limit = 200 } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, error: 'Faltan', attachments: [] });
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const lock = await client.getMailboxLock(folder, { readOnly: true });
    const attachments = [];
    try {
      const status = await client.status(folder, { messages: true });
      const total = status.messages || 0;
      const startSeq = Math.max(1, total - limit + 1);
      const iter = client.fetch(`${startSeq}:*`,
        { uid: true, envelope: true, bodyStructure: true },
        { uid: true });
      for await (const msg of iter) {
        const parts = collectAttachmentParts(msg.bodyStructure);
        if (parts.length === 0) continue;
        const from = msg.envelope?.from?.[0]?.address || '';
        const fromName = msg.envelope?.from?.[0]?.name || from;
        const subject = msg.envelope?.subject || '(Sin asunto)';
        const date = msg.envelope?.date ? new Date(msg.envelope.date).toISOString() : new Date().toISOString();
        for (const p of parts) {
          attachments.push({
            uid: msg.uid, folder, partId: p.partId,
            filename: p.filename, contentType: p.contentType, size: p.size,
            fromEmail: from, fromName, subject, date,
          });
        }
      }
    } finally { lock.release(); }
    await client.logout();
    attachments.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
    res.json({ success: true, attachments, total: attachments.length });
  } catch (err) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: err.message, attachments: [] });
  }
});

app.post('/api/download-attachment', async (req, res) => {
  const { email, password, host, folder = 'INBOX', uid, partId, filename } = req.body;
  if (!uid) return res.status(400).json({ success: false, error: 'Falta UID' });
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const lock = await client.getMailboxLock(folder, { readOnly: true });
    let parsed;
    try {
      const msg = await client.fetchOne(String(uid), { source: true }, { uid: true });
      if (msg?.source) parsed = await simpleParser(msg.source);
    } finally { lock.release(); }
    await client.logout();
    client = null;
    if (!parsed) return res.status(404).json({ success: false, error: 'No encontrado' });
    const attachments = parsed.attachments || [];
    if (attachments.length === 0) return res.status(404).json({ success: false, error: 'Sin adjuntos' });
    let target = null;
    if (filename) target = attachments.find((a) => a.filename === filename);
    if (!target && partId) {
      const parts = String(partId).trim().split('.').map((p) => parseInt(p, 10) || 0);
      const lastIdx = parts[parts.length - 1];
      if (lastIdx > 0 && lastIdx <= attachments.length) target = attachments[lastIdx - 1];
    }
    if (!target && attachments.length === 1) target = attachments[0];
    if (!target) return res.status(404).json({ success: false, error: 'Adjunto no encontrado' });
    const content = target.content || Buffer.from('');
    res.json({
      success: true,
      data: content.toString('base64'),
      filename: target.filename || filename || 'adjunto',
      contentType: target.contentType || 'application/octet-stream',
      size: content.length,
    });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

function analyzeIsSubscription({ from, subject, listUnsubscribe }) {
  if (listUnsubscribe && listUnsubscribe.trim().length > 0) return true;
  const text = `${from || ''} ${subject || ''}`.toLowerCase();
  const keywords = ['newsletter', 'boletin', 'boletín', 'suscripción', 'suscripcion',
    'newsletter@', 'marketing@', 'info@', 'noreply@', 'no-reply@',
    'no responder', 'no-responder', 'promociones', 'publicidad'];
  return keywords.some(k => text.includes(k));
}

function parseListUnsubscribe(raw) {
  if (!raw) return { mailto: [], http: [] };
  const result = { mailto: [], http: [] };
  const regex = /<([^>]+)>/g;
  let match;
  while ((match = regex.exec(raw)) !== null) {
    const url = match[1].trim();
    if (url.toLowerCase().startsWith('mailto:')) result.mailto.push(url);
    else if (url.toLowerCase().startsWith('http')) result.http.push(url);
  }
  return result;
}

app.post('/api/scan-subscriptions', async (req, res) => {
  const { email, password, host, maxMessages = 500 } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, error: 'Faltan', subscriptions: [] });
  let client;
  try {
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;
    const lock = await client.getMailboxLock('INBOX', { readOnly: true });
    const subscriptions = new Map();
    try {
      const status = await client.status('INBOX', { messages: true });
      const total = status.messages || 0;
      const startSeq = Math.max(1, total - maxMessages + 1);
      const iter = client.fetch(`${startSeq}:*`, {
        uid: true, envelope: true,
        headers: ['list-unsubscribe', 'list-unsubscribe-post'],
      });
      for await (const msg of iter) {
        const fromAddr = msg.envelope?.from?.[0]?.address || '';
        const fromName = msg.envelope?.from?.[0]?.name || '';
        const subject = msg.envelope?.subject || '';
        const date = msg.envelope?.date ? new Date(msg.envelope.date) : null;
        let headerListUnsub = '', headerListUnsubPost = '';
        if (msg.headers) {
          const raw = msg.headers.toString();
          const mUnsub = raw.match(/^List-Unsubscribe:\s*(.+)$/im);
          if (mUnsub) headerListUnsub = mUnsub[1].trim();
          const mPost = raw.match(/^List-Unsubscribe-Post:\s*(.+)$/im);
          if (mPost) headerListUnsubPost = mPost[1].trim();
        }
        if (!fromAddr) continue;
        const isSub = analyzeIsSubscription({
          from: fromAddr + ' ' + fromName, subject,
          listUnsubscribe: headerListUnsub,
        });
        if (!isSub) continue;
        const key = fromAddr.toLowerCase();
        const parsed = parseListUnsubscribe(headerListUnsub);
        if (!subscriptions.has(key)) {
          subscriptions.set(key, {
            email: fromAddr, name: fromName || fromAddr, subject,
            latestDate: date ? date.toISOString() : null,
            count: 1, listUnsubscribe: headerListUnsub,
            listUnsubscribePost: headerListUnsubPost,
            hasUnsubscribe: parsed.mailto.length > 0 || parsed.http.length > 0,
          });
        } else {
          const s = subscriptions.get(key);
          s.count++;
          if (date && (!s.latestDate || new Date(s.latestDate) < date)) {
            s.latestDate = date.toISOString();
            s.subject = subject;
            if (headerListUnsub) {
              s.listUnsubscribe = headerListUnsub;
              s.listUnsubscribePost = headerListUnsubPost;
              s.hasUnsubscribe = parsed.mailto.length > 0 || parsed.http.length > 0;
            }
          }
        }
      }
    } finally { lock.release(); }
    await client.logout();
    const list = Array.from(subscriptions.values()).sort((a, b) => (b.latestDate || '').localeCompare(a.latestDate || ''));
    res.json({ success: true, subscriptions: list, total: list.length });
  } catch (err) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: err.message, subscriptions: [] });
  }
});

app.post('/api/unsubscribe', async (req, res) => {
  const { email, password, listUnsubscribe, listUnsubscribePost } = req.body;
  if (!listUnsubscribe) return res.status(400).json({ success: false, error: 'Falta' });
  const parsed = parseListUnsubscribe(listUnsubscribe);
  const isOneClick = (listUnsubscribePost || '').toLowerCase().includes('one-click');
  try {
    let method = null, result = null;
    if (parsed.http.length > 0 && isOneClick) {
      try {
        const r = await fetch(parsed.http[0], {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'User-Agent': 'RSMail/3.0' },
          body: 'List-Unsubscribe=One-Click',
        });
        method = 'https-post';
        result = { status: r.status, ok: r.ok };
      } catch (_) {}
    }
    if (!result && parsed.http.length > 0) {
      try {
        const r = await fetch(parsed.http[0], { method: 'GET', headers: { 'User-Agent': 'RSMail/3.0' }, redirect: 'follow' });
        method = 'https-get';
        result = { status: r.status, ok: r.ok };
      } catch (_) {}
    }
    if (!result && parsed.mailto.length > 0) {
      const mailtoUrl = parsed.mailto[0].replace(/^mailto:/i, '');
      const [address, query] = mailtoUrl.split('?');
      const subjectMatch = (query || '').match(/subject=([^&]+)/i);
      const subject = subjectMatch ? decodeURIComponent(subjectMatch[1]) : 'unsubscribe';
      const auto = getAutoConfig(email);
      const transporter = await createSmtpTransporter({
        email, password,
        host: auto.smtpHost, port: auto.smtpPort, secure: auto.smtpSecure,
      });
      await transporter.sendMail({ from: email, to: address, subject, text: 'unsubscribe' });
      method = 'mailto';
      result = { ok: true };
    }
    if (!result) return res.status(400).json({ success: false, error: 'No hay método válido' });
    res.json({ success: true, method, result });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

app.delete('/api/account/:email', async (req, res) => {
  const email = decodeURIComponent(req.params.email);
  if (!email) return res.status(400).json({ success: false, error: 'Email requerido' });
  try {
    if (activeWorkers.has(email)) {
      const state = activeWorkers.get(email);
      state.active = false;
      activeWorkers.delete(email);
    }
    if (db) {
      await db.collection('user_accounts').doc(email).delete().catch(() => {});
      await db.collection('user_states').doc(email).delete().catch(() => {});
    }
    failedAccounts.delete(email);
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ------------------------------------------------------------
//  COMPARTIR EVENTO PÚBLICO
// ------------------------------------------------------------
app.get('/event/invite/:id', async (req, res) => {
  if (!db) return res.status(500).send('Firestore no disponible');
  try {
    const id = req.params.id;
    const doc = await db.collection('calendar_events').doc(id).get();
    if (!doc.exists) {
      return res.status(404).send(`<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:sans-serif;padding:40px;text-align:center;">
<h1 style="color:#c62828;">Evento no encontrado</h1></body></html>`);
    }

    const ev = doc.data() || {};
    const title = ev.title || 'Evento';
    const type = ev.type || 'cita';
    const isUrgent = ev.urgent === true;
    const desc = ev.description || '';
    const location = ev.location || '';

    let startDate = null;
    const raw = ev.eventTime || ev.startTime;
    if (raw && typeof raw.toDate === 'function') startDate = raw.toDate();
    else if (raw) startDate = new Date(raw);
    if (startDate && isNaN(startDate.getTime())) startDate = null;

    const fmtDisplay = (d) =>
      d.toLocaleString('es-ES', {
        day: '2-digit', month: 'long', year: 'numeric',
        hour: '2-digit', minute: '2-digit',
        timeZone: 'Europe/Madrid',
      });

    const timeStr = startDate ? fmtDisplay(startDate) : 'Fecha no especificada';

    let typeLabel = 'Evento', typeEmoji = '📌';
    if (type === 'cita') { typeLabel = 'Cita'; typeEmoji = '📅'; }
    else if (type === 'alarma') { typeLabel = 'Alarma'; typeEmoji = '⏰'; }
    else if (type === 'tarea') { typeLabel = 'Tarea'; typeEmoji = '✅'; }

    const esc = (s) => String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

    let googleUrl = '';
    if (startDate) {
      const end = new Date(startDate.getTime() + 60 * 60 * 1000);
      const fmtG = (d) =>
        d.getUTCFullYear().toString() +
        String(d.getUTCMonth() + 1).padStart(2, '0') +
        String(d.getUTCDate()).padStart(2, '0') + 'T' +
        String(d.getUTCHours()).padStart(2, '0') +
        String(d.getUTCMinutes()).padStart(2, '0') + '00Z';
      const params = new URLSearchParams({
        action: 'TEMPLATE', text: title,
        dates: `${fmtG(startDate)}/${fmtG(end)}`,
      });
      if (desc) params.set('details', desc);
      if (location) params.set('location', location);
      googleUrl = `https://calendar.google.com/calendar/render?${params.toString()}`;
    }

    const rsmailLink = `rsmail://event/invite/${id}`;
    res.send(`<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>RSMail · ${esc(title)}</title></head>
<body style="font-family:sans-serif;background:linear-gradient(135deg,#0D47A1,#1A73E8);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:16px;margin:0;">
<div style="background:#fff;border-radius:16px;max-width:640px;width:100%;padding:24px;">
<h1 style="font-size:22px;margin-bottom:14px;">${esc(title)}</h1>
<p style="font-size:14px;margin:6px 0;">🗓 ${esc(timeStr)}</p>
<p style="font-size:14px;margin:6px 0;">${typeEmoji} ${typeLabel}</p>
${isUrgent ? '<p style="color:#D32F2F;font-weight:bold;">🔴 URGENTE</p>' : ''}
${location ? `<p style="font-size:14px;margin:6px 0;">📍 ${esc(location)}</p>` : ''}
${desc ? `<p style="font-size:14px;margin:14px 0;white-space:pre-wrap;">${esc(desc)}</p>` : ''}
<div style="margin-top:20px;display:flex;flex-direction:column;gap:10px;">
${googleUrl ? `<a href="${googleUrl}" style="text-align:center;padding:14px;background:#1A73E8;color:#fff;text-decoration:none;border-radius:10px;font-weight:600;">📅 Añadir a Google Calendar</a>` : ''}
<a href="${rsmailLink}" style="text-align:center;padding:14px;border:2px solid #1A73E8;color:#1A73E8;text-decoration:none;border-radius:10px;font-weight:600;">📲 Abrir en RSMail</a>
</div></div></body></html>`);
  } catch (e) {
    res.status(500).send(`<!DOCTYPE html><html><body style="font-family:sans-serif;padding:40px;text-align:center;">
<h1 style="color:#c62828;">Error</h1><p>${e.message}</p></body></html>`);
  }
});

// ------------------------------------------------------------
//  🔥 IA — GROQ
// ------------------------------------------------------------
const GROQ_MODELS_FALLBACK = [
  'openai/gpt-oss-20b',
  'openai/gpt-oss-120b',
];

const RSMAIL_SYSTEM_PROMPT = `Eres RSMail AI, el asistente inteligente integrado en la aplicación RSMail.

REGLAS DE FORMATO (MUY IMPORTANTE):
- NUNCA uses markdown. Nada de **negrita**, ni ## títulos, ni \`código\`, ni listas con - o *.
- Escribe en texto plano, como si fuera una conversación normal.
- Puedes usar números para pasos: 1., 2., 3.
- Puedes usar emojis con moderación (1 o 2 por respuesta máximo).
- Máximo 120 palabras por respuesta salvo que el usuario pida detalle.

CONOCIMIENTO REAL DE LA APP (usa esta información para responder con precisión):
- Las reglas y filtros se configuran en: Mi Cuenta → Reglas y filtros → Gestionar reglas.
- Las firmas: Mi Cuenta → Firmas → Gestionar firmas.
- Las plantillas de correo: Mi Cuenta → Plantillas de correo.
- Los recordatorios de correo: Mi Cuenta → Recordatorios.
- Las ausencias/vacaciones: Mi Cuenta → Ausencias y vacaciones.
- La traducción automática: Mi Cuenta → Traducción automática.
- El backup: Mi Cuenta → Backup y restauración.
- Las suscripciones: Mi Cuenta → Suscripciones.
- Los datos personales: Mi Cuenta → Datos personales.
- La apariencia (tema): Mi Cuenta → Apariencia.
- Los permisos y notificaciones: Mi Cuenta → Permisos y notificaciones.
- Los contactos: menú lateral del correo → Contactos.
- Las carpetas nuevas: menú lateral del correo → Nueva carpeta.
- El calendario: pestaña Calendario (abajo). Tiene dos vistas: Semana y Calendario (mes).
- Crear un evento: botón + en la barra superior del calendario.
- Compartir un evento por WhatsApp: abrir el evento y pulsar "Compartir por WhatsApp".
- Compartir un evento por email: abrir el evento y pulsar "Compartir por email".
- El modo confidencial: se activa al redactar un correo.
- El chat interno: menú lateral del correo → Chat.
- Los adjuntos del chat van a Cloudinary automáticamente.

EJEMPLO DE RESPUESTA CORRECTA (cuando preguntan cómo crear una regla para mover facturas):
"Ve a Mi Cuenta, entra en Reglas y filtros y pulsa Gestionar reglas. Ahí crea una nueva regla: como condición pon Asunto contiene Factura (o Invoice), y como acción elige Mover a la carpeta que quieras. Guarda y listo."

Sé conciso, claro y directo. Si no sabes algo, dilo claramente en vez de inventarlo.`;

const AI_RULE_SYSTEM_PROMPT = `Eres un asistente que convierte instrucciones en español a reglas de correo electrónico.
Debes devolver SIEMPRE un JSON válido, sin markdown, sin texto extra, sin comentarios.

FORMATO EXACTO:
{
  "name": "nombre corto descriptivo en español",
  "enabled": true,
  "matchAll": true,
  "stopProcessing": false,
  "conditions": [
    { "field": "...", "operator": "...", "value": "..." }
  ],
  "actions": [
    { "type": "...", "value": "..." }
  ]
}

CAMPOS VÁLIDOS (field):
- "from": remitente
- "to": destinatario
- "cc": en copia
- "subject": asunto
- "body": cuerpo del mensaje
- "hasAttachment": tiene adjuntos (operator "isTrue" o "isFalse", SIN value)
- "sizeKb": tamaño en KB (operator "greaterThan" o "lessThan", value numérico)

OPERADORES VÁLIDOS:
- Para texto: "contains", "notContains", "equals", "notEquals", "startsWith", "endsWith", "regex"
- Solo para hasAttachment: "isTrue", "isFalse"
- Solo para sizeKb: "greaterThan", "lessThan"

ACCIONES VÁLIDAS (type):
- "moveTo": mover a carpeta (REQUIERE value = nombre exacto de una carpeta disponible)
- "markRead": marcar como leído (SIN value)
- "markUnread": marcar como no leído (SIN value)
- "star": marcar como importante (SIN value)
- "unstar": quitar importante (SIN value)
- "deleteMessage": mover a papelera (SIN value)
- "markSpam": marcar como spam (SIN value)
- "forward": reenviar a un email (REQUIERE value = dirección de email)

REGLAS IMPORTANTES:
1. Si el usuario dice "todas/todos/y" → "matchAll": true
2. Si dice "alguno/algunos/o" → "matchAll": false
3. "moveTo" SOLO puede usar nombres de la lista de CARPETAS DISPONIBLES.
4. Máximo 3 condiciones y 3 acciones por regla.
5. El "name" debe ser descriptivo y corto (max 40 caracteres).
6. Para texto, usa minúsculas en "value" cuando aplique salvo nombres propios.
7. Si el prompt NO describe una regla clara, devuelve: {"error": "No he entendido qué regla quieres crear"}

EJEMPLOS:

Prompt: "mueve todo lo de facturas a la carpeta Facturas"
Respuesta:
{"name":"Facturas a Facturas","enabled":true,"matchAll":true,"stopProcessing":false,"conditions":[{"field":"subject","operator":"contains","value":"factura"}],"actions":[{"type":"moveTo","value":"Facturas"}]}

Prompt: "los correos de juan lopez márcalos como importantes y leídos"
Respuesta:
{"name":"Juan López importante","enabled":true,"matchAll":true,"stopProcessing":false,"conditions":[{"field":"from","operator":"contains","value":"juan lopez"}],"actions":[{"type":"star"},{"type":"markRead"}]}

Prompt: "si viene de noreply o tiene newsletter en el asunto, borrar"
Respuesta:
{"name":"Borrar newsletters","enabled":true,"matchAll":false,"stopProcessing":true,"conditions":[{"field":"from","operator":"contains","value":"noreply"},{"field":"subject","operator":"contains","value":"newsletter"}],"actions":[{"type":"deleteMessage"}]}

Prompt: "los correos con adjuntos que pesen más de 5 MB, avisarme reenviando a backup@miempresa.com"
Respuesta:
{"name":"Adjuntos grandes a backup","enabled":true,"matchAll":true,"stopProcessing":false,"conditions":[{"field":"hasAttachment","operator":"isTrue"},{"field":"sizeKb","operator":"greaterThan","value":"5000"}],"actions":[{"type":"forward","value":"backup@miempresa.com"}]}`;

let groqEnabled = false;
try {
  const groqKey = process.env.GROQ_API_KEY;
  if (groqKey) {
    groqEnabled = true;
    console.log('✅ Groq AI configurado');
  } else {
    console.log('⚠️ GROQ_API_KEY no configurada');
  }
} catch (_) { groqEnabled = false; }

async function callGroqChat({ history, prompt }) {
  const groqKey = process.env.GROQ_API_KEY;
  if (!groqKey) throw new Error('GROQ_API_KEY no configurada');
  const messages = [
    { role: 'system', content: RSMAIL_SYSTEM_PROMPT },
    ...history,
    { role: 'user', content: prompt },
  ];
  let lastError = null;
  for (const model of GROQ_MODELS_FALLBACK) {
    try {
      const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${groqKey}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          model, messages, temperature: 0.5, max_tokens: 1024,
        }),
      });
      if (!res.ok) {
        lastError = new Error(`Groq ${res.status}`);
        if (res.status === 404 || res.status === 400) continue;
        if (res.status === 429) throw lastError;
        continue;
      }
      const data = await res.json();
      let reply = data?.choices?.[0]?.message?.content || '(Sin respuesta)';
      reply = reply.replace(/\*\*/g, '').replace(/^#+\s*/gm, '').replace(/`([^`]+)`/g, '$1');
      return { reply, model };
    } catch (e) {
      lastError = e;
      continue;
    }
  }
  throw lastError || new Error('Todos los modelos fallaron');
}

app.post('/api/ai/chat', async (req, res) => {
  if (!groqEnabled) return res.status(503).json({ success: false, error: 'IA no disponible' });
  try {
    const { message, history } = req.body || {};
    if (!message || typeof message !== 'string') {
      return res.status(400).json({ success: false, error: 'message_required' });
    }
    const historyMessages = (Array.isArray(history) ? history : []).map((m) => ({
      role: m.role === 'user' ? 'user' : 'assistant',
      content: m.content || '',
    }));
    const result = await callGroqChat({
      history: historyMessages.slice(-6),
      prompt: message,
    });
    res.json({ success: true, reply: result.reply, model: result.model });
  } catch (e) {
    res.status(500).json({
      success: false,
      error: /429|rate limit/i.test(e.message)
        ? 'El asistente está saturado. Inténtalo en unos segundos.'
        : e.message,
    });
  }
});

// ------------------------------------------------------------
//  🔥 IA — Preferencias
// ------------------------------------------------------------
app.post('/api/ai/preferences', async (req, res) => {
  if (!db) return res.status(500).json({ success: false, error: 'Firestore no configurado' });
  try {
    const {
      email, chat, naturalLanguageRules, detectUrgency,
      urgencyLevel, autoClassify, categories,
    } = req.body || {};
    if (!email || typeof email !== 'string') {
      return res.status(400).json({ success: false, error: 'email_required' });
    }
    const safeCategories = Array.isArray(categories)
      ? categories.map((c) => String(c).trim().toLowerCase()).filter((c) => c.length > 0).slice(0, 40)
      : undefined;

    const payload = { updatedAt: admin.firestore.FieldValue.serverTimestamp() };
    if (typeof chat === 'boolean') payload.chat = chat;
    if (typeof naturalLanguageRules === 'boolean') payload.naturalLanguageRules = naturalLanguageRules;
    if (typeof detectUrgency === 'boolean') payload.detectUrgency = detectUrgency;
    if (['low', 'medium', 'high'].includes(urgencyLevel)) payload.urgencyLevel = urgencyLevel;
    if (typeof autoClassify === 'boolean') payload.autoClassify = autoClassify;
    if (safeCategories && safeCategories.length > 0) payload.categories = safeCategories;

    await db.collection('ai_preferences').doc(email).set(payload, { merge: true });
    invalidateAiPreferencesCache(email);
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

app.get('/api/ai/preferences/:email', async (req, res) => {
  if (!db) return res.status(500).json({ success: false, error: 'Firestore no configurado' });
  try {
    const email = decodeURIComponent(req.params.email);
    const prefs = await getAiPreferences(email);
    res.json({ success: true, preferences: prefs });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ------------------------------------------------------------
//  🔥 IA — Reglas con lenguaje natural
// ------------------------------------------------------------
async function callGroqForRuleJson(prompt, folders) {
  const groqKey = process.env.GROQ_API_KEY;
  if (!groqKey) throw new Error('GROQ_API_KEY no configurada');
  const foldersList = (folders || []).filter((f) => f && f !== 'INBOX');
  const userMessage = `CARPETAS DISPONIBLES: ${JSON.stringify(foldersList)}

PROMPT DEL USUARIO:
"${prompt}"

Devuelve SOLO el JSON de la regla (o {"error": "..."}).`;

  const messages = [
    { role: 'system', content: AI_RULE_SYSTEM_PROMPT },
    { role: 'user', content: userMessage },
  ];

  let lastError = null;
  for (const model of GROQ_MODELS_FALLBACK) {
    try {
      const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${groqKey}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          model, messages, temperature: 0.2, max_tokens: 800,
          response_format: { type: 'json_object' },
        }),
      });
      if (!res.ok) {
        lastError = new Error(`Groq ${res.status}`);
        if (res.status === 404 || res.status === 400) continue;
        if (res.status === 429) throw lastError;
        continue;
      }
      const data = await res.json();
      const raw = data?.choices?.[0]?.message?.content || '{}';
      try { return JSON.parse(raw); }
      catch (_) { lastError = new Error('JSON inválido'); continue; }
    } catch (e) {
      lastError = e;
      continue;
    }
  }
  throw lastError || new Error('Todos los modelos fallaron');
}

function validateRuleShape(rule) {
  if (!rule || typeof rule !== 'object') return 'not_object';
  if (rule.error) return null;
  const validFields = ['from', 'to', 'cc', 'subject', 'body', 'hasAttachment', 'sizeKb'];
  const validOperators = ['contains', 'notContains', 'equals', 'notEquals',
    'startsWith', 'endsWith', 'regex', 'isTrue', 'isFalse', 'greaterThan', 'lessThan'];
  const validActions = ['moveTo', 'markRead', 'markUnread', 'star', 'unstar',
    'deleteMessage', 'markSpam', 'forward'];

  if (typeof rule.name !== 'string' || rule.name.trim().length === 0) return 'name';
  if (!Array.isArray(rule.conditions) || rule.conditions.length === 0) return 'conditions';
  if (!Array.isArray(rule.actions) || rule.actions.length === 0) return 'actions';

  for (const c of rule.conditions) {
    if (!c || !validFields.includes(c.field)) return 'condition_field';
    if (!validOperators.includes(c.operator)) return 'condition_operator';
    if (c.field !== 'hasAttachment' && (typeof c.value !== 'string' && typeof c.value !== 'number')) {
      return 'condition_value';
    }
  }
  for (const a of rule.actions) {
    if (!a || !validActions.includes(a.type)) return 'action_type';
    if ((a.type === 'moveTo' || a.type === 'forward') && (!a.value || String(a.value).trim().length === 0)) {
      return 'action_value';
    }
  }
  return null;
}

app.post('/api/ai/rules/parse', async (req, res) => {
  if (!groqEnabled) return res.status(503).json({ success: false, error: 'ai_unavailable' });
  try {
    const { email, prompt, folders } = req.body || {};
    if (!email || typeof email !== 'string') {
      return res.status(400).json({ success: false, error: 'email_required' });
    }
    if (!prompt || typeof prompt !== 'string' || prompt.trim().length < 4) {
      return res.status(400).json({ success: false, error: 'prompt_too_short' });
    }
    if (prompt.length > 500) {
      return res.status(400).json({ success: false, error: 'prompt_too_long' });
    }

    const prefs = await getAiPreferences(email);
    if (!prefs.naturalLanguageRules) {
      return res.status(403).json({
        success: false,
        error: 'nl_rules_disabled',
        message: 'Activa primero "Reglas con lenguaje natural" en Mi Perfil → Asistente IA.',
      });
    }

    const parsed = await callGroqForRuleJson(prompt.trim(), folders || []);

    if (parsed && parsed.error) {
      return res.json({ success: false, error: 'ai_parse_failed', message: String(parsed.error) });
    }

    const invalid = validateRuleShape(parsed);
    if (invalid) {
      return res.json({
        success: false,
        error: 'ai_invalid_rule',
        message: 'La IA no ha generado una regla válida. Intenta reformular la frase.',
      });
    }

    const rule = {
      name: String(parsed.name).trim().substring(0, 60),
      enabled: parsed.enabled !== false,
      matchAll: parsed.matchAll !== false,
      stopProcessing: parsed.stopProcessing === true,
      conditions: parsed.conditions.slice(0, 3).map((c) => ({
        field: c.field,
        operator: c.operator,
        value: c.field === 'hasAttachment' ? '' : String(c.value ?? ''),
      })),
      actions: parsed.actions.slice(0, 3).map((a) => ({
        type: a.type,
        value: a.value != null ? String(a.value) : '',
      })),
    };

    res.json({ success: true, rule });
  } catch (e) {
    res.status(500).json({
      success: false,
      error: /429|rate limit/i.test(e.message)
        ? 'El asistente está saturado. Inténtalo en unos segundos.'
        : e.message,
    });
  }
});

// ------------------------------------------------------------
//  DIAGNÓSTICO
// ------------------------------------------------------------
app.get('/api/ai/models', async (req, res) => {
  try {
    const key = process.env.GROQ_API_KEY;
    if (!key) return res.status(503).json({ success: false, error: 'GROQ_API_KEY no configurada' });
    const r = await fetch('https://api.groq.com/openai/v1/models', {
      headers: { 'Authorization': `Bearer ${key}` },
    });
    const data = await r.json();
    const models = (data.data || []).map((m) => ({
      id: m.id, contextWindow: m.context_window,
      ownedBy: m.owned_by, active: m.active,
    }));
    res.json({ success: true, total: models.length, models });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ------------------------------------------------------------
//  INICIAR SERVIDOR
// ------------------------------------------------------------
const PORT = process.env.PORT || 3000;
server.listen(PORT, async () => {
  console.log(`✅ Backend RSMAIL activo en puerto ${PORT}`);
  console.log(`   Firestore: ${db ? 'OK' : 'NO DISPONIBLE'}`);
  console.log(`   Firebase Storage: ${storageBucket ? storageBucket.name : 'NO DISPONIBLE'}`);
  console.log(`   Cloudinary: ${CLOUDINARY_ENABLED ? 'OK' : 'NO'}`);
  console.log(`   Microsoft OAuth: ${MICROSOFT_CLIENT_ID ? 'OK' : '❌'}`);
  console.log(`   Groq AI: ${groqEnabled ? 'OK' : 'NO CONFIGURADO'}`);
  await restoreWorkers();
});
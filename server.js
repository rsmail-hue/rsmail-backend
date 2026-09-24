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
  console.log('⚠️ Cloudinary sin configurar → fallback a almacenamiento local');
}

// ------------------------------------------------------------
//  FIREBASE ADMIN (FCM) — inicialización defensiva
// ------------------------------------------------------------
let db = null;
try {
  let adminVersion = 'desconocida';
  try { adminVersion = require('firebase-admin/package.json').version; } catch (_) {}
  console.log('🔧 firebase-admin versión:', adminVersion);
  console.log('🔧 typeof admin.apps:', typeof admin.apps);

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
        });
        console.log('✅ Firebase Admin inicializado con variables de entorno');
      } catch (e) {
        console.error('❌ Error al inicializar Firebase con env vars:', e.message);
      }
    } else {
      try {
        const serviceAccount = require('./serviceAccountKey.json');
        admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
        console.log('✅ Firebase Admin inicializado con serviceAccountKey.json');
      } catch (e) {
        console.error('⚠️ Sin credenciales de Firebase. Push deshabilitadas.');
      }
    }
  } else {
    console.log('ℹ️ Firebase Admin ya estaba inicializado');
  }

  if (admin && Array.isArray(admin.apps) && admin.apps.length > 0) {
    db = admin.firestore();
    console.log('✅ Firestore listo');
  } else {
    console.warn('⚠️ Firestore NO disponible.');
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
//  Microsoft OAuth2 — Helpers
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
  } catch (e) {
    console.error(`⚠️ Error guardando cuenta ${email}:`, e.message);
  }
}

async function restoreWorkers() {
  if (!db) return;
  try {
    const snapshot = await db.collection('user_accounts').get();
    console.log(`🔄 Restaurando ${snapshot.size} worker(s) desde Firestore...`);
    for (const doc of snapshot.docs) {
      const data = doc.data();
      if (data.email && data.password) {
        console.log(`  → Arrancando worker para ${data.email}`);
        startImapWorker(data.email, data.password, data.imapHost);
      }
    }
    console.log(`✅ Restauración de workers completada`);
  } catch (e) {
    console.error('⚠️ Error restaurando workers:', e.message);
  }
}

async function getSavedLastUid(email) {
  if (!db) return null;
  try {
    const doc = await db.collection('user_states').doc(email).get();
    if (doc.exists) return doc.data().lastUid || null;
  } catch (e) {
    console.error(`⚠️ Error al leer lastUid para ${email}:`, e.message);
  }
  return null;
}

async function saveLastUid(email, uid) {
  if (!db) return;
  try {
    await db.collection('user_states').doc(email).set({
      lastUid: uid,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    }, { merge: true });
  } catch (e) {
    console.error(`⚠️ Error al guardar lastUid para ${email}:`, e.message);
  }
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
      console.log(`🔌 Probando IMAP ${email} → ${c.host}:${c.port}${accessToken ? ' (OAuth2)' : ''}...`);
      const client = await connectImap(email, password, c.host, c.port, c.secure, accessToken);
      console.log(`✅ IMAP conectado: ${email} vía ${c.host}:${c.port}`);
      return { client, host: c.host, port: c.port };
    } catch (e) {
      console.log(`⚠️ IMAP ${c.host}:${c.port} falló: ${e.message}`);
      lastError = e;
    }
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
//  BREVO (fallback)
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
//  AUSENCIAS / VACACIONES
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
  } catch (e) {
    console.error('⚠️ Error leyendo absences_configs:', e.message);
    return null;
  }
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
    let lastError = null;
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
        console.log(`✅ Auto-reply enviado vía ${c.host}:${c.port}`);
        break;
      } catch (e) { lastError = e; }
    }
    if (!sent && process.env.BREVO_API_KEY) {
      try {
        await sendViaBrevo({ fromEmail: accountEmail, fromName: '', to: incomingFrom, subject: autoReplySubject, html });
        sent = true;
      } catch (e) { lastError = e; }
    }
    if (sent) {
      await sentRef.set({
        accountEmail, periodId: period.id, periodTitle: period.title,
        senderEmail: fromLower, originalSubject: incomingSubject,
        originalUid: incomingUid,
        sentAt: admin.firestore.FieldValue.serverTimestamp(),
      });
    }
  } catch (e) {
    console.error('❌ Error en checkAndSendAutoReply:', e.message);
  }
}

// ------------------------------------------------------------
//  MOTOR DE REGLAS / FILTROS
// ------------------------------------------------------------
const _rulesCache = new Map();
const RULES_CACHE_TTL_MS = 5 * 60 * 1000;

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
  } catch (e) {
    console.error('⚠️ Error cargando reglas:', e.message);
    return [];
  }
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
    console.log(`⚠️ Error ejecutando acciones: ${e.message}`);
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
  } catch (e) {
    console.error('❌ /api/translate error:', e.message);
    res.status(500).json({ error: e.message });
  }
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
              unique_filename: true, overwrite: false, access_mode: 'public',
            },
            (error, result) => { if (error) reject(error); else resolve(result); }
          );
          stream.end(req.file.buffer);
        });
        return res.json({
          success: true, url: uploadResult.secure_url,
          filename: originalName, size, provider: 'cloudinary',
          publicId: uploadResult.public_id,
        });
      } catch (cloudErr) {
        console.error('❌ Cloudinary falló:', cloudErr.message);
      }
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
  } catch (e) {
    console.error('❌ Error en /api/chat/upload:', e.message);
    res.status(500).json({ success: false, error: e.message });
  }
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
  console.log('🔌 Nuevo WebSocket conectado');
  let userEmail = null;
  ws.on('message', async (message) => {
    try {
      const data = JSON.parse(message);
      if (data.type === 'login') {
        userEmail = data.email;
        console.log(`📧 Login WebSocket para ${userEmail}`);
        if (activeWorkers.has(userEmail)) {
          activeWorkers.get(userEmail).ws = ws;
        } else {
          startImapWorker(data.email, data.password, data.imapHost, ws);
        }
      }
    } catch (e) { console.error('❌ Error WebSocket:', e.message); }
  });
  ws.on('close', () => {
    console.log(`🔌 WebSocket desconectado (${userEmail || '?'})`);
    if (userEmail && activeWorkers.has(userEmail)) {
      activeWorkers.get(userEmail).ws = null;
    }
  });
});

// ------------------------------------------------------------
//  MOTOR IMAP PERSISTENTE CON IDLE
// ------------------------------------------------------------
function startImapWorker(email, password, customHost, ws = null) {
  if (activeWorkers.has(email)) {
    const existing = activeWorkers.get(email);
    if (ws) existing.ws = ws;
    return;
  }
  if (!password) { console.error(`❌ startImapWorker: SIN PASSWORD para ${email}`); return; }
  if (isAccountBlocked(email)) { console.log(`🚫 Worker ${email} en cooldown, no se arranca`); return; }

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

        if (state.ws && state.ws.readyState === WebSocket.OPEN) {
          state.ws.send(JSON.stringify({
            type: 'new_email', email: state.email,
            timestamp: new Date().toISOString(),
            from, subject, uid: msg.uid,
          }));
        }

        console.log(`📤 Enviando push (data-only) para nuevo correo UID:${msg.uid}...`);
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
      console.log(`🔄 [BG] Conectando IMAP para ${state.email}...`);
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
            console.log(`⚡ IDLE: nuevo correo detectado. Rango ${state.lastUidNext} a ${currentUidNext - 1}`);
            await processEmailsInRange(state, state.lastUidNext, currentUidNext);
          }
        } catch (e) { console.error(`❌ Error en handleExists:`, e.message); }
      };

      state.client.on('exists', handleExists);

      const statusInit = await state.client.status('INBOX', { uidNext: true });
      const currentUidNext = statusInit.uidNext || 1;
      const savedUid = await getSavedLastUid(state.email);

      if (savedUid && savedUid > 0 && savedUid < currentUidNext) {
        console.log(`🔎 Recuperando correos entre UID ${savedUid} y ${currentUidNext - 1}...`);
        await processEmailsInRange(state, savedUid, currentUidNext);
      } else {
        state.lastUidNext = currentUidNext;
        await saveLastUid(state.email, currentUidNext);
      }

      console.log(`💤 Monitor IDLE activo para ${state.email}. Próximo UID: ${state.lastUidNext}`);

      let aliveLogCounter = 0;
      while (state.active && state.client.usable) {
        await new Promise(resolve => setTimeout(resolve, 5000));
        aliveLogCounter++;
        if (aliveLogCounter % 6 === 0) {
          state.lastAlive = Date.now();
          console.log(`💓 [${state.email}] IDLE activo (${Math.floor(aliveLogCounter * 5)}s)`);
        }
        if (aliveLogCounter % 12 === 0) {
          try {
            const s = await state.client.status('INBOX', { uidNext: true });
            if (s.uidNext && s.uidNext > state.lastUidNext) {
              console.log(`🔄 Fallback poll: nuevo correo detectado`);
              await processEmailsInRange(state, state.lastUidNext, s.uidNext);
            }
          } catch (e) {
            console.log(`⚠️ Fallback poll falló, cerrando conexión: ${e.message}`);
            break;
          }
        }
      }
    } catch (e) {
      if (state.active) {
        console.log(`⚠️ IMAP caído para ${state.email}: ${e.message}`);
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
        console.log(`🚫 Worker ${state.email} en cooldown. Esperando 30 min...`);
        await new Promise(r => setTimeout(r, 30 * 60 * 1000));
      } else {
        console.log(`⏳ Reconectando IMAP ${state.email} en 10s...`);
        await new Promise(r => setTimeout(r, 10000));
      }
    }
  }
}

// ------------------------------------------------------------
//  CRON: CALENDARIO (data-only)
// ------------------------------------------------------------
cron.schedule('* * * * *', async () => {
  if (!db) return;
  try {
    const now = new Date();
    const snapshot = await db.collection('calendar_events').get();
    for (const doc of snapshot.docs) {
      const event = doc.data();
      const rawTime = event.eventTime;
      const eventDate = rawTime?.toDate ? rawTime.toDate() : new Date(rawTime);
      if (isNaN(eventDate.getTime())) continue;
      const diffMs = eventDate.getTime() - now.getTime();
      const diffHours = diffMs / (1000 * 60 * 60);
      const formattedTime = eventDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      const recipientEmail = event.email || (Array.isArray(event.sharedEmails) && event.sharedEmails.length > 0 ? event.sharedEmails[0] : null);
      if (!recipientEmail) continue;

      if (!event.notified1Day && diffHours <= 25 && diffHours > 23) {
        console.log(`📅 Recordatorio 1 DÍA ANTES: ${event.title} → ${recipientEmail}`);
        await sendPushNotification(recipientEmail, {
          title: `📅 Mañana: ${event.title}`,
          body: `Tienes este evento programado para mañana a las ${formattedTime}`,
          data: {
            type: 'calendar_event',
            eventId: doc.id,
            eventTitle: event.title,
            eventTime: eventDate.toISOString(),
            eventDate: formattedTime,
            notice: '1day',
            text: `Tienes este evento programado para mañana a las ${formattedTime}`,
          },
        }, { dataOnly: true });
        await doc.ref.update({ notified1Day: true });
      }
      if (!event.notifiedEvent && diffMs <= 15 * 60 * 1000 && diffMs > 0) {
        console.log(`⏰ Recordatorio 15 MIN ANTES: ${event.title} → ${recipientEmail}`);
        await sendPushNotification(recipientEmail, {
          title: `⏰ Comienza pronto: ${event.title}`,
          body: event.description || `El evento comienza a las ${formattedTime}`,
          data: {
            type: 'calendar_event',
            eventId: doc.id,
            eventTitle: event.title,
            eventTime: eventDate.toISOString(),
            eventDate: formattedTime,
            notice: '15min',
            text: event.description || `El evento comienza a las ${formattedTime}`,
          },
        }, { dataOnly: true });
        await doc.ref.update({ notifiedEvent: true });
      }
    }
  } catch (e) { console.error('❌ Error en Cron Job:', e.message); }
});

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
        console.log(`♻️ Cron: reanimando worker caído para ${data.email}`);
        if (state) state.active = false;
        activeWorkers.delete(data.email);
        startImapWorker(data.email, data.password, data.imapHost);
      }
    }
  } catch (e) { console.error('⚠️ Error en cron de reanimación:', e.message); }
});

// ------------------------------------------------------------
//  CRON: RECORDATORIOS DE CORREO
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
    if (toSend.length === 0) return;
    console.log(`🔔 Procesando ${toSend.length} recordatorio(s)...`);
    for (const doc of toSend) {
      const data = doc.data();
      const docRef = doc.ref;
      try {
        const accountEmail = data.accountEmail;
        if (!accountEmail) { await docRef.update({ status: 'failed', error: 'Sin cuenta asociada' }); continue; }
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
            subject, sender: from,
            text: body,
          },
        }, { dataOnly: true });
        await docRef.update({ status: 'sent', sentAt: admin.firestore.FieldValue.serverTimestamp() });
        console.log(`✅ Recordatorio enviado: "${subject}"`);
      } catch (e) {
        console.error(`❌ Error recordatorio ${doc.id}:`, e.message);
        await docRef.update({ status: 'failed', error: e.message });
      }
    }
  } catch (e) { console.error('❌ Error en cron de recordatorios:', e.message); }
});

// ------------------------------------------------------------
//  CRON: LIMPIAR RESPUESTAS VACACIONES ANTIGUAS
// ------------------------------------------------------------
cron.schedule('0 */6 * * *', async () => {
  if (!db) return;
  try {
    const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 30);
    const snap = await db.collection('vacation_sent_replies').where('sentAt', '<', admin.firestore.Timestamp.fromDate(cutoff)).get();
    let deleted = 0;
    for (const doc of snap.docs) { await doc.ref.delete(); deleted++; }
    if (deleted > 0) console.log(`🗑️ Limpiados ${deleted} registros de auto-reply antiguos`);
  } catch (e) { console.error('⚠️ Error limpiando auto-replies:', e.message); }
});

// ------------------------------------------------------------
//  FCM PUSH — con soporte data-only
// ------------------------------------------------------------
async function sendPushNotification(email, payload, options = {}) {
  if (!db) return;
  const dataOnly = options.dataOnly === true;
  try {
    const tokensSnapshot = await db.collection('fcm_tokens').where('email', '==', email).get();
    if (tokensSnapshot.empty) { console.log(`📴 Sin token FCM para ${email}`); return; }
    const tokens = [];
    tokensSnapshot.forEach(doc => tokens.push(doc.data().token));
    console.log(`🚀 Enviando push${dataOnly ? ' (data-only)' : ''} a ${tokens.length} token(s) para ${email}`);

    const rawData = payload.data || { type: 'general' };
    const safeData = {};
    Object.keys(rawData).forEach((k) => {
      safeData[k] = rawData[k] == null ? '' : String(rawData[k]);
    });

    const message = {
      data: safeData,
      android: {
        priority: 'high',
        ttl: 60 * 60 * 1000,
      },
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
    console.log(`✅ Push enviado | éxito: ${response.successCount}, fallos: ${response.failureCount}`);
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
  } catch (e) { console.error('❌ Error creando confidencial:', e.message); res.status(500).json({ success: false, error: e.message }); }
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
    if (expiresAt < new Date()) return res.status(410).json({ success: false, error: 'Este enlace ha caducado' });
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
  res.send(`<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>RSMail · Mensaje confidencial</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; -webkit-user-select: none; user-select: none; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: linear-gradient(135deg, #0D47A1 0%, #1A73E8 100%); min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 16px; color: #333; }
  .container { background: #fff; border-radius: 16px; max-width: 720px; width: 100%; box-shadow: 0 10px 40px rgba(0,0,0,0.25); overflow: hidden; }
  .header { background: #1A73E8; color: #fff; padding: 20px; display: flex; align-items: center; gap: 12px; }
  .header h1 { font-size: 18px; font-weight: 600; }
  .header .lock { font-size: 24px; }
  .banner { background: #FFF3CD; border-left: 4px solid #FFC107; padding: 12px 20px; font-size: 13px; color: #856404; }
  .content { padding: 24px; }
  .subject { font-size: 20px; font-weight: 700; margin-bottom: 8px; color: #111; }
  .meta { font-size: 12px; color: #888; margin-bottom: 20px; }
  .body { font-size: 15px; line-height: 1.6; color: #222; white-space: pre-wrap; word-wrap: break-word; }
  .footer { padding: 16px 24px; background: #f5f5f5; font-size: 11px; color: #999; text-align: center; border-top: 1px solid #e0e0e0; }
  .password-box { padding: 40px 24px; text-align: center; }
  .password-box h2 { font-size: 18px; margin-bottom: 16px; }
  .password-box input { width: 100%; max-width: 300px; padding: 12px 16px; border-radius: 8px; border: 2px solid #1A73E8; font-size: 15px; outline: none; margin-bottom: 12px; -webkit-user-select: text; user-select: text; }
  .password-box button { background: #1A73E8; color: #fff; border: none; padding: 12px 32px; border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer; transition: background 0.2s; }
  .password-box button:hover { background: #0D47A1; }
  .error { padding: 40px 24px; text-align: center; color: #c62828; }
  .error h2 { font-size: 20px; margin-bottom: 8px; }
  .error p { font-size: 14px; color: #666; }
  .loading { padding: 40px; text-align: center; color: #1A73E8; font-size: 14px; }
  .spinner { width: 40px; height: 40px; border: 4px solid #e0e0e0; border-top-color: #1A73E8; border-radius: 50%; animation: spin 0.8s linear infinite; margin: 0 auto 16px; }
  @keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>
<div class="container" id="app">
  <div class="loading"><div class="spinner"></div>Cargando mensaje confidencial...</div>
</div>
<script>
const CONF_ID = '${id}';
document.addEventListener('contextmenu', e => e.preventDefault());
document.addEventListener('copy', e => e.preventDefault());
document.addEventListener('cut', e => e.preventDefault());
document.addEventListener('selectstart', e => e.preventDefault());
document.addEventListener('dragstart', e => e.preventDefault());
document.addEventListener('keydown', e => {
  if (e.ctrlKey && ['c','x','s','p','u','a'].includes(e.key.toLowerCase())) e.preventDefault();
  if (e.ctrlKey && e.shiftKey && ['i','j','c'].includes(e.key.toLowerCase())) e.preventDefault();
  if (e.key === 'F12') e.preventDefault();
  if (e.key === 'PrintScreen') { navigator.clipboard.writeText('RSMail · Contenido protegido'); e.preventDefault(); }
});
async function loadContent(password) {
  const body = { password: password || null };
  const res = await fetch('/api/confidential/open/' + CONF_ID, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  const data = await res.json();
  return { status: res.status, data };
}
async function init(password) {
  const app = document.getElementById('app');
  try {
    const { status, data } = await loadContent(password);
    if (status === 401) {
      app.innerHTML = '<div class="header"><span class="lock">🔒</span><h1>Mensaje confidencial</h1></div><div class="banner">Este mensaje está protegido. Introduce la contraseña para abrirlo.</div><div class="password-box"><h2>Introduce la contraseña</h2><input type="password" id="pwd" placeholder="Contraseña" autofocus><div><button onclick="submitPwd()">Abrir</button></div><p id="err" style="color:#c62828;margin-top:12px;font-size:13px;"></p></div><div class="footer">🔒 Enviado de forma segura con RSMail</div>';
      document.getElementById('pwd').addEventListener('keydown', e => { if (e.key === 'Enter') submitPwd(); });
      window.submitPwd = () => { const p = document.getElementById('pwd').value; if (!p) return; init(p); };
      return;
    }
    if (!data.success) {
      app.innerHTML = '<div class="header"><span class="lock">🔒</span><h1>Mensaje confidencial</h1></div><div class="error"><h2>' + (status === 410 ? 'Enlace caducado' : 'No disponible') + '</h2><p>' + (data.error || 'Este enlace ya no es válido') + '</p></div><div class="footer">RSMail</div>';
      return;
    }
    const exp = new Date(data.expiresAt).toLocaleDateString('es-ES');
    const safeBody = (data.body || '').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    app.innerHTML = '<div class="header"><span class="lock">🔒</span><h1>Mensaje confidencial</h1></div><div class="banner">Este mensaje es confidencial. Caduca el ' + exp + '.</div><div class="content"><div class="subject">' + (data.subject || '(Sin asunto)').replace(/</g,'&lt;') + '</div><div class="meta">De: ' + (data.from || '').replace(/</g,'&lt;') + ' · Para: ' + (data.to || '').replace(/</g,'&lt;') + '</div><div class="body">' + safeBody + '</div></div><div class="footer">🔒 Enviado de forma segura con RSMail · Caduca el ' + exp + '</div>';
  } catch (e) {
    app.innerHTML = '<div class="error"><h2>Error de conexión</h2><p>' + e.message + '</p></div>';
  }
}
init();
</script>
</body>
</html>`);
});

cron.schedule('*/30 * * * *', async () => {
  if (!db) return;
  try {
    const snap = await db.collection('confidential_emails').get();
    let deleted = 0;
    const now = Date.now();
    for (const doc of snap.docs) {
      const data = doc.data();
      const exp = data.expiresAt?.toDate ? data.expiresAt.toDate() : new Date(data.expiresAt);
      if (now - exp.getTime() > 30 * 24 * 60 * 60 * 1000) { await doc.ref.delete(); deleted++; }
    }
    if (deleted > 0) console.log(`🗑️ Eliminados ${deleted} confidenciales antiguos`);
  } catch (e) { console.error('❌ Error limpiando confidenciales:', e.message); }
});

// ------------------------------------------------------------
//  AUTH
// ------------------------------------------------------------
const handleAuth = async (req, res) => {
  const { email, password, host, port } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, error: 'Email y contraseña requeridos' });
  try {
    console.log(`🔐 Verificando cuenta ${email}...`);
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
    console.error(`❌ Auth falló para ${email}:`, e.message);
    return res.status(401).json({ success: false, error: 'Credenciales inválidas: ' + e.message });
  }
};

// ------------------------------------------------------------
//  MICROSOFT LOGIN
// ------------------------------------------------------------
app.post('/api/microsoft/login', async (req, res) => {
  const { email, refreshToken } = req.body;
  if (!email || !refreshToken) return res.status(400).json({ success: false, error: 'Email y refreshToken requeridos' });
  if (!MICROSOFT_CLIENT_ID) return res.status(500).json({ success: false, error: 'Servidor sin MICROSOFT_CLIENT_ID configurado' });
  try {
    console.log(`🔐 Microsoft login para ${email}...`);
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
    console.error(`❌ Microsoft login falló para ${email}:`, e.message);
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
      const deadEmails = [];
      for (const [email, state] of activeWorkers.entries()) {
        const isAlive = state.active && state.client && state.client.usable;
        if (!isAlive) deadEmails.push(email);
      }
      for (const email of deadEmails) {
        console.log(`♻️ /ping: reanimando worker caído para ${email}`);
        const state = activeWorkers.get(email);
        if (state) state.active = false;
        activeWorkers.delete(email);
        if (isAccountBlocked(email)) continue;
        const doc = await db.collection('user_accounts').doc(email).get();
        if (doc.exists) {
          const data = doc.data();
          if (data.email && data.password) startImapWorker(data.email, data.password, data.imapHost);
        }
      }
      if (activeWorkers.size === 0) {
        const snapshot = await db.collection('user_accounts').get();
        if (snapshot.size > 0) {
          console.log(`♻️ /ping: sin workers, arrancando ${snapshot.size}`);
          for (const d of snapshot.docs) {
            const data = d.data();
            if (data.email && data.password && !isAccountBlocked(data.email)) startImapWorker(data.email, data.password, data.imapHost);
          }
        }
      }
    } catch (e) { console.error('⚠️ Error en /ping:', e.message); }
  }
  res.json({ alive: true, ts: new Date().toISOString(), workers: activeWorkers.size, firestore: !!db, blockedAccounts: Array.from(failedAccounts.keys()) });
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
    blockedAccounts: Array.from(failedAccounts.entries()).map(([e, v]) => ({ email: e, count: v.count, until: v.until ? new Date(v.until).toISOString() : null })),
  });
});

app.get('/api/debug/auto-config/:email', (req, res) => {
  const email = req.params.email;
  const auto = getAutoConfig(email);
  const imapCandidates = getImapCandidates(email);
  const smtpCandidates = getSmtpCandidates(email);
  res.json({
    email, provider: auto?.provider || 'unknown',
    imap: { host: auto?.imapHost, port: auto?.imapPort },
    smtp: { host: auto?.smtpHost, port: auto?.smtpPort, secure: auto?.smtpSecure },
    imapCandidates, smtpCandidates,
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

app.get('/api/debug/translate/:email', async (req, res) => {
  if (!db) return res.status(500).json({ error: 'Firestore no configurado' });
  try {
    const doc = await db.collection('translation_configs').doc(req.params.email).get();
    res.json(doc.exists ? doc.data() : { enabled: false });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/debug/reminders', async (req, res) => {
  if (!db) return res.status(500).json({ error: 'Firestore no configurado' });
  try {
    const snap = await db.collection('email_reminders').get();
    const items = [];
    snap.forEach((doc) => {
      const d = doc.data();
      const ra = d.remindAt;
      const raDate = ra?.toDate ? ra.toDate() : (ra ? new Date(ra) : null);
      items.push({
        id: doc.id, accountEmail: d.accountEmail, subject: d.emailSubject,
        from: d.emailFrom, note: d.note, status: d.status,
        remindAt: raDate ? raDate.toISOString() : null,
        overdue: raDate ? (raDate.getTime() <= Date.now()) : false,
        error: d.error || null,
      });
    });
    res.json({ count: items.length, now: new Date().toISOString(), items });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/fcm-token', async (req, res) => {
  const { email, token, password, imapHost } = req.body;
  if (!email || !token) return res.status(400).json({ success: false, error: 'Email y token requeridos' });
  if (!db) return res.status(500).json({ success: false, error: 'Firestore no configurado' });
  try {
    const existing = await db.collection('fcm_tokens').where('email', '==', email).get();
    existing.forEach(doc => doc.ref.delete());
    await db.collection('fcm_tokens').add({
      email, token,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    console.log(`📱 Token FCM actualizado en Firestore para ${email}`);
    if (password) {
      console.log(`🔧 fcm-token: iniciando worker para ${email}`);
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

// ------------------------------------------------------------
//  CHAT — Notificar mensaje nuevo (data-only)
// ------------------------------------------------------------
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
          groupName: groupName || '',
          text: preview,
        },
      }, { dataOnly: true });
      notified.push(email);
    } catch (e) { console.log(`⚠️ No se pudo notificar a ${email}:`, e.message); }
  }
  res.json({ success: true, notified });
});

// ------------------------------------------------------------
//  SEND EMAIL (SMTP directo)
// ------------------------------------------------------------
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
    console.error('❌ Error /api/send-email:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ------------------------------------------------------------
//  SAVE TO SENT
// ------------------------------------------------------------
app.post('/api/save-to-sent', async (req, res) => {
  const { email, password, host, port, to, cc, bcc, subject, body } = req.body;
  if (!email || !password) {
    return res.status(400).json({ success: false, error: 'Parámetros insuficientes' });
  }

  let client;
  try {
    console.log(`💾 [save-to-sent] Guardando correo de ${email}...`);
    const conn = await connectImapAuto(email, password, host);
    client = conn.client;

    const list = await client.list();
    const folderNames = list.map(f => `${f.path} (name=${f.name}, su=${f.specialUse || '-'})`);
    console.log(`📁 [save-to-sent] Carpetas de ${email}: ${folderNames.join(' | ')}`);

    const candidates = list.filter(f => {
      const name = (f.name || '').toLowerCase();
      const path = (f.path || '').toLowerCase();
      const su = (f.specialUse || '').toLowerCase();
      return (
        su === '\\sent' ||
        name === 'sent' ||
        name === 'sent items' ||
        name === 'sent messages' ||
        name === 'enviados' ||
        name === 'elementos enviados' ||
        name.includes('sent') ||
        name.includes('enviad') ||
        path.includes('sent') ||
        path.includes('enviad')
      );
    });

    let sentFolder = candidates[0]?.path;
    if (sentFolder) {
      console.log(`✅ [save-to-sent] Carpeta elegida: "${sentFolder}"`);
    } else {
      sentFolder = 'Sent';
      console.log(`⚠️ [save-to-sent] No se encontró carpeta Enviados, probando "${sentFolder}"`);
    }

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
      '',
      body || '',
    ].join('\r\n');

    let saved = false;
    try {
      await client.append(sentFolder, Buffer.from(rawEmail, 'utf8'), ['\\Seen']);
      saved = true;
      console.log(`✅ [save-to-sent] Guardado en "${sentFolder}"`);
    } catch (e) {
      console.log(`⚠️ [save-to-sent] APPEND a "${sentFolder}" falló: ${e.message}`);
    }

    if (!saved) {
      const alternatives = [
        'Sent', 'Sent Items', 'Sent Messages',
        'Enviados', 'Elementos enviados',
        'INBOX.Sent', 'INBOX.Sent Items',
        '[Gmail]/Sent Mail', 'INBOX.Enviados',
      ];
      for (const alt of alternatives) {
        if (alt === sentFolder) continue;
        try {
          await client.append(alt, Buffer.from(rawEmail, 'utf8'), ['\\Seen']);
          saved = true;
          console.log(`✅ [save-to-sent] Guardado en "${alt}" (fallback)`);
          break;
        } catch (_) {}
      }
    }

    await client.logout();
    client = null;

    if (!saved) {
      console.error(`❌ [save-to-sent] No se pudo guardar en ninguna carpeta`);
      return res.status(500).json({
        success: false,
        error: 'No se pudo guardar en ninguna carpeta Enviados',
        triedFolder: sentFolder,
      });
    }

    res.json({ success: true, message: 'Guardado en Enviados', folder: sentFolder });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    console.error('❌ [save-to-sent] Error:', e.message);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ------------------------------------------------------------
//  FOLDERS
// ------------------------------------------------------------
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
        messages.push({
          uid: msg.uid, id: msg.uid.toString(),
          subject: msg.envelope?.subject || '(Sin asunto)',
          from: msg.envelope?.from?.[0]?.address || msg.envelope?.from?.[0]?.name || '',
          to: msg.envelope?.to?.[0]?.address || '',
          date: msg.envelope?.date ? new Date(msg.envelope.date).toISOString() : new Date().toISOString(),
          hasAttachments, flags,
          isRead: flags.includes('\\Seen'),
          isFlagged: flags.includes('\\Flagged'),
        });
      }
    } finally { lock.release(); }
    await client.logout();
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
      },
    });
  } catch (err) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/delete-message', async (req, res) => {
  const { email, password, host, port, uid, folder = 'INBOX' } = req.body;
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
  const { email, password, host, port, uid, folder = 'INBOX', read } = req.body;
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
  const { email, password, host, port, folder = 'INBOX' } = req.body;
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
  const { email, password, host, port, uid, folder = 'INBOX', flagged } = req.body;
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
//  CLOUD — ADJUNTOS
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
  const { email, password, host, port, folder = 'INBOX', limit = 200 } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, error: 'Email y contraseña requeridos', attachments: [] });
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
      console.log(`☁️ Escaneando adjuntos en ${email}: ${total} msgs, desde ${startSeq}`);
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
  const { email, password, host, port, folder = 'INBOX', uid, partId, filename } = req.body;
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
    if (!parsed) return res.status(404).json({ success: false, error: 'Correo no encontrado' });
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
    if (!target) {
      return res.status(404).json({
        success: false, error: 'Adjunto no encontrado',
        availableFilenames: attachments.map((a) => a.filename).filter(Boolean),
      });
    }
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

// ------------------------------------------------------------
//  SUSCRIPCIONES
// ------------------------------------------------------------
function analyzeIsSubscription({ from, subject, listUnsubscribe }) {
  if (listUnsubscribe && listUnsubscribe.trim().length > 0) return true;
  const text = `${from || ''} ${subject || ''}`.toLowerCase();
  const keywords = [
    'newsletter', 'boletin', 'boletín', 'suscripción', 'suscripcion',
    'newsletter@', 'marketing@', 'info@', 'noreply@', 'no-reply@',
    'no responder', 'no-responder', 'promociones', 'publicidad',
  ];
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
  const { email, password, host, port, maxMessages = 500 } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, error: 'Email y contraseña requeridos' });
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
      console.log(`📬 Escaneando suscripciones en ${email}: ${total} mensajes`);
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
  if (!listUnsubscribe) return res.status(400).json({ success: false, error: 'Falta la cabecera List-Unsubscribe' });
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
      } catch (e) { console.log('⚠️ POST one-click falló:', e.message); }
    }
    if (!result && parsed.http.length > 0) {
      try {
        const r = await fetch(parsed.http[0], { method: 'GET', headers: { 'User-Agent': 'RSMail/3.0' }, redirect: 'follow' });
        method = 'https-get';
        result = { status: r.status, ok: r.ok };
      } catch (e) { console.log('⚠️ GET falló:', e.message); }
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
    if (!result) return res.status(400).json({ success: false, error: 'No se encontró un método de baja válido' });
    res.json({ success: true, method, result });
  } catch (e) {
    console.error('❌ Error en /api/unsubscribe:', e.message);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ------------------------------------------------------------
//  ELIMINAR CUENTA
// ------------------------------------------------------------
app.delete('/api/account/:email', async (req, res) => {
  const email = decodeURIComponent(req.params.email);
  if (!email) return res.status(400).json({ success: false, error: 'Email requerido' });
  try {
    console.log(`🗑️ Borrando cuenta ${email} del backend...`);
    if (activeWorkers.has(email)) {
      const state = activeWorkers.get(email);
      state.active = false;
      activeWorkers.delete(email);
      console.log(`   ✓ Worker detenido`);
    }
    if (db) {
      await db.collection('user_accounts').doc(email).delete().catch(() => {});
      await db.collection('user_states').doc(email).delete().catch(() => {});
      console.log(`   ✓ Borrado de Firestore`);
    }
    failedAccounts.delete(email);
    res.json({ success: true, message: 'Cuenta eliminada del backend' });
  } catch (e) {
    console.error('❌ Error borrando cuenta:', e.message);
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
  console.log(`   Microsoft OAuth client_id: ${MICROSOFT_CLIENT_ID ? MICROSOFT_CLIENT_ID.substring(0, 12) + '...' : '❌ NO CONFIGURADO'}`);
  await restoreWorkers();
});
﻿const express = require('express');
const cors = require('cors');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Imap = require('imap');
const nodemailer = require('nodemailer');
const WebSocket = require('ws');
const http = require('http');
const admin = require('firebase-admin');
const cron = require('node-cron');

// ------------------------------------------------------------
//  FIREBASE ADMIN (FCM)
// ------------------------------------------------------------
if (!admin.apps.length) {
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
      console.error('❌ Error al inicializar Firebase:', e.message);
    }
  } else {
    try {
      const serviceAccount = require('./serviceAccountKey.json');
      admin.initializeApp({
        credential: admin.credential.cert(serviceAccount),
      });
      console.log('✅ Firebase Admin inicializado con serviceAccountKey.json');
    } catch (e) {
      console.error('⚠️ Sin credenciales de Firebase. Push deshabilitadas.');
    }
  }
}

const db = admin.apps.length ? admin.firestore() : null;

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const server = http.createServer(app);
const wss = new WebSocket.Server({ server });

// ------------------------------------------------------------
//  WORKERS IMAP PERSISTENTES (con IDLE)
// ------------------------------------------------------------
const activeWorkers = new Map();

// ------------------------------------------------------------
//  PERSISTENCIA DE CUENTAS (auto-arranque de workers)
// ------------------------------------------------------------
async function saveAccount(email, password, imapHost) {
  if (!db) return;
  try {
    await db.collection('user_accounts').doc(email).set({
      email,
      password,
      imapHost,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    }, { merge: true });
  } catch (e) {
    console.error(`⚠️ Error guardando cuenta ${email}:`, e.message);
  }
}

async function restoreWorkers() {
  if (!db) {
    console.log('⚠️ Sin Firestore, no se restauran workers');
    return;
  }
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
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
  } catch (e) {
    console.error(`⚠️ Error al guardar lastUid para ${email}:`, e.message);
  }
}

async function connectImap(email, password, host, port, secure) {
  const config = {
    host,
    port,
    secure,
    auth: { user: email, pass: password },
    logger: false,
    tls: {
      rejectUnauthorized: false,
      minVersion: 'TLSv1.2',
      servername: host
    },
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
          console.log(`ℹ️ Worker ya existente para ${userEmail}, ws actualizado`);
        } else {
          console.log(`🔧 WS login: creando worker para ${userEmail} (host=${data.imapHost || 'auto'})`);
          startImapWorker(data.email, data.password, data.imapHost, ws);
        }
      }
    } catch (e) {
      console.error('❌ Error WebSocket:', e.message);
    }
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
    console.log(`ℹ️ Worker ya existente para ${email}, no se reinicia`);
    return;
  }

  if (!password) {
    console.error(`❌ startImapWorker: SIN PASSWORD para ${email}, no se puede iniciar`);
    return;
  }

  const auto = getAutoConfig(email);
  const host = customHost || (auto ? auto.imapHost : 'mail.' + email.split('@')[1]);

  const workerState = {
    ws,
    email,
    password,
    host,
    lastUidNext: 0,
    client: null,
    active: true,
    lastAlive: Date.now(),
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

        if (state.ws && state.ws.readyState === WebSocket.OPEN) {
          state.ws.send(JSON.stringify({
            type: 'new_email',
            email: state.email,
            timestamp: new Date().toISOString(),
            from,
            subject,
            uid: msg.uid
          }));
        }

        console.log(`📤 Enviando push para nuevo correo UID:${msg.uid}...`);
        await sendPushNotification(state.email, {
          title: `📧 Nuevo correo de ${from}`,
          body: subject,
          data: {
            type: 'new_email',
            sender: from,
            subject,
            uid: String(msg.uid),
            folder: 'INBOX',
          }
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
    let idleStarted = false;
    try {
      console.log(`🔄 [BG] Conectando IMAP para ${state.email} (${state.host})...`);
      try {
        state.client = await connectImap(state.email, state.password, state.host, 993, true);
      } catch (err) {
        console.log(`⚠️ Puerto 993 falló, probando 143... (${err.message})`);
        state.client = await connectImap(state.email, state.password, state.host, 143, false);
      }

      await state.client.mailboxOpen('INBOX');

      const handleExists = async () => {
        try {
          state.lastAlive = Date.now();
          const status = await state.client.status('INBOX', { uidNext: true });
          const currentUidNext = status.uidNext || 1;

          if (currentUidNext > state.lastUidNext) {
            console.log(`⚡ IDLE: nuevo correo detectado al instante. Rango ${state.lastUidNext} a ${currentUidNext - 1}`);
            await processEmailsInRange(state, state.lastUidNext, currentUidNext);
          }
        } catch (e) {
          console.error(`❌ Error en handleExists (${state.email}):`, e.message);
        }
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
              console.log(`🔄 Fallback poll: nuevo correo detectado (IDLE no disparó)`);
              await processEmailsInRange(state, state.lastUidNext, s.uidNext);
            }
          } catch (e) {
            console.log(`⚠️ Fallback poll falló, cerrando conexión: ${e.message}`);
            break;
          }
        }
      }

      if (state.active && state.client.usable) {
        console.log(`⚠️ Bucle IDLE salió pero el cliente sigue usable. Reiniciando...`);
      }

    } catch (e) {
      if (state.active) {
        console.log(`⚠️ IMAP caído para ${state.email}: ${e.message}`);
      }
    } finally {
      if (state.client) {
        try { state.client.removeAllListeners('exists'); } catch (_) {}
        await state.client.logout().catch(() => {});
        state.client = null;
      }
    }

    if (state.active) {
      console.log(`⏳ Reconectando IMAP ${state.email} en 10s...`);
      await new Promise(r => setTimeout(r, 10000));
    }
  }
}

// ------------------------------------------------------------
//  CRON: RECORDATORIOS DE CALENDARIO
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

      const recipientEmail = event.email ||
        (Array.isArray(event.sharedEmails) && event.sharedEmails.length > 0
          ? event.sharedEmails[0]
          : null);
      if (!recipientEmail) continue;

      if (!event.notified1Day && diffHours <= 25 && diffHours > 23) {
        console.log(`📅 Recordatorio 1 DÍA ANTES para: ${event.title} → ${recipientEmail} (faltan ${diffHours.toFixed(1)}h)`);
        await sendPushNotification(recipientEmail, {
          title: `📅 Mañana: ${event.title}`,
          body: `Tienes este evento programado para mañana a las ${formattedTime}`,
          data: {
            type: 'calendar_event',
            eventId: doc.id,
            eventTitle: event.title,
            eventTime: eventDate.toISOString(),
            eventDate: formattedTime,
            notice: '1day'
          }
        });
        await doc.ref.update({ notified1Day: true });
      }

      if (!event.notifiedEvent && diffMs <= 15 * 60 * 1000 && diffMs > 0) {
        console.log(`⏰ Recordatorio 15 MIN ANTES para: ${event.title} → ${recipientEmail}`);
        await sendPushNotification(recipientEmail, {
          title: `⏰ Comienza pronto: ${event.title}`,
          body: event.description || `El evento comienza a las ${formattedTime}`,
          data: {
            type: 'calendar_event',
            eventId: doc.id,
            eventTitle: event.title,
            eventTime: eventDate.toISOString(),
            eventDate: formattedTime,
            notice: '15min'
          }
        });
        await doc.ref.update({ notifiedEvent: true });
      }
    }
  } catch (e) {
    console.error('❌ Error en Cron Job:', e.message);
  }
});

// ------------------------------------------------------------
//  CRON: REANIMACIÓN DE WORKERS CAÍDOS (cada 3 min)
// ------------------------------------------------------------
cron.schedule('*/3 * * * *', async () => {
  if (!db) return;
  try {
    const snapshot = await db.collection('user_accounts').get();
    for (const doc of snapshot.docs) {
      const data = doc.data();
      if (!data.email || !data.password) continue;

      const state = activeWorkers.get(data.email);
      const isAlive = state &&
                      state.active &&
                      state.client &&
                      state.client.usable;

      if (!isAlive) {
        console.log(`♻️ Cron: reanimando worker caído para ${data.email}`);
        if (state) state.active = false;
        activeWorkers.delete(data.email);
        startImapWorker(data.email, data.password, data.imapHost);
      }
    }
  } catch (e) {
    console.error('⚠️ Error en cron de reanimación:', e.message);
  }
});

// ------------------------------------------------------------
//  CRON: ENVÍO PROGRAMADO (cada minuto)
// ------------------------------------------------------------
cron.schedule('* * * * *', async () => {
  if (!db) return;

  try {
    const now = new Date();

    const snapshot = await db
      .collection('scheduled_emails')
      .where('status', '==', 'pending')
      .where('scheduledFor', '<=', admin.firestore.Timestamp.fromDate(now))
      .limit(10)
      .get();

    if (snapshot.empty) return;

    console.log(`⏰ Procesando ${snapshot.size} correo(s) programado(s)...`);

    for (const doc of snapshot.docs) {
      await processScheduledEmail(doc);
    }
  } catch (e) {
    console.error('❌ Error en cron de programados:', e.message);
  }
});

async function processScheduledEmail(doc) {
  const data = doc.data();
  const docRef = doc.ref;

  try {
    await docRef.update({
      status: 'processing',
      processingStartedAt: new Date(),
    });

    const accountEmail = data.accountEmail;
    if (!accountEmail) throw new Error('Sin cuenta emisora');

    const accountSnap = await db
      .collection('user_accounts')
      .doc(accountEmail)
      .get();
    if (!accountSnap.exists) throw new Error(`No hay cuenta para ${accountEmail}`);

    const account = accountSnap.data();
    const auto = getAutoConfig(accountEmail);
    const smtpHost = account.imapHost
      ? 'smtp.' + account.imapHost.replace(/^mail\./i, '')
      : auto.smtpHost;

    // 🔥 Adjuntos vienen como base64 en data.attachments
    const attachments = [];
    const rawAttachments = data.attachments || [];

    for (const att of rawAttachments) {
      try {
        if (att.content && att.filename) {
          attachments.push({
            filename: att.filename,
            content: Buffer.from(att.content, 'base64'),
            contentType: att.contentType || 'application/octet-stream',
          });
          console.log(`📎 Adjunto preparado: ${att.filename}`);
        }
      } catch (e) {
        console.error(`⚠️ Error procesando adjunto:`, e.message);
      }
    }

    const transporter = nodemailer.createTransport({
      host: smtpHost,
      port: 587,
      secure: false,
      auth: { user: accountEmail, pass: account.password },
      tls: { rejectUnauthorized: false },
    });

    await transporter.sendMail({
      from: accountEmail,
      to: data.to,
      cc: data.cc || undefined,
      bcc: data.bcc || undefined,
      subject: data.subject || '(Sin asunto)',
      html: data.body || '',
      attachments: attachments,
    });

    console.log(`✅ Programado enviado: "${data.subject}" → ${data.to}`);

    // Guardar en Enviados
    try {
      const imapAuto = getAutoConfig(accountEmail);
      const client = await connectImap(
        accountEmail,
        account.password,
        imapAuto.imapHost,
        993,
        true
      );

      const list = await client.list();
      const sentFolder =
        list.find(
          (f) =>
            f.specialUse === '\\Sent' ||
            /^sent$/i.test(f.name) ||
            /enviad/i.test(f.name)
        )?.path || 'INBOX.Sent';

      const rawEmail = [
        `From: ${accountEmail}`,
        `To: ${data.to || ''}`,
        `Subject: ${data.subject || '(Sin asunto)'}`,
        `Date: ${new Date().toUTCString()}`,
        `MIME-Version: 1.0`,
        `Content-Type: text/html; charset=utf-8`,
        '',
        data.body || '',
      ].join('\r\n');

      await client.append(sentFolder, Buffer.from(rawEmail), ['\\Seen']);
      await client.logout();
    } catch (e) {
      console.error('⚠️ No se pudo guardar en Enviados:', e.message);
    }

    await docRef.update({
      status: 'sent',
      sentAt: admin.firestore.FieldValue.serverTimestamp(),
    });

    await sendPushNotification(accountEmail, {
      title: '✅ Correo programado enviado',
      body: `"${data.subject}" se ha enviado a ${data.to}`,
      data: { type: 'scheduled_sent', docId: doc.id },
    });
  } catch (e) {
    console.error(`❌ Error enviando programado ${doc.id}:`, e.message);
    await docRef.update({
      status: 'failed',
      error: e.message,
      failedAt: admin.firestore.FieldValue.serverTimestamp(),
    });
  }
}

// ------------------------------------------------------------
//  FCM PUSH
// ------------------------------------------------------------
async function sendPushNotification(email, payload) {
  if (!db) {
    console.log('⚠️ sendPushNotification: sin db');
    return;
  }

  try {
    const tokensSnapshot = await db.collection('fcm_tokens').where('email', '==', email).get();
    if (tokensSnapshot.empty) {
      console.log(`📴 Sin token FCM para ${email}`);
      return;
    }

    const tokens = [];
    tokensSnapshot.forEach(doc => tokens.push(doc.data().token));

    console.log(`🚀 Enviando push a ${tokens.length} token(s) para ${email}`);

    const message = {
      notification: {
        title: payload.title || 'RSMAIL',
        body: payload.body || 'Nueva notificación',
      },
      data: payload.data || { type: 'general' },
      android: {
        priority: 'high',
        ttl: 60 * 60 * 1000,
        notification: {
          channelId: 'rsmail_high_importance_channel',
          sound: 'default',
          priority: 'max',
          visibility: 'public',
          defaultVibrateTimings: true,
          defaultSound: true,
        }
      },
      apns: {
        headers: {
          'apns-priority': '10',
          'apns-push-type': 'alert',
        },
        payload: {
          aps: {
            sound: 'default',
            badge: 1,
            'content-available': 1,
          }
        }
      },
      tokens,
    };

    const response = await admin.messaging().sendEachForMulticast(message);
    console.log(`✅ Push enviado a ${tokens.length} dispositivo(s) para ${email} | éxito: ${response.successCount}, fallos: ${response.failureCount}`);

    if (response.failureCount > 0) {
      const failedTokens = [];
      response.responses.forEach((resp, idx) => {
        if (!resp.success) {
          console.log(`❌ Token fallido: ${resp.error?.message}`);
          failedTokens.push(tokens[idx]);
        }
      });
      for (const token of failedTokens) {
        const snapshots = await db.collection('fcm_tokens').where('token', '==', token).get();
        snapshots.forEach(doc => doc.ref.delete());
      }
    }
  } catch (e) {
    console.error('❌ Error enviando Push:', e.message);
  }
}

// ------------------------------------------------------------
//  AUTO-CONFIG
// ------------------------------------------------------------
function getAutoConfig(email) {
  const domain = email.split('@')[1]?.toLowerCase();
  if (!domain) return null;
  if (domain.includes('gmail.com')) return { imapHost: 'imap.gmail.com', imapPort: 993, smtpHost: 'smtp.gmail.com', smtpPort: 587, secure: true };
  if (domain.includes('outlook.com') || domain.includes('hotmail.com') || domain.includes('live.com')) return { imapHost: 'outlook.office365.com', imapPort: 993, smtpHost: 'smtp.office365.com', smtpPort: 587, secure: true };
  if (domain.includes('yahoo.')) return { imapHost: 'imap.mail.yahoo.com', imapPort: 993, smtpHost: 'smtp.mail.yahoo.com', smtpPort: 465, secure: true };
  if (domain.includes('zoho.')) return { imapHost: 'imap.zoho.com', imapPort: 993, smtpHost: 'smtp.zoho.com', smtpPort: 465, secure: true };
  return { imapHost: 'mail.' + domain, imapPort: 993, smtpHost: 'mail.' + domain, smtpPort: 587, secure: true };
}

// ------------------------------------------------------------
//  AUTH
// ------------------------------------------------------------
const handleAuth = (req, res) => {
  const { email, password, host, port } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, error: 'Email y contraseña requeridos' });

  const auto = getAutoConfig(email);
  const targetHost = host || auto.imapHost;
  const targetPort = Number(port) || auto.imapPort;

  const imap = new Imap({
    user: email,
    password,
    host: targetHost,
    port: targetPort,
    tls: true,
    tlsOptions: { rejectUnauthorized: false }
  });

  let responded = false;

  imap.once('ready', () => {
    if (!responded) {
      responded = true;
      imap.end();
      saveAccount(email, password, targetHost);
      startImapWorker(email, password, targetHost);
      return res.json({
        success: true,
        message: 'Autenticación exitosa',
        account: {
          email,
          password,
          imapHost: targetHost,
          imapPort: targetPort,
          imapSecurity: 'ssl',
          smtpHost: auto.smtpHost,
          smtpPort: auto.smtpPort,
          smtpSecurity: 'starttls'
        }
      });
    }
  });

  imap.once('error', (err) => {
    if (!responded) {
      responded = true;
      imap.end();
      return res.status(401).json({ success: false, error: 'Credenciales inválidas: ' + err.message });
    }
  });

  setTimeout(() => {
    if (!responded) {
      responded = true;
      imap.end();
      return res.status(408).json({ success: false, error: 'Timeout IMAP' });
    }
  }, 15000);

  imap.connect();
};

// ------------------------------------------------------------
//  API REST
// ------------------------------------------------------------
app.post('/api/login', handleAuth);
app.post('/api/verify', handleAuth);

// /ping reanima workers caídos (aprovecha el ping de UptimeRobot)
app.get('/ping', async (req, res) => {
  if (db) {
    try {
      const deadEmails = [];
      for (const [email, state] of activeWorkers.entries()) {
        const isAlive = state.active &&
                        state.client &&
                        state.client.usable;
        if (!isAlive) {
          deadEmails.push(email);
        }
      }

      for (const email of deadEmails) {
        console.log(`♻️ /ping: reanimando worker caído para ${email}`);
        const state = activeWorkers.get(email);
        if (state) state.active = false;
        activeWorkers.delete(email);

        const doc = await db.collection('user_accounts').doc(email).get();
        if (doc.exists) {
          const data = doc.data();
          if (data.email && data.password) {
            startImapWorker(data.email, data.password, data.imapHost);
          }
        }
      }

      if (activeWorkers.size === 0) {
        const snapshot = await db.collection('user_accounts').get();
        if (snapshot.size > 0) {
          console.log(`♻️ /ping: sin workers, arrancando ${snapshot.size} desde Firestore`);
          for (const d of snapshot.docs) {
            const data = d.data();
            if (data.email && data.password) {
              startImapWorker(data.email, data.password, data.imapHost);
            }
          }
        }
      }
    } catch (e) {
      console.error('⚠️ Error en /ping reanimación:', e.message);
    }
  }

  res.json({
    alive: true,
    ts: new Date().toISOString(),
    workers: activeWorkers.size,
  });
});

// DEBUG: ver workers activos
app.get('/api/debug/workers', (req, res) => {
  const workers = [];
  for (const [email, state] of activeWorkers.entries()) {
    workers.push({
      email,
      host: state.host,
      active: state.active,
      lastUidNext: state.lastUidNext,
      clientUsable: state.client?.usable ?? false,
      wsConnected: state.ws?.readyState === 1,
      lastAliveAgo: Math.floor((Date.now() - (state.lastAlive || 0)) / 1000) + 's',
    });
  }
  res.json({
    count: workers.length,
    workers,
    uptime: process.uptime(),
    timestamp: new Date().toISOString(),
  });
});

app.post('/api/fcm-token', async (req, res) => {
  const { email, token, password, imapHost } = req.body;
  if (!email || !token) return res.status(400).json({ success: false, error: 'Email y token requeridos' });
  if (!db) return res.status(500).json({ success: false, error: 'Firestore no configurado' });

  try {
    const existing = await db.collection('fcm_tokens').where('email', '==', email).get();
    existing.forEach(doc => doc.ref.delete());

    await db.collection('fcm_tokens').add({
      email,
      token,
      createdAt: admin.firestore.FieldValue.serverTimestamp()
    });
    console.log(`📱 Token FCM actualizado en Firestore para ${email}`);

    if (password) {
      console.log(`🔧 fcm-token: iniciando worker para ${email} (host=${imapHost || 'auto'})`);
      saveAccount(email, password, imapHost);
      startImapWorker(email, password, imapHost);
    } else {
      console.log(`⚠️ fcm-token: sin password para ${email}, NO se inicia worker`);
    }

    res.json({ success: true });
  } catch (e) {
    console.error('❌ Error en /api/fcm-token:', e.message);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/send-notification', async (req, res) => {
  const { email, title, body, data } = req.body;
  if (!email) return res.status(400).json({ success: false, error: 'Email requerido' });

  try {
    await sendPushNotification(email, {
      title: title || '📅 Recordatorio',
      body: body || 'Evento programado',
      data: data || { type: 'calendar_event' }
    });
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/send-email', async (req, res) => {
  const { email, password, host, port, to, subject, body, attachments } = req.body;
  if (!email || !password || !to) return res.status(400).json({ success: false, error: 'Faltan campos' });

  const auto = getAutoConfig(email);
  const smtpHost = host || auto.smtpHost;
  const smtpPort = Number(port) || auto.smtpPort || 587;

  const transporter = nodemailer.createTransport({
    host: smtpHost,
    port: smtpPort,
    secure: smtpPort === 465,
    auth: { user: email, pass: password },
    tls: { rejectUnauthorized: false }
  });

  try {
    await transporter.sendMail({
      from: email,
      to,
      subject: subject || '(Sin asunto)',
      html: body || '',
      attachments: attachments ? attachments.map(att => ({
        filename: att.filename,
        content: Buffer.from(att.content, 'base64')
      })) : []
    });
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/save-to-sent', async (req, res) => {
  const { email, password, host, port, to, subject, body } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, error: 'Parámetros insuficientes' });

  const auto = getAutoConfig(email);
  const targetHost = host || auto.imapHost;
  const targetPort = Number(port) || auto.imapPort;

  let client;
  try {
    try {
      client = await connectImap(email, password, targetHost, targetPort, true);
    } catch (err) {
      client = await connectImap(email, password, targetHost, 143, false);
    }

    const list = await client.list();
    let sentFolder = list.find(f =>
      f.specialUse === '\\Sent' ||
      /^sent$/i.test(f.name) ||
      /enviad/i.test(f.name) ||
      /inbox\.sent/i.test(f.path)
    )?.path || 'INBOX.Sent';

    const messageId = `<${Date.now()}.${Math.random().toString(36).substring(2, 10)}@${email.split('@')[1]}>`;
    const date = new Date().toUTCString();

    const rawEmail = [
      `From: ${email}`,
      `To: ${to || ''}`,
      `Subject: ${subject || '(Sin asunto)'}`,
      `Date: ${date}`,
      `Message-ID: ${messageId}`,
      `MIME-Version: 1.0`,
      `Content-Type: text/html; charset=utf-8`,
      `Content-Transfer-Encoding: 7bit`,
      '',
      body || ''
    ].join('\r\n');

    await client.append(sentFolder, Buffer.from(rawEmail), ['\\Seen']);
    await client.logout();

    res.json({ success: true, message: 'Guardado en Enviados' });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/folders', async (req, res) => {
  const { email, password, host, port } = req.body;
  const auto = getAutoConfig(email);
  const targetHost = host || auto.imapHost;
  let client;
  try {
    try {
      client = await connectImap(email, password, targetHost, Number(port) || auto.imapPort, true);
    } catch (err) {
      client = await connectImap(email, password, targetHost, 143, false);
    }

    const list = await client.list();
    await client.logout();
    const folders = list.map(f => ({ name: f.name, path: f.path, specialUse: f.specialUse || '' }));
    res.json({ success: true, folders });
  } catch (err) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: err.message });
  }
});

function detectAttachmentsFromStructure(structure) {
  if (!structure) return false;
  const stack = [structure];

  while (stack.length) {
    const part = stack.pop();
    if (!part) continue;

    const disp = (part.disposition || '').toString().toLowerCase();
    if (disp === 'attachment') return true;

    if (part.dispositionParameters && part.dispositionParameters.filename) {
      return true;
    }

    if (part.parameters && part.parameters.name) {
      return true;
    }

    if (Array.isArray(part.childNodes)) {
      for (const child of part.childNodes) stack.push(child);
    }
  }
  return false;
}

app.post('/api/messages', async (req, res) => {
  const { email, password, host, port, folder = 'INBOX', limit = 20 } = req.body;
  const auto = getAutoConfig(email);
  const targetHost = host || auto.imapHost;

  let client;
  try {
    try {
      client = await connectImap(email, password, targetHost, Number(port) || 993, true);
    } catch (err) {
      client = await connectImap(email, password, targetHost, 143, false);
    }

    const lock = await client.getMailboxLock(folder);
    const messages = [];

    try {
      const iter = client.fetch(
        '1:*',
        { envelope: true, flags: true, bodyStructure: true },
        { max: limit, reverse: true }
      );

      for await (const msg of iter) {
        let flags = msg.flags;
        if (flags instanceof Set) {
          flags = Array.from(flags);
        } else if (!Array.isArray(flags)) {
          flags = [];
        }

        const hasAttachments = detectAttachmentsFromStructure(msg.bodyStructure);

        messages.push({
          uid: msg.uid,
          id: msg.uid.toString(),
          subject: msg.envelope?.subject || '(Sin asunto)',
          from: msg.envelope?.from?.[0]?.address || msg.envelope?.from?.[0]?.name || '',
          to: msg.envelope?.to?.[0]?.address || '',
          date: msg.envelope?.date ? new Date(msg.envelope.date).toISOString() : new Date().toISOString(),
          hasAttachments,
          flags,
          isRead: flags.includes('\\Seen'),
          isFlagged: flags.includes('\\Flagged'),
        });
      }
    } finally {
      lock.release();
    }

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
  const auto = getAutoConfig(email);
  const targetHost = host || auto.imapHost;

  let client;
  try {
    try {
      client = await connectImap(email, password, targetHost, Number(port) || auto.imapPort, true);
    } catch (err) {
      client = await connectImap(email, password, targetHost, 143, false);
    }

    const lock = await client.getMailboxLock(folder);
    let parsed;
    try {
      const msg = await client.fetchOne(String(uid), { source: true }, { uid: true });
      if (msg?.source) parsed = await simpleParser(msg.source);
    } finally {
      lock.release();
    }
    await client.logout();
    if (!parsed) return res.status(404).json({ success: false, error: 'Correo no encontrado' });

    const attachments = (parsed.attachments || []).map(att => ({
      filename: att.filename || 'adjunto',
      contentType: att.contentType,
      size: att.size,
      content: att.content ? att.content.toString('base64') : ''
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
        attachments
      }
    });
  } catch (err) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/delete-message', async (req, res) => {
  const { email, password, host, port, uid, folder = 'INBOX' } = req.body;
  if (!uid) return res.status(400).json({ success: false, error: 'Falta UID' });

  const auto = getAutoConfig(email);
  const targetHost = host || auto.imapHost;

  let client;
  try {
    try {
      client = await connectImap(email, password, targetHost, Number(port) || auto.imapPort, true);
    } catch (err) {
      client = await connectImap(email, password, targetHost, 143, false);
    }

    const lock = await client.getMailboxLock(folder);
    try {
      const isAlreadyTrash = folder.toLowerCase().includes('trash') || folder.toLowerCase().includes('papelera');
      if (isAlreadyTrash) {
        await client.messageDelete(String(uid), { uid: true });
      } else {
        const list = await client.list();
        let trashFolder = list.find(f =>
          f.specialUse === '\\Trash' ||
          /^trash$/i.test(f.name) ||
          /papelera/i.test(f.name) ||
          /inbox\.trash/i.test(f.path)
        )?.path || 'Trash';

        await client.messageMove(String(uid), trashFolder, { uid: true });
      }
    } finally {
      lock.release();
    }

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
  const auto = getAutoConfig(email);
  const targetHost = host || auto.imapHost;

  let client;
  try {
    try {
      client = await connectImap(email, password, targetHost, Number(port) || auto.imapPort, true);
    } catch (err) {
      client = await connectImap(email, password, targetHost, 143, false);
    }

    const lock = await client.getMailboxLock(folder);
    try {
      if (read) await client.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true });
      else await client.messageFlagsRemove(String(uid), ['\\Seen'], { uid: true });
    } finally {
      lock.release();
    }
    await client.logout();
    res.json({ success: true });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/mark-all-read', async (req, res) => {
  const { email, password, host, port, folder = 'INBOX' } = req.body;
  if (!email || !password) {
    return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  }
  const auto = getAutoConfig(email);
  const targetHost = host || auto.imapHost;

  let client;
  try {
    try {
      client = await connectImap(email, password, targetHost, Number(port) || auto.imapPort, true);
    } catch (err) {
      client = await connectImap(email, password, targetHost, 143, false);
    }

    const lock = await client.getMailboxLock(folder);
    let total = 0;
    try {
      const uids = [];
      const iter = client.fetch('1:*', { uid: true }, { uid: true });
      for await (const msg of iter) {
        if (msg.uid) uids.push(msg.uid);
      }
      total = uids.length;

      if (total > 0) {
        await client.messageFlagsAdd(uids, ['\\Seen'], { uid: true });
      }
    } finally {
      lock.release();
    }

    await client.logout();
    console.log(`✅ Marcados ${total} correos como leídos en ${folder} para ${email}`);
    res.json({ success: true, marked: total });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/toggle-flagged', async (req, res) => {
  const { email, password, host, port, uid, folder = 'INBOX', flagged } = req.body;
  if (uid == null) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  const auto = getAutoConfig(email);
  const targetHost = host || auto.imapHost;

  let client;
  try {
    try {
      client = await connectImap(email, password, targetHost, Number(port) || auto.imapPort, true);
    } catch (err) {
      client = await connectImap(email, password, targetHost, 143, false);
    }

    const lock = await client.getMailboxLock(folder);
    try {
      if (flagged) await client.messageFlagsAdd(String(uid), ['\\Flagged'], { uid: true });
      else await client.messageFlagsRemove(String(uid), ['\\Flagged'], { uid: true });
    } finally {
      lock.release();
    }
    await client.logout();
    res.json({ success: true });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/download-attachment', async (req, res) => {
  const { email, password, host, port, folder = 'INBOX', uid, partId } = req.body;
  if (!uid || !partId) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  const auto = getAutoConfig(email);
  const targetHost = host || auto.imapHost;

  let client;
  try {
    try {
      client = await connectImap(email, password, targetHost, Number(port) || auto.imapPort, true);
    } catch (err) {
      client = await connectImap(email, password, targetHost, 143, false);
    }

    const lock = await client.getMailboxLock(folder);
    let msg;
    try {
      msg = await client.fetchOne(String(uid), { bodyParts: [partId] }, { uid: true });
    } finally {
      lock.release();
    }
    await client.logout();
    if (!msg?.bodyParts?.[partId]) return res.status(404).json({ success: false, error: 'Adjunto no encontrado' });
    res.json({ success: true, data: msg.bodyParts[partId].toString('base64') });
  } catch (e) {
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: e.message });
  }
});

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
    if (url.toLowerCase().startsWith('mailto:')) {
      result.mailto.push(url);
    } else if (url.toLowerCase().startsWith('http')) {
      result.http.push(url);
    }
  }
  return result;
}

app.post('/api/scan-subscriptions', async (req, res) => {
  const { email, password, host, port, maxMessages = 500 } = req.body;
  if (!email || !password) {
    return res.status(400).json({ success: false, error: 'Email y contraseña requeridos' });
  }

  const auto = getAutoConfig(email);
  const targetHost = host || auto.imapHost;

  let client;
  try {
    try {
      client = await connectImap(email, password, targetHost, Number(port) || 993, true);
    } catch (err) {
      client = await connectImap(email, password, targetHost, 143, false);
    }

    const lock = await client.getMailboxLock('INBOX');
    const subscriptions = new Map();

    try {
      const status = await client.status('INBOX', { messages: true });
      const total = status.messages || 0;
      const startSeq = Math.max(1, total - maxMessages + 1);

      console.log(`📬 Escaneando suscripciones en ${email}: ${total} mensajes, desde ${startSeq}`);

      const iter = client.fetch(`${startSeq}:*`, {
        uid: true,
        envelope: true,
        headers: ['list-unsubscribe', 'list-unsubscribe-post'],
      });

      for await (const msg of iter) {
        const fromAddr = msg.envelope?.from?.[0]?.address || '';
        const fromName = msg.envelope?.from?.[0]?.name || '';
        const subject = msg.envelope?.subject || '';
        const date = msg.envelope?.date ? new Date(msg.envelope.date) : null;

        let headerListUnsub = '';
        let headerListUnsubPost = '';
        if (msg.headers) {
          const raw = msg.headers.toString();
          const mUnsub = raw.match(/^List-Unsubscribe:\s*(.+)$/im);
          if (mUnsub) headerListUnsub = mUnsub[1].trim();
          const mPost = raw.match(/^List-Unsubscribe-Post:\s*(.+)$/im);
          if (mPost) headerListUnsubPost = mPost[1].trim();
        }

        if (!fromAddr) continue;

        const isSub = analyzeIsSubscription({
          from: fromAddr + ' ' + fromName,
          subject,
          listUnsubscribe: headerListUnsub,
        });
        if (!isSub) continue;

        const key = fromAddr.toLowerCase();
        const parsed = parseListUnsubscribe(headerListUnsub);

        if (!subscriptions.has(key)) {
          subscriptions.set(key, {
            email: fromAddr,
            name: fromName || fromAddr,
            subject,
            latestDate: date ? date.toISOString() : null,
            count: 1,
            listUnsubscribe: headerListUnsub,
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
    } finally {
      lock.release();
    }

    await client.logout();

    const list = Array.from(subscriptions.values())
      .sort((a, b) => (b.latestDate || '').localeCompare(a.latestDate || ''));

    console.log(`✅ Detectadas ${list.length} suscripciones para ${email}`);
    res.json({ success: true, subscriptions: list, total: list.length });
  } catch (err) {
    console.error('❌ Error en /api/scan-subscriptions:', err.message);
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: err.message, subscriptions: [] });
  }
});

app.post('/api/unsubscribe', async (req, res) => {
  const {
    email, password,
    listUnsubscribe, listUnsubscribePost,
  } = req.body;

  if (!listUnsubscribe) {
    return res.status(400).json({ success: false, error: 'Falta la cabecera List-Unsubscribe' });
  }

  const parsed = parseListUnsubscribe(listUnsubscribe);
  const isOneClick = (listUnsubscribePost || '').toLowerCase().includes('one-click');

  try {
    let method = null;
    let result = null;

    if (parsed.http.length > 0 && isOneClick) {
      try {
        const r = await fetch(parsed.http[0], {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            'User-Agent': 'RSMail/3.0',
          },
          body: 'List-Unsubscribe=One-Click',
        });
        method = 'https-post';
        result = { status: r.status, ok: r.ok };
        console.log(`✅ Unsubscribe one-click: ${r.status} para ${parsed.http[0]}`);
      } catch (e) {
        console.log('⚠️ POST one-click falló:', e.message);
      }
    }

    if (!result && parsed.http.length > 0) {
      try {
        const r = await fetch(parsed.http[0], {
          method: 'GET',
          headers: { 'User-Agent': 'RSMail/3.0' },
          redirect: 'follow',
        });
        method = 'https-get';
        result = { status: r.status, ok: r.ok };
        console.log(`✅ Unsubscribe GET: ${r.status} para ${parsed.http[0]}`);
      } catch (e) {
        console.log('⚠️ GET falló:', e.message);
      }
    }

    if (!result && parsed.mailto.length > 0) {
      const mailtoUrl = parsed.mailto[0].replace(/^mailto:/i, '');
      const [address, query] = mailtoUrl.split('?');
      const subjectMatch = (query || '').match(/subject=([^&]+)/i);
      const subject = subjectMatch ? decodeURIComponent(subjectMatch[1]) : 'unsubscribe';

      const auto = getAutoConfig(email);
      const transporter = nodemailer.createTransport({
        host: auto.smtpHost,
        port: auto.smtpPort,
        secure: auto.smtpPort === 465,
        auth: { user: email, pass: password },
        tls: { rejectUnauthorized: false },
      });

      await transporter.sendMail({
        from: email,
        to: address,
        subject: subject,
        text: 'unsubscribe',
      });
      method = 'mailto';
      result = { ok: true };
      console.log(`✅ Unsubscribe por email a ${address}`);
    }

    if (!result) {
      return res.status(400).json({ success: false, error: 'No se encontró un método de baja válido' });
    }

    res.json({ success: true, method, result });
  } catch (e) {
    console.error('❌ Error en /api/unsubscribe:', e.message);
    res.status(500).json({ success: false, error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, async () => {
  console.log(`✅ Backend RSMAIL activo en puerto ${PORT}`);
  await restoreWorkers();
});
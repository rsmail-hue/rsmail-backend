const express = require('express');
const cors = require('cors');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Imap = require('imap');
const nodemailer = require('nodemailer');
const WebSocket = require('ws');
const http = require('http');
const admin = require('firebase-admin');

// ------------------------------------------------------------
//  FIREBASE ADMIN (FCM)
// ------------------------------------------------------------
if (!admin.apps.length) {
  if (process.env.FIREBASE_PRIVATE_KEY) {
    try {
      const formattedKey = process.env.FIREBASE_PRIVATE_KEY
        .replace(/^"|"$/g, '')
        .replace(/\\n/g, '\n');

      admin.initializeApp({
        credential: admin.credential.cert({
          projectId: process.env.FIREBASE_PROJECT_ID,
          clientEmail: process.env.FIREBASE_CLIENT_EMAIL,
          privateKey: formattedKey,
        }),
      });
      console.log('✅ Firebase Admin inicializado con variables de entorno');
    } catch (e) {
      console.error('❌ Error al inicializar Firebase con Variables de Entorno:', e.message);
    }
  } else {
    try {
      const serviceAccount = require('./serviceAccountKey.json');
      admin.initializeApp({
        credential: admin.credential.cert(serviceAccount),
      });
      console.log('✅ Firebase Admin inicializado con serviceAccountKey.json');
    } catch (e) {
      console.error('⚠️ No se encontró serviceAccountKey.json ni variables de entorno válidas.');
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
//  POLLING POR CUENTA (CORREGIDO CON LOCKS)
// ------------------------------------------------------------
const pollingStates = new Map();

// Conexión IMAP con fallback
async function connectImap(email, password, host, port, secure) {
  const config = {
    host,
    port,
    secure,
    auth: { user: email, pass: password },
    logger: false,
    tls: { rejectUnauthorized: false },
    connectionTimeout: 15000,
    authTimeout: 15000,
  };
  console.log(`🔄 Conectando IMAP a ${host}:${port} (secure=${secure})`);
  const client = new ImapFlow(config);
  await client.connect();
  return client;
}

// Obtener UIDs y sobres de todos los mensajes abriendo el Lock correctamente
async function fetchAllMessages(client) {
  const messages = [];
  let lock;
  try {
    lock = await client.getMailboxLock('INBOX');
    const iter = client.fetch('1:*', { uid: true, envelope: true });
    for await (const msg of iter) {
      messages.push(msg);
    }
  } catch (e) {
    console.log('⚠️ Error en fetchAllMessages:', e.message);
  } finally {
    if (lock) lock.release();
  }
  return messages;
}

wss.on('connection', (ws, req) => {
  console.log('🔌 Nuevo cliente WebSocket conectado');
  let email = null;
  let password = null;

  ws.on('message', async (message) => {
    try {
      const data = JSON.parse(message);
      if (data.type === 'login') {
        email = data.email;
        password = data.password;
        console.log(`📧 WebSocket login para ${email}`);
        startPolling(ws, email, password);
      }
    } catch (e) {
      console.error('❌ Error en WebSocket:', e);
    }
  });

  ws.on('close', () => {
    console.log('🔌 Cliente WebSocket desconectado');
    if (email && pollingStates.has(email)) {
      const state = pollingStates.get(email);
      if (state.interval) clearInterval(state.interval);
      pollingStates.delete(email);
    }
  });
});

async function startPolling(ws, email, password) {
  if (pollingStates.has(email)) {
    const old = pollingStates.get(email);
    if (old.interval) clearInterval(old.interval);
    pollingStates.delete(email);
  }

  const auto = getAutoConfig(email);
  const host = auto ? auto.imapHost : 'imap.gmail.com';

  // Conectar inicialmente para obtener UIDs
  let client;
  try {
    client = await connectImap(email, password, host, 993, true);
    console.log(`✅ IMAP conectado (SSL) para ${email}`);
  } catch (err) {
    console.log(`⚠️ Falló SSL directo para ${email}: ${err.message}`);
    try {
      client = await connectImap(email, password, host, 143, false);
      await client.startTls();
      console.log(`✅ IMAP conectado (STARTTLS) para ${email}`);
    } catch (err2) {
      console.error(`❌ No se pudo conectar IMAP para ${email}:`, err2.message);
      ws.send(JSON.stringify({ type: 'error', message: 'Error al conectar con el servidor de correo' }));
      return;
    }
  }

  let initialUids = [];
  try {
    const messages = await fetchAllMessages(client);
    initialUids = messages.map(m => m.uid).filter(u => u !== undefined);
    console.log(`📊 INBOX tiene ${initialUids.length} mensajes para ${email}`);
  } catch (e) {
    console.log(`⚠️ Error al obtener UIDs iniciales: ${e.message}`);
    initialUids = [];
  }

  await client.logout().catch(() => {});

  const state = {
    interval: null,
    ws,
    email,
    password,
    host,
    lastUids: new Set(initialUids),
  };
  pollingStates.set(email, state);
  console.log(`✅ Polling iniciado para ${email} (${initialUids.length} mensajes iniciales)`);

  const poll = async () => {
    try {
      let newClient;
      try {
        newClient = await connectImap(email, password, state.host, 993, true);
      } catch (err) {
        newClient = await connectImap(email, password, state.host, 143, false);
        await newClient.startTls();
      }

      const messages = await fetchAllMessages(newClient);
      await newClient.logout().catch(() => {});

      if (messages.length === 0) {
        console.log(`ℹ️ No hay mensajes en INBOX para ${email}`);
        return;
      }

      const currentUids = messages.map(m => m.uid).filter(u => u !== undefined);
      const newUids = currentUids.filter(uid => !state.lastUids.has(uid));

      if (newUids.length > 0) {
        console.log(`📨 ${newUids.length} nuevo(s) correo(s) detectado(s) para ${email}`);

        for (const uid of newUids) {
          const msg = messages.find(m => m.uid === uid);
          if (!msg) continue;

          const from = msg.envelope.from?.[0]?.address || 'Remitente desconocido';
          const subject = msg.envelope.subject || 'Nuevo correo';
          console.log(`📨 Nuevo correo UID:${uid} de ${from} - ${subject}`);

          if (ws.readyState === WebSocket.OPEN) {
            ws.send(JSON.stringify({
              type: 'new_email',
              email,
              timestamp: new Date().toISOString(),
              from,
              subject,
              uid
            }));
          }

          await sendPushNotification(email, {
            title: `📧 Nuevo correo de ${from}`,
            body: subject,
            data: { type: 'new_email', sender: from, subject, uid: String(uid) }
          });
        }

        state.lastUids = new Set(currentUids);
      } else {
        console.log(`🔍 Polling para ${email}... sin cambios (${currentUids.length} mensajes)`);
      }
    } catch (e) {
      console.log(`⚠️ Polling error para ${email}:`, e.message);
    }
  };

  await poll();
  state.interval = setInterval(poll, 10000);
  console.log(`✅ Polling loop iniciado para ${email} (cada 10s)`);
}

// ============================================================
//  FCM PUSH NOTIFICATIONS
// ============================================================
async function sendPushNotification(email, payload) {
  if (!db) return;
  try {
    const tokensSnapshot = await db.collection('fcm_tokens').where('email', '==', email).get();
    if (tokensSnapshot.empty) {
      console.log(`📴 No hay tokens FCM para ${email}`);
      return;
    }

    const tokens = [];
    tokensSnapshot.forEach(doc => tokens.push(doc.data().token));

    const message = {
      notification: {
        title: payload.title || 'RSMAIL',
        body: payload.body || 'Tienes una nueva notificación',
      },
      data: payload.data || { type: 'general' },
      tokens,
    };

    const response = await admin.messaging().sendEachForMulticast(message);
    console.log(`📨 Notificación push enviada a ${tokens.length} dispositivos para ${email}`);

    if (response.failureCount > 0) {
      const failedTokens = [];
      response.responses.forEach((resp, idx) => {
        if (!resp.success) {
          console.log(`❌ Token fallido: ${tokens[idx]} -> ${resp.error?.message}`);
          failedTokens.push(tokens[idx]);
        }
      });
      for (const token of failedTokens) {
        const snapshots = await db.collection('fcm_tokens').where('token', '==', token).get();
        snapshots.forEach(doc => doc.ref.delete());
      }
      console.log(`🧹 Eliminados ${failedTokens.length} tokens inválidos`);
    }
  } catch (e) {
    console.error('❌ Error enviando push notification:', e);
  }
}

// ============================================================
//  ENDPOINTS FCM
// ============================================================
app.post('/api/fcm-token', async (req, res) => {
  const { email, token } = req.body;
  if (!email || !token || !db) return res.status(400).json({ success: false, error: 'Email y token requeridos' });

  console.log('📥 Token recibido:', token);

  try {
    // Eliminar tokens anteriores del mismo email
    const existing = await db.collection('fcm_tokens').where('email', '==', email).get();
    existing.forEach(doc => doc.ref.delete());

    // Guardar el token nuevo
    await db.collection('fcm_tokens').add({
      email,
      token,
      createdAt: admin.firestore.FieldValue.serverTimestamp()
    });
    console.log(`📱 Token FCM guardado para ${email}`);
    res.json({ success: true });
  } catch (e) {
    console.error('❌ Error guardando token FCM:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/fcm-token/remove', async (req, res) => {
  const { token } = req.body;
  if (!token || !db) return res.status(400).json({ success: false, error: 'Token requerido' });

  try {
    const snapshot = await db.collection('fcm_tokens').where('token', '==', token).get();
    snapshot.forEach(doc => doc.ref.delete());
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ------------------------------------------------------------
//  ENDPOINT PARA LIMPIAR TOKENS ANTIGUOS
// ------------------------------------------------------------
app.post('/api/clean-fcm-tokens', async (req, res) => {
  if (!db) return res.status(500).json({ success: false, error: 'Firestore no disponible' });
  try {
    const snapshot = await db.collection('fcm_tokens').get();
    snapshot.forEach(doc => doc.ref.delete());
    res.json({ success: true, deleted: snapshot.size });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ------------------------------------------------------------
//  ENDPOINT TEMPORAL PARA PROBAR TOKEN FCM
// ------------------------------------------------------------
app.post('/api/test-token', async (req, res) => {
  const { token } = req.body;
  if (!token) return res.status(400).json({ success: false, error: 'Token requerido' });

  try {
    await admin.messaging().send({
      token,
      notification: {
        title: 'Test RSMail',
        body: 'Si ves esto, el token es válido',
      },
    });
    console.log('✅ Token válido y notificación enviada');
    res.json({ success: true, message: 'Notificación enviada correctamente' });
  } catch (e) {
    console.error('❌ Error enviando test:', e.message);
    res.status(400).json({ success: false, error: e.message });
  }
});

// ------------------------------------------------------------
//  AUTO-CONFIGURACIÓN DE SERVIDORES
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
//  LOGIN / VERIFICACIÓN
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
        },
        config: {
          imapHost: targetHost,
          imapPort: targetPort,
          smtpHost: auto.smtpHost,
          smtpPort: auto.smtpPort
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
      return res.status(408).json({ success: false, error: 'Timeout de conexión IMAP' });
    }
  }, 15000);

  imap.connect();
};

app.post('/api/login', handleAuth);
app.post('/api/verify', handleAuth);
app.get('/ping', (req, res) => res.json({ alive: true }));

// ------------------------------------------------------------
//  ENVÍO DE CORREO SMTP
// ------------------------------------------------------------
app.post('/api/send-email', async (req, res) => {
  const { email, password, host, port, to, subject, body, attachments } = req.body;
  if (!email || !password || !to) return res.status(400).json({ success: false, error: 'Faltan campos obligatorios' });

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
    const mailOptions = {
      from: email,
      to,
      subject: subject || '(Sin asunto)',
      html: body || '',
      attachments: attachments ? attachments.map(att => ({
        filename: att.filename,
        content: Buffer.from(att.content, 'base64')
      })) : []
    };

    await transporter.sendMail(mailOptions);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ------------------------------------------------------------
//  GUARDAR EN ENVIADOS
// ------------------------------------------------------------
app.post('/api/save-to-sent', async (req, res) => {
  const { email, password, host, port, to, subject, body } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, error: 'Parámetros insuficientes' });

  const auto = getAutoConfig(email);
  const client = new ImapFlow({
    host: host || auto.imapHost,
    port: Number(port) || auto.imapPort,
    secure: true,
    auth: { user: email, pass: password },
    logger: false,
    tls: { rejectUnauthorized: false }
  });

  try {
    await client.connect();
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
    console.error('❌ Error en save-to-sent:', e.message);
    res.status(500).json({ success: false, error: 'Error al guardar en Enviados: ' + e.message });
  }
});

// ------------------------------------------------------------
//  OBTENER CARPETAS
// ------------------------------------------------------------
app.post('/api/folders', async (req, res) => {
  const { email, password, host, port } = req.body;
  const auto = getAutoConfig(email);
  const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
  try {
    await client.connect();
    const list = await client.list();
    await client.logout();
    const folders = list.map(f => ({ name: f.name, path: f.path, specialUse: f.specialUse || '' }));
    res.json({ success: true, folders });
  } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

// ------------------------------------------------------------
//  OBTENER MENSAJES
// ------------------------------------------------------------
app.post('/api/messages', async (req, res) => {
  const { email, password, host, port, folder = 'INBOX', limit = 20 } = req.body;
  const auto = getAutoConfig(email);
  let client;
  try {
    try {
      client = new ImapFlow({
        host: host || auto.imapHost,
        port: Number(port) || 993,
        secure: true,
        auth: { user: email, pass: password },
        logger: false,
        tls: { rejectUnauthorized: false }
      });
      await client.connect();
    } catch (err) {
      client = new ImapFlow({
        host: host || auto.imapHost,
        port: 143,
        secure: false,
        auth: { user: email, pass: password },
        logger: false,
        tls: { rejectUnauthorized: false }
      });
      await client.connect();
      await client.startTls();
    }
    const lock = await client.getMailboxLock(folder);
    const messages = [];

    try {
      const iter = client.fetch('1:*', { envelope: true, flags: true }, { max: limit, reverse: true });
      for await (const msg of iter) {
        let flags = msg.flags || [];
        if (!Array.isArray(flags)) flags = Object.values(flags);
        messages.push({
          uid: msg.uid,
          id: msg.uid.toString(),
          subject: msg.envelope.subject || '(Sin asunto)',
          from: msg.envelope.from?.[0]?.address || msg.envelope.from?.[0]?.name || '',
          to: msg.envelope.to?.[0]?.address || '',
          date: msg.envelope.date ? new Date(msg.envelope.date).toISOString() : new Date().toISOString(),
          hasAttachments: false,
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
    console.error('❌ Error en /api/messages:', err.message);
    if (client) await client.logout().catch(() => {});
    res.status(500).json({ success: false, error: err.message, messages: [] });
  }
});

// ------------------------------------------------------------
//  DETALLE DEL CORREO
// ------------------------------------------------------------
app.post('/api/message-detail', async (req, res) => {
  const { email, password, host, port, folder = 'INBOX', uid } = req.body;
  if (!uid) return res.status(400).json({ success: false, error: 'UID requerido' });
  const auto = getAutoConfig(email);
  const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
  try {
    await client.connect();
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
  } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

// ------------------------------------------------------------
//  ELIMINAR / MOVER A PAPELERA
// ------------------------------------------------------------
app.post('/api/delete-message', async (req, res) => {
  const { email, password, host, port, uid, folder = 'INBOX' } = req.body;
  if (!uid) return res.status(400).json({ success: false, error: 'Falta UID' });

  const auto = getAutoConfig(email);
  const client = new ImapFlow({
    host: host || auto.imapHost,
    port: Number(port) || auto.imapPort,
    secure: true,
    auth: { user: email, pass: password },
    logger: false,
    tls: { rejectUnauthorized: false }
  });

  try {
    await client.connect();
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
          /inbox\.trash/i.test(f.path) ||
          /inbox\/trash/i.test(f.path)
        )?.path;

        if (!trashFolder) {
          const candidates = ['INBOX.Trash', 'Trash', 'Papelera'];
          for (const cand of candidates) {
            try {
              await client.mailboxCreate(cand);
              trashFolder = cand;
              break;
            } catch (createErr) {}
          }
          if (!trashFolder) throw new Error('No se encontró ni se pudo crear la carpeta de Papelera');
        }

        await client.messageMove(String(uid), trashFolder, { uid: true });
      }
    } finally {
      lock.release();
    }

    await client.logout();
    res.json({ success: true, message: 'Mensaje procesado correctamente' });
  } catch (e) {
    console.error('❌ Error en delete-message:', e.message);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ------------------------------------------------------------
//  MOVER MENSAJE GENÉRICO
// ------------------------------------------------------------
app.post('/api/move-message', async (req, res) => {
  const { email, password, host, port, uid, fromFolder, toFolder } = req.body;
  if (!uid || !fromFolder || !toFolder) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  const auto = getAutoConfig(email);
  const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
  try {
    await client.connect();
    const lock = await client.getMailboxLock(fromFolder);
    try {
      await client.messageMove(String(uid), toFolder, { uid: true });
    } finally {
      lock.release();
    }
    await client.logout();
    res.json({ success: true });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

// ------------------------------------------------------------
//  ACCIONES SOBRE MENSAJES (LEÍDO Y MARCADOS)
// ------------------------------------------------------------
app.post('/api/toggle-read', async (req, res) => {
  const { email, password, host, port, uid, folder = 'INBOX', read } = req.body;
  if (uid == null) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  const auto = getAutoConfig(email);
  const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
  try {
    await client.connect();
    const lock = await client.getMailboxLock(folder);
    try {
      if (read) await client.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true });
      else await client.messageFlagsRemove(String(uid), ['\\Seen'], { uid: true });
    } finally {
      lock.release();
    }
    await client.logout();
    res.json({ success: true });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.post('/api/toggle-flagged', async (req, res) => {
  const { email, password, host, port, uid, folder = 'INBOX', flagged } = req.body;
  if (uid == null) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  const auto = getAutoConfig(email);
  const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
  try {
    await client.connect();
    const lock = await client.getMailboxLock(folder);
    try {
      if (flagged) await client.messageFlagsAdd(String(uid), ['\\Flagged'], { uid: true });
      else await client.messageFlagsRemove(String(uid), ['\\Flagged'], { uid: true });
    } finally {
      lock.release();
    }
    await client.logout();
    res.json({ success: true });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.post('/api/create-folder', async (req, res) => {
  const { email, password, host, port, folderName } = req.body;
  if (!folderName) return res.status(400).json({ success: false, error: 'Falta folderName' });
  const auto = getAutoConfig(email);
  const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
  try { await client.connect(); await client.mailboxCreate(folderName); await client.logout(); res.json({ success: true }); } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.post('/api/delete-folder', async (req, res) => {
  const { email, password, host, port, folderName } = req.body;
  if (!folderName) return res.status(400).json({ success: false, error: 'Falta folderName' });
  const auto = getAutoConfig(email);
  const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
  try { await client.connect(); await client.mailboxDelete(folderName); await client.logout(); res.json({ success: true }); } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.post('/api/download-attachment', async (req, res) => {
  const { email, password, host, port, folder = 'INBOX', uid, partId } = req.body;
  if (!uid || !partId) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
  const auto = getAutoConfig(email);
  const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
  try {
    await client.connect();
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
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

// ------------------------------------------------------------
//  INICIAR SERVIDOR
// ------------------------------------------------------------
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => console.log(`✅ Backend RSMAIL en puerto ${PORT} activo`));
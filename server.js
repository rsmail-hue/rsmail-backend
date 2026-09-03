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
      admin.initializeApp({
        credential: admin.credential.cert({
          projectId: process.env.FIREBASE_PROJECT_ID,
          privateKey: process.env.FIREBASE_PRIVATE_KEY.replace(/\\n/g, '\n'),
          clientEmail: process.env.FIREBASE_CLIENT_EMAIL,
        }),
      });
      console.log('✅ Firebase Admin inicializado con variables de entorno');
    } catch (e) {
      console.error('❌ Error con variables de entorno:', e.message);
    }
  } else {
    try {
      const serviceAccount = require('./serviceAccountKey.json');
      admin.initializeApp({
        credential: admin.credential.cert(serviceAccount),
      });
      console.log('✅ Firebase Admin inicializado con serviceAccountKey.json');
    } catch (e) {
      console.error('⚠️ No se encontró serviceAccountKey.json. FCM no disponible.');
    }
  }
}
const db = admin.firestore();

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const server = http.createServer(app);
const wss = new WebSocket.Server({ server });

// ------------------------------------------------------------
//  WEBSOCKET + IMAP IDLE + POLLING
// ------------------------------------------------------------
const activeSessions = new Map(); // email -> { ws, imapClient, idleTimeout, lastUid, pollingInterval }

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
        await startImapIdle(ws, email, password);
      }
    } catch (e) {
      console.error('❌ Error en WebSocket:', e);
    }
  });

  ws.on('close', () => {
    console.log('🔌 Cliente WebSocket desconectado');
    if (email && activeSessions.has(email)) {
      const session = activeSessions.get(email);
      if (session.idleTimeout) clearTimeout(session.idleTimeout);
      if (session.pollingInterval) clearInterval(session.pollingInterval);
      if (session.imapClient) {
        session.imapClient.logout().catch(() => {});
      }
      activeSessions.delete(email);
    }
  });
});

// ------------------------------------------------------------
//  FUNCIÓN PARA CONECTAR IMAP CON FALLBACK
// ------------------------------------------------------------
async function connectImap(email, password, host, port, secure) {
  const config = {
    host: host,
    port: port,
    secure: secure,
    auth: { user: email, pass: password },
    logger: console,
    tls: {
      rejectUnauthorized: false,
      minVersion: 'TLSv1.2',
    },
    connectionTimeout: 15000,
    authTimeout: 15000,
  };
  console.log(`🔄 Conectando IMAP a ${host}:${port} (secure=${secure})`);
  const client = new ImapFlow(config);
  await client.connect();
  return client;
}

// ------------------------------------------------------------
//  START IMAP IDLE + POLLING
// ------------------------------------------------------------
async function startImapIdle(ws, email, password) {
  try {
    const auto = getAutoConfig(email);
    let client;
    let connected = false;

    // Si ya hay sesión, cerrarla
    if (activeSessions.has(email)) {
      const oldSession = activeSessions.get(email);
      if (oldSession.idleTimeout) clearTimeout(oldSession.idleTimeout);
      if (oldSession.pollingInterval) clearInterval(oldSession.pollingInterval);
      if (oldSession.imapClient) {
        await oldSession.imapClient.logout().catch(() => {});
      }
      activeSessions.delete(email);
      console.log(`♻️ Sesión IMAP anterior cerrada para ${email}`);
    }

    // Intentar STARTTLS (143)
    try {
      client = await connectImap(email, password, auto.imapHost, 143, false);
      await client.startTls();
      connected = true;
      console.log(`✅ IMAP IDLE conectado (STARTTLS) para ${email}`);
    } catch (err1) {
      console.log(`⚠️ Falló STARTTLS: ${err1.message}`);
      // Intentar SSL directo (993)
      try {
        client = await connectImap(email, password, auto.imapHost, 993, true);
        connected = true;
        console.log(`✅ IMAP IDLE conectado (SSL directo) para ${email}`);
      } catch (err2) {
        console.error(`❌ Todos los intentos de conexión IDLE fallaron para ${email}`);
        // Aun así, intentaremos usar polling sin IDLE
        connected = false;
      }
    }

    // Obtener el último UID conocido
    let lastUid = 0;
    if (connected) {
      try {
        await client.mailboxOpen('INBOX');
        const results = await client.fetch('1:*', { uid: true }, { max: 1, reverse: true });
        if (results && results.length > 0) {
          lastUid = results[0].uid || 0;
        }
      } catch (e) {
        console.log('⚠️ No se pudo obtener el último UID:', e.message);
      }
    }

    // Guardar sesión
    const session = {
      ws,
      imapClient: connected ? client : null,
      idleTimeout: null,
      pollingInterval: null,
      lastUid: lastUid,
      email: email,
      password: password,
      imapHost: auto.imapHost,
      isConnected: connected
    };
    activeSessions.set(email, session);

    // ------------------------------------------------------------
    //  POLLING DE RESPALDO (cada 15 segundos)
    // ------------------------------------------------------------
    const startPolling = () => {
      if (session.pollingInterval) clearInterval(session.pollingInterval);
      session.pollingInterval = setInterval(async () => {
        try {
          // Intentar conectar si no hay cliente o está cerrado
          let pollClient = session.imapClient;
          if (!pollClient || !session.isConnected) {
            try {
              // Intentar reconectar
              let newClient;
              try {
                newClient = await connectImap(email, password, auto.imapHost, 143, false);
                await newClient.startTls();
              } catch (err) {
                newClient = await connectImap(email, password, auto.imapHost, 993, true);
              }
              session.imapClient = newClient;
              session.isConnected = true;
              pollClient = newClient;
              console.log(`✅ Polling: reconexión IMAP exitosa para ${email}`);
            } catch (e) {
              console.log(`⚠️ Polling: no se pudo reconectar IMAP para ${email}:`, e.message);
              return;
            }
          }

          await pollClient.mailboxOpen('INBOX');
          const results = await pollClient.fetch('1:*', { uid: true, envelope: true }, { max: 1, reverse: true });
          if (results && results.length > 0) {
            const latest = results[0];
            const currentUid = latest.uid || 0;
            if (currentUid > session.lastUid) {
              // Nuevo correo detectado por polling
              session.lastUid = currentUid;
              const from = latest.envelope.from?.[0]?.address || 'Remitente desconocido';
              const subject = latest.envelope.subject || 'Nuevo correo';
              console.log(`📨 Polling: Nuevo correo detectado para ${email} (UID: ${currentUid})`);

              // Notificar por WebSocket
              if (ws.readyState === WebSocket.OPEN) {
                ws.send(JSON.stringify({
                  type: 'new_email',
                  email: email,
                  timestamp: new Date().toISOString(),
                  from: from,
                  subject: subject
                }));
              }

              // Notificación push
              await sendPushNotification(email, {
                title: `📧 Nuevo correo de ${from}`,
                body: subject,
                data: { type: 'new_email', from: from, subject: subject }
              });
            }
          }
        } catch (e) {
          console.log(`⚠️ Polling error para ${email}:`, e.message);
          session.isConnected = false;
        }
      }, 15000); // Cada 15 segundos
    };

    // Iniciar polling siempre (como respaldo)
    startPolling();

    // ------------------------------------------------------------
    //  IDLE (si la conexión inicial fue exitosa)
    // ------------------------------------------------------------
    if (connected && client) {
      const notifyNewEmail = async () => {
        try {
          // Obtener detalles del último correo
          const lock = await client.getMailboxLock('INBOX');
          let from = 'Remitente desconocido';
          let subject = 'Nuevo correo';
          let uid = 0;
          try {
            const results = await client.fetch('1:*', { envelope: true, uid: true }, { max: 1, reverse: true });
            if (results && results.length > 0) {
              const msg = results[0];
              uid = msg.uid || 0;
              from = msg.envelope.from?.[0]?.address || from;
              subject = msg.envelope.subject || subject;
              if (uid > session.lastUid) {
                session.lastUid = uid;
              }
            }
          } catch (fetchErr) {
            console.log('⚠️ IDLE: No se pudieron obtener detalles del correo:', fetchErr.message);
          } finally {
            lock.release();
          }

          // WebSocket
          if (ws.readyState === WebSocket.OPEN) {
            ws.send(JSON.stringify({
              type: 'new_email',
              email: email,
              timestamp: new Date().toISOString(),
              from: from,
              subject: subject
            }));
          }

          // Push notification
          await sendPushNotification(email, {
            title: `📧 Nuevo correo de ${from}`,
            body: subject,
            data: { type: 'new_email', from: from, subject: subject }
          });
        } catch (e) {
          console.error('❌ Error en notifyNewEmail (IDLE):', e.message);
        }
      };

      const idleLoop = async () => {
        try {
          await client.idle();
          await notifyNewEmail();
          const sessionData = activeSessions.get(email);
          if (sessionData && sessionData.imapClient) {
            sessionData.idleTimeout = setTimeout(idleLoop, 1000);
          }
        } catch (e) {
          console.log(`⚠️ IDLE interrumpido para ${email}:`, e.message);
          const sessionData = activeSessions.get(email);
          if (sessionData) {
            sessionData.isConnected = false;
            sessionData.idleTimeout = setTimeout(idleLoop, 5000);
          }
        }
      };

      // Iniciar IDLE (si la conexión existe)
      idleLoop();
    }

  } catch (e) {
    console.error(`❌ Error iniciando sesión para ${email}:`, e.message);
    ws.send(JSON.stringify({
      type: 'error',
      message: 'Error al conectar con IMAP: ' + e.message
    }));
  }
}

// ============================================================
//  FCM PUSH NOTIFICATIONS
// ============================================================
async function sendPushNotification(email, payload) {
  try {
    const tokensSnapshot = await db.collection('fcm_tokens')
      .where('email', '==', email)
      .get();

    if (tokensSnapshot.empty) {
      console.log(`📴 No hay tokens FCM para ${email}`);
      return;
    }

    const tokens = [];
    tokensSnapshot.forEach(doc => {
      tokens.push(doc.data().token);
    });

    const message = {
      notification: {
        title: payload.title || 'RSMAIL',
        body: payload.body || 'Tienes una nueva notificación',
      },
      data: payload.data || { type: 'general' },
      tokens: tokens,
    };

    const response = await admin.messaging().sendEachForMulticast(message);
    console.log(`📨 Notificación push enviada a ${tokens.length} dispositivos para ${email}`);

    if (response.failureCount > 0) {
      const failedTokens = [];
      response.responses.forEach((resp, idx) => {
        if (!resp.success) {
          failedTokens.push(tokens[idx]);
        }
      });
      for (const token of failedTokens) {
        const snapshots = await db.collection('fcm_tokens')
          .where('token', '==', token)
          .get();
        snapshots.forEach(doc => doc.ref.delete());
      }
      console.log(`🧹 Eliminados ${failedTokens.length} tokens inválidos`);
    }
  } catch (e) {
    console.error('❌ Error enviando push notification:', e);
  }
}

// ============================================================
//  ENDPOINTS FCM (MEJORADOS)
// ============================================================
app.post('/api/fcm-token', async (req, res) => {
  const { email, token } = req.body;
  if (!email || !token) {
    return res.status(400).json({ success: false, error: 'Email y token requeridos' });
  }

  try {
    const existing = await db.collection('fcm_tokens')
      .where('email', '==', email)
      .where('token', '==', token)
      .get();

    if (existing.empty) {
      await db.collection('fcm_tokens').add({
        email: email,
        token: token,
        createdAt: admin.firestore.FieldValue.serverTimestamp(),
      });
      console.log(`📱 Token FCM guardado para ${email}`);
    }
    res.json({ success: true });
  } catch (e) {
    console.error('❌ Error guardando token FCM:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/fcm-token/remove', async (req, res) => {
  const { token } = req.body;
  if (!token) {
    return res.status(400).json({ success: false, error: 'Token requerido' });
  }

  try {
    const snapshot = await db.collection('fcm_tokens')
      .where('token', '==', token)
      .get();
    snapshot.forEach(doc => doc.ref.delete());
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/calendar-notify', async (req, res) => {
  const { email, eventTitle, eventDate } = req.body;
  if (!email) return res.status(400).json({ success: false, error: 'Email requerido' });

  await sendPushNotification(email, {
    title: `📅 Recordatorio: ${eventTitle || 'Evento'}`,
    body: `Programado para ${eventDate || 'próximamente'}`,
    data: { type: 'calendar_event' }
  });
  res.json({ success: true });
});

app.post('/api/note-notify', async (req, res) => {
  const { email, noteTitle, senderEmail } = req.body;
  if (!email) return res.status(400).json({ success: false, error: 'Email requerido' });

  await sendPushNotification(email, {
    title: `📝 Nota compartida por ${senderEmail || 'Alguien'}`,
    body: `"${noteTitle || 'sin título'}"`,
    data: { type: 'note_shared' }
  });
  res.json({ success: true });
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
    return { imapHost: 'mail.' + domain, imapPort: 143, smtpHost: 'mail.' + domain, smtpPort: 587, secure: false };
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
//  4. OBTENER CARPETAS
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
//  5. OBTENER MENSAJES (CON CORRECCIÓN DE flags)
// ------------------------------------------------------------
app.post('/api/messages', async (req, res) => {
    const { email, password, host, port, folder = 'INBOX', limit = 20 } = req.body;
    const auto = getAutoConfig(email);
    let client;
    let connected = false;

    // PRIORIDAD 1: STARTTLS en puerto 143
    try {
        client = new ImapFlow({
            host: host || auto.imapHost,
            port: 143,
            secure: false,
            auth: { user: email, pass: password },
            logger: false,
            tls: { rejectUnauthorized: false }
        });
        await client.connect();
        await client.starttls();
        connected = true;
        console.log(`✅ IMAP conectado (STARTTLS) para ${email} en /api/messages`);
    } catch (err) {
        console.log(`⚠️ Falló STARTTLS en /api/messages: ${err.message}`);
        try {
            client = new ImapFlow({
                host: host || auto.imapHost,
                port: 993,
                secure: true,
                auth: { user: email, pass: password },
                logger: false,
                tls: { rejectUnauthorized: false }
            });
            await client.connect();
            connected = true;
            console.log(`✅ IMAP conectado (SSL directo) para ${email} en /api/messages`);
        } catch (err2) {
            console.error(`❌ Falló SSL directo en /api/messages: ${err2.message}`);
            return res.status(500).json({ success: false, error: err2.message, messages: [] });
        }
    }

    if (!connected) {
        return res.status(500).json({ success: false, error: 'No se pudo conectar al servidor IMAP', messages: [] });
    }

    try {
        const lock = await client.getMailboxLock(folder);
        const messages = [];

        try {
            for await (const msg of client.fetch('1:*', { envelope: true, flags: true }, { max: limit, reverse: true })) {
                let flags = msg.flags || [];
                if (!Array.isArray(flags)) {
                    flags = Object.values(flags);
                }
                messages.push({
                    uid: msg.uid,
                    id: msg.uid.toString(),
                    subject: msg.envelope.subject || '(Sin asunto)',
                    from: msg.envelope.from?.[0]?.address || msg.envelope.from?.[0]?.name || '',
                    to: msg.envelope.to?.[0]?.address || '',
                    date: msg.envelope.date ? new Date(msg.envelope.date).toISOString() : new Date().toISOString(),
                    hasAttachments: false,
                    flags: flags,
                    isRead: flags.includes('\\Seen'),
                    isFlagged: flags.includes('\\Flagged'),
                });
            }
        } finally {
            lock.release();
        }

        await client.logout();
        messages.sort((a, b) => new Date(b.date) - new Date(a.date));
        res.json({ success: true, messages: messages, total: messages.length });
    } catch (err) {
        console.error('❌ Error en /api/messages:', err.message);
        res.status(500).json({ success: false, error: err.message, messages: [] });
    }
});

// ------------------------------------------------------------
//  6. DETALLE DEL CORREO
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
//  7. ELIMINAR / MOVER A PAPELERA
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
            const isAlreadyTrash = folder.toLowerCase().includes('trash') ||
                                   folder.toLowerCase().includes('papelera');

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
                    if (!trashFolder) {
                        throw new Error('No se encontró ni se pudo crear la carpeta de Papelera');
                    }
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
//  8. MOVER MENSAJE GENÉRICO
// ------------------------------------------------------------
app.post('/api/move-message', async (req, res) => {
    const { email, password, host, port, uid, fromFolder, toFolder } = req.body;
    if (!uid || !fromFolder || !toFolder) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
    const auto = getAutoConfig(email);
    const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
    try { await client.connect(); await client.messageMove(uid, toFolder, { uid: true }); await client.logout(); res.json({ success: true }); } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

// ------------------------------------------------------------
//  9‑12. RESTO DE ENDPOINTS
// ------------------------------------------------------------
app.post('/api/toggle-read', async (req, res) => {
    const { email, password, host, port, uid, folder, read } = req.body;
    if (uid == null || !folder) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
    const auto = getAutoConfig(email);
    const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
    try { await client.connect(); if (read) await client.messageFlagsAdd(uid, ['\\Seen'], { uid: true }); else await client.messageFlagsRemove(uid, ['\\Seen'], { uid: true }); await client.logout(); res.json({ success: true }); } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.post('/api/toggle-flagged', async (req, res) => {
    const { email, password, host, port, uid, folder, flagged } = req.body;
    if (uid == null || !folder) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
    const auto = getAutoConfig(email);
    const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
    try { await client.connect(); if (flagged) await client.messageFlagsAdd(uid, ['\\Flagged'], { uid: true }); else await client.messageFlagsRemove(uid, ['\\Flagged'], { uid: true }); await client.logout(); res.json({ success: true }); } catch (e) { res.status(500).json({ success: false, error: e.message }); }
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
    const { email, password, host, port, folder, uid, partId } = req.body;
    if (!uid || !folder || !partId) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
    const auto = getAutoConfig(email);
    const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
    try { await client.connect(); await client.mailboxOpen(folder); const msg = await client.fetchOne(String(uid), { bodyParts: [partId] }, { uid: true }); await client.logout(); if (!msg?.bodyParts?.[partId]) return res.status(404).json({ success: false, error: 'Adjunto no encontrado' }); res.json({ success: true, data: msg.bodyParts[partId].toString('base64') }); } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

// ------------------------------------------------------------
//  INICIAR SERVIDOR
// ------------------------------------------------------------
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => console.log(`✅ Backend RSMAIL en puerto ${PORT} con WebSocket, IDLE y polling`));
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
//  WEBSOCKET + IMAP IDLE (UNA CONEXIÓN POR USUARIO)
// ------------------------------------------------------------
const activeSessions = new Map(); // email -> { ws, imapClient, idleTimeout }

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
      if (session.imapClient) {
        session.imapClient.logout().catch(() => {});
      }
      activeSessions.delete(email);
    }
  });
});

// ------------------------------------------------------------
//  FUNCIÓN PARA CONECTAR IMAP (SOLO SSL DIRECTO)
// ------------------------------------------------------------
async function connectImap(email, password, host, port) {
  const config = {
    host: host,
    port: port || 993,
    secure: true,
    auth: { user: email, pass: password },
    logger: false,
    tls: {
      rejectUnauthorized: false,
    },
    connectionTimeout: 15000,
    authTimeout: 15000,
  };
  console.log(`🔄 Conectando IMAP a ${host}:${port || 993} (SSL)`);
  const client = new ImapFlow(config);
  await client.connect();
  return client;
}

// ------------------------------------------------------------
//  START IMAP IDLE (CON CIERRE DE SESIONES ANTERIORES)
// ------------------------------------------------------------
async function startImapIdle(ws, email, password) {
  try {
    const auto = getAutoConfig(email);

    // 🔥 Si ya hay una sesión activa para este email, cerrarla primero
    if (activeSessions.has(email)) {
      const oldSession = activeSessions.get(email);
      if (oldSession.idleTimeout) clearTimeout(oldSession.idleTimeout);
      if (oldSession.imapClient) {
        await oldSession.imapClient.logout().catch(() => {});
      }
      activeSessions.delete(email);
      console.log(`♻️ Sesión IMAP anterior cerrada para ${email}`);
    }

    // 🔥 Conectar con SSL directo (puerto 993)
    const client = await connectImap(email, password, auto.imapHost, auto.imapPort || 993);
    console.log(`✅ IMAP IDLE conectado (SSL) para ${email}`);

    await client.mailboxOpen('INBOX');
    activeSessions.set(email, { ws, imapClient: client, idleTimeout: null });

    const notifyNewEmail = async () => {
      if (ws.readyState === WebSocket.OPEN) {
        ws.send(JSON.stringify({
          type: 'new_email',
          email: email,
          timestamp: new Date().toISOString()
        }));
      }
      await sendPushNotification(email, {
        title: '📧 Nuevo correo',
        body: 'Tienes un nuevo mensaje en tu bandeja',
        data: { type: 'new_email' }
      });
    };

    const idleLoop = async () => {
      try {
        await client.idle();
        await notifyNewEmail();
        const session = activeSessions.get(email);
        if (session && session.imapClient) {
          session.idleTimeout = setTimeout(idleLoop, 5000); // 5 segundos entre reintentos
        }
      } catch (e) {
        console.log(`⚠️ IDLE interrumpido para ${email}:`, e.message);
        const session = activeSessions.get(email);
        if (session) {
          session.idleTimeout = setTimeout(idleLoop, 15000); // Reintentar después de 15s
        }
      }
    };

    idleLoop();

    const heartbeat = setInterval(() => {
      if (ws.readyState === WebSocket.OPEN) {
        ws.send(JSON.stringify({ type: 'ping' }));
      }
    }, 45000); // Heartbeat cada 45 segundos (menos frecuente)

    ws.on('close', () => clearInterval(heartbeat));

  } catch (e) {
    console.error(`❌ Error iniciando IDLE para ${email}:`, e.message);
    // No enviamos error al cliente para no romper la conexión WebSocket
  }
}

// ============================================================
//  FCM PUSH NOTIFICATIONS (sin cambios)
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
//  ENDPOINTS FCM (sin cambios)
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
    title: '📅 Recordatorio de evento',
    body: `${eventTitle || 'Evento'} - ${eventDate || 'Próximamente'}`,
    data: { type: 'calendar_event', eventTitle: eventTitle || 'Evento', eventDate: eventDate || 'Próximamente' }
  });
  res.json({ success: true });
});

app.post('/api/note-notify', async (req, res) => {
  const { email, noteTitle, senderEmail } = req.body;
  if (!email) return res.status(400).json({ success: false, error: 'Email requerido' });

  await sendPushNotification(email, {
    title: '📝 Nota compartida',
    body: `${senderEmail || 'Alguien'} compartió "${noteTitle || 'una nota'}"`,
    data: { type: 'note_shared', noteTitle: noteTitle || 'nota', senderEmail: senderEmail || 'Alguien' }
  });
  res.json({ success: true });
});

// ------------------------------------------------------------
//  AUTO-CONFIGURACIÓN DE SERVIDORES (SOLO SSL DIRECTO)
// ------------------------------------------------------------
function getAutoConfig(email) {
    const domain = email.split('@')[1]?.toLowerCase();
    if (!domain) return null;
    if (domain.includes('gmail.com')) return { imapHost: 'imap.gmail.com', imapPort: 993, smtpHost: 'smtp.gmail.com', smtpPort: 587, secure: true };
    if (domain.includes('outlook.com') || domain.includes('hotmail.com') || domain.includes('live.com')) return { imapHost: 'outlook.office365.com', imapPort: 993, smtpHost: 'smtp.office365.com', smtpPort: 587, secure: true };
    if (domain.includes('yahoo.')) return { imapHost: 'imap.mail.yahoo.com', imapPort: 993, smtpHost: 'smtp.mail.yahoo.com', smtpPort: 465, secure: true };
    if (domain.includes('zoho.')) return { imapHost: 'imap.zoho.com', imapPort: 993, smtpHost: 'smtp.zoho.com', smtpPort: 465, secure: true };
    // Para dominios genéricos, usar SSL directo en 993
    return { imapHost: 'mail.' + domain, imapPort: 993, smtpHost: 'mail.' + domain, smtpPort: 587, secure: true };
}

// ------------------------------------------------------------
//  1. LOGIN / VERIFICACIÓN (sin cambios)
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
//  2. ENVÍO DE CORREO SMTP (sin cambios)
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
//  3. GUARDAR EN ENVIADOS (sin cambios)
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
//  4. OBTENER CARPETAS (SSL directo)
// ------------------------------------------------------------
app.post('/api/folders', async (req, res) => {
    const { email, password, host, port } = req.body;
    const auto = getAutoConfig(email);
    let client;
    try {
        client = new ImapFlow({
            host: host || auto.imapHost,
            port: Number(port) || auto.imapPort || 993,
            secure: true,
            auth: { user: email, pass: password },
            logger: false,
            tls: { rejectUnauthorized: false }
        });
        await client.connect();
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
//  5. OBTENER MENSAJES (SSL directo)
// ------------------------------------------------------------
app.post('/api/messages', async (req, res) => {
    const { email, password, host, port, folder = 'INBOX', limit = 20 } = req.body;
    const auto = getAutoConfig(email);
    let client;
    try {
        client = new ImapFlow({
            host: host || auto.imapHost,
            port: Number(port) || auto.imapPort || 993,
            secure: true,
            auth: { user: email, pass: password },
            logger: false,
            tls: { rejectUnauthorized: false }
        });
        await client.connect();
        const lock = await client.getMailboxLock(folder);
        const messages = [];

        try {
            for await (const msg of client.fetch('1:*', { envelope: true, flags: true }, { max: limit, reverse: true })) {
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
        if (client) await client.logout().catch(() => {});
        console.error('❌ Error en /api/messages:', err.message);
        res.status(500).json({ success: false, error: err.message, messages: [] });
    }
});

// ------------------------------------------------------------
//  6-12. RESTO DE ENDPOINTS (SSL directo y cierre)
// ------------------------------------------------------------
// (Mantener igual que antes, pero con la conexión SSL directa)
// ...

// ------------------------------------------------------------
//  INICIAR SERVIDOR
// ------------------------------------------------------------
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => console.log(`✅ Backend RSMAIL en puerto ${PORT} con WebSocket y FCM`));
const express = require('express');
const cors = require('cors');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const Imap = require('imap');
const nodemailer = require('nodemailer');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

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
//  LOGIN
// ------------------------------------------------------------
app.post('/api/login', (req, res) => {
    const { email, password } = req.body;
    if (!email || !password) return res.status(400).json({ success: false, error: 'Email y contraseña requeridos' });
    const auto = getAutoConfig(email);
    const imap = new Imap({ user: email, password, host: auto.imapHost, port: auto.imapPort, tls: true, tlsOptions: { rejectUnauthorized: false } });
    let responded = false;
    imap.once('ready', () => {
        if (!responded) { responded = true; imap.end(); res.json({ success: true, message: 'Autenticación exitosa', account: { email, password, imapHost: auto.imapHost, imapPort: auto.imapPort, imapSecurity: 'ssl', smtpHost: auto.smtpHost, smtpPort: auto.smtpPort, smtpSecurity: auto.smtpPort === 465 ? 'ssl' : 'starttls' } }); }
    });
    imap.once('error', (err) => { if (!responded) { responded = true; imap.end(); res.status(401).json({ success: false, error: 'Credenciales inválidas: ' + err.message }); } });
    const timeout = setTimeout(() => { if (!responded) { res.status(408).json({ success: false, error: 'Timeout' }); imap.end(); } }, 15000);
    imap.connect();
});

app.get('/ping', (req, res) => res.json({ alive: true }));

// ------------------------------------------------------------
//  OBTENER CARPETAS (devuelve path y specialUse)
// ------------------------------------------------------------
app.post('/api/folders', async (req, res) => {
    const { email, password, host, port } = req.body;
    const auto = getAutoConfig(email);
    const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
    try { await client.connect(); const list = await client.list(); await client.logout(); const folders = list.map(f => ({ name: f.name, path: f.path, specialUse: f.specialUse || '' })); res.json({ success: true, folders }); }
    catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

// ------------------------------------------------------------
//  OBTENER MENSAJES (ImapFlow con reverse, siempre recientes primero)
// ------------------------------------------------------------
app.post('/api/messages', async (req, res) => {
    const { email, password, host, port, folder = 'INBOX', limit = 30 } = req.body;
    const auto = getAutoConfig(email);
    const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
    try {
        await client.connect();
        const lock = await client.getMailboxLock(folder);
        const messages = [];
        try { for await (const msg of client.fetch('1:*', { envelope: true, bodyStructure: true }, { max: limit, reverse: true })) { messages.push({ uid: msg.uid, id: msg.uid.toString(), subject: msg.envelope.subject || '(Sin asunto)', from: msg.envelope.from?.[0]?.address || msg.envelope.from?.[0]?.name || '', to: msg.envelope.to?.[0]?.address || '', date: msg.envelope.date ? new Date(msg.envelope.date).toISOString() : new Date().toISOString(), hasAttachments: Boolean(msg.bodyStructure?.childNodes?.length) }); } } finally { lock.release(); }
        await client.logout();
        res.json({ success: true, messages, total: messages.length });
    } catch (err) { res.status(500).json({ success: false, error: err.message }); }
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
        await client.connect(); const lock = await client.getMailboxLock(folder); let parsed;
        try { const msg = await client.fetchOne(String(uid), { source: true }, { uid: true }); if (msg?.source) parsed = await simpleParser(msg.source); } finally { lock.release(); }
        await client.logout();
        if (!parsed) return res.status(404).json({ success: false, error: 'Correo no encontrado' });
        const attachments = (parsed.attachments || []).map(att => ({ filename: att.filename || 'adjunto', contentType: att.contentType, size: att.size, content: att.content ? att.content.toString('base64') : '' }));
        res.json({ success: true, message: { uid: Number(uid), subject: parsed.subject || '(Sin asunto)', from: parsed.from?.text || '', to: parsed.to?.text || '', date: parsed.date ? parsed.date.toISOString() : new Date().toISOString(), text: parsed.text || '', html: parsed.html || parsed.textAsHtml || parsed.text || '', attachments } });
    } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

// ------------------------------------------------------------
//  GUARDAR EN ENVIADOS
// ------------------------------------------------------------
app.post('/api/save-to-sent', async (req, res) => {
    const { email, password, host, rawMessage } = req.body;
    if (!email || !password || !rawMessage) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
    const auto = getAutoConfig(email);
    const imapHost = host || auto.imapHost;
    const imap = new Imap({ user: email, password, host: imapHost, port: 993, tls: true, tlsOptions: { rejectUnauthorized: false } });
    imap.once('ready', () => {
        imap.getBoxes((err, boxes) => {
            if (err) { imap.end(); return res.status(500).json({ error: 'Error carpetas' }); }
            const sentFolder = findSentFolder(boxes) || 'INBOX.Sent';
            imap.openBox(sentFolder, false, (openErr) => { if (openErr) { imap.end(); return res.status(500).json({ error: 'Error abriendo Enviados' }); } imap.append(rawMessage, { mailbox: sentFolder, flags: ['\\Seen'] }, (appendErr) => { imap.end(); if (appendErr) return res.status(500).json({ error: appendErr.message }); res.json({ success: true, message: 'Guardado en Enviados' }); }); });
        });
    });
    imap.once('error', (err) => res.status(500).json({ error: err.message }));
    imap.connect();
});

function findSentFolder(boxes) { const patterns = [/^sent$/i, /enviado/i, /inbox\.sent/i]; function search(boxObj, prefix = '') { for (const key in boxObj) { const fullPath = prefix ? `${prefix}${boxObj[key].delimiter}${key}` : key; if (boxObj[key].attribs && boxObj[key].attribs.map(a => a.toLowerCase()).includes('\\sent')) return fullPath; for (const reg of patterns) if (reg.test(key) || reg.test(fullPath)) return fullPath; if (boxObj[key].children) { const found = search(boxObj[key].children, fullPath); if (found) return found; } } return null; } return search(boxes); }

// ------------------------------------------------------------
//  MOVER MENSAJE
// ------------------------------------------------------------
app.post('/api/move-message', async (req, res) => {
    const { email, password, host, port, uid, fromFolder, toFolder } = req.body;
    if (!uid || !fromFolder || !toFolder) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
    const auto = getAutoConfig(email);
    const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
    try { await client.connect(); await client.messageMove(uid, toFolder, { uid: true }); await client.logout(); res.json({ success: true }); } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

// ------------------------------------------------------------
//  RESTO DE ENDPOINTS (delete-message, toggle-read, etc.)
// ------------------------------------------------------------
app.post('/api/delete-message', async (req, res) => {
    const { email, password, host, port, uid, folder } = req.body;
    if (!uid || !folder) return res.status(400).json({ success: false, error: 'Faltan parámetros' });
    const auto = getAutoConfig(email);
    const client = new ImapFlow({ host: host || auto.imapHost, port: Number(port) || auto.imapPort, secure: true, auth: { user: email, pass: password }, logger: false, tls: { rejectUnauthorized: false } });
    try { await client.connect(); await client.messageDelete(uid, { uid: true }); await client.expunge(); await client.logout(); res.json({ success: true }); } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Backend RSMAIL en puerto ${PORT}`));
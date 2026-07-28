const express = require('express');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const insecureTls = {
    rejectUnauthorized: false,
    checkServerIdentity: () => undefined
};

app.get('/ping', (req, res) => res.json({ alive: true, time: new Date().toISOString() }));

// ------------------------------------------------------------
//  OBTENER CARPETAS
// ------------------------------------------------------------
app.post('/api/folders', async (req, res) => {
    const { email, password, host, port, secure } = req.body;
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls
    });
    try {
        await client.connect();
        const mailboxes = await client.list();
        await client.logout();
        const folders = mailboxes.map(mbox => ({
            name: mbox.name,
            path: mbox.path,
            specialUse: mbox.specialUse
        }));
        res.json({ success: true, folders });
    } catch (error) {
        console.error('Error /api/folders:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------------------------------------------------
//  OBTENER LISTA DE MENSAJES
// ------------------------------------------------------------
app.post('/api/messages', async (req, res) => {
    const { email, password, host, port, secure, folder } = req.body;
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxOpen(folder || 'INBOX');
        const messages = [];
        for await (const msg of client.fetch('1:*', { envelope: true, bodyStructure: true })) {
            let hasAttachments = false;
            if (msg.bodyStructure) {
                const checkAttachments = (node) => {
                    if (!node) return;
                    if (node.disposition === 'attachment' ||
                        (node.type === 'application' && node.parameters && node.parameters.name) ||
                        node.type === 'image') {
                        hasAttachments = true;
                    }
                    if (node.childNodes) node.childNodes.forEach(checkAttachments);
                };
                checkAttachments(msg.bodyStructure);
            }
            messages.push({
                uid: msg.uid,
                subject: msg.envelope.subject || '(Sin asunto)',
                from: msg.envelope.from?.[0]?.address || email,
                date: msg.envelope.date || new Date(),
                hasAttachments: hasAttachments
            });
        }
        await client.logout();
        res.json({ success: true, messages });
    } catch (error) {
        console.error('Error /api/messages:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------------------------------------------------
//  DETALLE DE MENSAJE (CON mailparser y FALLBACK a texto plano)
// ------------------------------------------------------------
app.post('/api/message-detail', async (req, res) => {
    const { email, password, host, port, secure, folder, uid } = req.body;
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxOpen(folder || 'INBOX');

        const msg = await client.fetchOne(String(uid), {
            source: true,
            envelope: true,
            bodyStructure: true
        }, { uid: true });
        await client.logout();

        if (!msg || !msg.source) {
            return res.status(500).json({ success: false, error: 'No se pudo obtener el mensaje' });
        }

        const parsed = await simpleParser(msg.source);

        let htmlContent = '';
        let plainText = '';
        let to = '';
        let cc = '';
        let attachments = [];

        // Extraer HTML o texto plano
        if (parsed.html) {
            htmlContent = parsed.html;
        } else if (parsed.text) {
            plainText = parsed.text;
            // Envolver texto plano en <pre> para preservar formato
            htmlContent = `<pre style="font-family: monospace; white-space: pre-wrap; word-wrap: break-word;">${parsed.text}</pre>`;
        } else {
            htmlContent = '<p>No se pudo extraer contenido del mensaje.</p>';
        }

        // Convertir imágenes inline (cid) a base64
        if (parsed.attachments && parsed.attachments.length > 0) {
            for (const att of parsed.attachments) {
                if (att.contentId && att.contentType && att.contentType.startsWith('image/')) {
                    const base64 = att.content.toString('base64');
                    const cid = att.contentId.replace(/[<>]/g, '');
                    htmlContent = htmlContent.replace(new RegExp(`cid:${cid}`, 'g'), `data:${att.contentType};base64,${base64}`);
                }
            }
            // Adjuntos normales (no imágenes inline)
            for (const att of parsed.attachments) {
                if (!att.contentId || !att.contentType.startsWith('image/')) {
                    attachments.push({
                        filename: att.filename || 'adjunto',
                        contentType: att.contentType || 'application/octet-stream',
                        size: att.size || 0,
                        partId: att.contentId || att.partId || ''
                    });
                }
            }
        }

        if (parsed.to) to = parsed.to.map(a => a.address).join(', ');
        if (parsed.cc) cc = parsed.cc.map(a => a.address).join(', ');

        res.json({
            success: true,
            htmlBody: htmlContent,
            body: plainText,
            to,
            cc,
            attachments
        });
    } catch (error) {
        console.error('Error /api/message-detail:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------------------------------------------------
// RESTO DE ENDPOINTS (move, append, delete, toggle, etc.)
// ------------------------------------------------------------
app.post('/api/move-message', async (req, res) => {
    const { email, password, host, port, secure, uid, fromFolder, toFolder } = req.body;
    if (!uid || !fromFolder || !toFolder) {
        return res.status(400).json({ success: false, error: 'Faltan parámetros' });
    }
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxOpen(fromFolder);
        await client.messageMove(uid, toFolder, { uid: true });
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/move-message:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/append-sent', async (req, res) => {
    const { email, password, host, port, secure, rawMessage, sentFolderName } = req.body;
    if (!rawMessage) {
        return res.status(400).json({ success: false, error: 'Falta rawMessage' });
    }
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls
    });
    try {
        await client.connect();
        const folder = sentFolderName || 'Sent';
        await client.mailboxOpen(folder);
        await client.append(folder, rawMessage, ['\\Seen']);
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/append-sent:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/delete-message', async (req, res) => {
    const { email, password, host, port, secure, uid, folder } = req.body;
    if (!uid || !folder) {
        return res.status(400).json({ success: false, error: 'Faltan uid o folder' });
    }
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxOpen(folder);
        await client.messageDelete(uid, { uid: true });
        await client.expunge();
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/delete-message:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/toggle-read', async (req, res) => {
    const { email, password, host, port, secure, uid, folder, read } = req.body;
    if (uid == null || !folder) {
        return res.status(400).json({ success: false, error: 'Faltan uid o folder' });
    }
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxOpen(folder);
        if (read) {
            await client.messageSet(String(uid), ['\\Seen'], { uid: true });
        } else {
            await client.messageUnset(String(uid), ['\\Seen'], { uid: true });
        }
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/toggle-read:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/toggle-flagged', async (req, res) => {
    const { email, password, host, port, secure, uid, folder, flagged } = req.body;
    if (uid == null || !folder) {
        return res.status(400).json({ success: false, error: 'Faltan uid o folder' });
    }
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxOpen(folder);
        if (flagged) {
            await client.messageSet(String(uid), ['\\Flagged'], { uid: true });
        } else {
            await client.messageUnset(String(uid), ['\\Flagged'], { uid: true });
        }
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/toggle-flagged:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/create-folder', async (req, res) => {
    const { email, password, host, port, secure, folderName } = req.body;
    if (!folderName) {
        return res.status(400).json({ success: false, error: 'Falta folderName' });
    }
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxCreate(folderName);
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/create-folder:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/delete-folder', async (req, res) => {
    const { email, password, host, port, secure, folderName } = req.body;
    if (!folderName) {
        return res.status(400).json({ success: false, error: 'Falta folderName' });
    }
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxDelete(folderName);
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/delete-folder:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/download-attachment', async (req, res) => {
    const { email, password, host, port, secure, folder, uid, partId } = req.body;
    if (!uid || !folder || !partId) {
        return res.status(400).json({ success: false, error: 'Faltan uid, folder o partId' });
    }
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxOpen(folder);
        const msg = await client.fetchOne(String(uid), { bodyParts: [partId] }, { uid: true });
        await client.logout();
        const data = msg.bodyParts[partId]?.toString('base64') || '';
        res.json({ success: true, data });
    } catch (error) {
        console.error('Error /api/download-attachment:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`✅ Backend de RSMAIL corriendo en puerto ${PORT}`);
});
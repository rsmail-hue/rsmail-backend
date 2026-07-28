const express = require('express');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const insecureTls = { rejectUnauthorized: false, checkServerIdentity: () => undefined };

// ------------------------------------------------------------
//  PING
// ------------------------------------------------------------
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
//  DETALLE DE MENSAJE (con imágenes inline en base64)
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

        let html = '';
        let plainText = '';
        let to = '';
        let cc = '';
        let attachments = [];

        try {
            const parsed = await simpleParser(msg.source);

            if (parsed.html) {
                html = parsed.html;
            } else if (parsed.text) {
                const text = parsed.text;
                html = text
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/\r?\n/g, '<br>');
                html = '<div style="font-family: -apple-system, Roboto, sans-serif; font-size: 16px; max-width: 100%; word-wrap: break-word;">' + html + '</div>';
            }

            // ---------- PROCESAR IMÁGENES INLINE (cid) ----------
            if (parsed.attachments) {
                for (const att of parsed.attachments) {
                    // Solo si es imagen y tiene Content-ID
                    if (att.contentType && att.contentType.startsWith('image/') && att.cid) {
                        const base64 = att.content.toString('base64');
                        const cid = att.cid.replace(/[<>]/g, '');
                        // Reemplazar todas las ocurrencias de src="cid:..." en el HTML
                        const regex = new RegExp(`src="cid:${cid}"`, 'g');
                        html = html.replace(regex, `src="data:${att.contentType};base64,${base64}"`);
                    }
                }
                // Adjuntos normales (no imágenes inline)
                for (const att of parsed.attachments) {
                    if (!att.contentType.startsWith('image/') || !att.cid) {
                        attachments.push({
                            filename: att.filename || 'adjunto',
                            contentType: att.contentType || 'application/octet-stream',
                            size: att.size || 0,
                            partId: att.partId || ''
                        });
                    }
                }
            }

            if (parsed.to) to = parsed.to.map(a => a.address).join(', ');
            if (parsed.cc) cc = parsed.cc.map(a => a.address).join(', ');

        } catch (parseError) {
            console.error('Error mailparser:', parseError.message);
            // Fallback manual
            const sourceStr = msg.source.toString('utf-8');
            const htmlMatch = sourceStr.match(/Content-Type: text\/html[\s\S]*?\r\n\r\n([\s\S]*?)(?=\r\n--|$)/);
            if (htmlMatch) {
                html = htmlMatch[1].trim();
            } else {
                const textMatch = sourceStr.match(/Content-Type: text\/plain[\s\S]*?\r\n\r\n([\s\S]*?)(?=\r\n--|$)/);
                if (textMatch) {
                    plainText = textMatch[1].trim();
                    html = plainText.replace(/\n/g, '<br>');
                }
            }
        }

        if (!html) html = '<p>No se pudo extraer contenido del mensaje.</p>';

        res.json({
            success: true,
            htmlBody: html,
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
//  ENDPOINT DE DEPURACIÓN (opcional)
// ------------------------------------------------------------
app.post('/api/debug-html', async (req, res) => {
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
        const msg = await client.fetchOne(String(uid), { source: true }, { uid: true });
        await client.logout();
        if (msg && msg.source) {
            const sourceStr = msg.source.toString('utf-8');
            const htmlMatch = sourceStr.match(/Content-Type: text\/html[\s\S]*?\r\n\r\n([\s\S]*?)(?=\r\n--|$)/);
            const htmlPart = htmlMatch ? htmlMatch[1] : 'No se encontró parte HTML';
            res.json({
                success: true,
                raw: sourceStr.substring(0, 1000),
                htmlPart: htmlPart.substring(0, 500),
                fullHtml: htmlPart
            });
        } else {
            res.json({ success: false, error: 'No se pudo obtener el mensaje' });
        }
    } catch (error) {
        console.error('Error /api/debug-html:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------------------------------------------------
//  RESTO DE ENDPOINTS (move, append, delete, toggle, etc.)
// ------------------------------------------------------------
app.post('/api/move-message', async (req, res) => {
    const { email, password, host, port, secure, uid, fromFolder, toFolder } = req.body;
    if (!uid || !fromFolder || !toFolder) {
        return res.status(400).json({ success: false, error: 'Faltan parámetros (uid, fromFolder, toFolder)' });
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

// ------------------------------------------------------------
//  INICIAR SERVIDOR
// ------------------------------------------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`✅ Backend de RSMAIL corriendo en puerto ${PORT}`);
});
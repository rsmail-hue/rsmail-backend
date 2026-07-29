const express = require('express');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const cors = require('cors');
const Imap = require('imap');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const insecureTls = {
    rejectUnauthorized: false,
    checkServerIdentity: () => undefined
};

app.get('/ping', (req, res) => {
    res.json({ alive: true, time: new Date().toISOString() });
});

function createImapClient(email, password, host, port, secure) {
    return new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls,
        connectionTimeout: 60000,
        socketTimeout: 120000,
    });
}

app.post('/api/folders', async (req, res) => {
    const { email, password, host, port, secure } = req.body;
    const client = createImapClient(email, password, host, port, secure);
    let attempts = 0;
    while (attempts < 3) {
        try {
            await client.connect();
            const mailboxes = await client.list();
            await client.logout();
            const folders = mailboxes.map(mbox => ({
                name: mbox.name,
                path: mbox.path,
                specialUse: mbox.specialUse
            }));
            return res.json({ success: true, folders });
        } catch (error) {
            attempts++;
            console.error(`❌ Intento ${attempts} /api/folders falló:`, error.message);
            if (attempts >= 3) {
                return res.status(500).json({ success: false, error: error.message });
            }
            await new Promise(resolve => setTimeout(resolve, 1000 * attempts));
        }
    }
});

app.post('/api/messages', async (req, res) => {
    const { email, password, host, port, secure, folder } = req.body;
    const client = createImapClient(email, password, host, port, secure);
    let attempts = 0;
    while (attempts < 3) {
        try {
            await client.connect();
            await client.mailboxOpen(folder || 'INBOX');
            const messages = [];
            let count = 0;
            const maxMessages = 100;
            for await (const msg of client.fetch('1:*', { envelope: true, bodyStructure: true })) {
                if (count >= maxMessages) break;
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
                count++;
            }
            await client.logout();
            return res.json({ success: true, messages });
        } catch (error) {
            attempts++;
            console.error(`❌ Intento ${attempts} /api/messages falló:`, error.message);
            if (attempts >= 3) {
                return res.status(500).json({ success: false, error: error.message });
            }
            await new Promise(resolve => setTimeout(resolve, 1000 * attempts));
        }
    }
});

app.post('/api/message-detail', async (req, res) => {
    const { email, password, host, port, secure, folder, uid } = req.body;
    if (!email || !password || !uid) {
        return res.status(400).json({ error: 'Faltan parámetros requeridos' });
    }
    const imap = new Imap({
        user: email,
        password: password,
        host: host || 'imap.gmail.com',
        port: port || 993,
        tls: secure !== undefined ? secure : true,
        tlsOptions: { rejectUnauthorized: false },
        connTimeout: 60000,
        keepalive: true
    });
    let responded = false;
    let sendResponse = (data) => {
        if (!responded) {
            responded = true;
            res.json(data);
        }
    };
    let sendError = (msg) => {
        if (!responded) {
            responded = true;
            res.status(500).json({ error: msg });
        }
    };
    imap.once('ready', () => {
        imap.openBox(folder || 'INBOX', true, (err, box) => {
            if (err) {
                imap.end();
                console.error('Error abriendo carpeta:', err);
                return sendError('Error abriendo la carpeta');
            }
            const fetch = imap.fetch([parseInt(uid)], { bodies: '' });
            fetch.on('message', (msg) => {
                msg.on('body', (stream, info) => {
                    simpleParser(stream)
                        .then((parsed) => {
                            let htmlContent = parsed.html || parsed.textAsHtml;
                            if (!htmlContent && parsed.text) {
                                htmlContent = `<pre style="font-family: sans-serif; white-space: pre-wrap;">${parsed.text}</pre>`;
                            }
                            if (!htmlContent) {
                                htmlContent = '<p>(Este correo no contiene texto en el cuerpo)</p>';
                            }
                            if (parsed.attachments && parsed.attachments.length > 0) {
                                parsed.attachments.forEach((att) => {
                                    if (att.contentId && att.related) {
                                        const base64Src = `data:${att.contentType};base64,${att.content.toString('base64')}`;
                                        htmlContent = htmlContent.replace(new RegExp(`cid:${att.contentId}`, 'g'), base64Src);
                                    }
                                });
                            }
                            let to = '';
                            let cc = '';
                            if (parsed.to) {
                                if (Array.isArray(parsed.to)) {
                                    to = parsed.to.map(a => a.address).join(', ');
                                } else {
                                    to = parsed.to.address || parsed.to.text || '';
                                }
                            }
                            if (parsed.cc) {
                                if (Array.isArray(parsed.cc)) {
                                    cc = parsed.cc.map(a => a.address).join(', ');
                                } else {
                                    cc = parsed.cc.address || parsed.cc.text || '';
                                }
                            }
                            const attachmentsList = (parsed.attachments || []).map((att) => ({
                                filename: att.filename || 'adjunto',
                                size: att.size || 0,
                                contentType: att.contentType || 'application/octet-stream',
                                partId: att.contentId || att.filename || ''
                            }));
                            sendResponse({
                                success: true,
                                subject: parsed.subject || '',
                                from: parsed.from ? parsed.from.text : '',
                                to: to,
                                cc: cc,
                                htmlBody: htmlContent,
                                body: parsed.text || '',
                                attachments: attachmentsList
                            });
                            imap.end();
                        })
                        .catch((parseErr) => {
                            console.error('Error al parsear el correo:', parseErr);
                            sendError('Error al procesar el cuerpo del correo');
                            imap.end();
                        });
                });
            });
            fetch.once('error', (fetchErr) => {
                console.error('Fetch error:', fetchErr);
                sendError('Error al obtener el correo de IMAP');
                imap.end();
            });
        });
    });
    imap.once('error', (err) => {
        console.error('IMAP error:', err);
        sendError('Error de conexión IMAP: ' + err.message);
    });
    const timeout = setTimeout(() => {
        if (!responded) {
            sendError('Timeout al obtener el mensaje');
            imap.end();
        }
    }, 60000);
    const originalSend = sendResponse;
    sendResponse = (data) => {
        clearTimeout(timeout);
        originalSend(data);
    };
    sendError = (msg) => {
        clearTimeout(timeout);
        originalSend({ error: msg });
    };
    imap.connect();
});

app.post('/api/discover-folders', async (req, res) => {
    const { email, password, host, port, secure } = req.body;
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        const mailboxes = await client.list();
        await client.logout();
        const folderNames = mailboxes.map(m => m.name || m.path);
        const findFolder = (candidates) => {
            for (const cand of candidates) {
                const found = folderNames.find(f =>
                    f.toLowerCase() === cand.toLowerCase() ||
                    f.toLowerCase().includes(cand.toLowerCase())
                );
                if (found) return found;
            }
            return null;
        };
        const result = {
            trash: findFolder(['Trash', 'Deleted Items', 'Papelera', 'Deleted', 'INBOX.Trash', 'INBOX.Deleted']),
            spam: findFolder(['Spam', 'Junk', 'Junk Email', 'Correo no deseado', 'INBOX.Spam', 'INBOX.Junk']),
            sent: findFolder(['Sent', 'Sent Items', 'Enviados', 'INBOX.Sent']),
            drafts: findFolder(['Drafts', 'Borradores', 'INBOX.Drafts']),
            all: folderNames
        };
        console.log('📂 Carpetas descubiertas:', result);
        res.json({ success: true, folders: result });
    } catch (error) {
        console.error('❌ Error /api/discover-folders:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/move-message', async (req, res) => {
    const { email, password, host, port, secure, uid, fromFolder, toFolder } = req.body;
    if (!uid || !fromFolder || !toFolder) {
        return res.status(400).json({ success: false, error: 'Faltan parámetros' });
    }
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        const mailboxes = await client.list();
        const folderExists = mailboxes.some(m => m.path === toFolder || m.name === toFolder);
        if (!folderExists) {
            try {
                await client.mailboxCreate(toFolder);
                console.log(`✅ Carpeta "${toFolder}" creada.`);
            } catch (createError) {
                const alternatives = {
                    'INBOX.Trash': ['Trash', 'Deleted Items', 'Papelera', 'INBOX.Deleted'],
                    'INBOX.Spam': ['Spam', 'Junk', 'Junk Email', 'INBOX.Junk']
                };
                const altFolders = alternatives[toFolder] || [];
                let found = false;
                for (const alt of altFolders) {
                    if (mailboxes.some(m => m.path === alt || m.name === alt)) {
                        await client.messageMove(uid, alt, { uid: true });
                        found = true;
                        break;
                    }
                }
                if (!found) {
                    await client.messageMove(uid, 'INBOX', { uid: true });
                }
                await client.logout();
                return res.json({ success: true, warning: 'Carpeta no encontrada, se usó alternativa' });
            }
        }
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
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        let folder = sentFolderName || 'Sent';
        const mailboxes = await client.list();
        const folderExists = mailboxes.some(m => m.path === folder || m.name === folder);
        if (!folderExists) {
            const alternatives = ['Sent Items', 'Enviados', 'INBOX.Sent', 'Sent Mail'];
            let found = false;
            for (const alt of alternatives) {
                if (mailboxes.some(m => m.path === alt || m.name === alt)) {
                    folder = alt;
                    found = true;
                    break;
                }
            }
            if (!found) {
                try {
                    await client.mailboxCreate(folder);
                } catch (_) {
                    folder = 'INBOX';
                }
            }
        }
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
    const client = createImapClient(email, password, host, port, secure);
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
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        await client.mailboxOpen(folder);
        if (read) {
            await client.messageFlagsAdd(uid, ['\\Seen'], { uid: true });
        } else {
            await client.messageFlagsRemove(uid, ['\\Seen'], { uid: true });
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
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        await client.mailboxOpen(folder);
        if (flagged) {
            await client.messageFlagsAdd(uid, ['\\Flagged'], { uid: true });
        } else {
            await client.messageFlagsRemove(uid, ['\\Flagged'], { uid: true });
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
    const client = createImapClient(email, password, host, port, secure);
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
    const client = createImapClient(email, password, host, port, secure);
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
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        await client.mailboxOpen(folder);
        const msg = await client.fetchOne(String(uid), { bodyParts: [partId] }, { uid: true });
        await client.logout();
        if (!msg || !msg.bodyParts || !msg.bodyParts[partId]) {
            return res.status(404).json({ success: false, error: 'Adjunto no encontrado' });
        }
        const data = msg.bodyParts[partId].toString('base64');
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
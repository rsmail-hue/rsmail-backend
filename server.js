const express = require('express');
const Imap = require('imap');
const { simpleParser } = require('mailparser');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// ------------------------------------------------------------
//  PING (para verificar que el backend está vivo)
// ------------------------------------------------------------
app.get('/ping', (req, res) => {
    res.json({ alive: true, time: new Date().toISOString() });
});

// ------------------------------------------------------------
//  OBTENER CARPETAS (usando Imap)
// ------------------------------------------------------------
app.post('/api/folders', async (req, res) => {
    const { email, password, host, port, secure } = req.body;

    if (!email || !password) {
        return res.status(400).json({ error: 'Faltan credenciales' });
    }

    const imap = new Imap({
        user: email,
        password: password,
        host: host || 'imap.gmail.com',
        port: port || 993,
        tls: secure !== undefined ? secure : true,
        tlsOptions: { rejectUnauthorized: false }
    });

    imap.once('ready', () => {
        imap.getBoxes((err, boxes) => {
            imap.end();
            if (err) {
                console.error('Error al obtener carpetas:', err);
                return res.status(500).json({ error: 'Error al obtener carpetas' });
            }
            const folders = Object.keys(boxes).map(name => ({
                name: name,
                path: name,
                specialUse: boxes[name].specialUse
            }));
            res.json({ success: true, folders });
        });
    });

    imap.once('error', (err) => {
        console.error('IMAP error:', err);
        res.status(500).json({ error: 'Error de conexión IMAP: ' + err.message });
    });

    imap.connect();
});

// ------------------------------------------------------------
//  OBTENER LISTA DE MENSAJES (usando Imap)
// ------------------------------------------------------------
app.post('/api/messages', async (req, res) => {
    const { email, password, host, port, secure, folder } = req.body;

    if (!email || !password) {
        return res.status(400).json({ error: 'Faltan credenciales' });
    }

    const imap = new Imap({
        user: email,
        password: password,
        host: host || 'imap.gmail.com',
        port: port || 993,
        tls: secure !== undefined ? secure : true,
        tlsOptions: { rejectUnauthorized: false }
    });

    imap.once('ready', () => {
        imap.openBox(folder || 'INBOX', true, (err, box) => {
            if (err) {
                imap.end();
                console.error('Error al abrir carpeta:', err);
                return res.status(500).json({ error: 'Error al abrir carpeta' });
            }

            const messages = [];
            const fetch = imap.fetch('1:*', { bodies: '', struct: true });

            fetch.on('message', (msg, seqno) => {
                const message = { uid: 0, subject: '', from: '', date: '', hasAttachments: false };

                msg.on('attributes', (attrs) => {
                    message.uid = attrs.uid;
                    message.subject = attrs.subject || '(Sin asunto)';
                    message.from = attrs.from ? attrs.from[0].address : email;
                    message.date = attrs.date || new Date();
                    message.hasAttachments = attrs.struct ? attrs.struct.some(node => node.disposition === 'attachment') : false;
                });

                msg.once('end', () => {
                    messages.push(message);
                });
            });

            fetch.once('end', () => {
                imap.end();
                res.json({ success: true, messages });
            });

            fetch.once('error', (fetchErr) => {
                imap.end();
                console.error('Fetch error:', fetchErr);
                res.status(500).json({ error: 'Error al obtener mensajes' });
            });
        });
    });

    imap.once('error', (err) => {
        console.error('IMAP error:', err);
        res.status(500).json({ error: 'Error de conexión IMAP: ' + err.message });
    });

    imap.connect();
});

// ------------------------------------------------------------
//  DETALLE DE MENSAJE (con imap y mailparser - ¡LA CLAVE!)
// ------------------------------------------------------------
app.post('/api/message-detail', async (req, res) => {
    const { email, password, host, port, secure, folder, uid } = req.body;

    if (!email || !password || !uid) {
        return res.status(400).json({ error: 'Faltan parámetros requeridos (email, password, uid)' });
    }

    const imap = new Imap({
        user: email,
        password: password,
        host: host || 'imap.gmail.com',
        port: port || 993,
        tls: secure !== undefined ? secure : true,
        tlsOptions: { rejectUnauthorized: false }
    });

    imap.once('ready', () => {
        imap.openBox(folder || 'INBOX', true, (err, box) => {
            if (err) {
                imap.end();
                console.error('Error abriendo carpeta:', err);
                return res.status(500).json({ error: 'Error abriendo la carpeta' });
            }

            // Buscar el mensaje por su UID
            const fetch = imap.fetch([parseInt(uid)], { bodies: '' });

            fetch.on('message', (msg) => {
                msg.on('body', (stream, info) => {
                    // 🔥 Parsear el stream completo con mailparser usando Promesas
                    simpleParser(stream)
                        .then((parsed) => {
                            // Buscar HTML -> Texto en HTML -> Texto plano -> Mensaje de respaldo
                            let htmlContent = parsed.html || parsed.textAsHtml;
                            if (!htmlContent && parsed.text) {
                                htmlContent = `<pre style="font-family: sans-serif; white-space: pre-wrap;">${parsed.text}</pre>`;
                            }
                            if (!htmlContent) {
                                htmlContent = '<p>(Este correo no contiene texto en el cuerpo)</p>';
                            }

                            // Procesar imágenes inline (cid:) si existen
                            if (parsed.attachments && parsed.attachments.length > 0) {
                                parsed.attachments.forEach((att) => {
                                    if (att.contentId && att.related) {
                                        const base64Src = `data:${att.contentType};base64,${att.content.toString('base64')}`;
                                        htmlContent = htmlContent.replace(new RegExp(`cid:${att.contentId}`, 'g'), base64Src);
                                    }
                                });
                            }

                            // Mapear adjuntos
                            const attachmentsList = (parsed.attachments || []).map((att) => ({
                                filename: att.filename || 'adjunto',
                                size: att.size || 0,
                                contentType: att.contentType || 'application/octet-stream',
                                partId: att.contentId || att.filename || ''
                            }));

                            // Responder a Flutter
                            res.json({
                                success: true,
                                subject: parsed.subject || '',
                                from: parsed.from ? parsed.from.text : '',
                                to: parsed.to ? parsed.to.text : '',
                                cc: parsed.cc ? parsed.cc.text : '',
                                htmlBody: htmlContent,
                                body: parsed.text || '',
                                attachments: attachmentsList
                            });

                            imap.end();
                        })
                        .catch((parseErr) => {
                            console.error('Error al parsear el correo:', parseErr);
                            res.status(500).json({ error: 'Error al procesar el cuerpo del correo' });
                            imap.end();
                        });
                });
            });

            fetch.once('error', (fetchErr) => {
                console.error('Fetch error:', fetchErr);
                res.status(500).json({ error: 'Error al obtener el correo de IMAP' });
                imap.end();
            });
        });
    });

    imap.once('error', (err) => {
        console.error('IMAP error:', err);
        res.status(500).json({ error: 'Error de conexión IMAP: ' + err.message });
    });

    imap.connect();
});

// ------------------------------------------------------------
//  RESTO DE ENDPOINTS (move, append, delete, toggle, etc.)
//  (Se mantienen igual que antes, pero usando ImapFlow para simplicidad)
//  Si quieres, también puedes pasarlos a Imap, pero no es necesario
//  porque solo el detail necesita el stream completo.
// ------------------------------------------------------------
const { ImapFlow } = require('imapflow');
const insecureTls = {
    rejectUnauthorized: false,
    checkServerIdentity: () => undefined
};

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

// ------------------------------------------------------------
//  INICIAR SERVIDOR
// ------------------------------------------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`✅ Backend de RSMAIL corriendo en puerto ${PORT}`);
});
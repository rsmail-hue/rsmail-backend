const express = require('express');
const quotedPrintable = require('quoted-printable');
const { ImapFlow } = require('imapflow');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

const insecureTls = {
    rejectUnauthorized: false,
    checkServerIdentity: () => undefined
};

app.get('/ping', (req, res) => {
    res.json({ alive: true, time: new Date().toISOString() });
});

app.post('/api/folders', async (req, res) => {
    const { email, password, host, port, secure } = req.body;
    console.log('Peticion recibida en /api/folders', email);
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
        const folders = mailboxes.map(mbox => ({
            name: mbox.name,
            path: mbox.path,
            specialUse: mbox.specialUse
        }));
        await client.logout();
        res.json({ success: true, folders });
    } catch (error) {
        console.error('Error en /api/folders:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/messages', async (req, res) => {
    const { email, password, host, port, secure, folder } = req.body;
    console.log('Peticion recibida en /api/messages', email);
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
        for await (const msg of client.fetch('1:*', { envelope: true, flags: true })) {
            messages.push({
                uid: msg.uid,
                subject: msg.envelope.subject || '(Sin asunto)',
                from: (msg.envelope.from && msg.envelope.from[0]) ? msg.envelope.from[0].address : email,
                to: (msg.envelope.to && msg.envelope.to[0]) ? msg.envelope.to[0].address : '',
                date: msg.envelope.date || new Date().toISOString(),
                body: '',
                hasAttachments: false,
                attachments: []
            });
        }
        await client.logout();
        res.json({ success: true, messages });
    } catch (error) {
        console.error('Error en /api/messages:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/message-detail', async (req, res) => {
    const { email, password, host, port, secure, folder, uid } = req.body;
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxOpen(folder || 'INBOX');
        const msg = await client.fetchOne(uid.toString(), { 
            source: true, 
            envelope: true,
            bodyStructure: true 
        }, { uid: true });
        await client.logout();
        
        let body = '';
        let htmlBody = '';
        let hasAttachments = false;
        let attachments = [];
        
        if (msg && msg.source) {
            const fullSource = msg.source.toString();
            
            // Detectar si es multipart
            if (fullSource.includes('Content-Type: multipart/')) {
                // Extraer partes
                const boundaryMatch = fullSource.match(/boundary="([^"]+)"/);
                if (boundaryMatch) {
                    const boundary = boundaryMatch[1];
                    const parts = fullSource.split('--' + boundary);
                    
                    for (const part of parts) {
                        if (part.includes('Content-Type: text/html')) {
                            const htmlStart = part.indexOf('\r\n\r\n');
                            if (htmlStart !== -1) {
                                htmlBody = part.substring(htmlStart + 4).trim();
                            }
                        } else if (part.includes('Content-Type: text/plain') && !htmlBody) {
                            const textStart = part.indexOf('\r\n\r\n');
                            if (textStart !== -1) {
                                body = part.substring(textStart + 4).trim();
                            }
                        } else if (part.includes('Content-Disposition: attachment')) {
                            hasAttachments = true;
                            const nameMatch = part.match(/filename="([^"]+)"/);
                            attachments.push({
                                filename: nameMatch ? nameMatch[1] : 'adjunto',
                                size: 0
                            });
                        }
                    }
                }
            } else {
                // Mensaje simple (no multipart)
                const parts = fullSource.split('\r\n\r\n');
                if (parts.length > 1) {
                    body = parts.slice(1).join('\r\n\r\n').substring(0, 200000);
                }
            }
        }
        
        // Si no hay HTML, usar texto plano
        if (!htmlBody && body) {
            htmlBody = body.replace(/\n/g, '<br>');
        }
        
        res.json({ 
            success: true, 
            body: body,
            htmlBody: htmlBody,
            hasAttachments: hasAttachments,
            attachments: attachments
        });
    } catch (error) {
        console.error('Error en /api/message-detail:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/move-message', async (req, res) => {
    const { email, password, host, port, secure, uid, fromFolder, toFolder } = req.body;
    if (!uid || !fromFolder || !toFolder) return res.status(400).json({ success: false, error: 'Faltan parametros' });
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxOpen(fromFolder);
        await client.messageMove(uid, toFolder, { uid: true });
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error en /api/move-message:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/append-sent', async (req, res) => {
    const { email, password, host, port, secure, rawMessage, sentFolderName } = req.body;
    if (!rawMessage) return res.status(400).json({ success: false, error: 'Falta rawMessage' });
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
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
        console.error('Error en /api/append-sent:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/delete-message', async (req, res) => {
    const { email, password, host, port, secure, uid, folder } = req.body;
    if (!uid || !folder) return res.status(400).json({ success: false, error: 'Faltan parametros' });
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
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
        console.error('Error en /api/delete-message:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/toggle-read', async (req, res) => {
    const { email, password, host, port, secure, uid, folder, read } = req.body;
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        tls: insecureTls
    });
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
        console.error('Error en /api/toggle-read:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/toggle-flagged', async (req, res) => {
    const { email, password, host, port, secure, uid, folder, flagged } = req.body;
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        tls: insecureTls
    });
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
        console.error('Error en /api/toggle-flagged:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/create-folder', async (req, res) => {
    const { email, password, host, port, secure, folderName } = req.body;
    if (!folderName) return res.status(400).json({ success: false, error: 'Falta folderName' });
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxCreate(folderName);
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error en /api/create-folder:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/delete-folder', async (req, res) => {
    const { email, password, host, port, secure, folderName } = req.body;
    if (!folderName) return res.status(400).json({ success: false, error: 'Falta folderName' });
    const client = new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        tls: insecureTls
    });
    try {
        await client.connect();
        await client.mailboxDelete(folderName);
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error en /api/delete-folder:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log('Backend de correo corriendo en puerto ' + PORT);
});



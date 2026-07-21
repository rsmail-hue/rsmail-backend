const express = require('express');
const { ImapFlow } = require('imapflow');
const cors = require('cors');
const quotedPrintable = require('quoted-printable');
const utf8 = require('utf8');

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
        for await (const msg of client.fetch('1:*', { envelope: true })) {
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
            envelope: true
        }, { uid: true });
        await client.logout();
        
        let body = '';
        let htmlBody = '';
        let hasAttachments = false;
        let attachments = [];
        
        if (msg && msg.source) {
            const fullSource = msg.source.toString();
            
            // Extraer partes del mensaje multipart
            if (fullSource.includes('boundary=')) {
                const boundaryMatch = fullSource.match(/boundary="([^"]+)"/) || fullSource.match(/boundary=([^\s]+)/);
                if (boundaryMatch) {
                    const boundary = boundaryMatch[1].replace(/"/g, '');
                    const parts = fullSource.split('--' + boundary);
                    
                    for (const part of parts) {
                        // Detectar Content-Transfer-Encoding
                        let encoding = '';
                        if (part.includes('Content-Transfer-Encoding: quoted-printable')) {
                            encoding = 'quoted-printable';
                        } else if (part.includes('Content-Transfer-Encoding: base64')) {
                            encoding = 'base64';
                        }
                        
                        // Extraer HTML
                        if (part.includes('Content-Type: text/html')) {
                            const headerEnd = part.indexOf('\r\n\r\n');
                            if (headerEnd !== -1) {
                                let html = part.substring(headerEnd + 4).trim();
                                html = html.replace(/--\s*$/, '');
                                
                                if (encoding === 'quoted-printable') {
                                    try {
                                        html = quotedPrintable.decode(html);
                                    } catch(e) {
                                        html = html.replace(/=\r?\n/g, '').replace(/=3D/g, '=').replace(/=22/g, '"').replace(/=20/g, ' ');
                                    }
                                } else if (encoding === 'base64') {
                                    try {
                                        html = Buffer.from(html.replace(/\s/g, ''), 'base64').toString('utf-8');
                                    } catch(e) {}
                                }
                                
                                htmlBody = html;
                            }
                        }
                        
                        // Extraer texto plano
                        if (!htmlBody && part.includes('Content-Type: text/plain')) {
                            const headerEnd = part.indexOf('\r\n\r\n');
                            if (headerEnd !== -1) {
                                let text = part.substring(headerEnd + 4).trim();
                                text = text.replace(/--\s*$/, '');
                                
                                if (encoding === 'quoted-printable') {
                                    try {
                                        text = quotedPrintable.decode(text);
                                    } catch(e) {
                                        text = text.replace(/=\r?\n/g, '').replace(/=3D/g, '=');
                                    }
                                }
                                
                                body = text;
                            }
                        }
                    }
                }
            }
            
            // Si no se encontró HTML, usar el cuerpo simple
            if (!htmlBody && !body) {
                const headerEnd = fullSource.indexOf('\r\n\r\n');
                if (headerEnd !== -1) {
                    body = fullSource.substring(headerEnd + 4).trim();
                    if (fullSource.includes('Content-Transfer-Encoding: base64')) {
                        try {
                            body = Buffer.from(body.replace(/\s/g, ''), 'base64').toString('utf-8');
                        } catch(e) {}
                    }
                }
            }
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

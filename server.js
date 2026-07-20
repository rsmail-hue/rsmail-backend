const express = require('express');
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
        for await (const msg of client.fetch('1:*', { envelope: true, bodyStructure: true })) {
            messages.push({
                uid: msg.uid,
                subject: msg.envelope.subject,
                from: msg.envelope.from[0].address,
                date: msg.envelope.date
            });
        }
        await client.logout();
        res.json({ success: true, messages });
    } catch (error) {
        console.error('Error en /api/messages:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/move-message', async (req, res) => {
    const { email, password, host, port, secure, uid, fromFolder, toFolder } = req.body;
    if (!uid || !fromFolder || !toFolder) {
        return res.status(400).json({ success: false, error: 'Faltan parametros (uid, fromFolder, toFolder)' });
    }

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



// ------------------- IMAP: MENSAJES -------------------
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
        for await (const msg of client.fetch('1:*', { 
            envelope: true, 
            bodyStructure: true
        })) {
            let body = '(No se pudo cargar el contenido)';
            let htmlBody = '';
            let attachments = [];
            
            try {
                const textPart = await client.download(msg.uid, '1', { uid: true });
                body = textPart.toString().substring(0, 50000);
            } catch (e) {
                try {
                    const textPart = await client.download(msg.uid, 'TEXT', { uid: true });
                    body = textPart.toString().substring(0, 50000);
                } catch (e2) {}
            }
            
            try {
                const htmlPart = await client.download(msg.uid, '2', { uid: true });
                htmlBody = htmlPart.toString().substring(0, 50000);
            } catch (e) {}
            
            if (msg.bodyStructure && msg.bodyStructure.childNodes) {
                for (const node of msg.bodyStructure.childNodes) {
                    if (node.disposition === 'attachment' || node.type === 'application') {
                        attachments.push({
                            filename: node.dispositionParameters?.filename || node.parameters?.name || 'adjunto',
                            contentType: node.type + '/' + (node.subtype || 'octet-stream'),
                            size: node.size || 0,
                            partId: node.part
                        });
                    }
                }
            }

            messages.push({
                uid: msg.uid,
                subject: msg.envelope.subject || '(Sin asunto)',
                from: (msg.envelope.from && msg.envelope.from[0]) ? msg.envelope.from[0].address : email,
                to: (msg.envelope.to && msg.envelope.to[0]) ? msg.envelope.to[0].address : '',
                date: msg.envelope.date || new Date().toISOString(),
                body: body,
                htmlBody: htmlBody,
                hasAttachments: attachments.length > 0,
                attachments: attachments
            });
        }
        await client.logout();
        res.json({ success: true, messages });
    } catch (error) {
        console.error('Error en /api/messages:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

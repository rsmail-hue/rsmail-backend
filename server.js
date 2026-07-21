const express = require('express');
const { ImapFlow } = require('imapflow');
const cors = require('cors');
const quotedPrintable = require('quoted-printable');

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
                date: msg.envelope.date || new Date().toISOString(),
                body: '',
                hasAttachments: false
            });
        }
        await client.logout();
        res.json({ success: true, messages });
    } catch (error) {
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
        const msg = await client.fetchOne(uid.toString(), { source: true }, { uid: true });
        await client.logout();
        
        let htmlBody = '';
        if (msg && msg.source) {
            const fullSource = msg.source.toString();
            
            if (fullSource.includes('boundary=')) {
                const boundaryMatch = fullSource.match(/boundary="([^"]+)"/) || fullSource.match(/boundary=([^\s;]+)/);
                if (boundaryMatch) {
                    const boundary = boundaryMatch[1].replace(/"/g, '');
                    const parts = fullSource.split('--' + boundary);
                    
                    for (const part of parts) {
                        if (part.includes('Content-Type: text/html')) {
                            const headerEnd = part.indexOf('\r\n\r\n');
                            if (headerEnd !== -1) {
                                let html = part.substring(headerEnd + 4).trim();
                                html = html.replace(/--\s*$/, '');
                                
                                if (part.includes('quoted-printable')) {
                                    try {
                                        html = quotedPrintable.decode(html);
                                    } catch(e) {
                                        html = html.replace(/=\r?\n/g, '').replace(/=3D/g, '=').replace(/=22/g, '"').replace(/=20/g, ' ').replace(/=C3/g, 'Ã').replace(/=B3/g, '³');
                                    }
                                }
                                htmlBody = html;
                            }
                        }
                    }
                }
            }
        }
        
        res.json({ success: true, htmlBody: htmlBody });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log('Backend corriendo en puerto ' + PORT);
});

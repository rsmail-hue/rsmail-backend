const express = require('express');
const { ImapFlow } = require('imapflow');
const cors = require('cors');
const quotedPrintable = require('quoted-printable');

const app = express();
app.use(cors());
app.use(express.json());

const insecureTls = { rejectUnauthorized: false, checkServerIdentity: () => undefined };

app.get('/ping', (req, res) => res.json({ alive: true }));

app.post('/api/folders', async (req, res) => {
    try {
        const { email, password, host, port, secure } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        const mailboxes = await client.list();
        res.json({ success: true, folders: mailboxes.map(m => ({ name: m.name, path: m.path })) });
        await client.logout();
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/messages', async (req, res) => {
    try {
        const { email, password, host, port, secure, folder } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        await client.mailboxOpen(folder || 'INBOX');
        const msgs = [];
        for await (const m of client.fetch('1:*', { envelope: true })) {
            msgs.push({ uid: m.uid, subject: m.envelope.subject || '', from: m.envelope.from?.[0]?.address || email, date: m.envelope.date });
        }
        await client.logout();
        res.json({ success: true, messages: msgs });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// ------------------- MESSAGE-DETAIL MEJORADO -------------------
app.post('/api/message-detail', async (req, res) => {
    try {
        const { email, password, host, port, secure, folder, uid } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        await client.mailboxOpen(folder || 'INBOX');
        const msg = await client.fetchOne(String(uid), { source: true, envelope: true, bodyStructure: true }, { uid: true });
        await client.logout();

        let html = '';
        let plainText = '';
        let to = '';
        let cc = '';

        if (msg && msg.envelope) {
            if (msg.envelope.to) to = msg.envelope.to.map(a => a.address).join(', ');
            if (msg.envelope.cc) cc = msg.envelope.cc.map(a => a.address).join(', ');
        }

        if (msg && msg.bodyStructure) {
            // Funci?n recursiva para encontrar la parte text/html
            const findHtmlPart = (node) => {
                if (!node) return null;
                if (node.type === 'text' && node.subtype === 'html') return node;
                if (node.childNodes) {
                    for (const child of node.childNodes) {
                        const found = findHtmlPart(child);
                        if (found) return found;
                    }
                }
                return null;
            };
            const htmlPart = findHtmlPart(msg.bodyStructure);
            if (htmlPart && htmlPart.part) {
                try {
                    const { data } = await client.download(msg.uid, htmlPart.part, { uid: true });
                    let raw = data.toString();
                    if (htmlPart.encoding === 'quoted-printable') {
                        raw = quotedPrintable.decode(raw);
                    } else if (htmlPart.encoding === 'base64') {
                        raw = Buffer.from(raw, 'base64').toString('utf-8');
                    }
                    html = raw.substring(0, 200000); // l?mite 200KB
                } catch (e) { /* no se pudo descargar */ }
            }
        }

        // Si no se encontr? HTML con la estructura, intentar con el source
        if (!html && msg && msg.source) {
            const src = msg.source.toString();
            const bm = src.match(/boundary="([^"]+)"/) || src.match(/boundary=([^\s;]+)/);
            if (bm) {
                const boundary = bm[1].replace(/"/g, '');
                const parts = src.split('--' + boundary);
                for (const part of parts) {
                    if (part.includes('Content-Type: text/html')) {
                        const idx = part.indexOf('\r\n\r\n');
                        if (idx > -1) {
                            let raw = part.substring(idx + 4).replace(/--\s*$/, '').trim();
                            if (part.includes('quoted-printable')) {
                                try { raw = quotedPrintable.decode(raw); } catch(ex) { raw = raw.replace(/=\r?\n/g, '').replace(/=([0-9A-Fa-f]{2})/g, (m,c) => String.fromCharCode(parseInt(c,16))); }
                            }
                            html = raw.substring(0, 200000);
                        }
                    }
                    // Extraer texto plano como fallback
                    if (!plainText && part.includes('Content-Type: text/plain')) {
                        const idx = part.indexOf('\r\n\r\n');
                        if (idx > -1) {
                            plainText = part.substring(idx + 4).replace(/--\s*$/, '').trim();
                        }
                    }
                }
            }
        }

        // Si a?n no hay HTML, convertir texto plano en HTML b?sico
        if (!html && plainText) {
            html = '<div style="white-space: pre-wrap; font-family: sans-serif;">' +
                   plainText.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>') +
                   '</div>';
        }

        res.json({ success: true, htmlBody: html, body: plainText, to: to, cc: cc });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// ------------------- RESTO DE ENDPOINTS -------------------
app.post('/api/move-message', async (req, res) => {
    try {
        const { email, password, host, port, secure, uid, fromFolder, toFolder } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        await client.mailboxOpen(fromFolder);
        await client.messageMove(uid, toFolder, { uid: true });
        await client.logout();
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/append-sent', async (req, res) => {
    try {
        const { email, password, host, port, secure, rawMessage, sentFolderName } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        const folder = sentFolderName || 'Sent';
        await client.mailboxOpen(folder);
        await client.append(folder, rawMessage, ['\\Seen']);
        await client.logout();
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/delete-message', async (req, res) => {
    try {
        const { email, password, host, port, secure, uid, folder } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        await client.mailboxOpen(folder);
        await client.messageDelete(uid, { uid: true });
        await client.logout();
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/toggle-read', async (req, res) => {
    try {
        const { email, password, host, port, secure, uid, folder, read } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        await client.mailboxOpen(folder);
        if (read) await client.messageFlagsAdd(uid, ['\\Seen'], { uid: true });
        else await client.messageFlagsRemove(uid, ['\\Seen'], { uid: true });
        await client.logout();
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/toggle-flagged', async (req, res) => {
    try {
        const { email, password, host, port, secure, uid, folder, flagged } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        await client.mailboxOpen(folder);
        if (flagged) await client.messageFlagsAdd(uid, ['\\Flagged'], { uid: true });
        else await client.messageFlagsRemove(uid, ['\\Flagged'], { uid: true });
        await client.logout();
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/create-folder', async (req, res) => {
    try {
        const { email, password, host, port, secure, folderName } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        await client.mailboxCreate(folderName);
        await client.logout();
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/delete-folder', async (req, res) => {
    try {
        const { email, password, host, port, secure, folderName } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        await client.mailboxDelete(folderName);
        await client.logout();
        res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.listen(process.env.PORT || 3000, () => console.log('Backend OK'));

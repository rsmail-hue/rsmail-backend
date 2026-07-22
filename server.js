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

// ------------------- MESSAGE-DETAIL (OBTENCI?N ROBUSTA + CONVERSI?N A HTML ENRIQUECIDO) -------------------
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

        // 1. Obtener el texto plano del mensaje (siempre como fallback)
        if (msg && msg.source) {
            const src = msg.source.toString();
            const plainMatch = src.match(/Content-Type: text\/plain.*?\r\n\r\n([\s\S]*?)(?:\r\n--|$)/i);
            if (plainMatch) {
                plainText = plainMatch[1].trim();
                if (src.includes('quoted-printable')) {
                    try { plainText = quotedPrintable.decode(plainText); }
                    catch(ex) { plainText = plainText.replace(/=\r?\n/g, '').replace(/=([0-9A-Fa-f]{2})/g, (m,c) => String.fromCharCode(parseInt(c,16))); }
                }
            }
            if (!plainText) {
                plainText = src.substring(0, 100000);
            }
        }

        // 2. Intentar extraer HTML real (si existe)
        if (msg && msg.bodyStructure) {
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
                    html = raw.substring(0, 200000);
                } catch (e) {}
            }
        }

        // 3. Si no hay HTML, convertir el texto plano en HTML enriquecido
        if (!html && plainText) {
            let escaped = plainText
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;');
            escaped = escaped.replace(
                /(https?:\/\/[^\s<>"]+)/gi,
                '<a href="" style="color: #0066cc; word-break: break-all;"></a>'
            );
            escaped = escaped
                .split(/\r?\n\r?\n/)
                .map(para => '<p style="margin: 0 0 1em; line-height: 1.5;">' + para.replace(/\n/g, '<br>') + '</p>')
                .join('');
            html = '<div style="font-family: -apple-system, Roboto, sans-serif; font-size: 16px; max-width: 100%; word-wrap: break-word;">' +
                   escaped +
                   '</div>';
        }

        res.json({ success: true, htmlBody: html, body: plainText, to: to, cc: cc });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// ------------------- RESTO DE ENDPOINTS (sin cambios) -------------------
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

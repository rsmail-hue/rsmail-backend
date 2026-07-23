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

// ------------------- MENSAJES (con detecci?n de adjuntos) -------------------
app.post('/api/messages', async (req, res) => {
    try {
        const { email, password, host, port, secure, folder } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        await client.mailboxOpen(folder || 'INBOX');
        const msgs = [];
        for await (const m of client.fetch('1:*', { envelope: true, bodyStructure: true })) {
            let hasAttachments = false;
            if (m.bodyStructure) {
                const checkAttachments = (node) => {
                    if (!node) return;
                    if (node.disposition === 'attachment' || 
                        (node.type === 'application' && node.parameters && node.parameters.name) ||
                        node.type === 'image') {
                        hasAttachments = true;
                    }
                    if (node.childNodes) node.childNodes.forEach(checkAttachments);
                };
                checkAttachments(m.bodyStructure);
            }
            msgs.push({ 
                uid: m.uid, 
                subject: m.envelope.subject || '', 
                from: m.envelope.from?.[0]?.address || email, 
                date: m.envelope.date,
                hasAttachments: hasAttachments 
            });
        }
        await client.logout();
        res.json({ success: true, messages: msgs });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// ------------------- MESSAGE-DETAIL (HTML + ADJUNTOS) -------------------
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
        let attachments = [];

        if (msg && msg.envelope) {
            if (msg.envelope.to) to = msg.envelope.to.map(a => a.address).join(', ');
            if (msg.envelope.cc) cc = msg.envelope.cc.map(a => a.address).join(', ');
        }

        const src = msg?.source?.toString() || '';

        // Funci?n para decodificar QP
        const decodeQP = (text) => {
            try {
                return quotedPrintable.decode(text);
            } catch (e) {
                return text
                    .replace(/=\r?\n/g, '')
                    .replace(/=([0-9A-Fa-f]{2})/g, (m, c) => String.fromCharCode(parseInt(c, 16)));
            }
        };

        // Extraer adjuntos y HTML recursivamente
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
                        raw = decodeQP(raw);
                    } else if (htmlPart.encoding === 'base64') {
                        raw = Buffer.from(raw, 'base64').toString('utf-8');
                    }
                    html = raw.substring(0, 200000);
                } catch (e) {}
            }

            // Extraer adjuntos
            const extractAttachments = (node) => {
                if (!node) return;
                if (node.disposition === 'attachment' || 
                    (node.type === 'application' && node.parameters && node.parameters.name) ||
                    (node.type === 'image' && node.disposition === 'attachment')) {
                    attachments.push({
                        filename: node.dispositionParameters?.filename || node.parameters?.name || 'adjunto',
                        contentType: node.type + '/' + (node.subtype || 'octet-stream'),
                        size: node.size || 0,
                        partId: node.part
                    });
                }
                if (node.childNodes) node.childNodes.forEach(extractAttachments);
            };
            extractAttachments(msg.bodyStructure);
        }

        // Si no hay HTML, extraer del source
        if (!html && src) {
            const bm = src.match(/boundary="([^"]+)"/) || src.match(/boundary=([^\s;]+)/);
            if (bm) {
                const boundary = bm[1].replace(/"/g, '');
                const parts = src.split('--' + boundary);
                for (const part of parts) {
                    if (!html && part.includes('Content-Type: text/html')) {
                        const idx = part.indexOf('\r\n\r\n');
                        if (idx > -1) {
                            let raw = part.substring(idx + 4).replace(/--\s*$/, '').trim();
                            if (part.includes('quoted-printable')) {
                                raw = decodeQP(raw);
                            }
                            html = raw.substring(0, 200000);
                        }
                    }
                    if (!html && !plainText && part.includes('Content-Type: text/plain')) {
                        const idx = part.indexOf('\r\n\r\n');
                        if (idx > -1) {
                            let text = part.substring(idx + 4).replace(/--\s*$/, '').trim();
                            if (part.includes('quoted-printable')) {
                                text = decodeQP(text);
                            }
                            plainText = text;
                        }
                    }
                }
            } else {
                const headerEnd = src.indexOf('\r\n\r\n');
                if (headerEnd > -1) {
                    let body = src.substring(headerEnd + 4).trim();
                    if (src.includes('quoted-printable')) {
                        body = decodeQP(body);
                    }
                    plainText = body;
                }
            }
        }

        // Convertir texto plano en HTML si no hay HTML real
        if (!html && plainText) {
            if (/<(!DOCTYPE|html|head|body|div|table|style|script|p|br|hr|img|a|meta|link)/i.test(plainText)) {
                html = plainText.substring(0, 200000);
            } else {
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
        }

        if (!html) html = '<p>No se pudo extraer contenido.</p>';

        res.json({ success: true, htmlBody: html, body: plainText, to: to, cc: cc, attachments: attachments });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// ------------------- DESCARGA DE ADJUNTOS -------------------
app.post('/api/download-attachment', async (req, res) => {
    try {
        const { email, password, host, port, secure, folder, uid, partId } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        await client.mailboxOpen(folder || 'INBOX');
        const { data } = await client.download(parseInt(uid), partId, { uid: true });
        await client.logout();
        res.json({ success: true, data: data.toString('base64') });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// ------------------- RESTO DE ENDPOINTS (sin cambios) -------------------
app.post('/api/move-message', async (req, res) => { /* ... mismo c?digo ... */ });
app.post('/api/append-sent', async (req, res) => { /* ... */ });
app.post('/api/delete-message', async (req, res) => { /* ... */ });
app.post('/api/toggle-read', async (req, res) => { /* ... */ });
app.post('/api/toggle-flagged', async (req, res) => { /* ... */ });
app.post('/api/create-folder', async (req, res) => { /* ... */ });
app.post('/api/delete-folder', async (req, res) => { /* ... */ });

app.listen(process.env.PORT || 3000, () => console.log('Backend OK'));

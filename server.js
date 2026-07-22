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

// ---------- NUEVO message-detail mejorado ----------
app.post('/api/message-detail', async (req, res) => {
    try {
        const { email, password, host, port, secure, folder, uid } = req.body;
        const client = new ImapFlow({ host: host || 'imap.gmail.com', port: port || 993, secure: secure !== undefined ? secure : true, auth: { user: email, pass: password }, tls: insecureTls });
        await client.connect();
        await client.mailboxOpen(folder || 'INBOX');
        const msg = await client.fetchOne(String(uid), { source: true, envelope: true }, { uid: true });
        await client.logout();

        let html = '';
        let plainText = '';
        let to = '';
        let cc = '';

        if (msg && msg.envelope) {
            if (msg.envelope.to) to = msg.envelope.to.map(a => a.address).join(', ');
            if (msg.envelope.cc) cc = msg.envelope.cc.map(a => a.address).join(', ');
        }

        if (msg && msg.source) {
            const full = msg.source.toString();
            // Funci?n recursiva para extraer la parte text/html
            const extractHtml = (part) => {
                if (!part) return null;
                // Si esta parte es text/html
                if (part.type === 'text' && part.subtype === 'html') {
                    return part;
                }
                // Si es multipart, buscar en sus hijos
                if (part.childNodes) {
                    for (const child of part.childNodes) {
                        const found = extractHtml(child);
                        if (found) return found;
                    }
                }
                // Tambi?n buscar en partes con disposition inline
                return null;
            };
            if (msg.bodyStructure) {
                const htmlPart = extractHtml(msg.bodyStructure);
                if (htmlPart && htmlPart.part) {
                    try {
                        const { data } = await client.download(msg.uid, htmlPart.part, { uid: true });
                        let raw = data.toString();
                        // Decodificar si es necesario
                        if (htmlPart.encoding === 'quoted-printable') {
                            raw = quotedPrintable.decode(raw);
                        } else if (htmlPart.encoding === 'base64') {
                            raw = Buffer.from(raw, 'base64').toString('utf-8');
                        }
                        html = raw;
                    } catch (e) { /* no se pudo descargar esa parte */ }
                }
            }
            // Si no se encontr? HTML con la estructura, intentar parseando manualmente el source
            if (!html) {
                const bm = full.match(/boundary="([^"]+)"/) || full.match(/boundary=([^\s;]+)/);
                if (bm) {
                    const boundary = bm[1].replace(/"/g, '');
                    const parts = full.split('--' + boundary);
                    for (const part of parts) {
                        if (part.includes('Content-Type: text/html')) {
                            const idx = part.indexOf('\r\n\r\n');
                            if (idx > -1) {
                                let raw = part.substring(idx + 4).replace(/--\s*$/, '').trim();
                                if (part.includes('quoted-printable')) {
                                    try { raw = quotedPrintable.decode(raw); }
                                    catch(ex) { raw = raw.replace(/=\r?\n/g, '').replace(/=([0-9A-Fa-f]{2})/g, (m,c) => String.fromCharCode(parseInt(c,16))); }
                                }
                                html = raw;
                            }
                        }
                    }
                }
            }
            // Extraer texto plano como fallback
            const plainPart = full.match(/Content-Type: text\/plain.*?\r\n\r\n([\s\S]*?)(?:\r\n--|$)/i);
            if (plainPart) {
                plainText = plainPart[1].trim();
            }
        }

        // Si no hay HTML, convertir texto plano en HTML b?sico
        if (!html && plainText) {
            html = '<div style="white-space: pre-wrap; font-family: sans-serif;">' +
                   plainText.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>') +
                   '</div>';
        }

        res.json({ success: true, htmlBody: html, body: plainText, to: to, cc: cc });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// ... (el resto de endpoints se mantienen igual, no los modifiques)

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Backend OK'));

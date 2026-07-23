const express = require('express');
const { ImapFlow } = require('imapflow');
const cors = require('cors');
const { simpleParser } = require('mailparser');

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
        for await (const m of client.fetch('1:*', { envelope: true, bodyStructure: true })) {
            let hasAttachments = false;
            if (m.bodyStructure) {
                const checkAttachments = (node) => {
                    if (!node) return;
                    if (node.disposition === 'attachment' || 
                        (node.type === 'application' && node.parameters && node.parameters.name) ||
                        node.type === 'image') hasAttachments = true;
                    if (node.childNodes) node.childNodes.forEach(checkAttachments);
                };
                checkAttachments(m.bodyStructure);
            }
            msgs.push({ uid: m.uid, subject: m.envelope.subject || '', from: m.envelope.from?.[0]?.address || email, date: m.envelope.date, hasAttachments });
        }
        await client.logout();
        res.json({ success: true, messages: msgs });
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// ------------------- MESSAGE-DETAIL (con mailparser) -------------------
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

        // Usar mailparser para decodificar el mensaje completo
        if (msg && msg.source) {
            try {
                const parsed = await simpleParser(msg.source);
                // Si hay HTML, lo usamos; si no, texto plano
                if (parsed.html) {
                    html = parsed.html.substring(0, 200000);
                } else if (parsed.text) {
                    // Convertir texto plano a HTML b?sico (con saltos de l?nea)
                    html = parsed.text
                        .replace(/&/g, '&amp;')
                        .replace(/</g, '&lt;')
                        .replace(/>/g, '&gt;')
                        .replace(/\r?\n/g, '<br>');
                    html = '<div style="font-family: -apple-system, Roboto, sans-serif; font-size: 16px; max-width: 100%; word-wrap: break-word;">' + html + '</div>';
                }

                // Extraer adjuntos desde la estructura (m?s fiable que mailparser en algunos casos)
                if (msg.bodyStructure) {
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

                // Tambi?n extraer adjuntos desde parsed (si no se encontraron en bodyStructure)
                if (attachments.length === 0 && parsed.attachments) {
                    parsed.attachments.forEach(att => {
                        attachments.push({
                            filename: att.filename || 'adjunto',
                            contentType: att.contentType || 'application/octet-stream',
                            size: att.size || 0,
                            partId: att.partId || ''
                        });
                    });
                }
            } catch (e) {
                // Si falla mailparser, intentamos el m?todo antiguo como fallback
                console.error('mailparser error:', e.message);
            }
        }

        // Si mailparser no pudo obtener contenido, devolver un mensaje gen?rico
        if (!html) {
            html = '<p>No se pudo extraer contenido del mensaje.</p>';
        }

        res.json({ success: true, htmlBody: html, body: plainText, to: to, cc: cc, attachments });
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

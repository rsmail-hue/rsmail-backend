const express = require('express');
const cors = require('cors');
const { ImapFlow } = require('imapflow');
const nodemailer = require('nodemailer');
const { simpleParser } = require('mailparser');

const app = express();
app.use(cors());
app.use(express.json({ limit: '25mb' }));

// Helper de Autoconfiguración de dominios
function getAutoConfig(email) {
    const domain = email.split('@')[1]?.toLowerCase() || '';

    if (domain.includes('gmail.com')) {
        return { imapHost: 'imap.gmail.com', imapPort: 993, smtpHost: 'smtp.gmail.com', smtpPort: 465 };
    }
    if (domain.includes('outlook.com') || domain.includes('hotmail.com') || domain.includes('live.com')) {
        return { imapHost: 'outlook.office365.com', imapPort: 993, smtpHost: 'smtp.office365.com', smtpPort: 587 };
    }
    if (domain.includes('yahoo.')) {
        return { imapHost: 'imap.mail.yahoo.com', imapPort: 993, smtpHost: 'smtp.mail.yahoo.com', smtpPort: 465 };
    }
    if (domain.includes('mail.ru') || domain.includes('inbox.ru') || domain.includes('list.ru') || domain.includes('bk.ru')) {
        return { imapHost: 'imap.mail.ru', imapPort: 993, smtpHost: 'smtp.mail.ru', smtpPort: 465 };
    }

    return {
        imapHost: 'mail.' + domain,
        imapPort: 993,
        smtpHost: 'mail.' + domain,
        smtpPort: 587
    };
}

// Conexión IMAP
async function getImapClient(email, password, customHost, customPort) {
    const auto = getAutoConfig(email);
    const client = new ImapFlow({
        host: customHost || auto.imapHost,
        port: Number(customPort) || auto.imapPort,
        secure: true,
        auth: { user: email, pass: password },
        logger: false,
        tls: { rejectUnauthorized: false }
    });
    await client.connect();
    return client;
}

// ==========================================
// 1. VERIFICAR CREDENCIALES
// ==========================================
app.post('/api/verify', async (req, res) => {
    const { email, password, host, port } = req.body;
    try {
        const client = await getImapClient(email, password, host, port);
        await client.logout();
        const auto = getAutoConfig(email);
        res.json({
            success: true,
            config: {
                imapHost: host || auto.imapHost,
                imapPort: port || auto.imapPort,
                smtpHost: auto.smtpHost,
                smtpPort: auto.smtpPort
            }
        });
    } catch (err) {
        res.status(401).json({ success: false, error: err.message });
    }
});

// ==========================================
// 2. LISTAR CARPETAS
// ==========================================
app.post('/api/folders', async (req, res) => {
    const { email, password, host, port } = req.body;
    try {
        const client = await getImapClient(email, password, host, port);
        const list = await client.list();
        await client.logout();

        const folders = list.map(f => ({
            name: f.name,
            path: f.path,
            specialUse: f.specialUse || ''
        }));

        res.json({ success: true, folders });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// ==========================================
// 3. MENSAJES (MÁS NUEVOS PRIMERO)
// ==========================================
app.post('/api/messages', async (req, res) => {
    const { email, password, host, port, folder = 'INBOX', limit = 20 } = req.body;

    try {
        const client = await getImapClient(email, password, host, port);
        const lock = await client.getMailboxLock(folder);
        const messages = [];

        try {
            // "reverse: true" con "1:*" garantiza traer siempre los más recientes primero
            for await (let message of client.fetch('1:*', { envelope: true }, { max: Number(limit), reverse: true })) {
                messages.push({
                    uid: message.uid,
                    id: message.uid.toString(),
                    subject: message.envelope.subject || '(Sin asunto)',
                    from: message.envelope.from?.[0]?.address || message.envelope.from?.[0]?.name || '',
                    to: message.envelope.to?.[0]?.address || '',
                    date: message.envelope.date ? new Date(message.envelope.date).toISOString() : new Date().toISOString(),
                });
            }
        } finally {
            lock.release();
        }

        await client.logout();
        res.json({ success: true, messages, total: messages.length });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// ==========================================
// 4. DETALLE DEL MENSAJE
// ==========================================
app.post('/api/message-detail', async (req, res) => {
    const { email, password, host, port, folder = 'INBOX', uid } = req.body;

    if (!uid) return res.status(400).json({ success: false, error: 'UID requerido' });

    try {
        const client = await getImapClient(email, password, host, port);
        const lock = await client.getMailboxLock(folder);
        let parsedEmail = null;

        try {
            const message = await client.fetchOne(uid.toString(), { source: true }, { uid: true });
            if (message && message.source) {
                parsedEmail = await simpleParser(message.source);
            }
        } finally {
            lock.release();
        }

        await client.logout();

        if (!parsedEmail) return res.status(404).json({ success: false, error: 'Mensaje no encontrado' });

        const attachments = (parsedEmail.attachments || []).map(att => ({
            filename: att.filename || 'archivo',
            contentType: att.contentType,
            size: att.size,
            content: att.content ? att.content.toString('base64') : ''
        }));

        res.json({
            success: true,
            message: {
                uid: Number(uid),
                subject: parsedEmail.subject || '(Sin asunto)',
                from: parsedEmail.from?.text || '',
                to: parsedEmail.to?.text || '',
                date: parsedEmail.date ? parsedEmail.date.toISOString() : new Date().toISOString(),
                text: parsedEmail.text || '',
                html: parsedEmail.html || parsedEmail.textAsHtml || parsedEmail.text || '',
                attachments
            }
        });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// ==========================================
// 5. ENVIAR CORREO (SMTP MULTI-PUERTO BLINDADO)
// ==========================================
app.post('/api/send-email', async (req, res) => {
    const { email, password, smtpHost, smtpPort, to, subject, text, html, attachments } = req.body;

    const auto = getAutoConfig(email);
    const host = smtpHost || auto.smtpHost;

    // Intentar envío con los parámetros indicados o con fallback automático
    const portsToTry = smtpPort
        ? [Number(smtpPort), 587, 465]
        : [auto.smtpPort, 587, 465];

    let lastError = null;

    for (const port of [...new Set(portsToTry)]) {
        const isSecure = (port === 465);
        try {
            const transporter = nodemailer.createTransport({
                host: host,
                port: port,
                secure: isSecure,
                auth: { user: email, pass: password },
                tls: { rejectUnauthorized: false },
                connectionTimeout: 10000
            });

            const info = await transporter.sendMail({
                from: email,
                to,
                subject: subject || '(Sin asunto)',
                text: text || '',
                html: html || text || '',
                attachments: (attachments || []).map(att => ({
                    filename: att.filename,
                    content: Buffer.from(att.content, 'base64')
                }))
            });

            return res.json({ success: true, messageId: info.messageId });
        } catch (err) {
            console.error(`Falló puerto ${port}:`, err.message);
            lastError = err;
        }
    }

    return res.status(500).json({
        success: false,
        error: 'No se pudo enviar el correo por ningún puerto SMTP: ' + (lastError ? lastError.message : 'Error desconocido')
    });
});

// ==========================================
// 6. ELIMINAR / MOVER A LA PAPELERA
// ==========================================
app.post('/api/delete-message', async (req, res) => {
    const { email, password, host, port, folder = 'INBOX', uid } = req.body;

    if (!uid) return res.status(400).json({ success: false, error: 'UID requerido' });

    try {
        const client = await getImapClient(email, password, host, port);
        const lock = await client.getMailboxLock(folder);

        try {
            // Intentar buscar la carpeta Trash/Papelera
            const list = await client.list();
            const trashFolder = list.find(f => f.specialUse === '\\Trash' || f.name.toLowerCase().includes('trash') || f.name.toLowerCase().includes('papelera'));

            if (trashFolder && trashFolder.path !== folder) {
                // Mover a la papelera
                await client.messageMove(uid.toString(), trashFolder.path, { uid: true });
            } else {
                // Si ya estamos en la papelera o no se encuentra, marcar como borrado
                await client.messageDelete(uid.toString(), { uid: true });
            }
        } finally {
            lock.release();
        }

        await client.logout();
        res.json({ success: true, message: 'Mensaje eliminado o movido a la papelera' });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// INICIAR SERVIDOR
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Servidor listo en el puerto ${PORT}`);
});
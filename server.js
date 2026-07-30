const express = require('express');
const cors = require('cors');
const { ImapFlow } = require('imapflow');
const nodemailer = require('nodemailer');
const { simpleParser } = require('mailparser');

const app = express();
app.use(cors());
app.use(express.json({ limit: '25mb' }));

// Helper de configuración para cPanel / Servidores estándar
function getAutoConfig(email) {
    const domain = email.split('@')[1]?.toLowerCase() || '';

    if (domain.includes('gmail.com')) {
        return { imapHost: 'imap.gmail.com', imapPort: 993, smtpHost: 'smtp.gmail.com', smtpPort: 587, smtpSecure: false };
    }
    if (domain.includes('outlook.com') || domain.includes('hotmail.com')) {
        return { imapHost: 'outlook.office365.com', imapPort: 993, smtpHost: 'smtp.office365.com', smtpPort: 587, smtpSecure: false };
    }

    // Configuración cPanel para rsmicro.es y dominios propios
    return {
        imapHost: 'mail.' + domain,
        imapPort: 993,
        smtpHost: 'mail.' + domain,
        smtpPort: 587,
        smtpSecure: false // 587 usa STARTTLS (secure: false)
    };
}

// Conexión IMAP Robusta
async function getImapClient(email, password, customHost, customPort) {
    const auto = getAutoConfig(email);
    const client = new ImapFlow({
        host: customHost || auto.imapHost,
        port: Number(customPort) || auto.imapPort,
        secure: true, // IMAP siempre 993 TLS
        auth: { user: email, pass: password },
        logger: false,
        tls: { rejectUnauthorized: false }
    });
    await client.connect();
    return client;
}

// ==========================================
// 1. VERIFICAR Y AUTO-CONFIGURAR
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
// 2. LISTAR CARPETAS (Mapeado Correcto)
// ==========================================
app.post('/api/folders', async (req, res) => {
    const { email, password, host, port } = req.body;
    try {
        const client = await getImapClient(email, password, host, port);
        const list = await client.list();
        await client.logout();

        const folders = list.map(f => ({
            name: f.name,
            path: f.path, // 👈 Importante: Flutter debe enviar este "path" (ej: "INBOX.Sent")
            specialUse: f.specialUse || ''
        }));

        res.json({ success: true, folders });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// ==========================================
// 3. OBTENER MENSAJES (Los más recientes primero)
// ==========================================
app.post('/api/messages', async (req, res) => {
    const { email, password, host, port, folder = 'INBOX', limit = 25 } = req.body;

    try {
        const client = await getImapClient(email, password, host, port);
        const lock = await client.getMailboxLock(folder);
        const messages = [];

        try {
            // "reverse: true" obliga a traer los emails recién llegados primero
            for await (let message of client.fetch('1:*', { envelope: true, bodyStructure: true }, { max: Number(limit), reverse: true })) {
                messages.push({
                    uid: message.uid,
                    id: message.uid.toString(),
                    subject: message.envelope.subject || '(Sin asunto)',
                    from: message.envelope.from?.[0]?.address || message.envelope.from?.[0]?.name || 'Desconocido',
                    to: message.envelope.to?.[0]?.address || '',
                    date: message.envelope.date ? new Date(message.envelope.date).toISOString() : new Date().toISOString()
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
// 4. DETALLE DEL CORREO (Lee Texto/HTML real)
// ==========================================
app.post('/api/message-detail', async (req, res) => {
    const { email, password, host, port, folder = 'INBOX', uid } = req.body;

    if (!uid) return res.status(400).json({ success: false, error: 'UID de mensaje requerido' });

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
            filename: att.filename || 'adjunto',
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
// 5. ENVIAR CORREO (Solución Error 07 / cPanel)
// ==========================================
app.post('/api/send-email', async (req, res) => {
    const { email, password, smtpHost, smtpPort, to, subject, text, html, attachments } = req.body;

    const auto = getAutoConfig(email);
    const host = smtpHost || auto.smtpHost;

    // cPanel requiere probar SSL (465) y STARTTLS (587) alternativamente si uno falla
    const configurations = [
        { port: 587, secure: false }, // STARTTLS (Por defecto en cPanel)
        { port: 465, secure: true },  // SSL Directo
        { port: 25, secure: false }   // Fallback
    ];

    let lastError = null;

    for (const config of configurations) {
        try {
            const transporter = nodemailer.createTransport({
                host: host,
                port: config.port,
                secure: config.secure,
                auth: { user: email, pass: password },
                tls: { rejectUnauthorized: false },
                connectionTimeout: 8000
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
            console.log(`Intentó puerto ${config.port} pero falló: ${err.message}`);
            lastError = err;
        }
    }

    res.status(500).json({ success: false, error: 'Error al enviar por SMTP: ' + (lastError ? lastError.message : 'Error de conexión') });
});

// ==========================================
// 6. ELIMINAR / PAPELERA (Especial para cPanel)
// ==========================================
app.post('/api/delete-message', async (req, res) => {
    const { email, password, host, port, folder = 'INBOX', uid } = req.body;

    if (!uid) return res.status(400).json({ success: false, error: 'UID requerido' });

    try {
        const client = await getImapClient(email, password, host, port);
        const lock = await client.getMailboxLock(folder);

        try {
            const list = await client.list();
            // Buscar carpeta de papelera (en cPanel suele llamarse INBOX.Trash)
            const trashFolder = list.find(f =>
                f.specialUse === '\\Trash' ||
                f.path.toUpperCase().includes('TRASH') ||
                f.path.toUpperCase().includes('PAPELERA')
            );

            if (trashFolder && trashFolder.path !== folder) {
                // Mover a la papelera
                await client.messageMove(uid.toString(), trashFolder.path, { uid: true });
            } else {
                // Si ya estaba en la papelera, borrar definitivamente
                await client.messageDelete(uid.toString(), { uid: true });
            }
        } finally {
            lock.release();
        }

        await client.logout();
        res.json({ success: true, message: 'Mensaje movido o eliminado' });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// PUERTO
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Servidor cPanel/IMAP listo en el puerto ${PORT}`);
});
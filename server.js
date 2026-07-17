const express = require('express');
const { ImapFlow } = require('imapflow');
const cors = require('cors');
const nodemailer = require('nodemailer');
const multer = require('multer');
const upload = multer({ dest: 'uploads/' }); // carpeta temporal para adjuntos

const app = express();
app.use(cors());
app.use(express.json());

// Configuración TLS que ignora TODO: CA, caducidad, nombre del host
const insecureTls = {
    rejectUnauthorized: false,
    checkServerIdentity: () => undefined
};

// Ruta de prueba para verificar que Express está vivo
app.get('/ping', (req, res) => {
    res.json({ alive: true, time: new Date().toISOString() });
});

// ------------------- IMAP: CARPETAS -------------------
app.post('/api/folders', async (req, res) => {
    const { email, password, host, port, secure } = req.body;
    console.log('Petición recibida en /api/folders', email);

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

// ------------------- IMAP: MENSAJES -------------------
app.post('/api/messages', async (req, res) => {
    const { email, password, host, port, secure, folder } = req.body;
    console.log('Petición recibida en /api/messages', email);

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

// ------------------- NUEVO: ENVÍO SIMPLE (JSON) -------------------
app.post('/api/send', async (req, res) => {
    try {
        const { email, password, host, port, secure, to, subject, body,
                smtpHost, smtpPort, smtpSecure, sentFolderName } = req.body;

        // Construir transporte SMTP con las credenciales del usuario
        const transporter = nodemailer.createTransport({
            host: smtpHost || host || 'smtp.office365.com',
            port: parseInt(smtpPort || port, 10) || 587,
            secure: (smtpSecure || secure) === true,
            auth: {
                user: email,
                pass: password
            },
            tls: {
                rejectUnauthorized: false, // igual que en IMAP, ignora certificados
                checkServerIdentity: () => undefined
            }
        });

        const info = await transporter.sendMail({
            from: email,
            to: to,
            subject: subject,
            text: body
        });

        // Opcional: guardar en carpeta Enviados (requiere IMAP)
        if (sentFolderName) {
            try {
                const imapClient = new ImapFlow({
                    host: host || 'imap.gmail.com',
                    port: port || 993,
                    secure: secure !== undefined ? secure : true,
                    auth: { user: email, pass: password },
                    tls: insecureTls
                });
                await imapClient.connect();
                await imapClient.mailboxOpen(sentFolderName);
                await imapClient.append(body, { flags: ['\\Seen'] });
                await imapClient.logout();
            } catch (imapError) {
                console.warn('No se pudo guardar en Enviados:', imapError.message);
            }
        }

        res.json({ success: true, messageId: info.messageId });
    } catch (error) {
        console.error('Error en /api/send:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------- NUEVO: ENVÍO CON ADJUNTOS (multipart) -------------------
app.post('/api/send-attachments', upload.array('attachments', 10), async (req, res) => {
    try {
        const { email, password, host, port, secure, to, subject, body,
                smtpHost, smtpPort, smtpSecure, sentFolderName } = req.body;

        // Procesar archivos adjuntos subidos
        const attachments = (req.files || []).map(file => ({
            filename: file.originalname,
            path: file.path
        }));

        const transporter = nodemailer.createTransport({
            host: smtpHost || host || 'smtp.office365.com',
            port: parseInt(smtpPort || port, 10) || 587,
            secure: (smtpSecure || secure) === true,
            auth: {
                user: email,
                pass: password
            },
            tls: {
                rejectUnauthorized: false,
                checkServerIdentity: () => undefined
            }
        });

        const info = await transporter.sendMail({
            from: email,
            to: to,
            subject: subject,
            text: body,
            attachments: attachments
        });

        // Limpiar archivos temporales
        (req.files || []).forEach(file => {
            try { require('fs').unlinkSync(file.path); } catch (_) {}
        });

        res.json({ success: true, messageId: info.messageId });
    } catch (error) {
        console.error('Error en /api/send-attachments:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Backend de correo corriendo en puerto ${PORT}`);
});
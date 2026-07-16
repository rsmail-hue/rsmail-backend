const express = require('express');
const { ImapFlow } = require('imapflow');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

// Configuración TLS que ignora TODO: CA, caducidad, nombre del host
const insecureTls = {
    rejectUnauthorized: false,
    checkServerIdentity: () => undefined   // <-- ignora el mismatch de hostname
};

// Ruta de prueba para verificar que Express está vivo
app.get('/ping', (req, res) => {
    res.json({ alive: true, time: new Date().toISOString() });
});

// Endpoint para listar carpetas
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

// Endpoint para obtener mensajes
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Backend de correo corriendo en puerto ${PORT}`);
});
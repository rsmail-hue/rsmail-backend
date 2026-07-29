const express = require('express');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const cors = require('cors');
const Imap = require('imap');
const nodemailer = require('nodemailer');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const insecureTls = {
    rejectUnauthorized: false,
    checkServerIdentity: () => undefined
};

// ============================================================
//  CONNECTION POOL PARA SMTP (reutilizar conexiones)
// ============================================================
let transporterCache = {};

// ============================================================
//  ENVIAR CORREO + GUARDAR EN ENVIADOS (CORREGIDO)
// ============================================================
app.post('/api/send-email', async (req, res) => {
    let { email, password, smtpHost, smtpPort, secure, imapHost, to, subject, body, cc, bcc } = req.body;

    const domain = email.split('@')[1];
    if (!smtpHost || smtpHost.includes('smtp.')) smtpHost = `mail.${domain}`;
    if (!imapHost) imapHost = `mail.${domain}`;

    console.log(`📨 Enviando correo desde ${email} a ${to}`);

    // 1️⃣ Intentar SMTP en puertos 587 (STARTTLS) y 465 (SSL)
    const portsToTry = smtpPort ? [smtpPort] : [587, 465];
    let sendError = null;

    for (const port of portsToTry) {
        try {
            const transporter = nodemailer.createTransport({
                host: smtpHost,
                port: port,
                secure: port === 465,        // true solo para 465
                auth: { user: email, pass: password },
                tls: { rejectUnauthorized: false },
                connectionTimeout: 10000,
                greetingTimeout: 10000,
                socketTimeout: 10000,
            });

            const info = await transporter.sendMail({
                from: email,
                to: to,
                cc: cc || undefined,
                bcc: bcc || undefined,
                subject: subject,
                html: body,
            });

            console.log(`✅ Correo enviado por puerto ${port}: ${info.messageId}`);
            sendError = null;
            break; // Éxito, salir del bucle
        } catch (err) {
            console.log(`❌ Puerto ${port} falló: ${err.message}`);
            sendError = err;
        }
    }

    if (sendError) {
        return res.status(500).json({ success: false, error: 'Error SMTP: ' + sendError.message });
    }

    // 2️⃣ Guardar copia en Enviados (usando la carpeta real del servidor)
    const imap = new Imap({
        user: email,
        password: password,
        host: imapHost,
        port: 993,
        tls: true,
        tlsOptions: { rejectUnauthorized: false },
    });

    imap.once('ready', () => {
        imap.getBoxes((err, boxes) => {
            if (err) {
                console.error('Error obteniendo carpetas:', err);
                imap.end();
                return;
            }

            const sentFolder = findSentFolder(boxes);
            console.log(`📂 Carpeta Enviados detectada: ${sentFolder || 'ninguna, se omite guardado'}`);

            if (!sentFolder) {
                imap.end();
                return;
            }

            imap.openBox(sentFolder, false, (openErr) => {
                if (openErr) {
                    console.error('Error abriendo carpeta Enviados:', openErr);
                    imap.end();
                    return;
                }

                const rawMessage = [
                    `From: ${email}`,
                    `To: ${to}`,
                    `Subject: ${subject}`,
                    `Content-Type: text/html; charset=utf-8`,
                    '',
                    body
                ].join('\r\n');

                imap.append(rawMessage, { mailbox: sentFolder, flags: ['\\Seen'] }, (appendErr) => {
                    if (appendErr) console.error('Error al guardar en Enviados:', appendErr);
                    else console.log('✅ Copia guardada en Enviados');
                    imap.end();
                });
            });
        });
    });

    imap.once('error', (err) => {
        console.error('Error IMAP al guardar en Sent:', err.message);
        imap.end();
    });

    imap.connect();

    // Responder inmediatamente (el guardado en Enviados es en segundo plano)
    return res.json({ success: true, message: 'Correo enviado correctamente' });
});

// ------------------------------------------------------------
//  FUNCIÓN AUXILIAR: buscar carpeta de enviados
// ------------------------------------------------------------
function findSentFolder(boxes) {
    const patterns = [/^sent$/i, /enviado/i, /inbox\.sent/i];

    function search(boxObj, prefix = '') {
        for (const key in boxObj) {
            const fullPath = prefix ? `${prefix}${boxObj[key].delimiter}${key}` : key;
            // Atributos IMAP (\Sent)
            if (boxObj[key].attribs && boxObj[key].attribs.map(a => a.toLowerCase()).includes('\\sent')) {
                return fullPath;
            }
            // Coincidencia por nombre
            for (const reg of patterns) {
                if (reg.test(key) || reg.test(fullPath)) return fullPath;
            }
            if (boxObj[key].children) {
                const found = search(boxObj[key].children, fullPath);
                if (found) return found;
            }
        }
        return null;
    }

    return search(boxes);
}

// ============================================================
//  LOGIN (autenticación IMAP vía backend)
// ============================================================
app.post('/api/login', (req, res) => {
    const { email, password } = req.body;

    if (!email || !password) {
        return res.status(400).json({ success: false, error: 'Email y contraseña requeridos' });
    }

    const domain = email.split('@')[1];
    let imapHost = 'mail.' + domain;
    let smtpHost = 'mail.' + domain;

    if (domain === 'gmail.com') {
        imapHost = 'imap.gmail.com';
        smtpHost = 'smtp.gmail.com';
    } else if (['outlook.com', 'hotmail.com', 'live.com', 'msn.com'].includes(domain)) {
        imapHost = 'outlook.office365.com';
        smtpHost = 'smtp.office365.com';
    } else if (domain.includes('yahoo')) {
        imapHost = 'imap.mail.yahoo.com';
        smtpHost = 'smtp.mail.yahoo.com';
    } else if (['icloud.com', 'me.com'].includes(domain)) {
        imapHost = 'imap.mail.me.com';
        smtpHost = 'smtp.mail.me.com';
    }

    console.log(`🔐 Intentando login para ${email} en ${imapHost}:993`);

    const imap = new Imap({
        user: email,
        password: password,
        host: imapHost,
        port: 993,
        tls: true,
        tlsOptions: { rejectUnauthorized: false }
    });

    let responded = false;

    imap.once('ready', () => {
        if (!responded) {
            responded = true;
            imap.end();
            console.log(`✅ Login exitoso para ${email}`);
            res.json({
                success: true,
                message: 'Autenticación exitosa',
                account: {
                    email: email,
                    password: password,
                    imapHost: imapHost,
                    imapPort: 993,
                    imapSecurity: 'ssl',
                    smtpHost: smtpHost,
                    smtpPort: 465,
                    smtpSecurity: 'ssl'
                }
            });
        }
    });

    imap.once('error', (err) => {
        if (!responded) {
            responded = true;
            console.error('❌ Error de login IMAP:', err.message);
            res.status(401).json({ success: false, error: 'Credenciales inválidas o error de conexión: ' + err.message });
        }
        imap.end();
    });

    const timeout = setTimeout(() => {
        if (!responded) {
            responded = true;
            res.status(408).json({ success: false, error: 'Tiempo de espera agotado al conectar con el servidor' });
            imap.end();
        }
    }, 15000);

    imap.connect();
});

// ------------------------------------------------------------
//  PING
// ------------------------------------------------------------
app.get('/ping', (req, res) => {
    res.json({ alive: true, time: new Date().toISOString() });
});

// ------------------------------------------------------------
//  FUNCIÓN DE AYUDA: crear cliente IMAP con reconexión
// ------------------------------------------------------------
function createImapClient(email, password, host, port, secure) {
    return new ImapFlow({
        host: host || 'imap.gmail.com',
        port: port || 993,
        secure: secure !== undefined ? secure : true,
        auth: { user: email, pass: password },
        logger: false,
        tls: insecureTls,
        connectionTimeout: 60000,
        socketTimeout: 120000,
    });
}

// ------------------------------------------------------------
//  OBTENER CARPETAS
// ------------------------------------------------------------
app.post('/api/folders', async (req, res) => {
    const { email, password, host, port, secure } = req.body;
    const client = createImapClient(email, password, host, port, secure);

    let attempts = 0;
    while (attempts < 3) {
        try {
            await client.connect();
            const mailboxes = await client.list();
            await client.logout();
            const folders = mailboxes.map(mbox => ({
                name: mbox.name,
                path: mbox.path,
                specialUse: mbox.specialUse
            }));
            return res.json({ success: true, folders });
        } catch (error) {
            attempts++;
            console.error(`❌ Intento ${attempts} /api/folders falló:`, error.message);
            if (attempts >= 3) {
                return res.status(500).json({ success: false, error: error.message });
            }
            await new Promise(resolve => setTimeout(resolve, 1000 * attempts));
        }
    }
});

// ------------------------------------------------------------
//  OBTENER LISTA DE MENSAJES (MÁS RECIENTES PRIMERO)
// ------------------------------------------------------------
app.post('/api/messages', async (req, res) => {
    const { email, password, host, port, secure, folder, page = 1, limit = 20 } = req.body;

    const imap = new Imap({
        user: email,
        password: password,
        host: host || 'mail.' + email.split('@')[1],
        port: port || 993,
        tls: true,
        tlsOptions: { rejectUnauthorized: false }
    });

    imap.once('ready', () => {
        imap.openBox(folder || 'INBOX', true, (err, box) => {
            if (err) {
                imap.end();
                return res.status(500).json({ error: 'Error abriendo carpeta' });
            }

            const total = box.messages.total;
            if (total === 0) {
                imap.end();
                return res.json({ success: true, messages: [], total: 0 });
            }

            const end = total - (page - 1) * limit;
            const start = Math.max(1, end - limit + 1);

            if (end < 1) {
                imap.end();
                return res.json({ success: true, messages: [], total });
            }

            const fetch = imap.seq.fetch(`${start}:${end}`, {
                bodies: 'HEADER.FIELDS (FROM TO SUBJECT DATE)',
                struct: true
            });

            const messages = [];
            let count = 0;

            fetch.on('message', (msg, seqno) => {
                let attributes = null;
                msg.once('attributes', (attrs) => { attributes = attrs; });

                msg.on('body', (stream) => {
                    simpleParser(stream, (err, parsed) => {
                        if (!err && parsed) {
                            const uid = attributes ? attributes.uid : seqno;
                            messages.push({
                                uid: uid,
                                subject: parsed.subject || '(Sin asunto)',
                                from: parsed.from ? parsed.from.text : '',
                                date: parsed.date ? parsed.date.toISOString() : new Date().toISOString(),
                                hasAttachments: false
                            });
                            count++;
                        }
                    });
                });
            });

            fetch.once('end', () => {
                imap.end();
                messages.sort((a, b) => new Date(b.date) - new Date(a.date));
                res.json({ success: true, messages, total });
            });

            fetch.once('error', (err) => {
                imap.end();
                console.error('❌ Error en fetch:', err);
                res.status(500).json({ success: false, error: err.message });
            });
        });
    });

    imap.once('error', (err) => {
        console.error('❌ Error IMAP:', err);
        res.status(500).json({ success: false, error: err.message });
    });

    imap.connect();
});

// ------------------------------------------------------------
//  DETALLE DE MENSAJE
// ------------------------------------------------------------
app.post('/api/message-detail', async (req, res) => {
    const { email, password, host, port, secure, folder, uid } = req.body;

    if (!email || !password || !uid) {
        return res.status(400).json({ error: 'Faltan parámetros requeridos' });
    }

    const imap = new Imap({
        user: email,
        password: password,
        host: host || 'mail.' + email.split('@')[1],
        port: port || 993,
        tls: true,
        tlsOptions: { rejectUnauthorized: false },
        connTimeout: 60000,
        keepalive: true
    });

    let responded = false;

    let sendResponse = (data) => {
        if (!responded) {
            responded = true;
            res.json(data);
        }
    };

    let sendError = (msg) => {
        if (!responded) {
            responded = true;
            res.status(500).json({ error: msg });
        }
    };

    imap.once('ready', () => {
        imap.openBox(folder || 'INBOX', true, (err, box) => {
            if (err) {
                imap.end();
                console.error('Error abriendo carpeta:', err);
                return sendError('Error abriendo la carpeta');
            }

            const fetch = imap.fetch([parseInt(uid)], { bodies: '' });

            fetch.on('message', (msg) => {
                msg.on('body', (stream, info) => {
                    simpleParser(stream)
                        .then((parsed) => {
                            let htmlContent = parsed.html || parsed.textAsHtml;
                            if (!htmlContent && parsed.text) {
                                htmlContent = `<pre style="font-family: sans-serif; white-space: pre-wrap;">${parsed.text}</pre>`;
                            }
                            if (!htmlContent) {
                                htmlContent = '<p>(Este correo no contiene texto en el cuerpo)</p>';
                            }

                            if (parsed.attachments && parsed.attachments.length > 0) {
                                parsed.attachments.forEach((att) => {
                                    if (att.contentId && att.related) {
                                        const base64Src = `data:${att.contentType};base64,${att.content.toString('base64')}`;
                                        htmlContent = htmlContent.replace(new RegExp(`cid:${att.contentId}`, 'g'), base64Src);
                                    }
                                });
                            }

                            let to = '';
                            let cc = '';
                            if (parsed.to) {
                                if (Array.isArray(parsed.to)) {
                                    to = parsed.to.map(a => a.address).join(', ');
                                } else {
                                    to = parsed.to.address || parsed.to.text || '';
                                }
                            }
                            if (parsed.cc) {
                                if (Array.isArray(parsed.cc)) {
                                    cc = parsed.cc.map(a => a.address).join(', ');
                                } else {
                                    cc = parsed.cc.address || parsed.cc.text || '';
                                }
                            }

                            const attachmentsList = (parsed.attachments || []).map((att) => ({
                                filename: att.filename || 'adjunto',
                                size: att.size || 0,
                                contentType: att.contentType || 'application/octet-stream',
                                partId: att.contentId || att.filename || ''
                            }));

                            sendResponse({
                                success: true,
                                subject: parsed.subject || '',
                                from: parsed.from ? parsed.from.text : '',
                                to: to,
                                cc: cc,
                                htmlBody: htmlContent,
                                body: parsed.text || '',
                                attachments: attachmentsList
                            });

                            imap.end();
                        })
                        .catch((parseErr) => {
                            console.error('Error al parsear el correo:', parseErr);
                            sendError('Error al procesar el cuerpo del correo');
                            imap.end();
                        });
                });
            });

            fetch.once('error', (fetchErr) => {
                console.error('Fetch error:', fetchErr);
                sendError('Error al obtener el correo de IMAP');
                imap.end();
            });
        });
    });

    imap.once('error', (err) => {
        console.error('IMAP error:', err);
        sendError('Error de conexión IMAP: ' + err.message);
    });

    const timeout = setTimeout(() => {
        if (!responded) {
            sendError('Timeout al obtener el mensaje');
            imap.end();
        }
    }, 60000);

    const originalSend = sendResponse;
    sendResponse = (data) => {
        clearTimeout(timeout);
        originalSend(data);
    };
    sendError = (msg) => {
        clearTimeout(timeout);
        originalSend({ error: msg });
    };

    imap.connect();
});

// ------------------------------------------------------------
//  DESCUBRIR CARPETAS ESTÁNDAR
// ------------------------------------------------------------
app.post('/api/discover-folders', async (req, res) => {
    const { email, password, host, port, secure } = req.body;
    const client = createImapClient(email, password, host, port, secure);

    try {
        await client.connect();
        const mailboxes = await client.list();
        await client.logout();

        const folderNames = mailboxes.map(m => m.name || m.path);

        const findFolder = (candidates) => {
            for (const cand of candidates) {
                const found = folderNames.find(f =>
                    f.toLowerCase() === cand.toLowerCase() ||
                    f.toLowerCase().includes(cand.toLowerCase())
                );
                if (found) return found;
            }
            return null;
        };

        const result = {
            trash: findFolder(['Trash', 'Deleted Items', 'Papelera', 'Deleted', 'INBOX.Trash', 'INBOX.Deleted']),
            spam: findFolder(['Spam', 'Junk', 'Junk Email', 'Correo no deseado', 'INBOX.Spam', 'INBOX.Junk']),
            sent: findFolder(['Sent', 'Sent Items', 'Enviados', 'INBOX.Sent']),
            drafts: findFolder(['Drafts', 'Borradores', 'INBOX.Drafts']),
            all: folderNames
        };

        console.log('📂 Carpetas descubiertas:', result);
        res.json({ success: true, folders: result });
    } catch (error) {
        console.error('❌ Error /api/discover-folders:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------------------------------------------------
//  MOVER MENSAJE
// ------------------------------------------------------------
app.post('/api/move-message', async (req, res) => {
    const { email, password, host, port, secure, uid, fromFolder, toFolder } = req.body;

    console.log(`📤 Mover mensaje UID ${uid} de "${fromFolder}" a "${toFolder}"`);

    if (!uid || !fromFolder || !toFolder) {
        return res.status(400).json({ success: false, error: 'Faltan parámetros' });
    }

    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();

        const mailboxes = await client.list();
        const folderExists = mailboxes.some(m => m.path === toFolder || m.name === toFolder);

        if (!folderExists) {
            console.log(`⚠️ La carpeta "${toFolder}" no existe. Intentando crear...`);
            try {
                await client.mailboxCreate(toFolder);
                console.log(`✅ Carpeta "${toFolder}" creada.`);
            } catch (createError) {
                console.error(`❌ No se pudo crear "${toFolder}":`, createError.message);
                const alternatives = {
                    'INBOX.Trash': ['Trash', 'Deleted Items', 'Papelera', 'INBOX.Deleted'],
                    'INBOX.Spam': ['Spam', 'Junk', 'Junk Email', 'INBOX.Junk']
                };
                const altFolders = alternatives[toFolder] || [];
                let found = false;
                for (const alt of altFolders) {
                    if (mailboxes.some(m => m.path === alt || m.name === alt)) {
                        console.log(`✅ Usando carpeta alternativa: "${alt}"`);
                        await client.messageMove(uid, alt, { uid: true });
                        found = true;
                        break;
                    }
                }
                if (!found) {
                    console.log(`⚠️ Usando INBOX como destino fallback`);
                    await client.messageMove(uid, 'INBOX', { uid: true });
                }
                await client.logout();
                return res.json({ success: true, warning: 'Carpeta no encontrada, se usó alternativa' });
            }
        }

        await client.messageMove(uid, toFolder, { uid: true });
        await client.logout();
        console.log(`✅ Mensaje UID ${uid} movido a "${toFolder}"`);
        res.json({ success: true });
    } catch (error) {
        console.error('❌ Error /api/move-message:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------------------------------------------------
//  GUARDAR EN ENVIADOS (para envíos locales desde la app)
// ------------------------------------------------------------
app.post('/api/save-to-sent', async (req, res) => {
    const { email, password, host, rawMessage } = req.body;

    if (!email || !password || !rawMessage) {
        return res.status(400).json({ success: false, error: 'Faltan email, password o rawMessage' });
    }

    const domain = email.split('@')[1];
    const imapHost = host || `mail.${domain}`;

    const imap = new Imap({
        user: email,
        password: password,
        host: imapHost,
        port: 993,
        tls: true,
        tlsOptions: { rejectUnauthorized: false }
    });

    imap.once('ready', () => {
        imap.getBoxes((err, boxes) => {
            if (err) {
                imap.end();
                return res.status(500).json({ success: false, error: 'Error obteniendo carpetas' });
            }

            const sentFolder = findSentFolder(boxes) || 'INBOX.Sent';

            imap.openBox(sentFolder, false, (openErr) => {
                if (openErr) {
                    imap.end();
                    return res.status(500).json({ success: false, error: 'Error abriendo carpeta Enviados' });
                }

                imap.append(rawMessage, { mailbox: sentFolder, flags: ['\\Seen'] }, (appendErr) => {
                    imap.end();
                    if (appendErr) {
                        return res.status(500).json({ success: false, error: 'Error al guardar: ' + appendErr.message });
                    }
                    return res.json({ success: true, message: 'Guardado en Enviados correctamente' });
                });
            });
        });
    });

    imap.once('error', (err) => {
        if (!res.headersSent) {
            res.status(500).json({ success: false, error: 'Error IMAP: ' + err.message });
        }
    });

    imap.connect();
});

// ------------------------------------------------------------
//  ELIMINAR MENSAJE (PERMANENTE)
// ------------------------------------------------------------
app.post('/api/delete-message', async (req, res) => {
    const { email, password, host, port, secure, uid, folder } = req.body;
    if (!uid || !folder) {
        return res.status(400).json({ success: false, error: 'Faltan uid o folder' });
    }
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        await client.mailboxOpen(folder);
        await client.messageDelete(uid, { uid: true });
        await client.expunge();
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/delete-message:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------------------------------------------------
//  MARCAR LEÍDO / NO LEÍDO
// ------------------------------------------------------------
app.post('/api/toggle-read', async (req, res) => {
    const { email, password, host, port, secure, uid, folder, read } = req.body;
    if (uid == null || !folder) {
        return res.status(400).json({ success: false, error: 'Faltan uid o folder' });
    }
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        await client.mailboxOpen(folder);
        if (read) {
            await client.messageFlagsAdd(uid, ['\\Seen'], { uid: true });
        } else {
            await client.messageFlagsRemove(uid, ['\\Seen'], { uid: true });
        }
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/toggle-read:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------------------------------------------------
//  MARCAR IMPORTANTE
// ------------------------------------------------------------
app.post('/api/toggle-flagged', async (req, res) => {
    const { email, password, host, port, secure, uid, folder, flagged } = req.body;
    if (uid == null || !folder) {
        return res.status(400).json({ success: false, error: 'Faltan uid o folder' });
    }
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        await client.mailboxOpen(folder);
        if (flagged) {
            await client.messageFlagsAdd(uid, ['\\Flagged'], { uid: true });
        } else {
            await client.messageFlagsRemove(uid, ['\\Flagged'], { uid: true });
        }
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/toggle-flagged:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------------------------------------------------
//  CREAR / BORRAR CARPETA
// ------------------------------------------------------------
app.post('/api/create-folder', async (req, res) => {
    const { email, password, host, port, secure, folderName } = req.body;
    if (!folderName) {
        return res.status(400).json({ success: false, error: 'Falta folderName' });
    }
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        await client.mailboxCreate(folderName);
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/create-folder:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.post('/api/delete-folder', async (req, res) => {
    const { email, password, host, port, secure, folderName } = req.body;
    if (!folderName) {
        return res.status(400).json({ success: false, error: 'Falta folderName' });
    }
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        await client.mailboxDelete(folderName);
        await client.logout();
        res.json({ success: true });
    } catch (error) {
        console.error('Error /api/delete-folder:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------------------------------------------------
//  DESCARGAR ADJUNTO
// ------------------------------------------------------------
app.post('/api/download-attachment', async (req, res) => {
    const { email, password, host, port, secure, folder, uid, partId } = req.body;
    if (!uid || !folder || !partId) {
        return res.status(400).json({ success: false, error: 'Faltan uid, folder o partId' });
    }
    const client = createImapClient(email, password, host, port, secure);
    try {
        await client.connect();
        await client.mailboxOpen(folder);
        const msg = await client.fetchOne(String(uid), { bodyParts: [partId] }, { uid: true });
        await client.logout();
        if (!msg || !msg.bodyParts || !msg.bodyParts[partId]) {
            return res.status(404).json({ success: false, error: 'Adjunto no encontrado' });
        }
        const data = msg.bodyParts[partId].toString('base64');
        res.json({ success: true, data });
    } catch (error) {
        console.error('Error /api/download-attachment:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// ------------------------------------------------------------
//  INICIAR SERVIDOR
// ------------------------------------------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`✅ Backend de RSMAIL corriendo en puerto ${PORT}`);
});
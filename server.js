const express = require('express');
const session = require('express-session');
const Imap = require('imap');
const { simpleParser } = require('mailparser');
const dns = require('dns').promises;
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
  secret: process.env.SESSION_SECRET || 'raulsystem-secret-2026',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 2 * 60 * 60 * 1000 }
}));

const KNOWN_IMAP = {
  'gmail.com': { host: 'imap.gmail.com', port: 993 },
  'googlemail.com': { host: 'imap.gmail.com', port: 993 },
  'outlook.com': { host: 'outlook.office365.com', port: 993 },
  'hotmail.com': { host: 'outlook.office365.com', port: 993 },
  'yahoo.com': { host: 'imap.mail.yahoo.com', port: 993 },
};

async function resolveImapHost(domain) {
  if (KNOWN_IMAP[domain]) return KNOWN_IMAP[domain];
  const candidates = [`mail.${domain}`, `imap.${domain}`];
  for (const host of candidates) {
    try {
      await dns.lookup(host);
      return { host, port: 993 };
    } catch (_) {}
  }
  return { host: `mail.${domain}`, port: 993 };
}

function connectImap(email, password, imapConfig) {
  return new Promise((resolve, reject) => {
    const imap = new Imap({
      user: email,
      password: password,
      host: imapConfig.host,
      port: imapConfig.port,
      tls: true,
      tlsOptions: { rejectUnauthorized: false },
      connTimeout: 10000,
      authTimeout: 10000,
    });
    imap.once('ready', () => resolve(imap));
    imap.once('error', reject);
    imap.connect();
  });
}

function fetchMessages(imap, limit = 40) {
  return new Promise((resolve, reject) => {
    imap.openBox('INBOX', true, (err, box) => {
      if (err) return reject(err);
      const total = box.messages.total;
      if (total === 0) return resolve([]);
      const start = Math.max(1, total - limit + 1);
      const fetch = imap.seq.fetch(`${start}:*`, {
        bodies: ['HEADER.FIELDS (FROM TO SUBJECT DATE)'],
        struct: true,
      });
      const messages = [];
      fetch.on('message', (msg, seqno) => {
        const m = { seqno, headers: {}, uid: null };
        msg.on('body', (stream) => {
          let buf = '';
          stream.on('data', c => buf += c.toString('utf8'));
          stream.on('end', () => {
            const lines = buf.split('\r\n');
            lines.forEach(l => {
              const idx = l.indexOf(':');
              if (idx > 0) {
                const k = l.slice(0, idx).toLowerCase().trim();
                const v = l.slice(idx + 1).trim();
                m.headers[k] = v;
              }
            });
          });
        });
        msg.once('attributes', attrs => { m.uid = attrs.uid; });
        msg.once('end', () => messages.push(m));
      });
      fetch.once('error', reject);
      fetch.once('end', () => resolve(messages.reverse()));
    });
  });
}

app.post('/login', async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ error: 'Email e senha obrigatórios' });
  const domain = email.split('@')[1];
  if (!domain) return res.status(400).json({ error: 'Email inválido' });
  try {
    const imapConfig = await resolveImapHost(domain);
    const imap = await connectImap(email, password, imapConfig);
    imap.end();
    req.session.user = { email, password, imapConfig };
    res.json({ ok: true });
  } catch (err) {
    res.status(401).json({ error: 'Credenciais inválidas ou servidor IMAP não encontrado' });
  }
});

app.get('/inbox', async (req, res) => {
  if (!req.session.user) return res.status(401).json({ error: 'Não autenticado' });
  const { email, password, imapConfig } = req.session.user;
  try {
    const imap = await connectImap(email, password, imapConfig);
    const messages = await fetchMessages(imap, 40);
    imap.end();
    res.json({ email, messages });
  } catch (err) {
    res.status(500).json({ error: 'Erro ao conectar na caixa de entrada' });
  }
});

app.get('/message/:uid', async (req, res) => {
  if (!req.session.user) return res.status(401).json({ error: 'Não autenticado' });
  const { email, password, imapConfig } = req.session.user;
  const uid = parseInt(req.params.uid);
  try {
    const imap = await connectImap(email, password, imapConfig);
    const raw = await new Promise((resolve, reject) => {
      imap.openBox('INBOX', true, (err) => {
        if (err) return reject(err);
        const fetch = imap.fetch(uid, { bodies: '', struct: true });
        let buf = '';
        fetch.on('message', msg => {
          msg.on('body', stream => { stream.on('data', c => buf += c.toString()); });
        });
        fetch.once('end', () => resolve(buf));
        fetch.once('error', reject);
      });
    });
    imap.end();
    const parsed = await simpleParser(raw);
    res.json({
      from: parsed.from?.text || '',
      subject: parsed.subject || '',
      date: parsed.date || '',
      html: parsed.html || parsed.text || '',
    });
  } catch (err) {
    res.status(500).json({ error: 'Erro ao carregar mensagem' });
  }
});

app.post('/logout', (req, res) => {
  req.session.destroy();
  res.json({ ok: true });
});

app.get('/', (req, res) => {
  if (req.session.user) {
    res.sendFile(path.join(__dirname, 'public', 'inbox.html'));
  } else {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Raulsystem Webmail rodando na porta ${PORT}`));

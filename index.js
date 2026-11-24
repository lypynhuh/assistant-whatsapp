const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const { OpenAI } = require('openai');
const nodemailer = require('nodemailer');
const Imap = require('imap');

const openai = new OpenAI({ apiKey: process.env.OPENAI_KEY });

// --- OUTLOOK ---
const OUTLOOK_EMAIL = process.env.OUTLOOK_EMAIL; // ton email outlook
const OUTLOOK_PASS = process.env.OUTLOOK_PASS; // ton mot de passe d'application

async function lireMails() {
  return new Promise((resolve) => {
    const imap = new Imap({
      user: OUTLOOK_EMAIL,
      password: OUTLOOK_PASS,
      host: 'outlook.office365.com',
      port: 993,
      tls: true
    });
    imap.once('ready', () => {
      imap.openBox('INBOX', true, () => {
        imap.search(['UNSEEN'], (err, results) => {
          if (!results || results.length === 0) { imap.end(); return resolve("Aucun mail non lu."); }
          const f = imap.fetch(results.slice(-5), { bodies: 'HEADER.FIELDS (FROM SUBJECT DATE)' });
          let mails = "";
          f.on('message', (msg) => {
            msg.on('body', (stream) => {
              let buffer = '';
              stream.on('data', (chunk) => buffer += chunk.toString('utf8'));
              stream.on('end', () => mails += buffer + "\n---\n");
            });
          });
          f.once('end', () => { imap.end(); resolve(mails); });
        });
      });
    });
    imap.connect();
  });
}

async function envoyerMail(destinataire, sujet, texte) {
  const transporter = nodemailer.createTransport({
    host: 'smtp.office365.com', port: 587, secure: false,
    auth: { user: OUTLOOK_EMAIL, pass: OUTLOOK_PASS }
  });
  await transporter.sendMail({ from: OUTLOOK_EMAIL, to: destinataire, subject: sujet, text: texte });
  return "Mail envoyé !";
}

// --- WHATSAPP ---
const client = new Client({
    authStrategy: new LocalAuth(),
    authTimeoutMs: 120000,
    qrMaxRetries: 5,
     puppeteer: {
        headless: true,
        args: [   
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage',
            '--disable-accelerated-2d-canvas',
            '--no-first-run',
            '--no-zygote',
            '--disable-gpu'
        ]
    },
  });

client.on('qr', (qr) => {
  console.log('SCANNE CE QR CODE AVEC WHATSAPP');
  qrcode.generate(qr, { small: true });
});

client.on('ready', () => console.log('Assistant en ligne !'));

client.on('message', async (msg) => {
  if (msg.fromMe) return;
  const texte = msg.body.toLowerCase();

  // Commandes Outlook
  if (texte.includes('lis mes mails') || texte.includes('mes mails')) {
    const mails = await lireMails();
    msg.reply('Voici tes derniers mails non lus :\n\n' + mails);
    return;
  }

  // Sinon, ChatGPT répond
  const reponse = await openai.chat.completions.create({
    model: 'gpt-4o-mini',
    messages: [
      { role: 'system', content: 'Tu es un assistant personnel utile, drôle, qui parle français. Tu aides pour mails, agenda, code, fichiers.' },
      { role: 'user', content: msg.body }
    ]
  });
  msg.reply(reponse.choices[0].message.content);
});

client.initialize();

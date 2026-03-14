const express = require('express');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const path = require('path');
const helmet = require('helmet');

const app = express();
app.use(helmet());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

// Configurez ces variables d'environnement avant de lancer
// SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, DEST_EMAIL
const SMTP_HOST = process.env.SMTP_HOST;
const SMTP_PORT = process.env.SMTP_PORT;
const SMTP_USER = process.env.SMTP_USER;
const SMTP_PASS = process.env.SMTP_PASS;
const DEST_EMAIL = process.env.DEST_EMAIL;

if (!SMTP_HOST || !SMTP_PORT || !SMTP_USER || !SMTP_PASS || !DEST_EMAIL) {
  console.warn('Variables SMTP non configurées. Configurez SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, DEST_EMAIL.');
}

app.post('/submit', async (req, res) => {
  try {
    const payload = req.body;

    // Validation serveur minimale
    const id = String(payload.identifier || '').trim();
    if (!/^\d+$/.test(id)) return res.status(400).json({ error: 'Identifiant invalide (doit être numérique).' });

    const answers = payload.answers || {};
    const comment = String(payload.comment || '').trim();
    if (comment.length < 2 || comment.length > 3000) {
      return res.status(400).json({ error: 'Commentaire doit contenir entre 2 et 3000 caractères.' });
    }

    // Créer workbook Excel
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sondage');

    // En-têtes
    const headers = ['Identifiant', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Note (0-20)', 'Commentaire'];
    sheet.addRow(headers);

    // Ligne de réponses
    const row = [
      id,
      answers.q1 || '',
      Array.isArray(answers.q2) ? answers.q2.join('; ') : (answers.q2 || ''),
      answers.q3 || '',
      answers.q4 || '',
      answers.q5 || '',
      answers.q6 || '',
      answers.note !== undefined ? String(answers.note) : '',
      comment
    ];
    sheet.addRow(row);

    // Générer buffer
    const buffer = await workbook.xlsx.writeBuffer();

    // Préparer transporteur nodemailer
    const transporter = nodemailer.createTransport({
      host: SMTP_HOST,
      port: Number(SMTP_PORT),
      secure: Number(SMTP_PORT) === 465,
      auth: {
        user: SMTP_USER,
        pass: SMTP_PASS
      }
    });

    // Envoyer mail
    const mailOptions = {
      from: `"Sondage" <${SMTP_USER}>`,
      to: DEST_EMAIL,
      subject: `Réponse sondage - Identifiant ${id}`,
      text: `Identifiant: ${id}\nCommentaire: ${comment}`,
      attachments: [
        {
          filename: `sondage_${id}.xlsx`,
          content: buffer,
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      ]
    };

    await transporter.sendMail(mailOptions);

    return res.json({ ok: true });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'Erreur serveur lors de l\'envoi.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Sondage démarré sur http://localhost:${PORT}`);
});

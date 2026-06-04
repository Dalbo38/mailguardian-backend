require("dotenv").config();
const express = require("express");
const cors = require("cors");
const Imap = require("imap");
const { simpleParser } = require("mailparser");

const app = express();
app.use(cors());
app.use(express.json());

// ─── Route de test ────────────────────────────────────────────────────────────
app.get("/", (req, res) => {
  res.json({ status: "Mail Guardian Backend opérationnel 🛡️" });
});

// ─── Route : récupérer les emails via IMAP ────────────────────────────────────
app.post("/api/emails", async (req, res) => {
  const { email, password, limit = 20 } = req.body;

  if (!email || !password) {
    return res.status(400).json({ error: "Email et mot de passe requis" });
  }

  try {
    const emails = await fetchEmails(email, password, limit);
    res.json({ success: true, emails });
  } catch (err) {
    console.error("Erreur IMAP:", err.message);
    res.status(401).json({
      error: "Connexion impossible. Vérifiez vos identifiants et que le mot de passe d'application est correct.",
      detail: err.message,
    });
  }
});

// ─── Fonction IMAP ────────────────────────────────────────────────────────────
function fetchEmails(user, password, limit) {
  return new Promise((resolve, reject) => {
    const imap = new Imap({
      user,
      password,
      host: "outlook.office365.com",
      port: 993,
      tls: true,
      tlsOptions: { rejectUnauthorized: false },
      authTimeout: 10000,
    });

    const emails = [];

    imap.once("ready", () => {
      imap.openBox("INBOX", true, (err, box) => {
        if (err) { imap.end(); return reject(err); }

        const total = box.messages.total;
        if (total === 0) { imap.end(); return resolve([]); }

        // Récupérer les N derniers emails
        const start = Math.max(1, total - limit + 1);
        const fetch = imap.seq.fetch(`${start}:${total}`, {
          bodies: "",
          struct: true,
        });

        fetch.on("message", (msg, seqno) => {
          let buffer = "";
          msg.on("body", (stream) => {
            stream.on("data", (chunk) => { buffer += chunk.toString("utf8"); });
            stream.once("end", async () => {
              try {
                const parsed = await simpleParser(buffer);
                emails.push({
                  id: seqno,
                  from: parsed.from?.value?.[0]?.address || "",
                  name: parsed.from?.value?.[0]?.name || parsed.from?.value?.[0]?.address || "",
                  subject: parsed.subject || "(Sans objet)",
                  preview: (parsed.text || "").substring(0, 120).replace(/\n/g, " "),
                  body: (parsed.text || parsed.html || "").substring(0, 2000),
                  date: parsed.date ? new Date(parsed.date).toLocaleString("fr-FR") : "",
                  read: false,
                });
              } catch (e) {}
            });
          });
        });

        fetch.once("error", (err) => { imap.end(); reject(err); });
        fetch.once("end", () => { imap.end(); });
      });
    });

    imap.once("error", (err) => reject(err));
    imap.once("end", () => resolve(emails.reverse())); // Plus récents en premier
    imap.connect();
  });
}

// ─── Démarrage ────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`✅ Mail Guardian Backend démarré sur le port ${PORT}`);
});

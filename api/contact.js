// MissionXtra — Contact form handler
// Vercel Node serverless function. Sends contact-form submissions to admin@missionxtra.com
// via the Microsoft Graph API using app-only (client credentials) authentication.
//
// Required Vercel Environment Variables (Project → Settings → Environment Variables):
//   MSFT_TENANT_ID      — your Microsoft 365 tenant (directory) ID
//   MSFT_CLIENT_ID      — the Entra app registration's Application (client) ID
//   MSFT_CLIENT_SECRET  — the app registration's client secret VALUE  (mark as sensitive)
//   GRAPH_SENDER        — mailbox the message is SENT FROM, e.g. admin@missionxtra.com
//   CONTACT_TO          — mailbox the message is DELIVERED TO,  e.g. admin@missionxtra.com
//
// The app registration needs the Microsoft Graph *application* permission "Mail.Send"
// with admin consent granted. For security, scope it to GRAPH_SENDER only with an
// Exchange Online Application Access Policy (see the setup notes).

const clip = (s, n) => String(s == null ? "" : s).slice(0, n);
const esc = (s) =>
  String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

async function readJson(req) {
  if (req.body && typeof req.body === "object") return req.body;
  const chunks = [];
  for await (const c of req) chunks.push(c);
  const raw = Buffer.concat(chunks).toString("utf8") || "{}";
  try { return JSON.parse(raw); } catch { return {}; }
}

module.exports = async (req, res) => {
  res.setHeader("Content-Type", "application/json");

  if (req.method !== "POST") {
    res.statusCode = 405;
    return res.end(JSON.stringify({ error: "Method not allowed" }));
  }

  const {
    MSFT_TENANT_ID, MSFT_CLIENT_ID, MSFT_CLIENT_SECRET,
    GRAPH_SENDER, CONTACT_TO,
  } = process.env;

  if (!MSFT_TENANT_ID || !MSFT_CLIENT_ID || !MSFT_CLIENT_SECRET || !GRAPH_SENDER || !CONTACT_TO) {
    res.statusCode = 500;
    return res.end(JSON.stringify({ error: "Server email is not configured yet." }));
  }

  const body = await readJson(req);

  // Honeypot — bots fill hidden fields. Pretend success, send nothing.
  if (body.company_website) {
    res.statusCode = 200;
    return res.end(JSON.stringify({ ok: true }));
  }

  const name = clip(body.name, 120).trim();
  const email = clip(body.email, 160).trim();
  const organisation = clip(body.organisation, 160).trim();
  const subject = clip(body.subject, 180).trim();
  const message = clip(body.message, 5000).trim();

  if (!name || !email || !subject || !message) {
    res.statusCode = 400;
    return res.end(JSON.stringify({ error: "Missing required fields." }));
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    res.statusCode = 400;
    return res.end(JSON.stringify({ error: "Please provide a valid email address." }));
  }

  try {
    // 1) Get an app-only access token (client credentials)
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${MSFT_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: MSFT_CLIENT_ID,
          client_secret: MSFT_CLIENT_SECRET,
          scope: "https://graph.microsoft.com/.default",
          grant_type: "client_credentials",
        }),
      }
    );
    if (!tokenRes.ok) {
      const t = await tokenRes.text();
      console.error("Token error:", t);
      res.statusCode = 502;
      return res.end(JSON.stringify({ error: "Could not authenticate with mail service." }));
    }
    const { access_token } = await tokenRes.json();

    // 2) Send the mail as GRAPH_SENDER, delivered to CONTACT_TO, reply-to the visitor
    const html =
      `<h2 style="margin:0 0 12px">New enquiry from missionxtra.com</h2>` +
      `<table cellpadding="6" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px">` +
      `<tr><td><b>Name</b></td><td>${esc(name)}</td></tr>` +
      `<tr><td><b>Email</b></td><td>${esc(email)}</td></tr>` +
      (organisation ? `<tr><td><b>Organisation</b></td><td>${esc(organisation)}</td></tr>` : "") +
      `<tr><td><b>Subject</b></td><td>${esc(subject)}</td></tr>` +
      `</table>` +
      `<p style="font-family:Arial,sans-serif;font-size:14px;white-space:pre-wrap;margin-top:16px">${esc(message)}</p>`;

    const mail = {
      message: {
        subject: `[MissionXtra] ${subject}`,
        body: { contentType: "HTML", content: html },
        toRecipients: [{ emailAddress: { address: CONTACT_TO } }],
        replyTo: [{ emailAddress: { address: email, name } }],
      },
      saveToSentItems: true,
    };

    const sendRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(GRAPH_SENDER)}/sendMail`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${access_token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(mail),
      }
    );

    if (sendRes.status !== 202) {
      const t = await sendRes.text();
      console.error("Graph sendMail error:", sendRes.status, t);
      res.statusCode = 502;
      return res.end(JSON.stringify({ error: "The message could not be sent." }));
    }

    res.statusCode = 200;
    return res.end(JSON.stringify({ ok: true }));
  } catch (err) {
    console.error("Contact handler error:", err);
    res.statusCode = 500;
    return res.end(JSON.stringify({ error: "Unexpected server error." }));
  }
};

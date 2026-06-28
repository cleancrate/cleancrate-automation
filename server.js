const express = require('express');
const fetch = require('node-fetch');
const cors = require('cors');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Microsoft Auth ───────────────────────────────────────────────────────
app.post('/api/ms-token', async (req, res) => {
  const { clientId, tenantId, clientSecret } = req.body;
  try {
    const body = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: clientId,
      client_secret: clientSecret,
      scope: 'https://graph.microsoft.com/.default'
    });

    const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body
    });

    const data = await response.json();
    if (!response.ok) return res.status(400).json({ error: data.error_description || 'Auth failed' });
    res.json({ token: data.access_token });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── HubSpot: Get Contacts ────────────────────────────────────────────────
app.post('/api/hubspot/contacts', async (req, res) => {
  const { token } = req.body;
  try {
    const response = await fetch(
      'https://api.hubapi.com/crm/v3/objects/contacts?limit=100&properties=firstname,lastname,email',
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const data = await response.json();
    if (!response.ok) return res.status(400).json({ error: data.message || 'HubSpot error' });
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── HubSpot: Get Engagements for Contact ────────────────────────────────
app.post('/api/hubspot/engagements', async (req, res) => {
  const { token, contactId } = req.body;
  try {
    const response = await fetch(
      `https://api.hubapi.com/engagements/v1/engagements/associated/contact/${contactId}/paged?limit=50`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const data = await response.json();
    if (!response.ok) return res.status(400).json({ error: data.message || 'HubSpot error' });
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── Outlook: Get Sent Emails to a lead ──────────────────────────────────
app.post('/api/outlook/sent', async (req, res) => {
  const { msToken, senderEmail, leadEmail } = req.body;
  try {
    const filter = encodeURIComponent(`toRecipients/any(r:r/emailAddress/address eq '${leadEmail}')`);
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${senderEmail}/mailFolders/SentItems/messages?$filter=${filter}&$orderby=sentDateTime desc&$top=10`,
      { headers: { Authorization: `Bearer ${msToken}` } }
    );
    const data = await response.json();
    if (!response.ok) return res.status(400).json({ error: data.error?.message || 'Outlook error' });
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── Outlook: Check for replies from lead ────────────────────────────────
app.post('/api/outlook/replies', async (req, res) => {
  const { msToken, senderEmail, leadEmail } = req.body;
  try {
    const filter = encodeURIComponent(`from/emailAddress/address eq '${leadEmail}'`);
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${senderEmail}/mailFolders/Inbox/messages?$filter=${filter}&$top=5`,
      { headers: { Authorization: `Bearer ${msToken}` } }
    );
    const data = await response.json();
    if (!response.ok) return res.status(400).json({ error: data.error?.message || 'Outlook error' });
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── Outlook: Send Email ──────────────────────────────────────────────────
app.post('/api/outlook/send', async (req, res) => {
  const { msToken, senderEmail, toEmail, toName, subject, body } = req.body;
  try {
    const message = {
      message: {
        subject,
        body: { contentType: 'Text', content: body },
        toRecipients: [{ emailAddress: { address: toEmail, name: toName } }]
      },
      saveToSentItems: true
    };

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${msToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(message)
      }
    );

    if (response.status === 202) return res.json({ success: true });
    const data = await response.json();
    res.status(400).json({ error: data.error?.message || 'Send failed' });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Clean Crate server running on port ${PORT}`));

require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const fetch = require('node-fetch');
const msal = require('@azure/msal-node');
const jwt = require('jsonwebtoken');
const jwksRsa = require('jwks-rsa');
const cookieParser = require('cookie-parser');
const querystring = require('querystring');

const app = express();
app.use(cors({ origin: process.env.FRONTEND_URL || true, credentials: true }));
app.use(bodyParser.json());
app.use(cookieParser());

const msalConfig = {
  auth: {
    clientId: process.env.BACKEND_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.BACKEND_CLIENT_SECRET
  }
};
const cca = new msal.ConfidentialClientApplication(msalConfig);

function extractBearer(authHeader){
  if(!authHeader) return null;
  const parts = authHeader.split(' ');
  if(parts.length!==2) return null;
  return parts[1];
}

// JWT validation middleware using JWKS
const jwksClient = jwksRsa({
  jwksUri: `https://login.microsoftonline.com/${process.env.TENANT_ID}/discovery/v2.0/keys`
});
function validateJwtMiddleware(req, res, next){
  const token = extractBearer(req.headers.authorization);
  if(!token) return res.status(401).json({ error: 'Missing token' });
  const getKey = (header, callback) => {
    jwksClient.getSigningKey(header.kid, function(err, key) {
      if (err) return callback(err);
      const signingKey = key.getPublicKey();
      callback(null, signingKey);
    });
  };
  jwt.verify(token, getKey, { audience: process.env.BACKEND_CLIENT_ID, issuer: `https://login.microsoftonline.com/${process.env.TENANT_ID}/v2.0` }, (err, decoded) => {
    if (err) {
      console.error('JWT validation failed', err);
      return res.status(401).json({ error: 'Invalid token' });
    }
    req.auth = decoded;
    next();
  });
}

// Simple helper to call your tool API (replace with your logic)
async function callMyToolApi(path, method, bearerToken, body){
  const url = `${process.env.MY_TOOL_API_URL}${path}`;
  const opts = {
    method,
    headers: {
      'Authorization': `Bearer ${bearerToken}`,
      'Content-Type': 'application/json'
    }
  };
  if (body) opts.body = JSON.stringify(body);
  const r = await fetch(url, opts);
  const j = await r.json().catch(()=>null);
  return { ok: r.ok, status: r.status, json: j };
}

/**
 * POST /create-or-link
 * - Validates incoming token (JWT middleware)
 * - If timeId provided: optionally verify with your API
 * - If not: OBO exchange, call your API to create entry, return id
 */
app.post('/create-or-link', validateJwtMiddleware, async (req, res) => {
  try{
    const { subject, start, end, attendees, timeId } = req.body;

    if(timeId){
      // Optional: verify existence in your API
      // const check = await callMyToolApi(`/entries/${timeId}`, 'GET', '...');
      return res.json({ timeId, created: false });
    }

    // OBO: acquire token to call your API (scope must be configured in Azure AD)
    const incomingToken = extractBearer(req.headers.authorization);
    const oboRequest = {
      oboAssertion: incomingToken,
      scopes: [ process.env.API_SCOPE ]
    };
    const oboResp = await cca.acquireTokenOnBehalfOf(oboRequest);
    if(!oboResp || !oboResp.accessToken) return res.status(500).json({ error: 'Failed OBO' });
    const apiToken = oboResp.accessToken;

    // Call your API to create the entry
    const apiResp = await callMyToolApi('/entries', 'POST', apiToken, { title: subject, start, end, attendees });
    if(!apiResp.ok) return res.status(500).json({ error: 'MyTool API error', details: apiResp.json });

    const generatedId = (apiResp.json && (apiResp.json.id || apiResp.json.timeId)) || (Date.now().toString().slice(-6));
    return res.json({ timeId: generatedId, created: true, details: apiResp.json });

  } catch(err){
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

/**
 * Interactive OAuth fallback endpoints for dialog flow
 *  - /auth/start -> redirects user to Azure authorize endpoint
 *  - /auth/callback -> exchanges code for token, returns an HTML that posts a message to the opener/dialog
 */

app.get('/auth/start', (req, res) => {
  const state = Math.random().toString(36).substring(2);
  const nonce = Math.random().toString(36).substring(2);
  // Save state/nonce in cookie to validate later
  res.cookie('auth_state', state, { httpOnly: true });
  res.cookie('auth_nonce', nonce, { httpOnly: true });

  const params = {
    client_id: process.env.BACKEND_CLIENT_ID,
    response_type: 'code',
    redirect_uri: `${process.env.FRONTEND_URL}${process.env.AUTH_REDIRECT_PATH}`,
    response_mode: 'query',
    scope: 'openid profile email',
    state,
    nonce
  };
  const authorizeUrl = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/authorize?${querystring.stringify(params)}`;
  res.redirect(authorizeUrl);
});

app.get('/auth/callback', async (req, res) => {
  try {
    const { code, state } = req.query;
    const savedState = req.cookies['auth_state'];
    if (!code || state !== savedState) return res.status(400).send('Invalid auth response');

    // Exchange code for token (Authorization Code flow)
    const tokenRequest = {
      code,
      scopes: ["openid","profile","email"],
      redirectUri: `${process.env.FRONTEND_URL}${process.env.AUTH_REDIRECT_PATH}`,
    };
    // Use msal-node to acquire token by code
    const cca = new msal.ConfidentialClientApplication(msalConfig);
    const tokenResponse = await cca.acquireTokenByCode(tokenRequest);
    if(!tokenResponse || !tokenResponse.accessToken) return res.status(500).send('Token exchange failed');

    // Return an HTML page that posts message back to the Office dialog
    const html = `
<html><body>
<script>
  (function(){
    const message = ${JSON.stringify({ type: 'token', token: tokenResponse.accessToken })};
    try {
      // If running inside Office dialog, use Office common dialog messaging
      if (window.external && window.external.office) {
        // not standard; fallback to postMessage
      }
      // Post message to parent (Office dialog will capture via DialogMessageReceived)
      Office = window.Office || {};
      // Post using window.opener (in case)
      if (window.opener && window.opener.postMessage) {
        window.opener.postMessage(JSON.stringify(message), '*');
      }
      // Also use Microsoft Office Dialog API if available
      try { 
        Office.context.ui.messageParent(JSON.stringify(message));
      } catch(e) {}
    } catch(e) {}
    document.body.innerHTML = '<p>Authentification terminée. Vous pouvez fermer cette fenêtre.</p>';
  })();
</script>
</body></html>`;
    res.send(html);
  } catch(err){
    console.error(err);
    res.status(500).send('Auth callback error');
  }
});

const port = process.env.PORT || 3000;
app.listen(port, ()=> console.log(`Backend listening on ${port}`));

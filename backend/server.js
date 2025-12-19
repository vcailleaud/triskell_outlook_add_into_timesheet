/**
 * server.js
 * Backend OAuth Azure AD pour Outlook Calendar Add-in
 * Option A – OAuth via popup (Dialog API)
 */

require("dotenv").config();

const express = require("express");
const fetch = require("node-fetch");
const jwt = require("jsonwebtoken");
const bodyParser = require("body-parser");
const cors = require("cors");

const app = express();
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());
app.use(cors());

/**
 * =========================
 * CONFIGURATION
 * =========================
 */

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  REDIRECT_URI,
  APP_JWT_SECRET
} = process.env;

const AUTHORITY = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0`;
const SCOPES = "openid profile email User.Read";

/**
 * =========================
 * ROUTE 1 – START AUTH
 * =========================
 * Appelé depuis dialog.html
 */
app.get("/auth/start", (req, res) => {
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    response_type: "code",
    redirect_uri: REDIRECT_URI,
    response_mode: "query",
    scope: SCOPES,
    state: "outlook_addin"
  });

  const authUrl = `${AUTHORITY}/authorize?${params.toString()}`;
  res.redirect(authUrl);
});

/**
 * =========================
 * ROUTE 2 – CALLBACK
 * =========================
 * Azure AD redirige ici avec ?code=
 */
app.get("/auth/callback", async (req, res) => {
  try {
    const code = req.query.code;
    if (!code) {
      throw new Error("Missing authorization code");
    }

    /**
     * Exchange code → token
     */
    const tokenResponse = await fetch(`${AUTHORITY}/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        grant_type: "authorization_code",
        code,
        redirect_uri: REDIRECT_URI,
        scope: SCOPES
      })
    });

    const tokenData = await tokenResponse.json();

    if (!tokenData.access_token) {
      console.error(tokenData);
      throw new Error("Token exchange failed");
    }

    /**
     * Decode ID token (identity)
     */
    const idTokenPayload = jwt.decode(tokenData.id_token);

    /**
     * Generate INTERNAL JWT
     * → utilisé ensuite pour ton outil de temps
     */
    const appToken = jwt.sign(
      {
        oid: idTokenPayload.oid,
        email: idTokenPayload.preferred_username,
        name: idTokenPayload.name
      },
      APP_JWT_SECRET,
      { expiresIn: "1h" }
    );

    /**
     * Return token to Outlook dialog
     */
    res.send(`
      <html>
        <body>
          <script>
            Office.context.ui.messageParent(
              ${JSON.stringify(appToken)}
            );
          </script>
        </body>
      </html>
    `);

  } catch (err) {
    console.error(err);
    res.send(`
      <script>
        Office.context.ui.messageParent("AUTH_ERROR");
      </script>
    `);
  }
});

/**
 * =========================
 * HEALTH CHECK
 * =========================
 */
app.get("/", (req, res) => {
  res.send("Backend OAuth Outlook Add-in is running");
});

/**
 * =========================
 * START SERVER
 * =========================
 */
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Backend listening on ${PORT}`);
});

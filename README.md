# Outlook Add-in Kit — ZIP

Contenu du package :
- manifest.xml
- frontend/ (taskpane.html, taskpane.css, taskpane.js)
- backend/ (server.js, package.json, .env.example)

Features supplémentaires ajoutées :
- Fallback OAuth interactive via /auth/start and /auth/callback (dialog)
- JWT validation middleware using jwks-rsa before performing OBO
- OBO flow to exchange client token for API token

Quickstart:
1. Edit backend/.env with your AZURE values and FRONTEND_URL
2. Install backend deps: `npm install` puis `node server.js`
3. Host frontend on HTTPS and update manifest.xml (SourceLocation)
4. Deploy manifest in Exchange Admin Center or sideload for testing

Notes:
- The auth dialog flow uses Authorization Code flow and msal-node on server side
- Ensure redirect URI registered in Azure matches FRONTEND_URL + AUTH_REDIRECT_PATH

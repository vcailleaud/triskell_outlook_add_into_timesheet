Office.onReady(async () => {
  const output = document.getElementById("jwtToken");

  try {
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true
    });

    console.log("JWT SSO récupéré :", token);

    if (output) {
      output.textContent = token;
    }

  } catch (error) {
    console.error("Erreur SSO", error);

    if (output) {
      output.textContent =
        "Erreur lors de la récupération du token SSO\n\n" +
        JSON.stringify(error, null, 2);
    }
  }
});
/*
Office.onReady(() => {
  document.getElementById("btnSave").onclick = onSaveClick;
  document.getElementById("btnAuto").onclick = autoDetect;
  document.getElementById("btnAuthFallback").onclick = startInteractiveAuth;
  init();
});*/

async function init(){
  try {
    const item = Office.context.mailbox.item;
    const subject = await getProp(item, "subject");
    const start = await getProp(item, "start");
    const end = await getProp(item, "end");
    document.getElementById("subject").innerText = subject || "–";
    document.getElementById("range").innerText = formatRange(start, end);

    const existingId = await loadCustomProp("TimeEntryId");
    if (existingId) {
      document.getElementById("timeId").value = existingId;
      showStatus("ID chargé depuis CustomProperties");
      return;
    }

    const catId = await findCategoryId();
    if (catId) {
      document.getElementById("timeId").value = catId;
      showStatus("ID détecté depuis catégorie Outlook");
      return;
    }

    const auto = detectFromSubject(subject);
    if (auto) {
      document.getElementById("timeId").value = auto;
      showStatus("ID pré-rempli depuis le sujet");
    }
  } catch (e) { showError(e); }
}

function formatRange(start, end){
  if(!start && !end) return "–";
  return `${start || "?"} → ${end || "?"}`;
}

function showStatus(msg){ document.getElementById("status").innerText = msg; document.getElementById("error").innerText = ""; }
function showError(e){ document.getElementById("error").innerText = typeof e === "string" ? e : (e && e.message) || JSON.stringify(e); }

function getProp(item, propName){
  return new Promise((resolve, reject) => {
    const method = item[propName] ? null : propName + "Async";
    if (!method) { resolve(item[propName]); return; }
    if (!item[method]) return resolve(null);
    item[method]((res) => { if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value); else resolve(null); });
  });
}

/* CustomProperties helpers */
function loadCustomProp(key){
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.loadCustomPropertiesAsync((asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) { resolve(null); return; }
      const props = asyncResult.value;
      const v = props.get(key);
      resolve(v || null);
    });
  });
}

function saveCustomProp(key, value){
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.loadCustomPropertiesAsync((asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) { reject(asyncResult.error); return; }
      const props = asyncResult.value;
      props.set(key, value);
      props.saveAsync((saveResult) => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) resolve(true);
        else reject(saveResult.error);
      });
    });
  });
}

/* Categories helpers */
function findCategoryId(){
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    if (!item.categories || !item.categories.getAsync) { resolve(null); return; }
    item.categories.getAsync((res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) { resolve(null); return; }
      const cats = res.value || [];
      for (const c of cats) {
        const m = (c || "").match(/^TP-(\w+)$/i);
        if (m) return resolve(m[1]);
      }
      resolve(null);
    });
  });
}

function addCategoryId(id){
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item.categories || !item.categories.addAsync) { resolve(false); return; }
    const cat = `TP-${id}`;
    item.categories.addAsync([cat], (res) => { if (res.status === Office.AsyncResultStatus.Succeeded) resolve(true); else resolve(false); });
  });
}

/* Subject detection */
function detectFromSubject(subject){
  if(!subject) return null;
  let m = subject.match(/\bTP[-_]?(\d{3,})\b/i) || subject.match(/\b([A-Z]+-\d{1,6})\b/) || subject.match(/\b(\d{3,})\b/);
  return m ? m[1] : null;
}

/* Main save flow */
async function onSaveClick(){
  try {
    showStatus("Traitement en cours...");
    const item = Office.context.mailbox.item;
    const subject = await getProp(item, "subject");
    const start = await getProp(item, "start");
    const end = await getProp(item, "end");
    const attendeesProp = await getProp(item, "requiredAttendees");
    const attendees = Array.isArray(attendeesProp) ? attendeesProp.map(a => a.emailAddress || a.address || a) : [];

    let timeId = document.getElementById("timeId").value.trim();

    // Try SSO token first
    let aadToken = null;
    try {
      aadToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt:true });
    } catch (ssoErr){
      // SSO failed / not supported — prompt to use interactive fallback
      showError('SSO non disponible — utilisez "Se connecter (fallback OAuth)"');
      return;
    }

    const payload = { subject, start, end, attendees, timeId: timeId || null };
    const resp = await fetch("https://eb3fca99-8ba0-42c1-8dbb-6886917917f3/create-or-link", {
      method: "POST",
      headers: { "Content-Type":"application/json", "Authorization": `Bearer ${aadToken}` },
      body: JSON.stringify(payload)
    });
    const json = await resp.json();
    if (!resp.ok) throw new Error(json.error || JSON.stringify(json));

    const returnedId = json.timeId;
    if (!returnedId) throw new Error("Backend n'a pas retourné d'ID");

    await saveCustomProp("TimeEntryId", returnedId);
    await addCategoryId(returnedId);

    document.getElementById("timeId").value = returnedId;
    showStatus("Enregistré ✔️ ID: " + returnedId);

  } catch (e) { showError(e); }
}

/* Auto detect button */
async function autoDetect(){
  try {
    const existing = await loadCustomProp("TimeEntryId");
    if (existing) { document.getElementById("timeId").value = existing; showStatus("Chargé depuis CustomProperties"); return; }
    const cat = await findCategoryId();
    if (cat) { document.getElementById("timeId").value = cat; showStatus("Détecté depuis catégorie"); return; }
    const subj = document.getElementById("subject").innerText;
    const guessed = detectFromSubject(subj);
    if (guessed) { document.getElementById("timeId").value = guessed; showStatus("Pré-rempli depuis sujet"); return; }
    showStatus("Aucun ID détecté automatiquement");
  } catch (e) { showError(e); }
}

/* Fallback interactive auth using a dialog — opens backend /auth/start which redirects to Azure AD */
function startInteractiveAuth(){
  const dialogUrl = "https://eb3fca99-8ba0-42c1-8dbb-6886917917f3/auth/start";
  Office.context.ui.displayDialogAsync(dialogUrl, { height: 60, width: 40 }, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
        try {
          const data = JSON.parse(args.message);
          if (data.type === 'token') {
            // Received token from dialog
            showStatus('Authentifié (fallback)');
            // optionally store token for subsequent calls (not persisted)
            window._fallbackToken = data.token;
            dialog.close();
          } else if (data.type === 'error') {
            showError(data.error);
            dialog.close();
          }
        } catch(e) { console.error(e); }
      });
    } else {
      showError("Impossible d'ouvrir la fenêtre d'authentification");
    }
  });
}

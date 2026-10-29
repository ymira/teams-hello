import { app, nestedAppAuth } from "@microsoft/teams-js";
import {
  createNestablePublicClientApplication,
  createStandardPublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

const cfg = window.APP_CONFIG || {};
const SCOPES = ["User.Read", "GroupMember.Read.All"];
const $ = (id) => document.getElementById(id);

function status(msg, isError = false) {
  $("status").textContent = msg;
  $("status").className = isError ? "status error" : "status";
}

async function initTeams() {
  // Resolves with the Teams context, or null when opened in a normal browser.
  try {
    await Promise.race([
      app.initialize(),
      new Promise((_, rej) => setTimeout(() => rej(new Error("timeout")), 3000)),
    ]);
    return await app.getContext();
  } catch {
    return null;
  }
}

async function getToken(ctx) {
  const inTeamsWithNaa = ctx && nestedAppAuth.isNAAChannelRecommended();
  const auth = {
    clientId: cfg.clientId,
    authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
  };
  const pca = inTeamsWithNaa
    ? await createNestablePublicClientApplication({ auth: { ...auth, supportsNestedAppAuth: true } })
    : await createStandardPublicClientApplication({
        auth: { ...auth, redirectUri: window.location.origin + window.location.pathname },
        cache: { cacheLocation: "sessionStorage" },
      });

  let account = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
  if (!account && ctx?.user?.id) {
    account = pca.getAccount({ homeAccountId: undefined, localAccountId: ctx.user.id, tenantId: ctx.user.tenant?.id });
  }
  const request = { scopes: SCOPES, account: account || undefined, loginHint: ctx?.user?.loginHint };

  try {
    if (!account && !inTeamsWithNaa) throw new InteractionRequiredAuthError("no_account");
    return (await pca.acquireTokenSilent(request)).accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError || !account) {
      // Needs a click: browsers block popups that are not user-initiated.
      return await new Promise((resolve, reject) => {
        status("Please sign in to load your groups.");
        const btn = $("signin");
        btn.hidden = false;
        btn.onclick = async () => {
          btn.hidden = true;
          try { resolve((await pca.acquireTokenPopup(request)).accessToken); }
          catch (err) { reject(err); }
        };
      });
    }
    throw e;
  }
}

async function graph(token, url) {
  const items = [];
  while (url) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!r.ok) throw new Error(`Graph ${r.status}: ${(await r.text()).slice(0, 200)}`);
    const j = await r.json();
    if (j.value) items.push(...j.value); else return j;
    url = j["@odata.nextLink"];
  }
  return items;
}

function groupKind(g) {
  if (g.groupTypes?.includes("Unified")) return "Microsoft 365";
  if (g.securityEnabled && g.mailEnabled) return "Mail-enabled security";
  if (g.securityEnabled) return "Security";
  if (g.mailEnabled) return "Distribution list";
  return "Other";
}

async function main() {
  if (!cfg.clientId || cfg.clientId.startsWith("PASTE")) {
    status("Set clientId and tenantId in config.js first.", true);
    return;
  }
  const ctx = await initTeams();
  $("host").textContent = ctx ? `Running inside ${ctx.app?.host?.name || "Teams"}` : "Running in a browser (outside Teams)";
  if (ctx?.user) $("name").textContent = ctx.user.userPrincipalName || "";

  try {
    status("Signing in…");
    const token = await getToken(ctx);
    status("Loading your profile and groups…");
    const [me, groups] = await Promise.all([
      graph(token, "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,department"),
      graph(token, "https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.group?$select=id,displayName,description,mail,groupTypes,securityEnabled,mailEnabled&$top=999"),
    ]);

    $("hello").textContent = `Hello ${me.displayName} 👋`;
    $("name").textContent = [me.userPrincipalName, me.jobTitle, me.department].filter(Boolean).join(" · ");

    groups.sort((a, b) => (a.displayName || "").localeCompare(b.displayName || ""));
    $("count").textContent = `${groups.length} group${groups.length === 1 ? "" : "s"}`;
    const tbody = $("groups");
    tbody.innerHTML = "";
    for (const g of groups) {
      const tr = document.createElement("tr");
      for (const v of [g.displayName || "(no name)", groupKind(g), g.mail || "—"]) {
        const td = document.createElement("td");
        td.textContent = v;
        tr.appendChild(td);
      }
      tr.title = g.description || "";
      tbody.appendChild(tr);
    }
    $("filter").oninput = (e) => {
      const q = e.target.value.toLowerCase();
      for (const tr of tbody.rows) tr.hidden = !tr.textContent.toLowerCase().includes(q);
    };
    $("results").hidden = false;
    status("");
  } catch (e) {
    console.error(e);
    status(`Error: ${e.message || e}`, true);
  }
}

main();

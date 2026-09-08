#!/usr/bin/env node
import dns from "node:dns";
dns.setDefaultResultOrder("ipv4first");
import http from "node:http";
import { spawn } from "node:child_process";
import {
  chmodSync,
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  unlinkSync,
  writeFileSync,
} from "node:fs";
import { join, dirname } from "node:path";
import { randomBytes } from "node:crypto";

const HOME = process.env.HOME;
const EMAIL = "YOUR_BUSINESS_EMAIL";
const CRED_FILE = process.env.GOOGLE_OAUTH_CREDENTIALS
  || join(HOME, ".config/google-workspace-mcp/oauth_credentials.json");
const PRIMARY_DIR = process.env.CREDS_DIR || join(HOME, ".config/google-workspace-mcp");
const TOKEN_PATH = join(PRIMARY_DIR, `${EMAIL}.json`);
const CONSENT_TIMEOUT_MS = Number(process.env.CONSENT_TIMEOUT_MS || 600000);
const INCREMENTAL_GROUP_SCOPE = process.env.INCREMENTAL_GROUP_SCOPE === "1";
const INCREMENTAL_SCOPE_REPAIR = process.env.INCREMENTAL_SCOPE_REPAIR === "1";

const FULL_SCOPES = [
  "openid",
  "https://www.googleapis.com/auth/userinfo.email",
  "https://www.googleapis.com/auth/userinfo.profile",
  "https://www.googleapis.com/auth/gmail.modify",
  "https://www.googleapis.com/auth/gmail.compose",
  "https://www.googleapis.com/auth/gmail.send",
  "https://www.googleapis.com/auth/gmail.labels",
  "https://www.googleapis.com/auth/gmail.readonly",
  "https://www.googleapis.com/auth/gmail.settings.basic",
  "https://www.googleapis.com/auth/gmail.settings.sharing",
  "https://www.googleapis.com/auth/admin.directory.user.readonly",
  "https://www.googleapis.com/auth/admin.directory.user.alias.readonly",
  "https://www.googleapis.com/auth/admin.directory.user.alias",
  "https://www.googleapis.com/auth/admin.directory.group.readonly",
  "https://www.googleapis.com/auth/admin.directory.group",
  "https://www.googleapis.com/auth/apps.groups.settings",
  "https://www.googleapis.com/auth/calendar",
  "https://www.googleapis.com/auth/calendar.events",
  "https://www.googleapis.com/auth/calendar.readonly",
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/documents",
  "https://www.googleapis.com/auth/documents.readonly",
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/spreadsheets.readonly",
  "https://www.googleapis.com/auth/presentations",
  "https://www.googleapis.com/auth/presentations.readonly",
  "https://www.googleapis.com/auth/forms.body",
  "https://www.googleapis.com/auth/forms.body.readonly",
  "https://www.googleapis.com/auth/forms.responses.readonly",
  "https://www.googleapis.com/auth/tasks",
  "https://www.googleapis.com/auth/tasks.readonly",
  "https://www.googleapis.com/auth/contacts",
  "https://www.googleapis.com/auth/contacts.readonly",
  "https://www.googleapis.com/auth/chat.messages",
  "https://www.googleapis.com/auth/chat.messages.readonly",
  "https://www.googleapis.com/auth/chat.spaces",
  "https://www.googleapis.com/auth/chat.spaces.readonly",
  "https://www.googleapis.com/auth/cse",
  "https://www.googleapis.com/auth/script.projects",
  "https://www.googleapis.com/auth/script.projects.readonly",
  "https://www.googleapis.com/auth/script.deployments",
  "https://www.googleapis.com/auth/script.deployments.readonly",
  "https://www.googleapis.com/auth/script.processes",
  "https://www.googleapis.com/auth/script.metrics",
  "https://www.googleapis.com/auth/script.external_request",
  "https://www.googleapis.com/auth/script.scriptapp",
  "https://www.googleapis.com/auth/analytics.readonly",
  "https://www.googleapis.com/auth/webmasters.readonly",
  "https://www.googleapis.com/auth/content",
];
function loadExistingScopes() {
  if (!existsSync(TOKEN_PATH)) return [];
  const token = JSON.parse(readFileSync(TOKEN_PATH, "utf-8"));
  if (!Array.isArray(token.scopes) || !token.scopes.every(scope => typeof scope === "string")) {
    throw new Error(`Existing credential has no valid scopes array: ${TOKEN_PATH}`);
  }
  return [...new Set(token.scopes)];
}

const EXISTING_SCOPES = INCREMENTAL_GROUP_SCOPE ? loadExistingScopes() : [];
const GROUP_SCOPES = [
  "https://www.googleapis.com/auth/admin.directory.group.readonly",
  "https://www.googleapis.com/auth/admin.directory.group",
  "https://www.googleapis.com/auth/apps.groups.settings",
];
const INCREMENTAL_REPAIR_SCOPES = [
  "https://www.googleapis.com/auth/documents",
  "https://www.googleapis.com/auth/documents.readonly",
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/spreadsheets.readonly",
  "https://www.googleapis.com/auth/presentations",
  "https://www.googleapis.com/auth/presentations.readonly",
  "https://www.googleapis.com/auth/tasks",
  "https://www.googleapis.com/auth/tasks.readonly",
];
const SCOPES = INCREMENTAL_GROUP_SCOPE
  ? [...new Set([
      ...GROUP_SCOPES,
      ...(INCREMENTAL_SCOPE_REPAIR ? INCREMENTAL_REPAIR_SCOPES : []),
      ...FULL_SCOPES.filter(scope => !EXISTING_SCOPES.includes(scope)),
    ])]
  : FULL_SCOPES;
const REQUIRED_RESULT_SCOPES = [
  ...new Set([
    ...FULL_SCOPES,
    ...(INCREMENTAL_GROUP_SCOPE ? EXISTING_SCOPES : []),
  ]),
];

function loadClient() {
  if (!existsSync(CRED_FILE)) { console.error(`FATAL: client secrets not found: ${CRED_FILE}`); process.exit(2); }
  const raw = JSON.parse(readFileSync(CRED_FILE, "utf-8"));
  const c = raw.installed || raw.web || raw;
  if (!c.client_id || !c.client_secret) { console.error("FATAL: client_id/client_secret missing in client secrets"); process.exit(2); }
  return { client_id: c.client_id, client_secret: c.client_secret };
}

function writeToken(dir, tok) {
  if (!existsSync(dir)) mkdirSync(dir, { recursive: true, mode: 0o700 });
  const path = join(dir, `${EMAIL}.json`);
  if (existsSync(path)) {
    const backupPath = `${path}.prefullscope.bak`;
    copyFileSync(path, backupPath);
    chmodSync(backupPath, 0o600);
  }
  const tempPath = `${path}.tmp-${process.pid}-${randomBytes(6).toString("hex")}`;
  try {
    writeFileSync(tempPath, JSON.stringify(tok, null, 2), {
      mode: 0o600,
      flag: "wx",
    });
    renameSync(tempPath, path);
  } catch (error) {
    try { unlinkSync(tempPath); } catch {}
    throw error;
  }
  console.log(`  wrote ${path}`);
}

const { client_id, client_secret } = loadClient();
const state = randomBytes(16).toString("hex");

const server = http.createServer(async (req, res) => {
  const u = new URL(req.url, "http://localhost");
  if (!u.searchParams.has("code") && !u.searchParams.has("error")) {
    res.writeHead(204); res.end(); return;
  }
  const finish = (msg) => { res.writeHead(200, { "content-type": "text/html" }); res.end(`<html><body style="font-family:sans-serif;padding:2rem"><h2>${msg}</h2><p>You can close this tab and return to the terminal.</p></body></html>`); };
  if (u.searchParams.get("error")) { finish(`Consent error: ${u.searchParams.get("error")}`); console.error("CONSENT ERROR:", u.searchParams.get("error")); server.close(); process.exit(1); }
  if (u.searchParams.get("state") !== state) { finish("State mismatch — aborted."); console.error("FATAL: state mismatch"); server.close(); process.exit(1); }
  const code = u.searchParams.get("code");
  try {
    const body = new URLSearchParams({ code, client_id, client_secret, redirect_uri: REDIRECT, grant_type: "authorization_code" });
    const r = await fetch("https://oauth2.googleapis.com/token", { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body });
    const t = await r.json();
    if (!r.ok || !t.refresh_token) {
      finish("Token exchange failed — see terminal.");
      console.error("TOKEN EXCHANGE FAILED:", JSON.stringify({
        status: r.status,
        error: t?.error,
        error_description: t?.error_description,
        has_access_token: Boolean(t?.access_token),
        has_refresh_token: Boolean(t?.refresh_token),
      }));
      console.error(t.refresh_token ? "" : "No refresh_token returned — ensure prompt=consent + access_type=offline (it is) and that you fully re-approved.");
      server.close(); process.exit(1);
    }
    const expiry = new Date(Date.now() + (Number(t.expires_in || 3600) * 1000)).toISOString().replace(/\.\d+Z$/, ".000000Z");
    const tok = {
      token: t.access_token,
      refresh_token: t.refresh_token,
      token_uri: "https://oauth2.googleapis.com/token",
      client_id, client_secret,
      scopes: (t.scope ? t.scope.split(" ") : SCOPES),
      expiry,
    };
    console.log("\n✅ Consent complete. Granted scopes:");
    const granted = tok.scopes.map(s => s.replace("https://www.googleapis.com/auth/", ""));
    const required = [
      "admin.directory.group.readonly",
      "admin.directory.group",
      "apps.groups.settings",
      "analytics.readonly",
      "webmasters.readonly",
      "content",
    ];
    required.forEach(s => {
      console.log(`   ${granted.includes(s) ? "✓" : "✗ MISSING"} ${s}`);
    });
    console.log(`   (+${tok.scopes.length} scopes total)`);
    const missing = REQUIRED_RESULT_SCOPES.filter(
      scope => !tok.scopes.includes(scope),
    );
    if (missing.length) {
      finish("Consent was incomplete; the existing credential was left unchanged.");
      console.error(
        "FATAL: missing required scopes: "
        + missing.map(scope => scope.replace("https://www.googleapis.com/auth/", "")).join(", "),
      );
      server.close();
      process.exit(3);
      return;
    }
    writeToken(PRIMARY_DIR, tok);
    finish("Re-authorised successfully ✓");
    server.close();
    process.exit(0);
  } catch (e) { finish("Exchange error — see terminal."); console.error("FATAL", e.message); server.close(); process.exit(1); }
});

let REDIRECT;
server.listen(0, "127.0.0.1", () => {
  const port = server.address().port;
  REDIRECT = `http://localhost:${port}/`;
  const auth = new URL("https://accounts.google.com/o/oauth2/v2/auth");
  auth.searchParams.set("client_id", client_id);
  auth.searchParams.set("redirect_uri", REDIRECT);
  auth.searchParams.set("response_type", "code");
  auth.searchParams.set("scope", SCOPES.join(" "));
  auth.searchParams.set("access_type", "offline");
  auth.searchParams.set("prompt", "consent");
  auth.searchParams.set("login_hint", EMAIL);
  auth.searchParams.set("state", state);
  if (INCREMENTAL_GROUP_SCOPE) {
    auth.searchParams.set("include_granted_scopes", "true");
  }
  console.log(`\nListening for the OAuth callback on ${REDIRECT}\n`);
  console.log("OPEN THIS URL in a browser that can reach this host's localhost, sign in as");
  console.log(`${EMAIL}, and APPROVE EVERY permission:\n`);
  console.log(auth.toString());
  console.log(`\n(Waiting up to ${Math.ceil(CONSENT_TIMEOUT_MS / 60000)} minutes for you to approve…)\n`);
  if (process.env.NO_OPEN !== "1") {
    const opener = process.platform === "darwin" ? "open" : "xdg-open";
    const child = spawn(opener, [auth.toString()], {
      detached: true,
      stdio: "ignore",
    });
    child.on("error", () => {});
    child.unref();
  }
});
setTimeout(() => {
  console.error(`Timed out waiting for consent (${Math.ceil(CONSENT_TIMEOUT_MS / 60000)} min). Re-run when ready.`);
  process.exit(1);
}, CONSENT_TIMEOUT_MS);

// Pages Functions middleware — shared-password gate for Element Delivery Map.
//
// Runs on every request to map.elementdeliverymap.ca BEFORE the static asset is
// served. If the signed auth cookie is present + valid, passes through to the
// map. Otherwise shows the sign-in page. Deploys as part of the existing
// `wrangler pages deploy` flow — no separate Worker to manage.
//
// Deployed because Cloudflare Access one-time-PIN emails are eaten by Microsoft
// 365 Defender Safe Links for exec-tier users (April 2026 — D. Procopio case).
// Shared password sidesteps the email dependency entirely.
//
// Required environment variables (set in Cloudflare dashboard →
// Pages → element-delivery-map → Settings → Environment variables → Production,
// each marked "Encrypted"):
//   MAP_PASSWORD       — the shared password (distribute via Slack / 1Password).
//                        Rotate by updating this and redeploying.
//   MAP_COOKIE_SECRET  — random 32+ char hex string used to HMAC-sign the cookie.
//                        Generate once; rotating invalidates all active sessions.
//
// Cookie format:
//   map_auth = <expiryUnixSec>.<hmacHex>
//   HMAC-SHA256 signed so clients can't forge the expiry.
//   30-day lifetime — re-prompts on expiry. HttpOnly + Secure + SameSite=Lax.
//
// Routes added:
//   GET  /__auth/login   — login form (also shown when cookie absent/invalid)
//   POST /__auth/login   — validates password, sets cookie, redirects to /
//   GET  /__auth/logout  — clears cookie

const COOKIE_NAME = "map_auth";
const COOKIE_TTL_SEC = 60 * 60 * 24 * 30;   // 30 days
const LOGIN_PATH  = "/__auth/login";
const LOGOUT_PATH = "/__auth/logout";
const MIN_SECRET_LEN = 16;  // Guardrail: refuse to validate with a weak secret.

async function hmacHex(message, secret) {
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sig = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(message));
  return [...new Uint8Array(sig)]
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}

// Timing-safe string compare — constant-time XOR over max(a,b) length so
// neither the password value nor its length can be inferred via response timing.
function safeEq(a, b) {
  if (typeof a !== "string" || typeof b !== "string") return false;
  const maxLen = Math.max(a.length, b.length);
  // XOR lengths first so mismatched lengths always produce diff != 0.
  let diff = a.length ^ b.length;
  for (let i = 0; i < maxLen; i++) {
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }
  return diff === 0;
}

function parseCookie(header, name) {
  if (!header) return null;
  const parts = header.split(/;\s*/);
  for (const p of parts) {
    const eq = p.indexOf("=");
    if (eq < 0) continue;
    if (p.slice(0, eq) === name) return p.slice(eq + 1);
  }
  return null;
}

async function isAuthed(request, env) {
  // Fail closed if secrets are missing/weak — don't silently accept anything.
  if (!env.MAP_COOKIE_SECRET || env.MAP_COOKIE_SECRET.length < MIN_SECRET_LEN) return false;
  if (!env.MAP_PASSWORD) return false;

  const cookieHeader = request.headers.get("Cookie");
  const cookie = parseCookie(cookieHeader, COOKIE_NAME);
  if (!cookie) return false;

  const dot = cookie.indexOf(".");
  if (dot < 0) return false;
  const expStr = cookie.slice(0, dot);
  const sig    = cookie.slice(dot + 1);

  const exp = parseInt(expStr, 10);
  if (!Number.isFinite(exp)) return false;
  if (exp < Math.floor(Date.now() / 1000)) return false;

  const expected = await hmacHex(expStr, env.MAP_COOKIE_SECRET);
  return safeEq(sig, expected);
}

async function issueCookie(env) {
  const exp = Math.floor(Date.now() / 1000) + COOKIE_TTL_SEC;
  const sig = await hmacHex(String(exp), env.MAP_COOKIE_SECRET);
  return `${COOKIE_NAME}=${exp}.${sig}; Path=/; Max-Age=${COOKIE_TTL_SEC}; Secure; SameSite=Lax; HttpOnly`;
}

function clearCookie() {
  return `${COOKIE_NAME}=; Path=/; Max-Age=0; Secure; SameSite=Lax; HttpOnly`;
}

function loginHtml(errorMsg) {
  // Element brand — teal #2A5A63, off-white #F5F7F7, mint accent.
  const err = errorMsg
    ? `<div class="err" role="alert">${errorMsg}</div>`
    : "";
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>Element Delivery Map &mdash; Sign in</title>
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: #F5F7F7; color: #1E2F33;
    display: flex; align-items: center; justify-content: center;
    min-height: 100vh; padding: 20px;
  }
  .card {
    background: #fff; border-radius: 10px; padding: 32px;
    box-shadow: 0 2px 16px rgba(42,90,99,.08);
    width: 100%; max-width: 380px;
    border-top: 4px solid #2A5A63;
  }
  h1 { margin: 0 0 4px; font-size: 20px; color: #2A5A63; font-weight: 700; letter-spacing: -0.01em; }
  .sub { font-size: 13px; color: #6B7A7E; margin-bottom: 24px; }
  label { display: block; font-size: 12px; font-weight: 600; color: #6B7A7E; margin-bottom: 6px; }
  input[type=password] {
    width: 100%; padding: 10px 12px; font-size: 15px;
    border: 1px solid #D5DCDE; border-radius: 6px; background: #fff;
    color: #1E2F33; outline: none;
    transition: border-color .15s, box-shadow .15s;
  }
  input[type=password]:focus {
    border-color: #2A5A63;
    box-shadow: 0 0 0 3px rgba(42,90,99,.12);
  }
  button {
    width: 100%; margin-top: 16px; padding: 10px 12px;
    font-size: 14px; font-weight: 600; font-family: inherit;
    background: #2A5A63; color: #fff; border: none; border-radius: 6px;
    cursor: pointer; transition: background .15s;
  }
  button:hover  { background: #224750; }
  button:active { background: #1a383f; }
  .err {
    margin-bottom: 16px; padding: 8px 12px;
    background: #FDEBEA; color: #9A2925;
    border-radius: 6px; font-size: 12px;
    border: 1px solid #F5B8B5;
  }
  .footer {
    margin-top: 20px; font-size: 11px; color: #9AAAAF;
    text-align: center; letter-spacing: 0.02em;
  }
</style>
</head>
<body>
  <form class="card" method="POST" action="${LOGIN_PATH}">
    <h1>Element Delivery Map</h1>
    <div class="sub">Internal tool &mdash; sign in to continue</div>
    ${err}
    <label for="pw">Password</label>
    <input id="pw" name="password" type="password" autofocus autocomplete="current-password" required>
    <button type="submit">Sign in</button>
    <div class="footer">Element Event Solutions</div>
  </form>
</body>
</html>`;
}

function htmlResponse(body, status, extraHeaders = {}) {
  return new Response(body, {
    status,
    headers: {
      "Content-Type": "text/html; charset=utf-8",
      "Cache-Control": "no-store",
      "X-Frame-Options": "DENY",
      "X-Content-Type-Options": "nosniff",
      "Referrer-Policy": "no-referrer",
      ...extraHeaders,
    },
  });
}

export async function onRequest(context) {
  const { request, env, next } = context;
  const url = new URL(request.url);

  // ---- Auth endpoints (always reachable, no cookie required) ----

  // POST /__auth/login — validate password, set cookie, redirect to /.
  if (url.pathname === LOGIN_PATH && request.method === "POST") {
    // Fail closed if the app isn't configured.
    if (!env.MAP_PASSWORD || !env.MAP_COOKIE_SECRET ||
        env.MAP_COOKIE_SECRET.length < MIN_SECRET_LEN) {
      return htmlResponse(
        loginHtml("Service not configured. Contact the administrator."),
        503
      );
    }
    const form = await request.formData();
    const submitted = form.get("password") || "";
    if (safeEq(String(submitted), String(env.MAP_PASSWORD))) {
      const cookie = await issueCookie(env);
      return new Response("", {
        status: 303,
        headers: { "Location": "/", "Set-Cookie": cookie, "Cache-Control": "no-store" },
      });
    }
    return htmlResponse(loginHtml("Incorrect password"), 401);
  }

  // GET /__auth/login — show login form
  if (url.pathname === LOGIN_PATH) {
    return htmlResponse(loginHtml(""), 200);
  }

  // GET /__auth/logout — clear cookie, show login
  if (url.pathname === LOGOUT_PATH) {
    return htmlResponse(loginHtml(""), 200, { "Set-Cookie": clearCookie() });
  }

  // ---- All other paths require a valid cookie ----
  if (await isAuthed(request, env)) {
    return next();  // serve the static asset (index.html, JSON, etc.)
  }

  // Not authed. For HTML navigations show the login page; for asset fetches
  // (JSON, JS, images) return a 401 so the browser surfaces it cleanly.
  const accept = request.headers.get("Accept") || "";
  if (request.method === "GET" && accept.includes("text/html")) {
    return htmlResponse(loginHtml(""), 200);
  }
  return new Response("Unauthorized", {
    status: 401,
    headers: { "Cache-Control": "no-store" },
  });
}

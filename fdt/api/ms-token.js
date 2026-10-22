// Server-side half of the SharePoint sign-in. Microsoft limits sign-ins for browser-only apps to
// 24 hours, but gives ~90 days to apps that can prove who they are with a client secret — and a
// browser can't keep a secret. So the browser sends its one-time code (or its refresh token) here,
// this function adds the secret and passes the request on to Microsoft, and hands back the answer.
// The secret lives only in the MS_CLIENT_SECRET environment variable on Vercel.

const CLIENT_ID = "50108c90-8844-4fbc-96af-d4fb7e7fa4ca";
const TENANT_ID = "ba936175-44a3-4888-b75e-6f814421a09c";
const SCOPES = "Files.ReadWrite.All Sites.ReadWrite.All User.Read offline_access";
const ALLOWED_ORIGIN = "https://www.doctracker.org";
// Fixed here on purpose: the browser doesn't get to choose where Microsoft sends the user back.
const REDIRECT_URI = "https://www.doctracker.org/auth/callback";
const MAX_FIELD_LENGTH = 8192;

const send = (res, status, body) => {
  res.status(status);
  res.setHeader("Cache-Control", "no-store");
  res.setHeader("Content-Type", "application/json");
  res.send(JSON.stringify(body));
};

const isShortString = (v) => typeof v === "string" && v.length > 0 && v.length <= MAX_FIELD_LENGTH;

module.exports = async function handler(req, res) {
  // Setup check: confirms the function deployed and the secret is set, without revealing it.
  if (req.method === "GET") {
    return send(res, 200, { ok: true, secretConfigured: !!process.env.MS_CLIENT_SECRET });
  }
  if (req.method !== "POST") return send(res, 405, { error: "method_not_allowed" });
  if (req.headers.origin !== ALLOWED_ORIGIN) return send(res, 403, { error: "forbidden_origin" });
  if (!process.env.MS_CLIENT_SECRET) return send(res, 500, { error: "server_not_configured" });

  let body = req.body;
  if (typeof body === "string") {
    try { body = JSON.parse(body); } catch { return send(res, 400, { error: "bad_request" }); }
  }
  if (!body || typeof body !== "object") return send(res, 400, { error: "bad_request" });

  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: process.env.MS_CLIENT_SECRET,
    scope: SCOPES,
  });

  if (body.grant_type === "authorization_code") {
    if (!isShortString(body.code) || !isShortString(body.code_verifier)) return send(res, 400, { error: "bad_request" });
    params.set("grant_type", "authorization_code");
    params.set("code", body.code);
    params.set("code_verifier", body.code_verifier);
    params.set("redirect_uri", REDIRECT_URI);
  } else if (body.grant_type === "refresh_token") {
    if (!isShortString(body.refresh_token)) return send(res, 400, { error: "bad_request" });
    params.set("grant_type", "refresh_token");
    params.set("refresh_token", body.refresh_token);
  } else {
    return send(res, 400, { error: "unsupported_grant_type" });
  }

  try {
    const ms = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params,
    });
    const data = await ms.json();
    return send(res, ms.status, data);
  } catch {
    // Deliberately no detail — and nothing from the request is logged anywhere.
    return send(res, 502, { error: "upstream_unreachable" });
  }
};

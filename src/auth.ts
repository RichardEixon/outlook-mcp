/**
 * Microsoft OAuth 2.0 authentication using the standard Authorization Code flow.
 * Tokens are obtained via REST calls to the Microsoft identity platform directly,
 * keeping dependencies minimal and making the refresh-token flow fully transparent.
 */

import dotenv from "dotenv";
dotenv.config();

// ── Constants ─────────────────────────────────────────────────────────────────

/** Microsoft Graph permissions this server requires */
const SCOPES = [
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "https://graph.microsoft.com/Mail.Read",
  "https://graph.microsoft.com/Mail.ReadWrite",
  "https://graph.microsoft.com/Tasks.ReadWrite",
  "offline_access",
  "openid",
  "profile",
].join(" ");

function tenantId(): string {
  return process.env.MICROSOFT_TENANT_ID || "common";
}

function tokenUrl(): string {
  return `https://login.microsoftonline.com/${tenantId()}/oauth2/v2.0/token`;
}

function authUrl(): string {
  return `https://login.microsoftonline.com/${tenantId()}/oauth2/v2.0/authorize`;
}

function getRedirectUri(): string {
  return (
    process.env.REDIRECT_URI ||
    `http://localhost:${process.env.PORT || 3000}/auth/callback`
  );
}

// ── In-memory access-token cache ──────────────────────────────────────────────

let tokenCache: { accessToken: string; expiresAt: number } | null = null;

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Returns the URL the user must visit to start the OAuth login flow.
 */
export function getAuthorizationUrl(): string {
  const clientId = requireEnv("MICROSOFT_CLIENT_ID");
  const params = new URLSearchParams({
    client_id: clientId,
    response_type: "code",
    redirect_uri: getRedirectUri(),
    scope: SCOPES,
    response_mode: "query",
    // A random state value helps prevent CSRF — keep it simple for now
    state: Date.now().toString(),
  });
  return `${authUrl()}?${params.toString()}`;
}

/**
 * Exchanges a one-time authorization code (from the callback URL) for
 * access + refresh tokens.
 */
export async function exchangeCodeForTokens(code: string): Promise<{
  accessToken: string;
  refreshToken: string;
  expiresIn: number;
}> {
  const body = new URLSearchParams({
    client_id: requireEnv("MICROSOFT_CLIENT_ID"),
    client_secret: requireEnv("MICROSOFT_CLIENT_SECRET"),
    code,
    redirect_uri: getRedirectUri(),
    grant_type: "authorization_code",
    scope: SCOPES,
  });

  const data = await postToTokenEndpoint(body);
  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresIn: data.expires_in,
  };
}

/**
 * Returns a valid access token, using the in-memory cache or refreshing
 * automatically when the token is close to expiry (< 5 minutes remaining).
 *
 * Requires `MICROSOFT_REFRESH_TOKEN` to be set in the environment.
 */
export async function getAccessToken(): Promise<string> {
  // Serve from cache with a 5-minute safety buffer
  if (tokenCache && Date.now() < tokenCache.expiresAt - 5 * 60 * 1000) {
    return tokenCache.accessToken;
  }

  const refreshToken = process.env.MICROSOFT_REFRESH_TOKEN;
  if (!refreshToken) {
    throw new Error(
      "Not authenticated. Please visit /auth on the server to log in with Microsoft."
    );
  }

  const body = new URLSearchParams({
    client_id: requireEnv("MICROSOFT_CLIENT_ID"),
    client_secret: requireEnv("MICROSOFT_CLIENT_SECRET"),
    refresh_token: refreshToken,
    grant_type: "refresh_token",
    scope: SCOPES,
  });

  const data = await postToTokenEndpoint(body);

  // Update cache
  tokenCache = {
    accessToken: data.access_token,
    expiresAt: Date.now() + data.expires_in * 1000,
  };

  // Microsoft sometimes rotates the refresh token — keep it current
  if (data.refresh_token && data.refresh_token !== refreshToken) {
    process.env.MICROSOFT_REFRESH_TOKEN = data.refresh_token;
    console.log(
      "🔄 Refresh token rotated. Update MICROSOFT_REFRESH_TOKEN in your Replit Secrets."
    );
    console.log(`   New value: ${data.refresh_token}`);
  }

  return data.access_token;
}

/** Invalidates the in-memory token cache (useful after re-authentication). */
export function clearTokenCache(): void {
  tokenCache = null;
}

// ── Internal helpers ──────────────────────────────────────────────────────────

async function postToTokenEndpoint(
  body: URLSearchParams
): Promise<Record<string, string>> {
  const response = await fetch(tokenUrl(), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  const data = (await response.json()) as Record<string, string>;

  if (!response.ok) {
    const detail = data.error_description || data.error || JSON.stringify(data);
    throw new Error(`Microsoft token request failed: ${detail}`);
  }

  return data;
}

function requireEnv(key: string): string {
  const value = process.env[key];
  if (!value) {
    throw new Error(
      `Missing required environment variable: ${key}. Check your .env file.`
    );
  }
  return value;
}

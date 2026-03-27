# Outlook MCP Server

A **Model Context Protocol (MCP)** server that connects Claude to your Microsoft Outlook — letting you query your calendar, read emails, and manage tasks just by chatting with Claude.

---

## What You Can Do

Once connected, you can ask Claude things like:

- _"What meetings do I have this week?"_
- _"Show me unread emails from John"_
- _"Create a meeting tomorrow at 2 PM with sarah@example.com"_
- _"What tasks are due this Friday?"_
- _"Mark task XYZ as done"_

---

## Architecture

```
Claude (web) ──SSE──► Replit (this server) ──Graph API──► Microsoft Outlook
```

The server runs permanently on Replit. Both your work PC and home PC connect to it via Claude's web interface — no local installation needed.

---

## Step 1 — Register an Azure App

You need an Azure app to get API access to your Outlook.

1. Go to [portal.azure.com](https://portal.azure.com) and sign in with your Microsoft account.
2. Search for **"App registrations"** → click **New registration**.
3. Fill in:
   - **Name**: `Outlook MCP` (or anything you like)
   - **Supported account types**: _Accounts in any organisational directory and personal Microsoft accounts_
   - **Redirect URI**: Select **Web** and enter:
     - For Replit: `https://YOUR-REPL-NAME.YOUR-USERNAME.repl.co/auth/callback`
     - For local: `http://localhost:3000/auth/callback`
4. Click **Register**.

### Get Your Credentials

After registration:

- **Client ID**: On the app overview page → copy **Application (client) ID**
- **Tenant ID**: Copy **Directory (tenant) ID**
  _(Use `common` if you have a personal Outlook.com account)_

### Create a Client Secret

1. Left sidebar → **Certificates & secrets** → **New client secret**
2. Add a description (e.g. `MCP Server`), choose **24 months** expiry
3. Click **Add** → **immediately copy the `Value`** (it's hidden after you leave the page)

### Set API Permissions

1. Left sidebar → **API permissions** → **Add a permission** → **Microsoft Graph**
2. Choose **Delegated permissions** and add:
   - `Calendars.ReadWrite`
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Tasks.ReadWrite`
   - `offline_access`
3. Click **Grant admin consent** (if you have admin rights) — otherwise these are granted when you log in.

---

## Step 2 — Deploy on Replit

1. Go to [replit.com](https://replit.com) → **Create Repl** → **Import from GitHub** (or upload files manually).
2. Upload all files from this project.
3. Open the **Secrets** panel (lock icon in the sidebar) and add:

   | Key | Value |
   |-----|-------|
   | `MICROSOFT_CLIENT_ID` | Your Azure App's client ID |
   | `MICROSOFT_CLIENT_SECRET` | Your client secret value |
   | `MICROSOFT_TENANT_ID` | `common` (personal) or your tenant ID (work) |
   | `REDIRECT_URI` | `https://your-repl-name.your-username.repl.co/auth/callback` |
   | `PORT` | `3000` |

4. In the Replit shell, run:
   ```bash
   npm install
   npm run build
   npm start
   ```

5. Replit will show your public URL (e.g. `https://outlook-mcp.username.repl.co`).

### Keep It Running

Replit free tier sleeps after inactivity. To keep the server awake:
- Use [UptimeRobot](https://uptimerobot.com) — add a free HTTP monitor pointing to `https://your-repl.repl.co/health`
- Or upgrade to Replit's paid plan.

---

## Step 3 — Authenticate with Microsoft

1. Visit `https://your-repl.repl.co/auth` in your browser.
2. Log in with your Microsoft / Outlook account.
3. After login, you'll see a **Refresh Token** displayed on screen.
4. Copy it and add it as a Replit Secret:

   | Key | Value |
   |-----|-------|
   | `MICROSOFT_REFRESH_TOKEN` | _(the long token string)_ |

5. Restart the Replit server. The status page at `/` should now show **Authenticated: Yes**.

> **Note**: The refresh token lasts 90 days for personal accounts. When it expires, visit `/auth` again.

---

## Step 4 — Connect Claude (Web)

1. Open [claude.ai](https://claude.ai) → click your profile → **Settings** → **Integrations** (or **MCP Servers**).
2. Add a new MCP server:
   - **Name**: `Outlook`
   - **URL**: `https://your-repl-name.your-username.repl.co/sse`
3. Save. Claude will connect and show the available tools.
4. Start a new conversation and try: _"What's on my calendar today?"_

---

## Running Locally (Optional)

```bash
# 1. Install dependencies
npm install

# 2. Copy and fill in env file
cp .env.example .env
# Edit .env with your Azure credentials

# 3. Run in dev mode (hot reload)
npm run dev

# 4. Visit http://localhost:3000/auth to authenticate
# 5. Copy the refresh token into your .env as MICROSOFT_REFRESH_TOKEN

# 6. Build for production
npm run build
npm start
```

---

## Project Structure

```
outlook-mcp/
├── src/
│   ├── index.ts          # Express server + MCP server setup
│   ├── auth.ts           # Microsoft OAuth 2.0 token management
│   ├── graph.ts          # Graph API client factory
│   ├── types.ts          # TypeScript types + response helpers
│   └── tools/
│       ├── calendar.ts   # 6 calendar tools
│       ├── email.ts      # 4 email tools
│       └── tasks.ts      # 5 task tools
├── .env.example          # Environment variable template
├── .gitignore
├── package.json
├── tsconfig.json
├── README.md
└── TOOLS.md              # Detailed tool reference
```

---

## Troubleshooting

### "Not authenticated. Please visit /auth"
The server has no refresh token. Visit `/auth` on your Replit URL and log in.

### "Microsoft token request failed: invalid_grant"
Your refresh token has expired (happens after 90 days of inactivity, or if you re-authenticated elsewhere).
→ Visit `/auth` again and save the new refresh token to Replit Secrets.

### "Missing required environment variable: MICROSOFT_CLIENT_ID"
The env vars aren't set. Double-check your Replit Secrets panel — keys are case-sensitive.

### "AADSTS700016: Application not found"
The `MICROSOFT_CLIENT_ID` is wrong, or the app was registered in a different tenant.
→ Check the Azure portal and use `common` as the tenant ID for personal accounts.

### "AADSTS65001: The user or administrator has not consented"
The API permissions haven't been granted yet.
→ In Azure Portal → your app → API permissions → Grant admin consent.
→ Or delete the app and re-register, then log in fresh via `/auth`.

### Claude can't connect to the MCP
- Check the Replit URL is correct and the server is running (visit `/health`)
- Make sure the URL ends in `/sse` in Claude's MCP settings
- Check the Replit server isn't sleeping (add an UptimeRobot monitor)

### "CORS" errors in browser console
These are expected for SSE connections — Claude's client handles this internally.

---

## Security Notes

- Never commit your `.env` file — it's in `.gitignore`
- Store secrets in Replit Secrets, not in code
- The refresh token gives full access to your Outlook — treat it like a password
- Client secrets expire: set a reminder to rotate yours before it expires in Azure

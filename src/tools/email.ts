/**
 * Email tools for the Outlook MCP Server.
 * Uses Microsoft Graph API /me/messages and /me/mailFolders endpoints.
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getGraphClient } from "../graph";
import { GraphMessage, McpToolResponse, ok, err } from "../types";

// ── Tool Definitions ──────────────────────────────────────────────────────────

export const emailTools: Tool[] = [
  {
    name: "search_emails",
    description:
      "Search Outlook emails by keyword, sender, or date range. Results are sorted newest-first.",
    inputSchema: {
      type: "object",
      properties: {
        keyword: {
          type: "string",
          description: "Search term — checked against subject and body.",
        },
        from: {
          type: "string",
          description: "Filter by sender's email address (exact match) or display name (partial).",
        },
        startDate: {
          type: "string",
          description: "Return only emails received on or after this date (YYYY-MM-DD).",
        },
        endDate: {
          type: "string",
          description: "Return only emails received on or before this date (YYYY-MM-DD).",
        },
        folder: {
          type: "string",
          description: "Folder to search: inbox | sentitems | drafts | deleteditems (default: inbox).",
        },
        limit: {
          type: "number",
          description: "Max emails to return (default: 20, max: 50).",
        },
      },
    },
  },
  {
    name: "get_email_details",
    description:
      "Get the full content, recipients, and metadata of a specific email.",
    inputSchema: {
      type: "object",
      properties: {
        emailId: {
          type: "string",
          description: "The email message ID (from search_emails or get_unread_emails).",
        },
      },
      required: ["emailId"],
    },
  },
  {
    name: "mark_email_as_read",
    description: "Mark an email as read or unread.",
    inputSchema: {
      type: "object",
      properties: {
        emailId: {
          type: "string",
          description: "The email message ID.",
        },
        isRead: {
          type: "boolean",
          description: "true = mark as read, false = mark as unread. Defaults to true.",
        },
      },
      required: ["emailId"],
    },
  },
  {
    name: "get_unread_emails",
    description: "Get unread emails from a folder, sorted by received date (newest first).",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Folder to check (default: inbox).",
        },
        limit: {
          type: "number",
          description: "Max emails to return (default: 20, max: 50).",
        },
      },
    },
  },
];

// ── Formatting Helpers ────────────────────────────────────────────────────────

const SG_LOCALE = "en-SG";
const SG_TZ = "Asia/Singapore";

function formatDateTime(iso: string): string {
  return new Date(iso).toLocaleString(SG_LOCALE, { timeZone: SG_TZ });
}

function stripHtml(html: string): string {
  return html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}

function formatEmailSummary(msg: GraphMessage): string {
  const received = formatDateTime(msg.receivedDateTime);
  const to = (msg.toRecipients || []).map((r) => r.emailAddress.address).join(", ");
  const unread = msg.isRead ? "" : " 📬";

  const lines: string[] = [
    `📧${unread} **${msg.subject || "(No Subject)"}**`,
    `   ID:       ${msg.id}`,
    `   From:     ${msg.from?.emailAddress.name || ""} <${msg.from?.emailAddress.address || ""}>`,
    `   To:       ${to}`,
    `   Received: ${received}`,
    `   Status:   ${msg.isRead ? "Read" : "Unread"}`,
    `   Priority: ${msg.importance || "normal"}`,
  ];
  if (msg.bodyPreview) lines.push(`   Preview:  ${msg.bodyPreview.slice(0, 150)}`);
  return lines.join("\n");
}

function formatEmailDetails(msg: GraphMessage): string {
  const received = formatDateTime(msg.receivedDateTime);
  const to = (msg.toRecipients || [])
    .map((r) => `${r.emailAddress.name} <${r.emailAddress.address}>`)
    .join(", ");
  const cc = (msg.ccRecipients || [])
    .map((r) => `${r.emailAddress.name} <${r.emailAddress.address}>`)
    .join(", ");

  const bodyText = msg.body
    ? msg.body.contentType === "html"
      ? stripHtml(msg.body.content).slice(0, 1500)
      : msg.body.content.slice(0, 1500)
    : "(no body)";

  const lines: string[] = [
    `📧 **${msg.subject || "(No Subject)"}**`,
    ``,
    `ID:         ${msg.id}`,
    `From:       ${msg.from?.emailAddress.name} <${msg.from?.emailAddress.address}>`,
    `To:         ${to}`,
  ];
  if (cc) lines.push(`CC:         ${cc}`);
  lines.push(
    `Received:   ${received}`,
    `Status:     ${msg.isRead ? "Read" : "📬 Unread"}`,
    `Importance: ${msg.importance || "normal"}`,
    `Has Files:  ${msg.hasAttachments ? "Yes" : "No"}`,
    ``,
    `── Body ─────────────────────────────────────────────`,
    bodyText
  );
  if (msg.webLink) lines.push(``, `Open in Outlook: ${msg.webLink}`);
  return lines.join("\n");
}

// ── Tool Handler ──────────────────────────────────────────────────────────────

export async function handleEmailTool(
  name: string,
  args: Record<string, unknown>
): Promise<McpToolResponse> {
  try {
    const client = await getGraphClient();

    switch (name) {
      // ── search_emails ─────────────────────────────────────────────────────
      case "search_emails": {
        const folder = (args.folder as string) || "inbox";
        const limit = Math.min((args.limit as number) || 20, 50);
        const filters: string[] = [];

        if (args.from) {
          // Try matching on email address; Graph $search handles display name
          filters.push(`from/emailAddress/address eq '${args.from}'`);
        }
        if (args.startDate) {
          filters.push(`receivedDateTime ge ${args.startDate}T00:00:00Z`);
        }
        if (args.endDate) {
          filters.push(`receivedDateTime le ${args.endDate}T23:59:59Z`);
        }

        const select =
          "id,subject,from,toRecipients,receivedDateTime,isRead,importance,bodyPreview";

        let apiCall = client
          .api(`/me/mailFolders/${folder}/messages`)
          .select(select)
          .top(limit)
          .orderby("receivedDateTime desc");

        if (args.keyword) {
          // $search and $filter cannot be combined — keyword takes precedence
          apiCall = apiCall.search(`"${args.keyword}"`);
        } else if (filters.length) {
          apiCall = apiCall.filter(filters.join(" and "));
        }

        const result = await apiCall.get();
        const messages: GraphMessage[] = result.value || [];

        if (!messages.length) return ok("📭 No emails found matching the criteria.");

        // If we used $search, apply date/sender filters client-side
        const filtered =
          args.keyword && filters.length
            ? messages.filter((m) => {
                if (args.from && !m.from?.emailAddress.address.includes(args.from as string)) return false;
                if (args.startDate && m.receivedDateTime < `${args.startDate}T00:00:00Z`) return false;
                if (args.endDate && m.receivedDateTime > `${args.endDate}T23:59:59Z`) return false;
                return true;
              })
            : messages;

        const body = filtered.map(formatEmailSummary).join("\n\n---\n\n");
        return ok(`Found ${filtered.length} email(s):\n\n${body}`);
      }

      // ── get_email_details ─────────────────────────────────────────────────
      case "get_email_details": {
        const msg: GraphMessage = await client
          .api(`/me/messages/${args.emailId}`)
          .select(
            "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,importance,body,hasAttachments,webLink"
          )
          .get();

        return ok(formatEmailDetails(msg));
      }

      // ── mark_email_as_read ────────────────────────────────────────────────
      case "mark_email_as_read": {
        const isRead = args.isRead !== undefined ? (args.isRead as boolean) : true;
        await client.api(`/me/messages/${args.emailId}`).patch({ isRead });
        return ok(`✅ Email marked as ${isRead ? "read" : "unread"}.`);
      }

      // ── get_unread_emails ─────────────────────────────────────────────────
      case "get_unread_emails": {
        const folder = (args.folder as string) || "inbox";
        const limit = Math.min((args.limit as number) || 20, 50);

        const result = await client
          .api(`/me/mailFolders/${folder}/messages`)
          .filter("isRead eq false")
          .select("id,subject,from,toRecipients,receivedDateTime,isRead,importance,bodyPreview")
          .top(limit)
          .orderby("receivedDateTime desc")
          .get();

        const messages: GraphMessage[] = result.value || [];
        if (!messages.length) return ok("✅ No unread emails — inbox is clear!");

        const body = messages.map(formatEmailSummary).join("\n\n---\n\n");
        return ok(`📬 ${messages.length} unread email(s):\n\n${body}`);
      }

      default:
        return err(`Unknown email tool: ${name}`);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`[Email:${name}]`, error);
    return err(message);
  }
}

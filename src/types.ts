/**
 * Shared TypeScript types for the Outlook MCP Server.
 * These mirror the relevant Microsoft Graph API response shapes.
 */

// ── Calendar ──────────────────────────────────────────────────────────────────

export interface GraphEvent {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location?: { displayName: string };
  attendees?: Array<{
    emailAddress: { name: string; address: string };
    status: { response: string };
    type: string;
  }>;
  body?: { content: string; contentType: "html" | "text" };
  bodyPreview?: string;
  organizer?: { emailAddress: { name: string; address: string } };
  isAllDay?: boolean;
  isCancelled?: boolean;
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
  webLink?: string;
  recurrence?: Record<string, unknown>;
}

// ── Email ─────────────────────────────────────────────────────────────────────

export interface GraphMessage {
  id: string;
  subject: string;
  from: { emailAddress: { name: string; address: string } };
  toRecipients: Array<{ emailAddress: { name: string; address: string } }>;
  ccRecipients?: Array<{ emailAddress: { name: string; address: string } }>;
  receivedDateTime: string;
  sentDateTime: string;
  isRead: boolean;
  importance: "low" | "normal" | "high";
  body?: { content: string; contentType: "html" | "text" };
  bodyPreview?: string;
  hasAttachments?: boolean;
  webLink?: string;
}

// ── Tasks ─────────────────────────────────────────────────────────────────────

export type TaskStatus =
  | "notStarted"
  | "inProgress"
  | "completed"
  | "waitingOnOthers"
  | "deferred";

export type TaskImportance = "low" | "normal" | "high";

export interface GraphTask {
  id: string;
  title: string;
  status: TaskStatus;
  importance: TaskImportance;
  isReminderOn?: boolean;
  dueDateTime?: { dateTime: string; timeZone: string };
  reminderDateTime?: { dateTime: string; timeZone: string };
  completedDateTime?: { dateTime: string; timeZone: string };
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  body?: { content: string; contentType: "html" | "text" };
}

export interface GraphTaskList {
  id: string;
  displayName: string;
  isOwner: boolean;
  isShared: boolean;
  wellknownListName?: string;
}

// ── MCP Responses ─────────────────────────────────────────────────────────────

export interface McpTextContent {
  type: "text";
  text: string;
}

export interface McpToolResponse {
  content: McpTextContent[];
  isError?: boolean;
}

/** Convenience: build a successful MCP tool response */
export function ok(text: string): McpToolResponse {
  return { content: [{ type: "text", text }] };
}

/** Convenience: build an error MCP tool response */
export function err(message: string): McpToolResponse {
  return {
    content: [{ type: "text", text: `❌ Error: ${message}` }],
    isError: true,
  };
}

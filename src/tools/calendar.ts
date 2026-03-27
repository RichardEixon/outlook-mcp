/**
 * Calendar tools for the Outlook MCP Server.
 * Uses Microsoft Graph API /me/events and /me/calendarView endpoints.
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getGraphClient } from "../graph";
import { GraphEvent, McpToolResponse, ok, err } from "../types";

// ── Tool Definitions ──────────────────────────────────────────────────────────

export const calendarTools: Tool[] = [
  {
    name: "get_calendar_events",
    description:
      "Get upcoming Outlook calendar events. Returns events sorted by start time. Defaults to the next 7 days.",
    inputSchema: {
      type: "object",
      properties: {
        startDate: {
          type: "string",
          description:
            "Start of the date range in ISO 8601 format (e.g. 2024-03-25T00:00:00). Defaults to now.",
        },
        endDate: {
          type: "string",
          description:
            "End of the date range in ISO 8601 format. Defaults to 7 days from start.",
        },
        limit: {
          type: "number",
          description: "Max events to return (default: 20, max: 50).",
        },
      },
    },
  },
  {
    name: "get_event_details",
    description:
      "Get full details for a single calendar event — attendees, meeting links, description.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The event ID (from get_calendar_events or search_events_by_date).",
        },
      },
      required: ["eventId"],
    },
  },
  {
    name: "create_calendar_event",
    description: "Create a new event in the Outlook calendar.",
    inputSchema: {
      type: "object",
      properties: {
        subject: { type: "string", description: "Event title." },
        startDateTime: {
          type: "string",
          description: "Start time in ISO 8601 format (e.g. 2024-03-25T10:00:00).",
        },
        endDateTime: {
          type: "string",
          description: "End time in ISO 8601 format.",
        },
        timeZone: {
          type: "string",
          description: "IANA timezone name (e.g. Asia/Singapore). Defaults to Asia/Singapore.",
        },
        location: { type: "string", description: "Physical or virtual location." },
        description: { type: "string", description: "Event body / notes." },
        attendees: {
          type: "array",
          items: { type: "string" },
          description: "List of attendee email addresses.",
        },
        isOnlineMeeting: {
          type: "boolean",
          description: "Set to true to generate a Microsoft Teams meeting link.",
        },
      },
      required: ["subject", "startDateTime", "endDateTime"],
    },
  },
  {
    name: "update_calendar_event",
    description: "Update the subject, time, location, or description of an existing event.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "The event ID to update." },
        subject: { type: "string", description: "New event title." },
        startDateTime: { type: "string", description: "New start time (ISO 8601)." },
        endDateTime: { type: "string", description: "New end time (ISO 8601)." },
        timeZone: { type: "string", description: "Timezone for new times." },
        location: { type: "string", description: "New location." },
        description: { type: "string", description: "New body / notes." },
      },
      required: ["eventId"],
    },
  },
  {
    name: "delete_calendar_event",
    description: "Permanently delete a calendar event.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "The event ID to delete." },
      },
      required: ["eventId"],
    },
  },
  {
    name: "search_events_by_date",
    description:
      "Search calendar events within a date range, with an optional keyword filter on title or location.",
    inputSchema: {
      type: "object",
      properties: {
        startDate: {
          type: "string",
          description: "Range start (YYYY-MM-DD or ISO 8601).",
        },
        endDate: {
          type: "string",
          description: "Range end (YYYY-MM-DD or ISO 8601).",
        },
        keyword: {
          type: "string",
          description: "Optional: filter events whose title, location, or preview contains this text.",
        },
      },
      required: ["startDate", "endDate"],
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

function formatEventSummary(event: GraphEvent): string {
  const start = formatDateTime(event.start.dateTime);
  const end = formatDateTime(event.end.dateTime);
  const lines: string[] = [
    `📅 **${event.subject}**`,
    `   ID:         ${event.id}`,
    `   Start:      ${start}`,
    `   End:        ${end}`,
  ];
  if (event.location?.displayName) lines.push(`   Location:   ${event.location.displayName}`);
  if (event.isOnlineMeeting) lines.push(`   Teams:      ${event.onlineMeetingUrl || "Yes"}`);
  if (event.organizer) lines.push(`   Organiser:  ${event.organizer.emailAddress.name}`);
  if (event.attendees?.length) {
    lines.push(`   Attendees:  ${event.attendees.map((a) => a.emailAddress.address).join(", ")}`);
  }
  if (event.bodyPreview) lines.push(`   Preview:    ${event.bodyPreview.slice(0, 120)}`);
  return lines.join("\n");
}

function formatEventDetails(event: GraphEvent): string {
  const start = formatDateTime(event.start.dateTime);
  const end = formatDateTime(event.end.dateTime);
  const attendees =
    event.attendees
      ?.map(
        (a) =>
          `  • ${a.emailAddress.name} <${a.emailAddress.address}> — ${a.status?.response || "none"}`
      )
      .join("\n") || "  None";

  const bodyText = event.body
    ? event.body.contentType === "html"
      ? stripHtml(event.body.content).slice(0, 800)
      : event.body.content.slice(0, 800)
    : "(no description)";

  const lines: string[] = [
    `📅 **${event.subject}**`,
    ``,
    `ID:         ${event.id}`,
    `Start:      ${start}`,
    `End:        ${end}`,
  ];
  if (event.isAllDay) lines.push(`All Day:    Yes`);
  if (event.isCancelled) lines.push(`⚠️ Status:  CANCELLED`);
  if (event.location?.displayName) lines.push(`Location:   ${event.location.displayName}`);
  if (event.isOnlineMeeting) lines.push(`Teams Link: ${event.onlineMeetingUrl || "Yes"}`);
  if (event.organizer)
    lines.push(`Organiser:  ${event.organizer.emailAddress.name} <${event.organizer.emailAddress.address}>`);
  lines.push(``, `Attendees:`, attendees);
  lines.push(``, `Description:`, bodyText);
  if (event.webLink) lines.push(``, `Open in Outlook: ${event.webLink}`);
  return lines.join("\n");
}

// ── Tool Handler ──────────────────────────────────────────────────────────────

export async function handleCalendarTool(
  name: string,
  args: Record<string, unknown>
): Promise<McpToolResponse> {
  try {
    const client = await getGraphClient();

    switch (name) {
      // ── get_calendar_events ───────────────────────────────────────────────
      case "get_calendar_events": {
        const start = (args.startDate as string) || new Date().toISOString();
        const end =
          (args.endDate as string) ||
          new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
        const limit = Math.min((args.limit as number) || 20, 50);

        const result = await client
          .api("/me/calendarView")
          .query({
            startDateTime: start,
            endDateTime: end,
            $top: limit,
            $orderby: "start/dateTime asc",
            $select:
              "id,subject,start,end,location,attendees,organizer,isOnlineMeeting,onlineMeetingUrl,bodyPreview,isAllDay,isCancelled",
          })
          .get();

        const events: GraphEvent[] = result.value || [];
        if (!events.length) return ok("📭 No events found in the specified date range.");

        const body = events.map(formatEventSummary).join("\n\n---\n\n");
        return ok(`Found ${events.length} event(s):\n\n${body}`);
      }

      // ── get_event_details ─────────────────────────────────────────────────
      case "get_event_details": {
        const event: GraphEvent = await client
          .api(`/me/events/${args.eventId}`)
          .select(
            "id,subject,start,end,location,attendees,organizer,body,isOnlineMeeting,onlineMeetingUrl,webLink,isCancelled,isAllDay,recurrence"
          )
          .get();

        return ok(formatEventDetails(event));
      }

      // ── create_calendar_event ─────────────────────────────────────────────
      case "create_calendar_event": {
        const tz = (args.timeZone as string) || SG_TZ;
        const payload: Record<string, unknown> = {
          subject: args.subject,
          start: { dateTime: args.startDateTime, timeZone: tz },
          end: { dateTime: args.endDateTime, timeZone: tz },
        };

        if (args.location) payload.location = { displayName: args.location };
        if (args.description) payload.body = { content: args.description, contentType: "text" };
        if (Array.isArray(args.attendees) && args.attendees.length) {
          payload.attendees = (args.attendees as string[]).map((email) => ({
            emailAddress: { address: email },
            type: "required",
          }));
        }
        if (args.isOnlineMeeting) payload.isOnlineMeeting = true;

        const created: GraphEvent = await client.api("/me/events").post(payload);
        return ok(
          `✅ Event created!\n\nID:    ${created.id}\nTitle: ${created.subject}\nStart: ${formatDateTime(created.start.dateTime)}`
        );
      }

      // ── update_calendar_event ─────────────────────────────────────────────
      case "update_calendar_event": {
        const tz = (args.timeZone as string) || SG_TZ;
        const updates: Record<string, unknown> = {};

        if (args.subject) updates.subject = args.subject;
        if (args.startDateTime) updates.start = { dateTime: args.startDateTime, timeZone: tz };
        if (args.endDateTime) updates.end = { dateTime: args.endDateTime, timeZone: tz };
        if (args.location) updates.location = { displayName: args.location };
        if (args.description) updates.body = { content: args.description, contentType: "text" };

        await client.api(`/me/events/${args.eventId}`).patch(updates);
        return ok(`✅ Event updated successfully.`);
      }

      // ── delete_calendar_event ─────────────────────────────────────────────
      case "delete_calendar_event": {
        await client.api(`/me/events/${args.eventId}`).delete();
        return ok(`✅ Event deleted.`);
      }

      // ── search_events_by_date ─────────────────────────────────────────────
      case "search_events_by_date": {
        const toIso = (d: string) => (d.includes("T") ? d : `${d}T00:00:00`);
        const start = toIso(args.startDate as string);
        const end = toIso(args.endDate as string).replace("T00:00:00", "T23:59:59");

        const result = await client
          .api("/me/calendarView")
          .query({
            startDateTime: start,
            endDateTime: end,
            $top: 50,
            $orderby: "start/dateTime asc",
            $select:
              "id,subject,start,end,location,attendees,organizer,bodyPreview,isOnlineMeeting,isCancelled",
          })
          .get();

        let events: GraphEvent[] = result.value || [];

        if (args.keyword) {
          const kw = (args.keyword as string).toLowerCase();
          events = events.filter(
            (e) =>
              e.subject?.toLowerCase().includes(kw) ||
              e.bodyPreview?.toLowerCase().includes(kw) ||
              e.location?.displayName?.toLowerCase().includes(kw)
          );
        }

        if (!events.length) return ok("📭 No events found for the specified criteria.");
        const body = events.map(formatEventSummary).join("\n\n---\n\n");
        return ok(`Found ${events.length} event(s):\n\n${body}`);
      }

      default:
        return err(`Unknown calendar tool: ${name}`);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`[Calendar:${name}]`, error);
    return err(message);
  }
}

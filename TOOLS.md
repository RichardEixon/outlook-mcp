# Tool Reference

Complete reference for all 15 tools in the Outlook MCP Server.

---

## Calendar Tools (6)

### `get_calendar_events`
Get upcoming calendar events sorted by start time.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `startDate` | string | No | ISO 8601 start of range. Defaults to now. |
| `endDate` | string | No | ISO 8601 end of range. Defaults to 7 days from start. |
| `limit` | number | No | Max events to return (default: 20, max: 50). |

**Example prompts**
- _"Show me my calendar for this week"_
- _"What meetings do I have today?"_
- _"List events from March 25 to March 31"_

**Returns**: List of events with ID, title, start/end time, location, organiser, attendees, and preview.

---

### `get_event_details`
Get full details for a single event including the complete description, all attendees and their RSVP status, and Teams link.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | **Yes** | Event ID from `get_calendar_events`. |

**Example prompts**
- _"Tell me more about that 2 PM meeting"_ (after listing events)
- _"Show full details for event ID ABC123"_

**Returns**: Full event details — description, attendee list with RSVP status, online meeting URL, Outlook web link.

---

### `create_calendar_event`
Create a new event in your Outlook calendar.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `subject` | string | **Yes** | Event title. |
| `startDateTime` | string | **Yes** | Start time (ISO 8601, e.g. `2024-03-25T14:00:00`). |
| `endDateTime` | string | **Yes** | End time (ISO 8601). |
| `timeZone` | string | No | IANA timezone (default: `Asia/Singapore`). |
| `location` | string | No | Physical or virtual location. |
| `description` | string | No | Event body / notes. |
| `attendees` | string[] | No | List of email addresses to invite. |
| `isOnlineMeeting` | boolean | No | `true` to generate a Teams meeting link. |

**Example prompts**
- _"Create a meeting tomorrow at 3 PM called Project Sync"_
- _"Schedule a 1-hour Teams call with bob@company.com on Friday at 10 AM"_

**Returns**: Confirmation with the new event's ID, title, and start time.

---

### `update_calendar_event`
Update an existing event's title, time, location, or description.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | **Yes** | Event ID to update. |
| `subject` | string | No | New title. |
| `startDateTime` | string | No | New start time. |
| `endDateTime` | string | No | New end time. |
| `timeZone` | string | No | Timezone for new times. |
| `location` | string | No | New location. |
| `description` | string | No | New body/notes. |

**Example prompts**
- _"Move my 3 PM meeting to 4 PM"_
- _"Change the location of tomorrow's standup to Conference Room B"_

**Returns**: Success confirmation.

---

### `delete_calendar_event`
Permanently delete a calendar event.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | **Yes** | Event ID to delete. |

**Example prompts**
- _"Cancel/delete that Friday meeting"_

> **Warning**: This cannot be undone.

---

### `search_events_by_date`
Search events within a date range with an optional keyword filter.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `startDate` | string | **Yes** | Range start (`YYYY-MM-DD` or ISO 8601). |
| `endDate` | string | **Yes** | Range end (`YYYY-MM-DD` or ISO 8601). |
| `keyword` | string | No | Filter by title, location, or body preview. |

**Example prompts**
- _"Find all meetings with 'budget' in the title next month"_
- _"What events do I have between April 1 and April 15?"_

**Returns**: Filtered list of matching events.

---

## Email Tools (4)

### `search_emails`
Search emails by keyword, sender, or date. Results sorted newest-first.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `keyword` | string | No | Search in subject and body. |
| `from` | string | No | Sender's email address. |
| `startDate` | string | No | Return emails received on/after this date (`YYYY-MM-DD`). |
| `endDate` | string | No | Return emails received on/before this date (`YYYY-MM-DD`). |
| `folder` | string | No | `inbox` \| `sentitems` \| `drafts` \| `deleteditems` (default: `inbox`). |
| `limit` | number | No | Max results (default: 20, max: 50). |

**Example prompts**
- _"Find emails from alice@company.com this week"_
- _"Search for emails about 'invoice' in March"_
- _"Show me emails sent to my drafts folder"_

**Returns**: Email list with ID, sender, subject, received time, read/unread status, and body preview.

---

### `get_email_details`
Get the full content and all metadata of a specific email.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `emailId` | string | **Yes** | Email message ID (from `search_emails` or `get_unread_emails`). |

**Example prompts**
- _"Show me the full content of that email"_
- _"Read email ID XYZ"_

**Returns**: Full email with body (HTML stripped), all recipients, CC, attachments flag, Outlook web link.

---

### `mark_email_as_read`
Mark a specific email as read or unread.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `emailId` | string | **Yes** | Email message ID. |
| `isRead` | boolean | No | `true` = read, `false` = unread (default: `true`). |

**Example prompts**
- _"Mark that email as read"_
- _"Mark email XYZ as unread"_

**Returns**: Success confirmation.

---

### `get_unread_emails`
Get a list of unread emails, sorted newest-first.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `folder` | string | No | Folder to check (default: `inbox`). |
| `limit` | number | No | Max results (default: 20, max: 50). |

**Example prompts**
- _"Show me my unread emails"_
- _"How many unread messages do I have?"_

**Returns**: List of unread emails with sender, subject, received time, and preview.

---

## Task Tools (5)

### `get_tasks`
Get tasks from Microsoft To Do / Outlook Tasks.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `listName` | string | No | Task list name (default: your default To Do list). |
| `status` | string | No | `notStarted` \| `inProgress` \| `completed` \| `waitingOnOthers` \| `deferred` \| `all` (default: `notStarted`). |
| `importance` | string | No | `low` \| `normal` \| `high`. |
| `limit` | number | No | Max tasks (default: 20). |

**Example prompts**
- _"Show me my open tasks"_
- _"What high-priority tasks do I have?"_
- _"List all completed tasks"_

**Returns**: Task list with ID, list ID, title, status, priority, due date, and notes. The **Task ID** and **List ID** are both needed for update/complete/delete operations.

---

### `create_task`
Create a new task.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `title` | string | **Yes** | Task title. |
| `dueDate` | string | No | Due date in `YYYY-MM-DD` format. |
| `importance` | string | No | `low` \| `normal` \| `high` (default: `normal`). |
| `description` | string | No | Task notes. |
| `listName` | string | No | Which list to add it to (default list if omitted). |

**Example prompts**
- _"Create a task: Submit expense report by Friday"_
- _"Add a high-priority task to review the contract by 2024-04-01"_

**Returns**: Confirmation with the created task's ID and list ID.

---

### `update_task`
Update a task's title, due date, priority, status, or notes.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `taskId` | string | **Yes** | Task ID (from `get_tasks`). |
| `listId` | string | **Yes** | List ID (from `get_tasks`). |
| `title` | string | No | New title. |
| `dueDate` | string | No | New due date (`YYYY-MM-DD`). |
| `importance` | string | No | New priority. |
| `status` | string | No | New status. |
| `description` | string | No | New notes. |

**Example prompts**
- _"Change the due date of task XYZ to next Monday"_
- _"Set the priority of that task to high"_

**Returns**: Success confirmation.

---

### `complete_task`
Mark a task as completed.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `taskId` | string | **Yes** | Task ID. |
| `listId` | string | **Yes** | List ID. |

**Example prompts**
- _"Mark the expense report task as done"_
- _"Complete task XYZ"_

**Returns**: Success confirmation.

---

### `delete_task`
Permanently delete a task.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `taskId` | string | **Yes** | Task ID. |
| `listId` | string | **Yes** | List ID. |

**Example prompts**
- _"Delete the old grocery task"_

> **Warning**: This cannot be undone.

---

## Tips

- **IDs are long strings** — you don't need to type them. Claude will use the ID from the previous tool result automatically when you say "that event" or "that email".
- **Timezone**: All times are displayed in Singapore time (Asia/Singapore) by default.
- **Task List ID**: Always shown in `get_tasks` results — you need both `taskId` and `listId` to modify a task.
- **Email search**: If you provide both `keyword` and other filters, keyword search takes precedence and other filters are applied client-side.

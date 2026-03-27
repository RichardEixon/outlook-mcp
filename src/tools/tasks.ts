/**
 * Tasks tools for the Outlook MCP Server.
 * Uses Microsoft Graph API /me/todo/lists and /me/todo/lists/{id}/tasks endpoints.
 * Works with both Microsoft To Do and Outlook Tasks (they share the same backend).
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { Client } from "@microsoft/microsoft-graph-client";
import { getGraphClient } from "../graph";
import {
  GraphTask,
  GraphTaskList,
  TaskStatus,
  TaskImportance,
  McpToolResponse,
  ok,
  err,
} from "../types";

// ── Tool Definitions ──────────────────────────────────────────────────────────

export const taskTools: Tool[] = [
  {
    name: "get_tasks",
    description:
      "Get tasks from Microsoft To Do / Outlook Tasks. Defaults to showing active (not-started) tasks from the default list.",
    inputSchema: {
      type: "object",
      properties: {
        listName: {
          type: "string",
          description:
            "Task list name (default: uses the default To Do list). Case-insensitive.",
        },
        status: {
          type: "string",
          enum: ["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred", "all"],
          description: "Filter by status. Defaults to 'notStarted'.",
        },
        importance: {
          type: "string",
          enum: ["low", "normal", "high"],
          description: "Filter by priority.",
        },
        limit: {
          type: "number",
          description: "Max tasks to return (default: 20).",
        },
      },
    },
  },
  {
    name: "create_task",
    description: "Create a new task in Microsoft To Do / Outlook Tasks.",
    inputSchema: {
      type: "object",
      properties: {
        title: { type: "string", description: "Task title." },
        dueDate: {
          type: "string",
          description: "Due date in YYYY-MM-DD format.",
        },
        importance: {
          type: "string",
          enum: ["low", "normal", "high"],
          description: "Priority (default: normal).",
        },
        description: { type: "string", description: "Task notes." },
        listName: {
          type: "string",
          description: "Task list to add to (default: default list).",
        },
      },
      required: ["title"],
    },
  },
  {
    name: "update_task",
    description: "Update a task's title, due date, priority, status, or notes.",
    inputSchema: {
      type: "object",
      properties: {
        taskId: { type: "string", description: "Task ID (from get_tasks)." },
        listId: { type: "string", description: "Task list ID (from get_tasks)." },
        title: { type: "string", description: "New title." },
        dueDate: { type: "string", description: "New due date (YYYY-MM-DD)." },
        importance: {
          type: "string",
          enum: ["low", "normal", "high"],
        },
        status: {
          type: "string",
          enum: ["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"],
        },
        description: { type: "string", description: "New notes." },
      },
      required: ["taskId", "listId"],
    },
  },
  {
    name: "complete_task",
    description: "Mark a task as completed.",
    inputSchema: {
      type: "object",
      properties: {
        taskId: { type: "string", description: "Task ID." },
        listId: { type: "string", description: "Task list ID." },
      },
      required: ["taskId", "listId"],
    },
  },
  {
    name: "delete_task",
    description: "Permanently delete a task.",
    inputSchema: {
      type: "object",
      properties: {
        taskId: { type: "string", description: "Task ID." },
        listId: { type: "string", description: "Task list ID." },
      },
      required: ["taskId", "listId"],
    },
  },
];

// ── Formatting Helpers ────────────────────────────────────────────────────────

const SG_LOCALE = "en-SG";
const SG_TZ = "Asia/Singapore";

const STATUS_EMOJI: Record<string, string> = {
  notStarted: "⬜",
  inProgress: "🔄",
  completed: "✅",
  waitingOnOthers: "⏳",
  deferred: "📌",
};

const IMPORTANCE_LABEL: Record<string, string> = {
  low: "🔽 Low",
  normal: "➡️ Normal",
  high: "🔼 High",
};

function formatTask(task: GraphTask, listId: string): string {
  const statusEmoji = STATUS_EMOJI[task.status] || "⬜";
  const importance = IMPORTANCE_LABEL[task.importance] || task.importance;

  const due = task.dueDateTime
    ? new Date(task.dueDateTime.dateTime).toLocaleDateString(SG_LOCALE, {
        timeZone: SG_TZ,
      })
    : null;

  const lines: string[] = [
    `${statusEmoji} **${task.title}**`,
    `   Task ID:  ${task.id}`,
    `   List ID:  ${listId}`,
    `   Status:   ${task.status}`,
    `   Priority: ${importance}`,
  ];
  if (due) lines.push(`   Due:      ${due}`);
  if (task.body?.content) lines.push(`   Notes:    ${task.body.content.slice(0, 120)}`);
  return lines.join("\n");
}

// ── Internal Helpers ──────────────────────────────────────────────────────────

/**
 * Resolves a human-readable list name to a Graph task list ID.
 * Falls back to the default list if no name is provided.
 */
async function resolveListId(client: Client, listName?: string): Promise<string> {
  const result = await client.api("/me/todo/lists").get();
  const lists: GraphTaskList[] = result.value || [];

  if (!lists.length) {
    throw new Error("No task lists found in your account.");
  }

  if (listName) {
    const match = lists.find(
      (l) => l.displayName.toLowerCase() === listName.toLowerCase()
    );
    if (!match) {
      const available = lists.map((l) => `"${l.displayName}"`).join(", ");
      throw new Error(
        `Task list "${listName}" not found. Available lists: ${available}`
      );
    }
    return match.id;
  }

  // Default: prefer the "Tasks" well-known list, then the first list
  const defaultList =
    lists.find((l) => l.wellknownListName === "defaultList") || lists[0];
  return defaultList.id;
}

// ── Tool Handler ──────────────────────────────────────────────────────────────

export async function handleTaskTool(
  name: string,
  args: Record<string, unknown>
): Promise<McpToolResponse> {
  try {
    const client = await getGraphClient();

    switch (name) {
      // ── get_tasks ─────────────────────────────────────────────────────────
      case "get_tasks": {
        const listId = await resolveListId(client, args.listName as string);
        const limit = (args.limit as number) || 20;
        const statusFilter = (args.status as string) || "notStarted";

        let apiCall = client
          .api(`/me/todo/lists/${listId}/tasks`)
          .top(limit)
          .select("id,title,status,importance,dueDateTime,body,completedDateTime,createdDateTime");

        if (statusFilter !== "all") {
          apiCall = apiCall.filter(`status eq '${statusFilter}'`);
        }
        // Note: Graph API doesn't support $filter on importance directly in all versions;
        // we filter client-side if needed
        const result = await apiCall.get();
        let tasks: GraphTask[] = result.value || [];

        if (args.importance) {
          tasks = tasks.filter((t) => t.importance === args.importance);
        }

        if (!tasks.length) return ok("📋 No tasks found with the specified filters.");

        const body = tasks.map((t) => formatTask(t, listId)).join("\n\n");
        return ok(`Found ${tasks.length} task(s) (List ID: ${listId}):\n\n${body}`);
      }

      // ── create_task ───────────────────────────────────────────────────────
      case "create_task": {
        const listId = await resolveListId(client, args.listName as string);

        const payload: Partial<GraphTask> = {
          title: args.title as string,
          importance: (args.importance as TaskImportance) || "normal",
          status: "notStarted",
        };

        if (args.dueDate) {
          payload.dueDateTime = {
            dateTime: `${args.dueDate}T23:59:59`,
            timeZone: SG_TZ,
          };
        }
        if (args.description) {
          payload.body = { content: args.description as string, contentType: "text" };
        }

        const created: GraphTask = await client
          .api(`/me/todo/lists/${listId}/tasks`)
          .post(payload);

        return ok(`✅ Task created!\n\n${formatTask(created, listId)}`);
      }

      // ── update_task ───────────────────────────────────────────────────────
      case "update_task": {
        const updates: Record<string, unknown> = {};

        if (args.title) updates.title = args.title;
        if (args.importance) updates.importance = args.importance;
        if (args.status) updates.status = args.status;
        if (args.dueDate) {
          updates.dueDateTime = {
            dateTime: `${args.dueDate}T23:59:59`,
            timeZone: SG_TZ,
          };
        }
        if (args.description) {
          updates.body = { content: args.description, contentType: "text" };
        }

        await client
          .api(`/me/todo/lists/${args.listId}/tasks/${args.taskId}`)
          .patch(updates);

        return ok(`✅ Task updated successfully.`);
      }

      // ── complete_task ─────────────────────────────────────────────────────
      case "complete_task": {
        await client
          .api(`/me/todo/lists/${args.listId}/tasks/${args.taskId}`)
          .patch({
            status: "completed" as TaskStatus,
            completedDateTime: {
              dateTime: new Date().toISOString().replace("Z", ""),
              timeZone: "UTC",
            },
          });

        return ok(`✅ Task marked as completed!`);
      }

      // ── delete_task ───────────────────────────────────────────────────────
      case "delete_task": {
        await client
          .api(`/me/todo/lists/${args.listId}/tasks/${args.taskId}`)
          .delete();

        return ok(`✅ Task deleted.`);
      }

      default:
        return err(`Unknown task tool: ${name}`);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`[Tasks:${name}]`, error);
    return err(message);
  }
}

#!/usr/bin/env npx tsx
/**
 * Google Workspace Manager CLI
 *
 * Zod-validated CLI for Google Workspace operations via MCP.
 */

import { z, createCommand, runCli, cacheCommands, cliTypes } from "@local/cli-utils";
import { GoogleWorkspaceMCPClient } from "./mcp-client.js";

// Define commands with Zod schemas
const commands = {
  "list-tools": createCommand(
    z.object({}),
    async (_args, client: GoogleWorkspaceMCPClient) => {
      const tools = await client.listTools();
      return tools.map((t: { name: string; description?: string }) => ({
        name: t.name,
        description: t.description,
      }));
    },
    "List all available MCP tools"
  ),

  // ==================== Gmail ====================
  "search-gmail": createCommand(
    z.object({
      query: z.string().min(1).describe("Search query"),
      limit: cliTypes.int(1, 500).optional().describe("Max results"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { query, limit } = args as { query: string; limit?: number };
      return client.searchGmailMessages(query, limit);
    },
    "Search Gmail messages"
  ),

  "get-gmail-message": createCommand(
    z.object({
      id: z.string().min(1).describe("Message ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.getGmailMessage(id);
    },
    "Get a specific email"
  ),

  "get-gmail-thread": createCommand(
    z.object({
      id: z.string().min(1).describe("Thread ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.getGmailThread(id);
    },
    "Get a full email thread"
  ),

  "list-gmail-labels": createCommand(
    z.object({}),
    async (_args, client: GoogleWorkspaceMCPClient) => client.listGmailLabels(),
    "List Gmail labels"
  ),

  "send-gmail": createCommand(
    z.object({
      to: z.string().min(1).describe("Recipient email"),
      subject: z.string().min(1).describe("Email subject"),
      body: z.string().min(1).describe("Email body"),
      cc: z.string().optional().describe("CC recipients"),
      bcc: z.string().optional().describe("BCC recipients"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { to, subject, body, cc, bcc } = args as {
        to: string; subject: string; body: string; cc?: string; bcc?: string;
      };
      return client.sendGmailMessage(to, subject, body, cc, bcc);
    },
    "Send an email"
  ),

  "create-gmail-draft": createCommand(
    z.object({
      to: z.string().min(1).describe("Recipient email"),
      subject: z.string().min(1).describe("Email subject"),
      body: z.string().min(1).describe("Email body"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { to, subject, body } = args as { to: string; subject: string; body: string };
      return client.createGmailDraft(to, subject, body);
    },
    "Create a draft"
  ),

  // ==================== Calendar ====================
  "list-calendars": createCommand(
    z.object({}),
    async (_args, client: GoogleWorkspaceMCPClient) => client.listCalendars(),
    "List available calendars"
  ),

  "get-events": createCommand(
    z.object({
      calendarId: z.string().optional().describe("Calendar ID"),
      timeMin: z.string().optional().describe("Events after this time (ISO 8601)"),
      timeMax: z.string().optional().describe("Events before this time (ISO 8601)"),
      eventId: z.string().optional().describe("Filter to specific event ID"),
      limit: cliTypes.int(1, 2500).optional().describe("Max results"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { calendarId, timeMin, timeMax, eventId, limit } = args as {
        calendarId?: string; timeMin?: string; timeMax?: string;
        eventId?: string; limit?: number;
      };
      let result = await client.getEvents({
        calendarId,
        timeMin,
        timeMax,
        maxResults: limit,
      });
      if (eventId && result.events) {
        result.events = result.events.filter((e: { id: string }) => e.id === eventId);
      }
      return result;
    },
    "Get calendar events"
  ),

  "create-event": createCommand(
    z.object({
      summary: z.string().min(1).describe("Event title"),
      start: z.string().min(1).describe("Event start (ISO 8601, include offset e.g. +00:00)"),
      end: z.string().min(1).describe("Event end (ISO 8601, include offset e.g. +00:00)"),
      description: z.string().optional().describe("Event description"),
      location: z.string().optional().describe("Event location"),
      attendees: z.string().optional().describe("Comma-separated attendee emails"),
      timezone: z.string().optional().describe("IANA timezone (e.g. Europe/London). Only needed if start/end lack offset"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { summary, start, end, description, location, attendees, timezone } = args as {
        summary: string; start: string; end: string;
        description?: string; location?: string; attendees?: string; timezone?: string;
      };
      // Validate: datetimes must include offset OR timezone must be provided
      const hasOffset = (dt: string) => /[+-]\d{2}:\d{2}$|Z$/.test(dt);
      if (!hasOffset(start) && !hasOffset(end) && !timezone) {
        throw new Error("Start/end times must include timezone offset (e.g. +00:00) or use --timezone flag");
      }
      return client.createEvent(summary, start, end, { description, location, attendees, timezone });
    },
    "Create a new event"
  ),

  "delete-event": createCommand(
    z.object({
      id: z.string().min(1).describe("Event ID"),
      calendarId: z.string().optional().describe("Calendar ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, calendarId } = args as { id: string; calendarId?: string };
      return client.deleteEvent(id, calendarId);
    },
    "Delete an event"
  ),

  // ==================== Drive ====================
  "search-drive": createCommand(
    z.object({
      query: z.string().min(1).describe("Search query"),
      limit: cliTypes.int(1, 1000).optional().describe("Max results"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { query, limit } = args as { query: string; limit?: number };
      return client.searchDriveFiles(query, limit);
    },
    "Search Drive files"
  ),

  "get-drive-content": createCommand(
    z.object({
      id: z.string().min(1).describe("File ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.getDriveFileContent(id);
    },
    "Get file content"
  ),

  "list-drive-items": createCommand(
    z.object({
      folderId: z.string().optional().describe("Folder ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { folderId } = args as { folderId?: string };
      return client.listDriveItems(folderId);
    },
    "List files in folder"
  ),

  // ==================== Docs ====================
  "search-docs": createCommand(
    z.object({
      query: z.string().min(1).describe("Search query"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { query } = args as { query: string };
      return client.searchDocs(query);
    },
    "Search Google Docs"
  ),

  "get-doc-content": createCommand(
    z.object({
      id: z.string().min(1).describe("Document ID"),
      suggestionsMode: z.string().optional().describe("Suggestions view mode"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, suggestionsMode } = args as { id: string; suggestionsMode?: string };
      return client.getDocContent(id, suggestionsMode);
    },
    "Get document content"
  ),

  "create-doc": createCommand(
    z.object({
      title: z.string().min(1).describe("Document title"),
      content: z.string().optional().describe("Initial content"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { title, content } = args as { title: string; content?: string };
      return client.createDoc(title, content);
    },
    "Create new document"
  ),

  "modify-doc-text": createCommand(
    z.object({
      id: z.string().min(1).describe("Document ID"),
      operation: z.enum(["insert", "replace", "delete"]).describe("Operation type"),
      index: cliTypes.int(1).optional().describe("Insert index"),
      text: z.string().optional().describe("Text to insert/replace"),
      startIndex: cliTypes.int(1).optional().describe("Start index for replace/delete"),
      endIndex: cliTypes.int(1).optional().describe("End index for replace/delete"),
      bold: cliTypes.bool().optional().describe("Bold formatting"),
      italic: cliTypes.bool().optional().describe("Italic formatting"),
      underline: cliTypes.bool().optional().describe("Underline formatting"),
      fontSize: cliTypes.int(1).optional().describe("Font size in points"),
      fontFamily: z.string().optional().describe("Font family name"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, operation, index, text, startIndex, endIndex, bold, italic, underline, fontSize, fontFamily } = args as {
        id: string; operation: "insert" | "replace" | "delete";
        index?: number; text?: string; startIndex?: number; endIndex?: number;
        bold?: boolean; italic?: boolean; underline?: boolean;
        fontSize?: number; fontFamily?: string;
      };
      return client.modifyDocText(id, operation, {
        index, text, startIndex, endIndex, bold, italic, underline, fontSize, fontFamily,
      });
    },
    "Insert/replace/delete text"
  ),

  "find-replace-doc": createCommand(
    z.object({
      id: z.string().min(1).describe("Document ID"),
      find: z.string().min(1).describe("Text to find"),
      replace: z.string().min(1).describe("Replacement text"),
      replaceAll: cliTypes.bool().optional().describe("Replace all occurrences"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, find, replace, replaceAll } = args as {
        id: string; find: string; replace: string; replaceAll?: boolean;
      };
      return client.findAndReplaceDoc(id, find, replace, replaceAll !== false);
    },
    "Find and replace in document"
  ),

  // ==================== Sheets ====================
  "list-spreadsheets": createCommand(
    z.object({}),
    async (_args, client: GoogleWorkspaceMCPClient) => client.listSpreadsheets(),
    "List spreadsheets"
  ),

  "get-spreadsheet-info": createCommand(
    z.object({
      id: z.string().min(1).describe("Spreadsheet ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.getSpreadsheetInfo(id);
    },
    "Get spreadsheet metadata"
  ),

  "read-sheet": createCommand(
    z.object({
      id: z.string().min(1).describe("Spreadsheet ID"),
      range: z.string().min(1).describe("Sheet range (e.g., A1:D10)"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, range } = args as { id: string; range: string };
      return client.readSheetValues(id, range);
    },
    "Read sheet values"
  ),

  "write-sheet": createCommand(
    z.object({
      id: z.string().min(1).describe("Spreadsheet ID"),
      range: z.string().min(1).describe("Sheet range"),
      values: z.string().min(1).describe("JSON array of values"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, range, values } = args as { id: string; range: string; values: string };
      return client.writeSheetValues(id, range, JSON.parse(values));
    },
    "Write sheet values"
  ),

  "write-rich-text": createCommand(
    z.object({
      id: z.string().min(1).describe("Spreadsheet ID"),
      cell: z.string().min(1).describe("Cell reference (e.g., AD2)"),
      segments: z.string().min(1).describe("JSON array of {text, url?, bold?, ...}"),
      sheetName: z.string().optional().describe("Sheet name"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, cell, segments, sheetName } = args as {
        id: string; cell: string; segments: string; sheetName?: string;
      };
      return client.writeRichTextCell(id, cell, JSON.parse(segments), sheetName);
    },
    "Write rich text with formatting/links to a cell"
  ),

  "write-rich-text-batch": createCommand(
    z.object({
      id: z.string().min(1).describe("Spreadsheet ID"),
      cells: z.string().min(1).describe("JSON array of {cell, segments}"),
      sheetName: z.string().optional().describe("Sheet name"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, cells, sheetName } = args as { id: string; cells: string; sheetName?: string };
      return client.writeRichTextCells(id, JSON.parse(cells), sheetName);
    },
    "Write rich text to multiple cells"
  ),

  "create-spreadsheet": createCommand(
    z.object({
      title: z.string().min(1).describe("Spreadsheet title"),
      sheetNames: z.string().optional().describe("JSON array of sheet names"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { title, sheetNames } = args as { title: string; sheetNames?: string };
      return client.createSpreadsheet(title, sheetNames ? JSON.parse(sheetNames) : undefined);
    },
    "Create new spreadsheet"
  ),

  // ==================== Tasks ====================
  "list-task-lists": createCommand(
    z.object({}),
    async (_args, client: GoogleWorkspaceMCPClient) => client.listTaskLists(),
    "List task lists"
  ),

  "list-tasks": createCommand(
    z.object({
      id: z.string().min(1).describe("Task list ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.listTasks(id);
    },
    "List tasks"
  ),

  "create-task": createCommand(
    z.object({
      listId: z.string().min(1).describe("Task list ID"),
      title: z.string().min(1).describe("Task title"),
      notes: z.string().optional().describe("Task notes"),
      due: z.string().optional().describe("Due date (ISO 8601)"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { listId, title, notes, due } = args as {
        listId: string; title: string; notes?: string; due?: string;
      };
      return client.createTask(listId, title, notes, due);
    },
    "Create a task"
  ),

  "complete-task": createCommand(
    z.object({
      listId: z.string().min(1).describe("Task list ID"),
      id: z.string().min(1).describe("Task ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { listId, id } = args as { listId: string; id: string };
      return client.completeTask(listId, id);
    },
    "Mark task complete"
  ),

  // ==================== Document Comments ====================
  "get-doc-comments": createCommand(
    z.object({
      id: z.string().min(1).describe("Document ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.getDocumentComments(id);
    },
    "Get document comments"
  ),

  "create-doc-comment": createCommand(
    z.object({
      id: z.string().min(1).describe("Document ID"),
      text: z.string().min(1).describe("Comment text"),
      location: z.string().optional().describe("Location as JSON"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, text, location } = args as { id: string; text: string; location?: string };
      return client.createDocumentComment(id, text, location ? JSON.parse(location) : undefined);
    },
    "Create document comment"
  ),

  "reply-doc-comment": createCommand(
    z.object({
      id: z.string().min(1).describe("Document ID"),
      commentId: z.string().min(1).describe("Comment ID"),
      text: z.string().min(1).describe("Reply text"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, commentId, text } = args as { id: string; commentId: string; text: string };
      return client.replyToDocumentComment(id, commentId, text);
    },
    "Reply to document comment"
  ),

  "resolve-doc-comment": createCommand(
    z.object({
      id: z.string().min(1).describe("Document ID"),
      commentId: z.string().min(1).describe("Comment ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, commentId } = args as { id: string; commentId: string };
      return client.resolveDocumentComment(id, commentId);
    },
    "Resolve document comment"
  ),

  // ==================== Spreadsheet Comments ====================
  "get-sheet-comments": createCommand(
    z.object({
      id: z.string().min(1).describe("Spreadsheet ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.getSpreadsheetComments(id);
    },
    "Get spreadsheet comments"
  ),

  "create-sheet-comment": createCommand(
    z.object({
      id: z.string().min(1).describe("Spreadsheet ID"),
      sheetId: cliTypes.int(0).describe("Sheet ID"),
      rowIndex: cliTypes.int(0).describe("Row index"),
      columnIndex: cliTypes.int(0).describe("Column index"),
      text: z.string().min(1).describe("Comment text"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, sheetId, rowIndex, columnIndex, text } = args as {
        id: string; sheetId: number; rowIndex: number; columnIndex: number; text: string;
      };
      return client.createSpreadsheetComment(id, sheetId, rowIndex, columnIndex, text);
    },
    "Create spreadsheet comment"
  ),

  "reply-sheet-comment": createCommand(
    z.object({
      id: z.string().min(1).describe("Spreadsheet ID"),
      commentId: z.string().min(1).describe("Comment ID"),
      text: z.string().min(1).describe("Reply text"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, commentId, text } = args as { id: string; commentId: string; text: string };
      return client.replyToSpreadsheetComment(id, commentId, text);
    },
    "Reply to spreadsheet comment"
  ),

  "resolve-sheet-comment": createCommand(
    z.object({
      id: z.string().min(1).describe("Spreadsheet ID"),
      commentId: z.string().min(1).describe("Comment ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, commentId } = args as { id: string; commentId: string };
      return client.resolveSpreadsheetComment(id, commentId);
    },
    "Resolve spreadsheet comment"
  ),

  // ==================== Presentation Comments ====================
  "get-presentation-comments": createCommand(
    z.object({
      id: z.string().min(1).describe("Presentation ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.getPresentationComments(id);
    },
    "Get presentation comments"
  ),

  "create-presentation-comment": createCommand(
    z.object({
      id: z.string().min(1).describe("Presentation ID"),
      slideId: z.string().min(1).describe("Slide ID"),
      text: z.string().min(1).describe("Comment text"),
      location: z.string().optional().describe("Location as JSON"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, slideId, text, location } = args as {
        id: string; slideId: string; text: string; location?: string;
      };
      return client.createPresentationComment(id, slideId, text, location ? JSON.parse(location) : undefined);
    },
    "Create presentation comment"
  ),

  "reply-presentation-comment": createCommand(
    z.object({
      id: z.string().min(1).describe("Presentation ID"),
      commentId: z.string().min(1).describe("Comment ID"),
      text: z.string().min(1).describe("Reply text"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, commentId, text } = args as { id: string; commentId: string; text: string };
      return client.replyToPresentationComment(id, commentId, text);
    },
    "Reply to presentation comment"
  ),

  "resolve-presentation-comment": createCommand(
    z.object({
      id: z.string().min(1).describe("Presentation ID"),
      commentId: z.string().min(1).describe("Comment ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, commentId } = args as { id: string; commentId: string };
      return client.resolvePresentationComment(id, commentId);
    },
    "Resolve presentation comment"
  ),

  // Pre-built cache commands
  ...cacheCommands<GoogleWorkspaceMCPClient>(),
};

// Run CLI
runCli(commands, GoogleWorkspaceMCPClient, {
  programName: "google-workspace-cli",
  description: "Google Workspace operations via MCP",
});

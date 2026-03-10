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
      calendarId: z.string().optional().describe("Calendar ID (default: primary)"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { summary, start, end, description, location, attendees, timezone, calendarId } = args as {
        summary: string; start: string; end: string;
        description?: string; location?: string; attendees?: string; timezone?: string; calendarId?: string;
      };
      // Validate: datetimes must include offset OR timezone must be provided
      const hasOffset = (dt: string) => /[+-]\d{2}:\d{2}$|Z$/.test(dt);
      if (!hasOffset(start) && !hasOffset(end) && !timezone) {
        throw new Error("Start/end times must include timezone offset (e.g. +00:00) or use --timezone flag");
      }
      return client.createEvent(summary, start, end, { description, location, attendees, timezone, calendarId });
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

  // ==================== Gmail Filters & Labels ====================
  "list-gmail-filters": createCommand(
    z.object({}),
    async (_args, client: GoogleWorkspaceMCPClient) => client.listGmailFilters(),
    "List Gmail filters"
  ),

  "create-gmail-filter": createCommand(
    z.object({
      criteria: z.string().min(1).describe('Filter criteria JSON (e.g. {"from":"user@example.com"})'),
      action: z.string().min(1).describe('Filter action JSON (e.g. {"addLabelIds":["LABEL_ID"],"removeLabelIds":["INBOX"]})'),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { criteria, action } = args as { criteria: string; action: string };
      return client.manageGmailFilter("create", {
        criteria: JSON.parse(criteria),
        filterAction: JSON.parse(action),
      });
    },
    "Create a Gmail filter"
  ),

  "delete-gmail-filter": createCommand(
    z.object({
      id: z.string().min(1).describe("Filter ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.manageGmailFilter("delete", { filterId: id });
    },
    "Delete a Gmail filter"
  ),

  "manage-gmail-label": createCommand(
    z.object({
      action: z.enum(["create", "update", "delete"]).describe("Action"),
      name: z.string().optional().describe("Label name (for create/update)"),
      labelId: z.string().optional().describe("Label ID (for update/delete)"),
      labelListVisibility: z.enum(["labelShow", "labelHide"]).optional().describe("Show in label list"),
      messageListVisibility: z.enum(["show", "hide"]).optional().describe("Show in message list"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { action, name, labelId, labelListVisibility, messageListVisibility } = args as {
        action: "create" | "update" | "delete"; name?: string; labelId?: string;
        labelListVisibility?: string; messageListVisibility?: string;
      };
      return client.manageGmailLabel(action, { name, labelId, labelListVisibility, messageListVisibility });
    },
    "Create/update/delete a Gmail label"
  ),

  "modify-message-labels": createCommand(
    z.object({
      id: z.string().min(1).describe("Message ID"),
      add: z.string().optional().describe("Comma-separated label IDs to add"),
      remove: z.string().optional().describe("Comma-separated label IDs to remove"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, add, remove } = args as { id: string; add?: string; remove?: string };
      return client.modifyGmailMessageLabels(
        id,
        add ? add.split(",").map(s => s.trim()) : undefined,
        remove ? remove.split(",").map(s => s.trim()) : undefined,
      );
    },
    "Add/remove labels on a message"
  ),

  "get-gmail-messages-batch": createCommand(
    z.object({
      ids: z.string().min(1).describe("Comma-separated message IDs (max 25)"),
      format: z.enum(["full", "metadata"]).optional().describe("Message format"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { ids, format } = args as { ids: string; format?: "full" | "metadata" };
      return client.getGmailMessagesBatch(ids.split(",").map(s => s.trim()), format);
    },
    "Get multiple emails in one batch"
  ),

  // ==================== Contacts ====================
  "list-contacts": createCommand(
    z.object({
      limit: cliTypes.int(1, 1000).optional().describe("Max results"),
      sortOrder: z.string().optional().describe("Sort: LAST_MODIFIED_ASCENDING, FIRST_NAME_ASCENDING, etc."),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { limit, sortOrder } = args as { limit?: number; sortOrder?: string };
      return client.listContacts(limit, sortOrder);
    },
    "List contacts"
  ),

  "get-contact": createCommand(
    z.object({
      id: z.string().min(1).describe("Contact ID (e.g. c1234567890)"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.getContact(id);
    },
    "Get contact details"
  ),

  "search-contacts": createCommand(
    z.object({
      query: z.string().min(1).describe("Search query (name, email, phone)"),
      limit: cliTypes.int(1, 30).optional().describe("Max results (max 30)"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { query, limit } = args as { query: string; limit?: number };
      return client.searchContacts(query, limit);
    },
    "Search contacts"
  ),

  "create-contact": createCommand(
    z.object({
      givenName: z.string().optional().describe("First name"),
      familyName: z.string().optional().describe("Last name"),
      email: z.string().optional().describe("Email address"),
      phone: z.string().optional().describe("Phone number"),
      organization: z.string().optional().describe("Company"),
      jobTitle: z.string().optional().describe("Job title"),
      notes: z.string().optional().describe("Notes"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { givenName, familyName, email, phone, organization, jobTitle, notes } = args as {
        givenName?: string; familyName?: string; email?: string; phone?: string;
        organization?: string; jobTitle?: string; notes?: string;
      };
      return client.manageContact("create", { givenName, familyName, email, phone, organization, jobTitle, notes });
    },
    "Create a contact"
  ),

  "update-contact": createCommand(
    z.object({
      id: z.string().min(1).describe("Contact ID"),
      givenName: z.string().optional().describe("First name"),
      familyName: z.string().optional().describe("Last name"),
      email: z.string().optional().describe("Email address"),
      phone: z.string().optional().describe("Phone number"),
      organization: z.string().optional().describe("Company"),
      jobTitle: z.string().optional().describe("Job title"),
      notes: z.string().optional().describe("Notes"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, givenName, familyName, email, phone, organization, jobTitle, notes } = args as {
        id: string; givenName?: string; familyName?: string; email?: string; phone?: string;
        organization?: string; jobTitle?: string; notes?: string;
      };
      return client.manageContact("update", { contactId: id, givenName, familyName, email, phone, organization, jobTitle, notes });
    },
    "Update a contact"
  ),

  "delete-contact": createCommand(
    z.object({
      id: z.string().min(1).describe("Contact ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.manageContact("delete", { contactId: id });
    },
    "Delete a contact"
  ),

  "list-contact-groups": createCommand(
    z.object({
      limit: cliTypes.int(1, 1000).optional().describe("Max results"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { limit } = args as { limit?: number };
      return client.listContactGroups(limit);
    },
    "List contact groups"
  ),

  // ==================== Chat ====================
  "list-chat-spaces": createCommand(
    z.object({
      type: z.enum(["all", "room", "dm"]).optional().describe("Space type filter"),
      limit: cliTypes.int(1, 1000).optional().describe("Max results"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { type, limit } = args as { type?: string; limit?: number };
      return client.listChatSpaces(type, limit);
    },
    "List Chat spaces"
  ),

  "get-chat-messages": createCommand(
    z.object({
      spaceId: z.string().min(1).describe("Space ID (e.g. spaces/XXXXXXXXX)"),
      limit: cliTypes.int(1, 1000).optional().describe("Max messages"),
      orderBy: z.string().optional().describe("Order (e.g. createTime desc)"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { spaceId, limit, orderBy } = args as { spaceId: string; limit?: number; orderBy?: string };
      return client.getChatMessages(spaceId, limit, orderBy);
    },
    "Get messages from a Chat space"
  ),

  "send-chat-message": createCommand(
    z.object({
      spaceId: z.string().min(1).describe("Space ID"),
      text: z.string().min(1).describe("Message text"),
      threadName: z.string().optional().describe("Thread resource name for reply"),
      threadKey: z.string().optional().describe("App-defined thread key"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { spaceId, text, threadName, threadKey } = args as {
        spaceId: string; text: string; threadName?: string; threadKey?: string;
      };
      return client.sendChatMessage(spaceId, text, threadName, threadKey);
    },
    "Send a Chat message"
  ),

  "search-chat-messages": createCommand(
    z.object({
      query: z.string().min(1).describe("Search query"),
      spaceId: z.string().optional().describe("Limit to specific space"),
      limit: cliTypes.int(1, 100).optional().describe("Max results"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { query, spaceId, limit } = args as { query: string; spaceId?: string; limit?: number };
      return client.searchChatMessages(query, spaceId, limit);
    },
    "Search Chat messages"
  ),

  // ==================== Advanced Drive ====================
  "copy-drive-file": createCommand(
    z.object({
      id: z.string().min(1).describe("File ID to copy"),
      name: z.string().optional().describe("New name for the copy"),
      parentId: z.string().optional().describe("Destination folder ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, name, parentId } = args as { id: string; name?: string; parentId?: string };
      return client.copyDriveFile(id, name, parentId);
    },
    "Copy a Drive file"
  ),

  "create-drive-folder": createCommand(
    z.object({
      name: z.string().min(1).describe("Folder name"),
      parentId: z.string().optional().describe("Parent folder ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { name, parentId } = args as { name: string; parentId?: string };
      return client.createDriveFolder(name, parentId);
    },
    "Create a Drive folder"
  ),

  "get-drive-share-link": createCommand(
    z.object({
      id: z.string().min(1).describe("File or folder ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.getDriveShareLink(id);
    },
    "Get shareable link"
  ),

  "share-drive-file": createCommand(
    z.object({
      id: z.string().min(1).describe("File or folder ID"),
      action: z.enum(["grant", "revoke", "update", "transfer_owner"]).describe("Access action"),
      shareWith: z.string().optional().describe("Email to share with"),
      role: z.enum(["reader", "commenter", "writer"]).optional().describe("Permission role"),
      shareType: z.enum(["user", "group", "domain", "anyone"]).optional().describe("Share type"),
      permissionId: z.string().optional().describe("Permission ID (for update/revoke)"),
      newOwnerEmail: z.string().optional().describe("New owner email (for transfer)"),
      sendNotification: cliTypes.bool().optional().describe("Send notification email"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, action, shareWith, role, shareType, permissionId, newOwnerEmail, sendNotification } = args as {
        id: string; action: "grant" | "revoke" | "update" | "transfer_owner";
        shareWith?: string; role?: string; shareType?: string;
        permissionId?: string; newOwnerEmail?: string; sendNotification?: boolean;
      };
      return client.manageDriveAccess(id, action, { shareWith, role, shareType, permissionId, newOwnerEmail, sendNotification });
    },
    "Share/unshare a Drive file"
  ),

  "get-drive-permissions": createCommand(
    z.object({
      id: z.string().min(1).describe("File ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id } = args as { id: string };
      return client.getDriveFilePermissions(id);
    },
    "Get file permissions"
  ),

  // ==================== Advanced Calendar ====================
  "modify-event": createCommand(
    z.object({
      action: z.enum(["create", "update", "delete"]).describe("Action"),
      summary: z.string().optional().describe("Event title"),
      start: z.string().optional().describe("Start time (RFC3339)"),
      end: z.string().optional().describe("End time (RFC3339)"),
      eventId: z.string().optional().describe("Event ID (for update/delete)"),
      calendarId: z.string().optional().describe("Calendar ID"),
      description: z.string().optional().describe("Event description"),
      location: z.string().optional().describe("Event location"),
      attendees: z.string().optional().describe("Comma-separated attendee emails"),
      timezone: z.string().optional().describe("IANA timezone"),
      addGoogleMeet: cliTypes.bool().optional().describe("Add Google Meet"),
      transparency: z.enum(["opaque", "transparent"]).optional().describe("Busy/free status"),
      visibility: z.enum(["default", "public", "private", "confidential"]).optional().describe("Visibility"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { action, summary, start, end, eventId, calendarId, description, location,
        attendees, timezone, addGoogleMeet, transparency, visibility } = args as {
        action: "create" | "update" | "delete"; summary?: string; start?: string; end?: string;
        eventId?: string; calendarId?: string; description?: string; location?: string;
        attendees?: string; timezone?: string; addGoogleMeet?: boolean;
        transparency?: string; visibility?: string;
      };
      return client.manageEvent(action, {
        summary, startTime: start, endTime: end, eventId, calendarId,
        description, location,
        attendees: attendees ? attendees.split(",").map(s => s.trim()) : undefined,
        timezone, addGoogleMeet, transparency, visibility,
      });
    },
    "Create/update/delete calendar event"
  ),

  "query-freebusy": createCommand(
    z.object({
      timeMin: z.string().min(1).describe("Start of interval (RFC3339)"),
      timeMax: z.string().min(1).describe("End of interval (RFC3339)"),
      calendarIds: z.string().optional().describe("Comma-separated calendar IDs"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { timeMin, timeMax, calendarIds } = args as { timeMin: string; timeMax: string; calendarIds?: string };
      return client.queryFreebusy(
        timeMin, timeMax,
        calendarIds ? calendarIds.split(",").map(s => s.trim()) : undefined,
      );
    },
    "Check free/busy status"
  ),

  // ==================== Advanced Docs ====================
  "export-doc-pdf": createCommand(
    z.object({
      id: z.string().min(1).describe("Document ID"),
      filename: z.string().optional().describe("PDF filename"),
      folderId: z.string().optional().describe("Destination folder ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, filename, folderId } = args as { id: string; filename?: string; folderId?: string };
      return client.exportDocToPdf(id, filename, folderId);
    },
    "Export document to PDF"
  ),

  "get-doc-markdown": createCommand(
    z.object({
      id: z.string().min(1).describe("Document ID or URL"),
      includeComments: cliTypes.bool().optional().describe("Include comments (default: true)"),
      commentMode: z.enum(["inline", "appendix", "none"]).optional().describe("Comment display mode"),
      includeResolved: cliTypes.bool().optional().describe("Include resolved comments"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, includeComments, commentMode, includeResolved } = args as {
        id: string; includeComments?: boolean; commentMode?: string; includeResolved?: boolean;
      };
      return client.getDocAsMarkdown(id, { includeComments, commentMode, includeResolved });
    },
    "Get document as Markdown"
  ),

  "list-docs-in-folder": createCommand(
    z.object({
      folderId: z.string().optional().describe("Folder ID (default: root)"),
      limit: cliTypes.int(1, 1000).optional().describe("Max results"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { folderId, limit } = args as { folderId?: string; limit?: number };
      return client.listDocsInFolder(folderId, limit);
    },
    "List Docs in a folder"
  ),

  // ==================== Advanced Sheets ====================
  "format-sheet-range": createCommand(
    z.object({
      id: z.string().min(1).describe("Spreadsheet ID"),
      range: z.string().min(1).describe("Range in A1 notation"),
      backgroundColor: z.string().optional().describe("Background hex color"),
      textColor: z.string().optional().describe("Text hex color"),
      numberFormatType: z.string().optional().describe("Format type (NUMBER, CURRENCY, DATE, PERCENT)"),
      numberFormatPattern: z.string().optional().describe("Custom format pattern"),
      wrapStrategy: z.enum(["WRAP", "CLIP", "OVERFLOW_CELL"]).optional().describe("Text wrapping"),
      horizontalAlignment: z.enum(["LEFT", "CENTER", "RIGHT"]).optional().describe("H alignment"),
      verticalAlignment: z.enum(["TOP", "MIDDLE", "BOTTOM"]).optional().describe("V alignment"),
      bold: cliTypes.bool().optional().describe("Bold text"),
      italic: cliTypes.bool().optional().describe("Italic text"),
      fontSize: cliTypes.int(1).optional().describe("Font size in points"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, range, ...opts } = args as {
        id: string; range: string; backgroundColor?: string; textColor?: string;
        numberFormatType?: string; numberFormatPattern?: string;
        wrapStrategy?: string; horizontalAlignment?: string; verticalAlignment?: string;
        bold?: boolean; italic?: boolean; fontSize?: number;
      };
      return client.formatSheetRange(id, range, opts);
    },
    "Format a sheet range"
  ),

  "add-sheet": createCommand(
    z.object({
      id: z.string().min(1).describe("Spreadsheet ID"),
      name: z.string().min(1).describe("New sheet name"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { id, name } = args as { id: string; name: string };
      return client.createSheet(id, name);
    },
    "Add a sheet tab"
  ),

  // ==================== Advanced Tasks ====================
  "get-task": createCommand(
    z.object({
      listId: z.string().min(1).describe("Task list ID"),
      id: z.string().min(1).describe("Task ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { listId, id } = args as { listId: string; id: string };
      return client.getTask(listId, id);
    },
    "Get task details"
  ),

  "update-task": createCommand(
    z.object({
      listId: z.string().min(1).describe("Task list ID"),
      id: z.string().min(1).describe("Task ID"),
      title: z.string().optional().describe("New title"),
      notes: z.string().optional().describe("New notes"),
      status: z.enum(["needsAction", "completed"]).optional().describe("Status"),
      due: z.string().optional().describe("Due date (RFC3339)"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { listId, id, title, notes, status, due } = args as {
        listId: string; id: string; title?: string; notes?: string; status?: string; due?: string;
      };
      return client.manageTask("update", listId, { taskId: id, title, notes, status, due });
    },
    "Update a task"
  ),

  "delete-task": createCommand(
    z.object({
      listId: z.string().min(1).describe("Task list ID"),
      id: z.string().min(1).describe("Task ID"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { listId, id } = args as { listId: string; id: string };
      return client.manageTask("delete", listId, { taskId: id });
    },
    "Delete a task"
  ),

  "move-task": createCommand(
    z.object({
      listId: z.string().min(1).describe("Source task list ID"),
      id: z.string().min(1).describe("Task ID"),
      destinationListId: z.string().optional().describe("Destination task list ID"),
      parent: z.string().optional().describe("Parent task ID (make subtask)"),
    }),
    async (args, client: GoogleWorkspaceMCPClient) => {
      const { listId, id, destinationListId, parent } = args as {
        listId: string; id: string; destinationListId?: string; parent?: string;
      };
      return client.manageTask("move", listId, { taskId: id, destinationTaskList: destinationListId, parent });
    },
    "Move a task"
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

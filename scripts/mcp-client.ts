/**
 * Google Workspace MCP Client
 *
 * Wrapper client for Google Workspace APIs via MCP server.
 * Provides access to Gmail, Calendar, Drive, Docs, Sheets, Tasks, and Comments.
 *
 * Key features:
 * - Gmail: search, read, send, draft
 * - Calendar: events, CRUD operations
 * - Drive: file search, content retrieval
 * - Docs: content, text modification, find/replace
 * - Sheets: read, write, rich text cells with hyperlinks
 * - Tasks: task lists, task management
 * - Comments: read/write/reply/resolve across Docs, Sheets, Slides
 */

import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { PluginCache, TTL, createCacheKey } from "@local/plugin-cache";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

interface MCPConfig {
  mcpServer: {
    command: string;
    args: string[];
    env?: Record<string, string>;
  };
  userEmail?: string;
}

/**
 * Rich text segment with optional formatting and hyperlink
 */
export interface RichTextSegment {
  text: string;
  url?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  sproductthrough?: boolean;
  fontSize?: number;
  fontFamily?: string;
  foregroundColor?: string;  // Hex color, e.g., "#FF0000"
}

/**
 * Cell definition for batch rich text writing
 */
export interface RichTextCellDef {
  cell: string;  // A1 notation, e.g., "AD2" or "Sheet1!B5"
  segments: RichTextSegment[];
}

// Initialize cache with namespace
const cache = new PluginCache({
  namespace: "google-workspace-manager",
  defaultTTL: TTL.FIVE_MINUTES,
});

export class GoogleWorkspaceMCPClient {
  private client: Client | null = null;
  private transport: StdioClientTransport | null = null;
  private config: MCPConfig;
  private connected: boolean = false;
  private cacheDisabled: boolean = false;

  constructor() {
    // When compiled, __dirname is dist/, so look in parent for config.json
    const configPath = join(__dirname, "..", "config.json");
    this.config = JSON.parse(readFileSync(configPath, "utf-8"));
  }

  // ============================================
  // CACHE CONTROL
  // ============================================

  /** Disables caching for all subsequent requests. */
  disableCache(): void {
    this.cacheDisabled = true;
    cache.disable();
  }

  /** Re-enables caching after it was disabled. */
  enableCache(): void {
    this.cacheDisabled = false;
    cache.enable();
  }

  /** Returns cache statistics including hit/miss counts. */
  getCacheStats() {
    return cache.getStats();
  }

  /** Clears all cached data. @returns Number of cache entries cleared */
  clearCache(): number {
    return cache.clear();
  }

  /** Invalidates a specific cache entry by key. */
  invalidateCacheKey(key: string): boolean {
    return cache.invalidate(key);
  }

  // ============================================
  // CONNECTION MANAGEMENT
  // ============================================

  /** Establishes connection to the MCP server. */
  async connect(): Promise<void> {
    if (this.connected) return;

    const env = {
      ...process.env,
      ...this.config.mcpServer.env,
    };

    // Ensure required env vars are set for workspace-mcp
    if (!env.GOOGLE_OAUTH_CLIENT_ID) {
      throw new Error(
        "GOOGLE_OAUTH_CLIENT_ID environment variable is not set."
      );
    }
    if (!env.GOOGLE_OAUTH_CLIENT_SECRET) {
      throw new Error(
        "GOOGLE_OAUTH_CLIENT_SECRET environment variable is not set."
      );
    }
    // Derive GOOGLE_MCP_CREDENTIALS_DIR from GOOGLE_OAUTH_TOKEN if not set
    if (!env.GOOGLE_MCP_CREDENTIALS_DIR && env.GOOGLE_OAUTH_TOKEN) {
      env.GOOGLE_MCP_CREDENTIALS_DIR = dirname(env.GOOGLE_OAUTH_TOKEN);
    }
    if (!env.GOOGLE_MCP_CREDENTIALS_DIR) {
      throw new Error(
        "GOOGLE_MCP_CREDENTIALS_DIR environment variable is not set."
      );
    }

    this.transport = new StdioClientTransport({
      command: this.config.mcpServer.command,
      args: this.config.mcpServer.args,
      env: env as Record<string, string>,
    });

    this.client = new Client(
      { name: "google-workspace-cli", version: "1.0.0" },
      { capabilities: {} }
    );

    await this.client.connect(this.transport);
    this.connected = true;
  }

  /** Disconnects from the MCP server. */
  async disconnect(): Promise<void> {
    if (this.client && this.connected) {
      await this.client.close();
      this.connected = false;
    }
  }

  // ============================================
  // MCP TOOLS
  // ============================================

  /** Lists available MCP tools. @returns Array of tool definitions */
  async listTools(): Promise<any[]> {
    await this.connect();
    const result = await this.client!.listTools();
    return result.tools;
  }

  /** Calls an MCP tool with arguments. Automatically injects user email if configured. */
  async callTool(name: string, args: Record<string, any> = {}): Promise<any> {
    await this.connect();

    // Add user email if configured
    if (this.config.userEmail && !args.user_google_email) {
      args.user_google_email = this.config.userEmail;
    }

    const result = await this.client!.callTool({ name, arguments: args });
    const content = result.content as Array<{ type: string; text?: string }>;

    if (result.isError) {
      const errorContent = content.find((c) => c.type === "text");
      throw new Error(errorContent?.text || "Tool call failed");
    }

    const textContent = content.find((c) => c.type === "text");
    if (textContent?.text) {
      try {
        return JSON.parse(textContent.text);
      } catch {
        return textContent.text;
      }
    }

    return content;
  }

  // ============================================
  // GMAIL OPERATIONS
  // ============================================

  /**
   * Searches Gmail messages.
   * @param query - Gmail search query (e.g., "from:user@example.com")
   * @param maxResults - Max messages to return
   * @returns Search results with message metadata
   * @cached TTL: 5 minutes
   */
  async searchGmailMessages(query: string, maxResults?: number): Promise<any> {
    const cacheKey = createCacheKey("gmail_search", { query, maxResults });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { query };
        if (maxResults) args.maxResults = maxResults;
        return this.callTool("search_gmail_messages", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Gets full content of a Gmail message. @cached TTL: 15 minutes */
  async getGmailMessage(messageId: string): Promise<any> {
    const cacheKey = createCacheKey("gmail_message", { id: messageId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_gmail_message_content", { message_id: messageId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Lists all Gmail labels. @cached TTL: 1 hour */
  async listGmailLabels(): Promise<any> {
    return cache.getOrFetch(
      "gmail_labels",
      () => this.callTool("list_gmail_labels", {}),
      { ttl: TTL.HOUR, bypassCache: this.cacheDisabled }
    );
  }

  /** Sends a Gmail message. @invalidates gmail_search/* */
  async sendGmailMessage(to: string, subject: string, body: string, cc?: string, bcc?: string): Promise<any> {
    const args: Record<string, any> = { to, subject, body };
    if (cc) args.cc = cc;
    if (bcc) args.bcc = bcc;
    const result = await this.callTool("send_gmail_message", args);
    cache.invalidatePattern(/^gmail_search/);
    return result;
  }

  /** Creates a Gmail draft. */
  async createGmailDraft(to: string, subject: string, body: string): Promise<any> {
    return this.callTool("draft_gmail_message", { to, subject, body });
  }

  /** Gets all messages in a Gmail thread. @cached TTL: 15 minutes */
  async getGmailThread(threadId: string): Promise<any> {
    const cacheKey = createCacheKey("gmail_thread", { id: threadId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_gmail_thread_content", { thread_id: threadId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  // ============================================
  // CALENDAR OPERATIONS
  // ============================================

  /** Lists all calendars. @cached TTL: 1 hour */
  async listCalendars(): Promise<any> {
    return cache.getOrFetch(
      "calendars",
      () => this.callTool("list_calendars", {}),
      { ttl: TTL.HOUR, bypassCache: this.cacheDisabled }
    );
  }

  /** Gets calendar events with optional date range filtering. @cached TTL: 15 minutes */
  async getEvents(options?: { calendarId?: string; timeMin?: string; timeMax?: string; maxResults?: number }): Promise<any> {
    const cacheKey = createCacheKey("calendar_events", options || {});
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (options?.calendarId) args.calendar_id = options.calendarId;
        if (options?.timeMin) args.time_min = options.timeMin;
        if (options?.timeMax) args.time_max = options.timeMax;
        if (options?.maxResults) args.max_results = options.maxResults;
        return this.callTool("get_events", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Creates a calendar event. @invalidates calendar_events/* */
  async createEvent(summary: string, start: string, end: string, options?: { description?: string; location?: string; attendees?: string; timezone?: string }): Promise<any> {
    const args: Record<string, any> = { summary, start_time: start, end_time: end };
    if (options?.description) args.description = options.description;
    if (options?.location) args.location = options.location;
    if (options?.attendees) args.attendees = options.attendees;
    if (options?.timezone) args.timezone = options.timezone;
    const result = await this.callTool("create_event", args);
    cache.invalidatePattern(/^calendar_events/);
    return result;
  }

  /** Deletes a calendar event. @invalidates calendar_events/* */
  async deleteEvent(eventId: string, calendarId?: string): Promise<any> {
    const args: Record<string, any> = { event_id: eventId };
    if (calendarId) args.calendar_id = calendarId;
    const result = await this.callTool("delete_event", args);
    cache.invalidatePattern(/^calendar_events/);
    return result;
  }

  // ============================================
  // DRIVE OPERATIONS
  // ============================================

  /** Searches Drive files by name/content. @cached TTL: 15 minutes */
  async searchDriveFiles(query: string, maxResults?: number): Promise<any> {
    const cacheKey = createCacheKey("drive_search", { query, maxResults });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { query };
        if (maxResults) args.maxResults = maxResults;
        return this.callTool("search_drive_files", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Gets content of a Drive file. @cached TTL: 5 minutes */
  async getDriveFileContent(fileId: string): Promise<any> {
    const cacheKey = createCacheKey("drive_file", { id: fileId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_drive_file_content", { file_id: fileId }),
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Lists items in a Drive folder. @cached TTL: 15 minutes */
  async listDriveItems(folderId?: string): Promise<any> {
    const cacheKey = createCacheKey("drive_items", { folder: folderId || "root" });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (folderId) args.folderId = folderId;
        return this.callTool("list_drive_items", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  // ============================================
  // DOCS OPERATIONS
  // ============================================

  /** Searches Google Docs. @cached TTL: 15 minutes */
  async searchDocs(query: string): Promise<any> {
    const cacheKey = createCacheKey("docs_search", { query });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("search_docs", { query }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Gets Google Doc content with optional suggestions view. @cached TTL: 5 minutes */
  async getDocContent(documentId: string, suggestionsViewMode?: string): Promise<any> {
    const cacheKey = createCacheKey("doc_content", { id: documentId, mode: suggestionsViewMode });
    return cache.getOrFetch(
      cacheKey,
      () => {
        const args: Record<string, any> = { document_id: documentId };
        if (suggestionsViewMode) args.suggestions_view_mode = suggestionsViewMode;
        return this.callTool("get_doc_content", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Creates a new Google Doc. @invalidates docs_search/* */
  async createDoc(title: string, content?: string): Promise<any> {
    const args: Record<string, any> = { title };
    if (content) args.content = content;
    const result = await this.callTool("create_doc", args);
    cache.invalidatePattern(/^docs_search/);
    return result;
  }

  /** Modifies text in a Google Doc (insert, replace, or delete). @invalidates doc_content/{documentId} */
  async modifyDocText(
    documentId: string,
    operation: "insert" | "replace" | "delete",
    options: {
      index?: number;
      text?: string;
      startIndex?: number;
      endIndex?: number;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      fontSize?: number;
      fontFamily?: string;
    }
  ): Promise<any> {
    const args: Record<string, any> = {
      document_id: documentId,
      operation
    };
    if (options.index !== undefined) args.index = options.index;
    if (options.text) args.text = options.text;
    if (options.startIndex !== undefined) args.start_index = options.startIndex;
    if (options.endIndex !== undefined) args.end_index = options.endIndex;
    if (options.bold !== undefined) args.bold = options.bold;
    if (options.italic !== undefined) args.italic = options.italic;
    if (options.underline !== undefined) args.underline = options.underline;
    if (options.fontSize !== undefined) args.font_size = options.fontSize;
    if (options.fontFamily) args.font_family = options.fontFamily;

    const result = await this.callTool("modify_doc_text", args);
    cache.invalidate(createCacheKey("doc_content", { id: documentId }));
    return result;
  }

  /** Find and replace text in a Google Doc. @invalidates doc_content/{documentId} */
  async findAndReplaceDoc(
    documentId: string,
    findText: string,
    replaceText: string,
    replaceAll: boolean = true
  ): Promise<any> {
    const result = await this.callTool("find_and_replace_doc", {
      document_id: documentId,
      find_text: findText,
      replace_text: replaceText,
      replace_all: replaceAll
    });
    cache.invalidate(createCacheKey("doc_content", { id: documentId }));
    return result;
  }

  // ============================================
  // SHEETS OPERATIONS
  // ============================================

  /** Lists all spreadsheets. @cached TTL: 15 minutes */
  async listSpreadsheets(): Promise<any> {
    return cache.getOrFetch(
      "spreadsheets_list",
      () => this.callTool("list_spreadsheets", {}),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Gets spreadsheet metadata and sheet names. @cached TTL: 15 minutes */
  async getSpreadsheetInfo(spreadsheetId: string): Promise<any> {
    const cacheKey = createCacheKey("spreadsheet_info", { id: spreadsheetId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_spreadsheet_info", { spreadsheet_id: spreadsheetId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Reads values from a spreadsheet range. @cached TTL: 5 minutes */
  async readSheetValues(spreadsheetId: string, range: string): Promise<any> {
    const cacheKey = createCacheKey("sheet_values", { id: spreadsheetId, range });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("read_sheet_values", { spreadsheet_id: spreadsheetId, range_name: range }),
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Writes values to a spreadsheet range. @invalidates sheet_values/{spreadsheetId}/{range} */
  async writeSheetValues(spreadsheetId: string, range: string, values: any[][]): Promise<any> {
    const result = await this.callTool("modify_sheet_values", { spreadsheet_id: spreadsheetId, range_name: range, values });
    cache.invalidate(createCacheKey("sheet_values", { id: spreadsheetId, range }));
    return result;
  }

  /**
   * Write rich text with multiple hyperlinks to a single cell.
   *
   * @param spreadsheetId - The spreadsheet ID
   * @param cell - Cell reference in A1 notation (e.g., "AD2", "Sheet1!B5")
   * @param segments - Array of RichTextSegment objects (text + optional url/formatting)
   * @param sheetName - Optional sheet name (default: first sheet)
   *
   * @example
   * await client.writeRichTextCell(spreadsheetId, "AD2", [
   *   { text: "WARNING: ", bold: true, foregroundColor: "#FF0000" },
   *   { text: "See " },
   *   { text: "ticket #123", url: "https://gorgias.com/ticket/123" }
   * ]);
   */
  async writeRichTextCell(
    spreadsheetId: string,
    cell: string,
    segments: RichTextSegment[],
    sheetName?: string
  ): Promise<any> {
    const args: Record<string, any> = {
      spreadsheet_id: spreadsheetId,
      cell: cell,
      segments: segments,
    };

    if (sheetName) {
      args.sheet_name = sheetName;
    }

    const result = await this.callTool("write_rich_text_cell", args);

    // Invalidate cache for this spreadsheet (all ranges since we don't know the exact range)
    cache.invalidatePattern(new RegExp(`^sheet_values:.*${spreadsheetId}`));

    return result;
  }

  /**
   * Write rich text to multiple cells in a single API call.
   *
   * @param spreadsheetId - The spreadsheet ID
   * @param cells - Array of {cell: string, segments: RichTextSegment[]} objects
   * @param sheetName - Optional sheet name (default: first sheet)
   *
   * @example
   * await client.writeRichTextCells(spreadsheetId, [
   *   { cell: "AD2", segments: [{ text: "Bold", bold: true }] },
   *   { cell: "AD3", segments: [{ text: "Link", url: "https://..." }] },
   *   { cell: "AD4", segments: [{ text: "Red", foregroundColor: "#FF0000" }] }
   * ]);
   */
  async writeRichTextCells(
    spreadsheetId: string,
    cells: RichTextCellDef[],
    sheetName?: string
  ): Promise<any> {
    const args: Record<string, any> = {
      spreadsheet_id: spreadsheetId,
      cells: cells,
    };

    if (sheetName) {
      args.sheet_name = sheetName;
    }

    const result = await this.callTool("write_rich_text_cells", args);

    // Invalidate cache for this spreadsheet
    cache.invalidatePattern(new RegExp(`^sheet_values:.*${spreadsheetId}`));

    return result;
  }

  /** Creates a new spreadsheet with optional named sheets. */
  async createSpreadsheet(title: string, sheetNames?: string[]): Promise<any> {
    const args: Record<string, any> = { title };
    if (sheetNames) args.sheet_names = sheetNames;
    return this.callTool("create_spreadsheet", args);
  }

  // ============================================
  // TASKS OPERATIONS
  // ============================================

  /** Lists all task lists. @cached TTL: 15 minutes */
  async listTaskLists(): Promise<any> {
    return cache.getOrFetch(
      "task_lists",
      () => this.callTool("list_task_lists", {}),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Lists tasks in a task list. @cached TTL: 5 minutes */
  async listTasks(taskListId: string): Promise<any> {
    const cacheKey = createCacheKey("tasks", { listId: taskListId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("list_tasks", { task_list_id: taskListId }),
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Creates a new task. @invalidates tasks/{taskListId} */
  async createTask(taskListId: string, title: string, notes?: string, due?: string): Promise<any> {
    const args: Record<string, any> = { task_list_id: taskListId, title };
    if (notes) args.notes = notes;
    if (due) args.due = due;
    const result = await this.callTool("create_task", args);
    cache.invalidate(createCacheKey("tasks", { listId: taskListId }));
    return result;
  }

  /** Marks a task as completed. @invalidates tasks/{taskListId} */
  async completeTask(taskListId: string, taskId: string): Promise<any> {
    const result = await this.callTool("update_task", { tasklist_id: taskListId, task_id: taskId, status: "completed" });
    cache.invalidate(createCacheKey("tasks", { listId: taskListId }));
    return result;
  }

  // ============================================
  // DOCUMENT COMMENTS OPERATIONS
  // ============================================

  /** Gets comments on a Google Doc. @cached TTL: 15 minutes */
  async getDocumentComments(documentId: string): Promise<any> {
    const cacheKey = createCacheKey("doc_comments", { id: documentId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("read_document_comments", { document_id: documentId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /** Creates a comment on a Google Doc. @invalidates doc_comments/{documentId} */
  async createDocumentComment(documentId: string, text: string, location?: Record<string, any>): Promise<any> {
    const args: Record<string, any> = { document_id: documentId, text };
    if (location) args.location = location;
    const result = await this.callTool("create_document_comment", args);
    cache.invalidate(createCacheKey("doc_comments", { id: documentId }));
    return result;
  }

  /** Replies to a document comment. @invalidates doc_comments/{documentId} */
  async replyToDocumentComment(documentId: string, commentId: string, text: string): Promise<any> {
    const result = await this.callTool("reply_to_document_comment", { document_id: documentId, comment_id: commentId, text });
    cache.invalidate(createCacheKey("doc_comments", { id: documentId }));
    return result;
  }

  /** Resolves a document comment. @invalidates doc_comments/{documentId} */
  async resolveDocumentComment(documentId: string, commentId: string): Promise<any> {
    const result = await this.callTool("resolve_document_comment", { document_id: documentId, comment_id: commentId });
    cache.invalidate(createCacheKey("doc_comments", { id: documentId }));
    return result;
  }

  // ============================================
  // SPREADSHEET COMMENTS OPERATIONS
  // ============================================

  /** Gets comments on a spreadsheet. @cached TTL: 15 minutes */
  async getSpreadsheetComments(spreadsheetId: string): Promise<any> {
    const cacheKey = createCacheKey("sheet_comments", { id: spreadsheetId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("read_spreadsheet_comments", { spreadsheet_id: spreadsheetId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /**
   * Creates a comment on a spreadsheet cell.
   *
   * @param spreadsheetId - The spreadsheet ID
   * @param sheetId - Numeric sheet ID (from getSpreadsheetInfo)
   * @param rowIndex - Zero-based row index
   * @param columnIndex - Zero-based column index
   * @param text - Comment text
   * @returns Created comment details
   *
   * @invalidates sheet_comments/{spreadsheetId}
   */
  async createSpreadsheetComment(spreadsheetId: string, sheetId: number, rowIndex: number, columnIndex: number, text: string): Promise<any> {
    const result = await this.callTool("create_spreadsheet_comment", { spreadsheet_id: spreadsheetId, sheet_id: sheetId, row_index: rowIndex, column_index: columnIndex, text });
    cache.invalidate(createCacheKey("sheet_comments", { id: spreadsheetId }));
    return result;
  }

  /**
   * Replies to a spreadsheet comment.
   *
   * @param spreadsheetId - The spreadsheet ID
   * @param commentId - ID of the comment to reply to
   * @param text - Reply text
   * @returns Reply details
   *
   * @invalidates sheet_comments/{spreadsheetId}
   */
  async replyToSpreadsheetComment(spreadsheetId: string, commentId: string, text: string): Promise<any> {
    const result = await this.callTool("reply_to_spreadsheet_comment", { spreadsheet_id: spreadsheetId, comment_id: commentId, text });
    cache.invalidate(createCacheKey("sheet_comments", { id: spreadsheetId }));
    return result;
  }

  /**
   * Resolves (closes) a spreadsheet comment.
   *
   * @param spreadsheetId - The spreadsheet ID
   * @param commentId - ID of the comment to resolve
   * @returns Updated comment details
   *
   * @invalidates sheet_comments/{spreadsheetId}
   */
  async resolveSpreadsheetComment(spreadsheetId: string, commentId: string): Promise<any> {
    const result = await this.callTool("resolve_spreadsheet_comment", { spreadsheet_id: spreadsheetId, comment_id: commentId });
    cache.invalidate(createCacheKey("sheet_comments", { id: spreadsheetId }));
    return result;
  }

  // ============================================
  // PRESENTATION COMMENTS OPERATIONS
  // ============================================

  /** Gets comments on a Google Slides presentation. @cached TTL: 15 minutes */
  async getPresentationComments(presentationId: string): Promise<any> {
    const cacheKey = createCacheKey("presentation_comments", { id: presentationId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("read_presentation_comments", { presentation_id: presentationId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  /**
   * Creates a comment on a presentation slide.
   *
   * @param presentationId - The presentation ID
   * @param slideId - ID of the slide to comment on
   * @param text - Comment text
   * @param location - Optional location anchor within the slide
   * @returns Created comment details
   *
   * @invalidates presentation_comments/{presentationId}
   */
  async createPresentationComment(presentationId: string, slideId: string, text: string, location?: Record<string, any>): Promise<any> {
    const args: Record<string, any> = { presentation_id: presentationId, slide_id: slideId, text };
    if (location) args.location = location;
    const result = await this.callTool("create_presentation_comment", args);
    cache.invalidate(createCacheKey("presentation_comments", { id: presentationId }));
    return result;
  }

  /**
   * Replies to a presentation comment.
   *
   * @param presentationId - The presentation ID
   * @param commentId - ID of the comment to reply to
   * @param text - Reply text
   * @returns Reply details
   *
   * @invalidates presentation_comments/{presentationId}
   */
  async replyToPresentationComment(presentationId: string, commentId: string, text: string): Promise<any> {
    const result = await this.callTool("reply_to_presentation_comment", { presentation_id: presentationId, comment_id: commentId, text });
    cache.invalidate(createCacheKey("presentation_comments", { id: presentationId }));
    return result;
  }

  /**
   * Resolves (closes) a presentation comment.
   *
   * @param presentationId - The presentation ID
   * @param commentId - ID of the comment to resolve
   * @returns Updated comment details
   *
   * @invalidates presentation_comments/{presentationId}
   */
  async resolvePresentationComment(presentationId: string, commentId: string): Promise<any> {
    const result = await this.callTool("resolve_presentation_comment", { presentation_id: presentationId, comment_id: commentId });
    cache.invalidate(createCacheKey("presentation_comments", { id: presentationId }));
    return result;
  }
}

export default GoogleWorkspaceMCPClient;

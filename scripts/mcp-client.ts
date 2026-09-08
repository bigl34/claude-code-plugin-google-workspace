
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import type { RequestOptions } from "@modelcontextprotocol/sdk/shared/protocol.js";
import type { IOType } from "node:child_process";
import { existsSync } from "fs";
import { homedir } from "os";
import { basename, dirname, join } from "path";
import { loadServiceConfig, z } from "@local/cli-utils";
import { PluginCache, TTL, createCacheKey } from "@local/plugin-cache";

const GoogleWorkspaceConfigSchema = z.object({
  mcpServer: z.object({
    command: z.string().min(1),
    args: z.array(z.string()),
    env: z.record(z.string(), z.string()).optional(),
  }),
  userEmail: z.string().email().optional(),
});

type MCPConfig = z.infer<typeof GoogleWorkspaceConfigSchema>;

export interface DriveFileRawEnvelope {
  encoding?: string;
  data?: string;
  sizeBytes?: number;
  fileId?: string;
  fileName?: string;
  exportMimeType?: string;
  sourceMimeType?: string;
}

export function sheetValuesCachePattern(spreadsheetId: string): RegExp {
  const escapedSpreadsheetId = spreadsheetId.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`^sheet_values\\?(?:[^&]+&)*id=${escapedSpreadsheetId}(?:&|$)`); // nosemgrep: detect-non-literal-regexp
}

const DEFAULT_LOCAL_MCP_SOURCE = join(
  homedir(),
  "repos",
  "work",
  "mcp-servers",
  "google_workspace_mcp",
);

export function resolveGoogleWorkspaceMcpLaunch(
  command: string,
  args: string[],
  env: Record<string, string | undefined>,
  pathExists: (path: string) => boolean = existsSync,
  enforceReadOnlyIndexing: boolean = false,
): { command: string; args: string[]; cwd?: string } {
  if (enforceReadOnlyIndexing) {
    const conflictingFlags = new Set(["--tools", "--tool-tier", "--permissions"]);
    const conflict = args.find((arg) => conflictingFlags.has(arg.split("=", 1)[0]));
    if (conflict) {
      throw new Error(
        `Read-only indexing refuses conflicting MCP launch flag '${conflict}'`,
      );
    }
  }

  const explicitSource = env.GOOGLE_WORKSPACE_MCP_SOURCE?.trim();
  const source = explicitSource || DEFAULT_LOCAL_MCP_SOURCE;
  const localSourceRequired = Boolean(explicitSource) || enforceReadOnlyIndexing;

  if (!pathExists(source)) {
    if (localSourceRequired) {
      throw new Error(
        explicitSource
          ? `GOOGLE_WORKSPACE_MCP_SOURCE does not exist: ${explicitSource}`
          : `Read-only indexing requires the canonical local Google Workspace MCP: ${source}`,
      );
    }
    return { command, args: [...args] };
  }

  const fromIndex = args.findIndex((arg) => arg === "--from");
  const configuredSource = fromIndex >= 0 ? args[fromIndex + 1] : undefined;
  if (!configuredSource || !configuredSource.includes("google_workspace_mcp")) {
    if (localSourceRequired) {
      throw new Error(
        "Read-only indexing requires a Google Workspace MCP --from launch source",
      );
    }
    return { command, args: [...args] };
  }

  const entrypointIndex = args.findIndex((arg) => arg === "workspace-mcp");
  const commandName = basename(command);
  if (entrypointIndex < 0 || (commandName !== "uv" && commandName !== "uvx")) {
    if (localSourceRequired) {
      throw new Error(
        "Local Google Workspace MCP requires a uv/uvx launch with the workspace-mcp entrypoint",
      );
    }
    return { command, args: [...args] };
  }

  const expectedPrefix =
    commandName === "uvx"
      ? ["--from", configuredSource]
      : ["tool", "uvx", "--from", configuredSource];
  const actualPrefix = args.slice(0, entrypointIndex);
  if (
    actualPrefix.length !== expectedPrefix.length ||
    actualPrefix.some((arg, index) => arg !== expectedPrefix[index])
  ) {
    throw new Error(
      "Google Workspace MCP launch has unsupported uv/uvx arguments before workspace-mcp",
    );
  }

  const uvCommand =
    commandName === "uvx" ? join(dirname(command), "uv") : command;
  return {
    command: uvCommand,
    args: [
      "run",
      "--project",
      source,
      "workspace-mcp",
      ...args.slice(entrypointIndex + 1),
    ],
    cwd: source,
  };
}

const ACCOUNT_OAUTH_ENV_PATTERN =
  /^GOOGLE_(?:OAUTH_CLIENT_(?:ID|SECRET)|MCP_(?:CREDENTIALS_DIR|ACCOUNT_PROFILE))_[A-Z0-9_]+$/;

interface AccountOAuthEnvironmentOptions {
  interactiveAuth?: boolean;
  requireDedicatedIndexingClient?: boolean;
  forceReadOnlyIndexing?: boolean;
}

export function resolveAccountOAuthEnvironment(
  sourceEnv: Record<string, string | undefined>,
  accountName?: string,
  options: AccountOAuthEnvironmentOptions = {},
): Record<string, string | undefined> {
  const env = { ...sourceEnv };
  let accountProfile: string | undefined;

  if (accountName) {
    const suffix = accountName.trim().toUpperCase();
    if (!/^[A-Z0-9_]+$/.test(suffix)) {
      throw new Error(
        `Account name '${accountName}' cannot be used for OAuth client selection`,
      );
    }

    const clientIdKey = `GOOGLE_OAUTH_CLIENT_ID_${suffix}`;
    const clientSecretKey = `GOOGLE_OAUTH_CLIENT_SECRET_${suffix}`;
    const credentialsDirKey = `GOOGLE_MCP_CREDENTIALS_DIR_${suffix}`;
    const profileKey = `GOOGLE_MCP_ACCOUNT_PROFILE_${suffix}`;
    const clientId = env[clientIdKey]?.trim();
    const clientSecret = env[clientSecretKey]?.trim();

    if (Boolean(clientId) !== Boolean(clientSecret)) {
      throw new Error(
        `Account '${accountName}' requires both ${clientIdKey} and ${clientSecretKey}`,
      );
    }
    if (options.requireDedicatedIndexingClient && (!clientId || !clientSecret)) {
      throw new Error(
        `Account '${accountName}' requires a dedicated OAuth client for this operation`,
      );
    }
    if (
      options.requireDedicatedIndexingClient &&
      clientId === env.GOOGLE_OAUTH_CLIENT_ID?.trim()
    ) {
      throw new Error(
        `Account '${accountName}' dedicated OAuth client must differ from the shared client`,
      );
    }

    if (clientId && clientSecret) {
      env.GOOGLE_OAUTH_CLIENT_ID = clientId;
      env.GOOGLE_OAUTH_CLIENT_SECRET = clientSecret;
    }

    const credentialsDir = env[credentialsDirKey]?.trim();
    if (options.requireDedicatedIndexingClient && !credentialsDir) {
      throw new Error(
        `Account '${accountName}' requires an isolated credential directory for this operation`,
      );
    }
    if (options.requireDedicatedIndexingClient && credentialsDir) {
      const genericCredentialDirs = new Set(
        [
          env.WORKSPACE_MCP_CREDENTIALS_DIR?.trim(),
          env.GOOGLE_MCP_CREDENTIALS_DIR?.trim(),
          env.GOOGLE_OAUTH_TOKEN?.trim()
            ? dirname(env.GOOGLE_OAUTH_TOKEN.trim())
            : undefined,
        ].filter((value): value is string => Boolean(value)),
      );
      if (genericCredentialDirs.has(credentialsDir)) {
        throw new Error(
          `Account '${accountName}' credential directory must differ from the shared directory`,
        );
      }
    }
    if (credentialsDir) {
      env.GOOGLE_MCP_CREDENTIALS_DIR = credentialsDir;
      env.WORKSPACE_MCP_CREDENTIALS_DIR = credentialsDir;
    }

    accountProfile = env[profileKey]?.trim().toLowerCase();
    if (accountProfile && accountProfile !== "indexing_readonly") {
      throw new Error(
        `Account '${accountName}' has unsupported OAuth profile '${accountProfile}'`,
      );
    }
    if (options.requireDedicatedIndexingClient && accountProfile !== "indexing_readonly") {
      throw new Error(
        `Account '${accountName}' requires the indexing_readonly OAuth profile for this operation`,
      );
    }
  }

  if (accountProfile === "indexing_readonly" || options.forceReadOnlyIndexing) {
    env.WORKSPACE_MCP_TOOLS = "gmail,drive,docs,sheets";
    env.WORKSPACE_MCP_READ_ONLY = "true";
    env.WORKSPACE_MCP_NONINTERACTIVE =
      accountProfile === "indexing_readonly" && options.interactiveAuth ? "0" : "1";
    delete env.WORKSPACE_MCP_PERMISSIONS;
    delete env.WORKSPACE_MCP_TOOL_TIER;
  }

  delete env.WORKSPACE_MCP_INDEXING_PROFILE;

  for (const key of Object.keys(env)) {
    if (ACCOUNT_OAUTH_ENV_PATTERN.test(key)) {
      delete env[key];
    }
  }

  return env;
}

export interface RichTextSegment {
  text: string;
  url?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  sproductthrough?: boolean;
  fontSize?: number;
  fontFamily?: string;
  foregroundColor?: string;
}

export interface RichTextCellDef {
  cell: string;
  segments: RichTextSegment[];
}

const cache = new PluginCache({
  namespace: "google-workspace-manager",
  defaultTTL: TTL.FIVE_MINUTES,
});

const GOOGLE_WORKSPACE_MCP_REQUEST_OPTIONS: RequestOptions = {
  timeout: 300_000,
  resetTimeoutOnProgress: true,
  maxTotalTimeout: 300_000,
};

export class GoogleWorkspaceMCPClient {
  private client: Client | null = null;
  private transport: StdioClientTransport | null = null;
  private config: MCPConfig;
  private connected: boolean = false;
  private cacheDisabled: boolean = false;
  private accountName?: string;
  private interactiveAuth: boolean = false;
  private requireDedicatedIndexingClient: boolean = false;
  private stderr: IOType;

  constructor(opts?: { config?: MCPConfig; stderr?: IOType }) {
    this.config =
      opts?.config ??
      loadServiceConfig("google-workspace-manager", {
        schema: GoogleWorkspaceConfigSchema,
      });
    this.stderr = opts?.stderr ?? "inherit";
  }


  disableCache(): void {
    this.cacheDisabled = true;
    cache.disable();
  }

  enableCache(): void {
    this.cacheDisabled = false;
    cache.enable();
  }

  getCacheStats() {
    return cache.getStats();
  }

  clearCache(): number {
    return cache.clear();
  }

  invalidateCacheKey(key: string): boolean {
    return cache.invalidate(key);
  }


  useAccount(
    accountName: string,
    requireDedicatedIndexingClient: boolean = false,
  ): void {
    if (this.connected) {
      throw new Error("OAuth account must be selected before the MCP connection opens");
    }
    if (this.accountName && this.accountName !== accountName) {
      throw new Error(
        `OAuth account already selected as '${this.accountName}'`,
      );
    }
    this.accountName = accountName;
    this.requireDedicatedIndexingClient ||= requireDedicatedIndexingClient;
  }

  useInteractiveAuth(requireDedicatedIndexingClient: boolean = false): void {
    if (this.connected) {
      throw new Error("Interactive auth must be selected before the MCP connection opens");
    }
    this.interactiveAuth = true;
    this.requireDedicatedIndexingClient ||= requireDedicatedIndexingClient;
  }

  async connect(): Promise<void> {
    if (this.connected) return;

    const mergedEnv = {
        ...process.env,
        ...this.config.mcpServer.env,
    };
    const accountSuffix = this.accountName?.trim().toUpperCase();
    const selectedAccountProfile = accountSuffix
      ? mergedEnv[`GOOGLE_MCP_ACCOUNT_PROFILE_${accountSuffix}`]?.trim().toLowerCase()
      : undefined;
    const forceReadOnlyIndexing =
      process.env.WORKSPACE_MCP_INDEXING_PROFILE === "readonly";
    const enforceReadOnlyIndexing =
      forceReadOnlyIndexing || selectedAccountProfile === "indexing_readonly";

    const env = resolveAccountOAuthEnvironment(
      mergedEnv,
      this.accountName,
      {
        interactiveAuth: this.interactiveAuth,
        requireDedicatedIndexingClient: this.requireDedicatedIndexingClient,
        forceReadOnlyIndexing,
      },
    );

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
    if (!env.GOOGLE_MCP_CREDENTIALS_DIR && env.GOOGLE_OAUTH_TOKEN) {
      env.GOOGLE_MCP_CREDENTIALS_DIR = dirname(env.GOOGLE_OAUTH_TOKEN);
    }
    if (!env.GOOGLE_MCP_CREDENTIALS_DIR) {
      throw new Error(
        "GOOGLE_MCP_CREDENTIALS_DIR environment variable is not set."
      );
    }

    const launch = resolveGoogleWorkspaceMcpLaunch(
      this.config.mcpServer.command,
      this.config.mcpServer.args,
      env,
      existsSync,
      enforceReadOnlyIndexing,
    );

    this.transport = new StdioClientTransport({
      command: launch.command,
      args: launch.args,
      cwd: launch.cwd,
      env: env as Record<string, string>,
      stderr: this.stderr,
    });

    this.client = new Client(
      { name: "google-workspace-cli", version: "1.0.0" },
      { capabilities: {} }
    );

    await this.client.connect(this.transport, GOOGLE_WORKSPACE_MCP_REQUEST_OPTIONS);
    this.connected = true;
  }

  async disconnect(): Promise<void> {
    if (this.client && this.connected) {
      await this.client.close();
      this.connected = false;
    }
  }


  async listTools(): Promise<any[]> {
    await this.connect();
    const result = await this.client!.listTools(undefined, GOOGLE_WORKSPACE_MCP_REQUEST_OPTIONS);
    return result.tools;
  }

  async callTool(name: string, args: Record<string, any> = {}): Promise<any> {
    await this.connect();

    if (this.config.userEmail && !args.user_google_email) {
      args.user_google_email = this.config.userEmail;
    }

    const result = await this.client!.callTool(
      { name, arguments: args },
      undefined,
      GOOGLE_WORKSPACE_MCP_REQUEST_OPTIONS
    );
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


  async searchGmailMessages(
    query: string,
    maxResults?: number,
    pageToken?: string,
    accountEmail?: string,
  ): Promise<any> {
    const cacheKey = createCacheKey("gmail_search", {
      query,
      maxResults,
      pageToken,
      account: accountEmail,
    });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { query };
        if (maxResults) args.page_size = maxResults;
        if (pageToken) args.page_token = pageToken;
        if (accountEmail) args.user_google_email = accountEmail;
        return this.callTool("search_gmail_messages", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async getGmailMessage(messageId: string, userEmail?: string): Promise<any> {
    const cacheKey = createCacheKey("gmail_message", { id: messageId, account: userEmail });
    return cache.getOrFetch(
      cacheKey,
      () => {
        const args: Record<string, any> = { message_id: messageId };
        if (userEmail) args.user_google_email = userEmail;
        return this.callTool("get_gmail_message_content", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async listGmailLabels(): Promise<any> {
    return cache.getOrFetch(
      "gmail_labels",
      () => this.callTool("list_gmail_labels", {}),
      { ttl: TTL.HOUR, bypassCache: this.cacheDisabled }
    );
  }

  async sendGmailMessage(
    to: string,
    subject: string,
    body: string,
    cc?: string,
    bcc?: string,
    accountEmail?: string
  ): Promise<any> {
    const args: Record<string, any> = { to, subject, body };
    if (cc) args.cc = cc;
    if (bcc) args.bcc = bcc;
    if (accountEmail) args.user_google_email = accountEmail;
    const result = await this.callTool("send_gmail_message", args);
    cache.invalidatePattern(/^gmail_search/);
    return result;
  }

  async createGmailDraft(to: string, subject: string, body: string, threadId?: string): Promise<any> {
    const params: Record<string, string> = { to, subject, body };
    if (threadId) params.thread_id = threadId;
    return this.callTool("draft_gmail_message", params);
  }

  async getGmailThread(threadId: string): Promise<any> {
    const cacheKey = createCacheKey("gmail_thread", { id: threadId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_gmail_thread_content", { thread_id: threadId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }


  async listCalendars(): Promise<any> {
    return cache.getOrFetch(
      "calendars",
      () => this.callTool("list_calendars", {}),
      { ttl: TTL.HOUR, bypassCache: this.cacheDisabled }
    );
  }

  async getEvents(options?: { calendarId?: string; eventId?: string; timeMin?: string; timeMax?: string; maxResults?: number }): Promise<any> {
    const cacheKey = createCacheKey("calendar_events", options || {});
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (options?.calendarId) args.calendar_id = options.calendarId;
        if (options?.eventId) args.event_id = options.eventId;
        if (options?.timeMin) args.time_min = options.timeMin;
        if (options?.timeMax) args.time_max = options.timeMax;
        if (options?.maxResults) args.max_results = options.maxResults;
        return this.callTool("get_events", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async createEvent(summary: string, start: string, end: string, options?: { description?: string; location?: string; attendees?: string; timezone?: string; calendarId?: string; sendUpdates?: string }): Promise<any> {
    const args: Record<string, any> = { action: "create", summary, start_time: start, end_time: end };
    if (options?.description) args.description = options.description;
    if (options?.location) args.location = options.location;
    if (options?.attendees) args.attendees = Array.isArray(options.attendees) ? options.attendees : options.attendees.split(",").map((s: string) => s.trim());
    if (options?.timezone) args.timezone = options.timezone;
    if (options?.calendarId) args.calendar_id = options.calendarId;
    if (options?.sendUpdates) args.send_updates = options.sendUpdates;
    const result = await this.callTool("manage_event", args);
    cache.invalidatePattern(/^calendar_events/);
    return result;
  }

  async deleteEvent(eventId: string, calendarId?: string): Promise<any> {
    const args: Record<string, any> = { action: "delete", event_id: eventId };
    if (calendarId) args.calendar_id = calendarId;
    const result = await this.callTool("manage_event", args);
    cache.invalidatePattern(/^calendar_events/);
    return result;
  }


  async searchDriveFiles(query: string, maxResults?: number): Promise<any> {
    const cacheKey = createCacheKey("drive_search", { query, maxResults });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { query };
        if (maxResults) args.page_size = maxResults;
        return this.callTool("search_drive_files", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async getDriveFileContent(fileId: string): Promise<any> {
    const cacheKey = createCacheKey("drive_file", { id: fileId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_drive_file_content", { file_id: fileId }),
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async getDriveFileRaw(
    fileId: string,
    exportFormat?: string,
    maxBytes?: number,
  ): Promise<DriveFileRawEnvelope> {
    const args: Record<string, unknown> = { file_id: fileId, raw_base64: true };
    if (exportFormat) args.export_format = exportFormat;
    if (maxBytes) args.max_bytes = maxBytes;
    return this.callTool("get_drive_file_content", args);
  }

  async listDriveItems(folderId?: string): Promise<any> {
    const cacheKey = createCacheKey("drive_items", { folder: folderId || "root" });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (folderId) args.folder_id = folderId;
        return this.callTool("list_drive_items", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }


  async searchDocs(query: string): Promise<any> {
    const cacheKey = createCacheKey("docs_search", { query });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("search_docs", { query }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async getDocContent(documentId: string, suggestionsViewMode?: string, userEmail?: string): Promise<any> {
    const cacheKey = createCacheKey("doc_content", { id: documentId, mode: suggestionsViewMode, account: userEmail });
    return cache.getOrFetch(
      cacheKey,
      () => {
        const args: Record<string, any> = { document_id: documentId };
        if (suggestionsViewMode) args.suggestions_view_mode = suggestionsViewMode;
        if (userEmail) args.user_google_email = userEmail;
        return this.callTool("get_doc_content", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async createDoc(title: string, content?: string): Promise<any> {
    const args: Record<string, any> = { title };
    if (content) args.content = content;
    const result = await this.callTool("create_doc", args);
    cache.invalidatePattern(/^docs_search/);
    return result;
  }

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


  async listSpreadsheets(accountEmail?: string): Promise<any> {
    const cacheKey = createCacheKey("spreadsheets_list", { account: accountEmail });
    return cache.getOrFetch(
      cacheKey,
      () => {
        const args: Record<string, any> = {};
        if (accountEmail) args.user_google_email = accountEmail;
        return this.callTool("list_spreadsheets", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async getSpreadsheetInfo(spreadsheetId: string, accountEmail?: string): Promise<any> {
    const cacheKey = createCacheKey("spreadsheet_info", { id: spreadsheetId, account: accountEmail });
    return cache.getOrFetch(
      cacheKey,
      () => {
        const args: Record<string, any> = { spreadsheet_id: spreadsheetId };
        if (accountEmail) args.user_google_email = accountEmail;
        return this.callTool("get_spreadsheet_info", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async readSheetValues(spreadsheetId: string, range: string, accountEmail?: string): Promise<any> {
    const cacheKey = createCacheKey("sheet_values", { id: spreadsheetId, range, account: accountEmail });
    return cache.getOrFetch(
      cacheKey,
      () => {
        const args: Record<string, any> = { spreadsheet_id: spreadsheetId, range_name: range };
        if (accountEmail) args.user_google_email = accountEmail;
        return this.callTool("read_sheet_values", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async writeSheetValues(spreadsheetId: string, range: string, values: any[][]): Promise<any> {
    const result = await this.callTool("modify_sheet_values", { spreadsheet_id: spreadsheetId, range_name: range, values });
    cache.invalidate(createCacheKey("sheet_values", { id: spreadsheetId, range }));
    return result;
  }

  async expandSheetGrid(
    spreadsheetId: string,
    sheetName: string,
    insertRows: number,
    insertColumns: number,
  ): Promise<unknown> {
    return this.callTool("resize_sheet_dimensions", {
      spreadsheet_id: spreadsheetId,
      sheet_name: sheetName,
      insert_rows: insertRows,
      insert_columns: insertColumns,
    });
  }

  async getForm(formId: string, accountEmail?: string): Promise<unknown> {
    const args: Record<string, unknown> = { form_id: formId };
    if (accountEmail) args.user_google_email = accountEmail;
    return this.callTool("get_form", args);
  }

  async getFormResponse(formId: string, responseId: string, accountEmail?: string): Promise<unknown> {
    const args: Record<string, unknown> = { form_id: formId, response_id: responseId };
    if (accountEmail) args.user_google_email = accountEmail;
    return this.callTool("get_form_response", args);
  }

  async listFormResponses(
    formId: string,
    pageSize?: number,
    pageToken?: string,
    accountEmail?: string,
  ): Promise<unknown> {
    const args: Record<string, unknown> = { form_id: formId };
    if (pageSize) args.page_size = pageSize;
    if (pageToken) args.page_token = pageToken;
    if (accountEmail) args.user_google_email = accountEmail;
    return this.callTool("list_form_responses", args);
  }

  async listDirectoryUsers(options: Record<string, unknown> = {}): Promise<unknown> {
    return this.callTool("list_directory_users", options);
  }

  async getDirectoryUser(userKey: string): Promise<unknown> {
    return this.callTool("get_directory_user", { user_key: userKey });
  }

  async listDirectoryUserAliases(userKey: string): Promise<unknown> {
    return this.callTool("list_directory_user_aliases", { user_key: userKey });
  }

  async listDirectoryGroups(
    options: Record<string, unknown> = {},
  ): Promise<unknown> {
    return this.callTool("list_directory_groups", options);
  }

  async getDirectoryGroup(groupKey: string): Promise<unknown> {
    return this.callTool("get_directory_group", { group_key: groupKey });
  }

  async listDirectoryGroupAliases(groupKey: string): Promise<unknown> {
    return this.callTool("list_directory_group_aliases", { group_key: groupKey });
  }

  async listDirectoryGroupMembers(
    groupKey: string,
    options: Record<string, unknown> = {},
  ): Promise<unknown> {
    return this.callTool("list_directory_group_members", {
      group_key: groupKey,
      ...options,
    });
  }

  async getDirectoryGroupMember(
    groupKey: string,
    memberKey: string,
  ): Promise<unknown> {
    return this.callTool("get_directory_group_member", {
      group_key: groupKey,
      member_key: memberKey,
    });
  }

  async getGroupSettings(groupEmail: string): Promise<unknown> {
    return this.callTool("get_group_settings", { group_email: groupEmail });
  }

  async createDirectoryGroup(
    groupEmail: string,
    name: string,
    description: string,
    dryRun: boolean,
    confirmation?: string,
  ): Promise<unknown> {
    return this.callTool("create_directory_group", {
      group_email: groupEmail,
      name,
      description,
      dry_run: dryRun,
      confirmation,
    });
  }

  async insertDirectoryGroupMember(
    groupKey: string,
    memberEmail: string,
    role: "OWNER" | "MANAGER" | "MEMBER",
    deliverySettings: "ALL_MAIL" | "DAILY" | "DIGEST" | "DISABLED" | "NONE",
    dryRun: boolean,
    confirmation?: string,
  ): Promise<unknown> {
    return this.callTool("insert_directory_group_member", {
      group_key: groupKey,
      member_email: memberEmail,
      role,
      delivery_settings: deliverySettings,
      dry_run: dryRun,
      confirmation,
    });
  }

  async patchGroupSettings(
    groupEmail: string,
    settings: Record<string, unknown>,
    dryRun: boolean,
    confirmation?: string,
  ): Promise<unknown> {
    return this.callTool("patch_group_settings", {
      group_email: groupEmail,
      settings,
      dry_run: dryRun,
      confirmation,
    });
  }

  async insertDirectoryUserAlias(
    userKey: string,
    alias: string,
    dryRun: boolean,
    confirmation?: string,
  ): Promise<unknown> {
    return this.callTool("insert_directory_user_alias", {
      user_key: userKey,
      alias,
      dry_run: dryRun,
      confirmation,
    });
  }

  async deleteDirectoryUserAlias(
    userKey: string,
    alias: string,
    dryRun: boolean,
    confirmation?: string,
  ): Promise<unknown> {
    return this.callTool("delete_directory_user_alias", {
      user_key: userKey,
      alias,
      dry_run: dryRun,
      confirmation,
    });
  }

  async listGmailSendAs(accountEmail?: string): Promise<unknown> {
    const args: Record<string, unknown> = {};
    if (accountEmail) args.user_google_email = accountEmail;
    return this.callTool("list_gmail_send_as", args);
  }

  async updateGmailSendAs(sendAsEmail: string, options: Record<string, unknown>): Promise<unknown> {
    return this.callTool("update_gmail_send_as", {
      send_as_email: sendAsEmail,
      ...options,
    });
  }

  async deleteGmailSendAs(
    sendAsEmail: string,
    dryRun: boolean,
    confirmation?: string,
  ): Promise<unknown> {
    return this.callTool("delete_gmail_send_as", {
      send_as_email: sendAsEmail,
      dry_run: dryRun,
      confirmation,
    });
  }

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

    cache.invalidatePattern(sheetValuesCachePattern(spreadsheetId));

    return result;
  }

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

    cache.invalidatePattern(sheetValuesCachePattern(spreadsheetId));

    return result;
  }

  async createSpreadsheet(title: string, sheetNames?: string[]): Promise<any> {
    const args: Record<string, any> = { title };
    if (sheetNames) args.sheet_names = sheetNames;
    return this.callTool("create_spreadsheet", args);
  }


  async listTaskLists(): Promise<any> {
    return cache.getOrFetch(
      "task_lists",
      () => this.callTool("list_task_lists", {}),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async listTasks(taskListId: string): Promise<any> {
    const cacheKey = createCacheKey("tasks", { listId: taskListId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("list_tasks", { task_list_id: taskListId }),
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async createTask(taskListId: string, title: string, notes?: string, due?: string): Promise<any> {
    const args: Record<string, any> = { task_list_id: taskListId, title };
    if (notes) args.notes = notes;
    if (due) args.due = due;
    const result = await this.callTool("create_task", args);
    cache.invalidate(createCacheKey("tasks", { listId: taskListId }));
    return result;
  }

  async completeTask(taskListId: string, taskId: string): Promise<any> {
    const result = await this.callTool("update_task", { task_list_id: taskListId, task_id: taskId, status: "completed" });
    cache.invalidate(createCacheKey("tasks", { listId: taskListId }));
    return result;
  }


  async getDocumentComments(documentId: string, userEmail?: string): Promise<any> {
    const cacheKey = createCacheKey("doc_comments", { id: documentId, account: userEmail });
    return cache.getOrFetch(
      cacheKey,
      () => {
        const args: Record<string, any> = { document_id: documentId };
        if (userEmail) args.user_google_email = userEmail;
        return this.callTool("read_document_comments", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async createDocumentComment(documentId: string, text: string, location?: Record<string, any>): Promise<any> {
    const args: Record<string, any> = { document_id: documentId, text };
    if (location) args.location = location;
    const result = await this.callTool("create_document_comment", args);
    cache.invalidate(createCacheKey("doc_comments", { id: documentId }));
    return result;
  }

  async replyToDocumentComment(documentId: string, commentId: string, text: string): Promise<any> {
    const result = await this.callTool("reply_to_document_comment", { document_id: documentId, comment_id: commentId, text });
    cache.invalidate(createCacheKey("doc_comments", { id: documentId }));
    return result;
  }

  async resolveDocumentComment(documentId: string, commentId: string): Promise<any> {
    const result = await this.callTool("resolve_document_comment", { document_id: documentId, comment_id: commentId });
    cache.invalidate(createCacheKey("doc_comments", { id: documentId }));
    return result;
  }


  async getSpreadsheetComments(spreadsheetId: string): Promise<any> {
    const cacheKey = createCacheKey("sheet_comments", { id: spreadsheetId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("read_spreadsheet_comments", { spreadsheet_id: spreadsheetId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async createSpreadsheetComment(spreadsheetId: string, sheetId: number, rowIndex: number, columnIndex: number, text: string): Promise<any> {
    const result = await this.callTool("create_spreadsheet_comment", { spreadsheet_id: spreadsheetId, sheet_id: sheetId, row_index: rowIndex, column_index: columnIndex, text });
    cache.invalidate(createCacheKey("sheet_comments", { id: spreadsheetId }));
    return result;
  }

  async replyToSpreadsheetComment(spreadsheetId: string, commentId: string, text: string): Promise<any> {
    const result = await this.callTool("reply_to_spreadsheet_comment", { spreadsheet_id: spreadsheetId, comment_id: commentId, text });
    cache.invalidate(createCacheKey("sheet_comments", { id: spreadsheetId }));
    return result;
  }

  async resolveSpreadsheetComment(spreadsheetId: string, commentId: string): Promise<any> {
    const result = await this.callTool("resolve_spreadsheet_comment", { spreadsheet_id: spreadsheetId, comment_id: commentId });
    cache.invalidate(createCacheKey("sheet_comments", { id: spreadsheetId }));
    return result;
  }


  async listGmailFilters(): Promise<any> {
    return cache.getOrFetch(
      "gmail_filters",
      () => this.callTool("list_gmail_filters", {}),
      { ttl: TTL.HOUR, bypassCache: this.cacheDisabled }
    );
  }

  async manageGmailFilter(
    action: "create" | "delete",
    options: { criteria?: Record<string, any>; filterAction?: Record<string, any>; filterId?: string }
  ): Promise<any> {
    const args: Record<string, any> = { action };
    if (options.criteria) args.criteria = options.criteria;
    if (options.filterAction) args.filter_action = options.filterAction;
    if (options.filterId) args.filter_id = options.filterId;
    const result = await this.callTool("manage_gmail_filter", args);
    cache.invalidate("gmail_filters");
    return result;
  }

  async manageGmailLabel(
    action: "create" | "update" | "delete",
    options: { name?: string; labelId?: string; labelListVisibility?: string; messageListVisibility?: string }
  ): Promise<any> {
    const args: Record<string, any> = { action };
    if (options.name) args.name = options.name;
    if (options.labelId) args.label_id = options.labelId;
    if (options.labelListVisibility) args.label_list_visibility = options.labelListVisibility;
    if (options.messageListVisibility) args.message_list_visibility = options.messageListVisibility;
    const result = await this.callTool("manage_gmail_label", args);
    cache.invalidate("gmail_labels");
    return result;
  }

  async modifyGmailMessageLabels(
    messageId: string,
    addLabelIds?: string[],
    removeLabelIds?: string[]
  ): Promise<any> {
    const args: Record<string, any> = { message_id: messageId };
    if (addLabelIds) args.add_label_ids = addLabelIds;
    if (removeLabelIds) args.remove_label_ids = removeLabelIds;
    const result = await this.callTool("modify_gmail_message_labels", args);
    cache.invalidatePattern(/^gmail_search/);
    return result;
  }

  async getGmailMessagesBatch(
    messageIds: string[],
    format: "full" | "metadata" = "full",
    bodyFormat?: "text" | "html" | "raw"
  ): Promise<any> {
    const cacheKey = createCacheKey("gmail_batch", { ids: messageIds.join(","), format, bodyFormat });
    const toolArgs: Record<string, any> = { message_ids: messageIds, format };
    if (bodyFormat) {
      toolArgs.body_format = bodyFormat;
    }
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_gmail_messages_content_batch", toolArgs),
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async getGmailMessageRaw(messageId: string, userEmail?: string): Promise<any> {
    const cacheKey = createCacheKey("gmail_raw_message", { id: messageId, userEmail });
    const toolArgs: Record<string, any> = {
      message_ids: [messageId],
      format: "full",
      body_format: "raw",
    };
    if (userEmail) {
      toolArgs.user_google_email = userEmail;
    }
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_gmail_messages_content_batch", toolArgs),
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }


  async listContacts(pageSize?: number, sortOrder?: string, pageToken?: string): Promise<any> {
    const cacheKey = createCacheKey("contacts_list", { pageSize, sortOrder, pageToken });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (pageSize) args.page_size = pageSize;
        if (sortOrder) args.sort_order = sortOrder;
        if (pageToken) args.page_token = pageToken;
        return this.callTool("list_contacts", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async getContact(contactId: string): Promise<any> {
    const cacheKey = createCacheKey("contact", { id: contactId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_contact", { contact_id: contactId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async searchContacts(query: string, pageSize?: number): Promise<any> {
    const cacheKey = createCacheKey("contacts_search", { query, pageSize });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { query };
        if (pageSize) args.page_size = pageSize;
        return this.callTool("search_contacts", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async manageContact(
    action: "create" | "update" | "delete",
    options: {
      contactId?: string; givenName?: string; familyName?: string;
      email?: string; phone?: string; organization?: string;
      jobTitle?: string; notes?: string;
    }
  ): Promise<any> {
    const args: Record<string, any> = { action };
    if (options.contactId) args.contact_id = options.contactId;
    if (options.givenName) args.given_name = options.givenName;
    if (options.familyName) args.family_name = options.familyName;
    if (options.email) args.email = options.email;
    if (options.phone) args.phone = options.phone;
    if (options.organization) args.organization = options.organization;
    if (options.jobTitle) args.job_title = options.jobTitle;
    if (options.notes) args.notes = options.notes;
    const result = await this.callTool("manage_contact", args);
    cache.invalidatePattern(/^contacts_/);
    cache.invalidatePattern(/^contact:/);
    return result;
  }

  async listContactGroups(pageSize?: number): Promise<any> {
    const cacheKey = createCacheKey("contact_groups", { pageSize });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (pageSize) args.page_size = pageSize;
        return this.callTool("list_contact_groups", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }


  async listChatSpaces(spaceType?: string, pageSize?: number): Promise<any> {
    const cacheKey = createCacheKey("chat_spaces", { spaceType, pageSize });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (spaceType) args.space_type = spaceType;
        if (pageSize) args.page_size = pageSize;
        return this.callTool("list_spaces", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async getChatMessages(spaceId: string, pageSize?: number, orderBy?: string): Promise<any> {
    const cacheKey = createCacheKey("chat_messages", { spaceId, pageSize, orderBy });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { space_id: spaceId };
        if (pageSize) args.page_size = pageSize;
        if (orderBy) args.order_by = orderBy;
        return this.callTool("get_messages", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async sendChatMessage(spaceId: string, text: string, threadName?: string, threadKey?: string): Promise<any> {
    const args: Record<string, any> = { space_id: spaceId, message_text: text };
    if (threadName) args.thread_name = threadName;
    if (threadKey) args.thread_key = threadKey;
    const result = await this.callTool("send_message", args);
    cache.invalidatePattern(/^chat_messages/);
    return result;
  }

  async searchChatMessages(query: string, spaceId?: string, pageSize?: number): Promise<any> {
    const cacheKey = createCacheKey("chat_search", { query, spaceId, pageSize });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { query };
        if (spaceId) args.space_id = spaceId;
        if (pageSize) args.page_size = pageSize;
        return this.callTool("search_messages", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }


  async copyDriveFile(fileId: string, newName?: string, parentFolderId?: string): Promise<any> {
    const args: Record<string, any> = { file_id: fileId };
    if (newName) args.new_name = newName;
    if (parentFolderId) args.parent_folder_id = parentFolderId;
    return this.callTool("copy_drive_file", args);
  }

  async createDriveFolder(folderName: string, parentFolderId?: string): Promise<any> {
    const args: Record<string, any> = { folder_name: folderName };
    if (parentFolderId) args.parent_folder_id = parentFolderId;
    const result = await this.callTool("create_drive_folder", args);
    this.invalidateDriveListings();
    return result;
  }

  private invalidateDriveListings(): void {
    cache.invalidatePattern(/^drive_search/);
    cache.invalidatePattern(/^drive_items/);
  }

  async createDriveFile(options: {
    fileName: string;
    content?: string;
    fileUrl?: string;
    folderId?: string;
    mimeType?: string;
    accountEmail?: string;
  }): Promise<unknown> {
    const args: Record<string, unknown> = { file_name: options.fileName };
    if (options.content !== undefined) args.content = options.content;
    if (options.fileUrl !== undefined) args.fileUrl = options.fileUrl;
    if (options.folderId !== undefined) args.folder_id = options.folderId;
    if (options.mimeType !== undefined) args.mime_type = options.mimeType;
    if (options.accountEmail !== undefined) args.user_google_email = options.accountEmail;
    const result = await this.callTool("create_drive_file", args);
    this.invalidateDriveListings();
    return result;
  }

  async trashDriveFile(fileId: string, accountEmail?: string): Promise<unknown> {
    const args: Record<string, unknown> = {
      file_id: fileId,
      trashed: true,
    };
    if (accountEmail !== undefined) args.user_google_email = accountEmail;
    const result = await this.callTool("update_drive_file", args);
    cache.invalidatePattern(/^drive_/);
    return result;
  }

  async getDriveShareLink(fileId: string): Promise<any> {
    const cacheKey = createCacheKey("drive_share_link", { id: fileId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_drive_shareable_link", { file_id: fileId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async manageDriveAccess(
    fileId: string,
    action: "grant" | "grant_batch" | "update" | "revoke" | "transfer_owner",
    options: {
      shareWith?: string; role?: string; shareType?: string;
      permissionId?: string; recipients?: unknown[];
      sendNotification?: boolean; emailMessage?: string;
      expirationTime?: string; newOwnerEmail?: string;
    } = {}
  ): Promise<any> {
    const args: Record<string, any> = { file_id: fileId, action };
    if (options.shareWith) args.share_with = options.shareWith;
    if (options.role) args.role = options.role;
    if (options.shareType) args.share_type = options.shareType;
    if (options.permissionId) args.permission_id = options.permissionId;
    if (options.recipients) args.recipients = options.recipients;
    if (options.sendNotification !== undefined) args.send_notification = options.sendNotification;
    if (options.emailMessage) args.email_message = options.emailMessage;
    if (options.expirationTime) args.expiration_time = options.expirationTime;
    if (options.newOwnerEmail) args.new_owner_email = options.newOwnerEmail;
    const result = await this.callTool("manage_drive_access", args);
    cache.invalidatePattern(/^drive_/);
    return result;
  }

  async getDriveFilePermissions(fileId: string): Promise<any> {
    const cacheKey = createCacheKey("drive_permissions", { id: fileId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_drive_file_permissions", { file_id: fileId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }


  async manageEvent(
    action: "create" | "update" | "delete",
    options: {
      summary?: string; startTime?: string; endTime?: string;
      eventId?: string; calendarId?: string; description?: string;
      location?: string; attendees?: unknown; timezone?: string;
      addGoogleMeet?: boolean; transparency?: string; visibility?: string;
      sendUpdates?: string;
    } = {}
  ): Promise<any> {
    const args: Record<string, any> = { action };
    if (options.summary) args.summary = options.summary;
    if (options.startTime) args.start_time = options.startTime;
    if (options.endTime) args.end_time = options.endTime;
    if (options.eventId) args.event_id = options.eventId;
    if (options.calendarId) args.calendar_id = options.calendarId;
    if (options.description) args.description = options.description;
    if (options.location) args.location = options.location;
    if (options.attendees) args.attendees = options.attendees;
    if (options.timezone) args.timezone = options.timezone;
    if (options.addGoogleMeet !== undefined) args.add_google_meet = options.addGoogleMeet;
    if (options.transparency) args.transparency = options.transparency;
    if (options.visibility) args.visibility = options.visibility;
    if (options.sendUpdates) args.send_updates = options.sendUpdates;
    const result = await this.callTool("manage_event", args);
    cache.invalidatePattern(/^calendar_events/);
    return result;
  }

  async queryFreebusy(timeMin: string, timeMax: string, calendarIds?: string[]): Promise<any> {
    const cacheKey = createCacheKey("freebusy", { timeMin, timeMax, cals: calendarIds?.join(",") });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { time_min: timeMin, time_max: timeMax };
        if (calendarIds) args.calendar_ids = calendarIds;
        return this.callTool("query_freebusy", args);
      },
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }


  async exportDocToPdf(documentId: string, pdfFilename?: string, folderId?: string): Promise<any> {
    const args: Record<string, any> = { document_id: documentId };
    if (pdfFilename) args.pdf_filename = pdfFilename;
    if (folderId) args.folder_id = folderId;
    return this.callTool("export_doc_to_pdf", args);
  }

  async getDocAsMarkdown(
    documentId: string,
    options: {
      includeComments?: boolean;
      commentMode?: string;
      includeResolved?: boolean;
      suggestionsViewMode?: string;
    } = {}
  ): Promise<any> {
    const cacheKey = createCacheKey("doc_markdown", { id: documentId, ...options });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = { document_id: documentId };
        if (options.includeComments !== undefined) args.include_comments = options.includeComments;
        if (options.commentMode) args.comment_mode = options.commentMode;
        if (options.includeResolved !== undefined) args.include_resolved = options.includeResolved;
        if (options.suggestionsViewMode) args.suggestions_view_mode = options.suggestionsViewMode;
        return this.callTool("get_doc_as_markdown", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async listDocsInFolder(folderId?: string, pageSize?: number): Promise<any> {
    const cacheKey = createCacheKey("docs_in_folder", { folder: folderId || "root", pageSize });
    return cache.getOrFetch(
      cacheKey,
      async () => {
        const args: Record<string, any> = {};
        if (folderId) args.folder_id = folderId;
        if (pageSize) args.page_size = pageSize;
        return this.callTool("list_docs_in_folder", args);
      },
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }


  async formatSheetRange(
    spreadsheetId: string,
    rangeName: string,
    options: {
      backgroundColor?: string; textColor?: string;
      numberFormatType?: string; numberFormatPattern?: string;
      wrapStrategy?: string; horizontalAlignment?: string;
      verticalAlignment?: string; bold?: boolean;
      italic?: boolean; fontSize?: number;
    }
  ): Promise<any> {
    const args: Record<string, any> = { spreadsheet_id: spreadsheetId, range_name: rangeName };
    if (options.backgroundColor) args.background_color = options.backgroundColor;
    if (options.textColor) args.text_color = options.textColor;
    if (options.numberFormatType) args.number_format_type = options.numberFormatType;
    if (options.numberFormatPattern) args.number_format_pattern = options.numberFormatPattern;
    if (options.wrapStrategy) args.wrap_strategy = options.wrapStrategy;
    if (options.horizontalAlignment) args.horizontal_alignment = options.horizontalAlignment;
    if (options.verticalAlignment) args.vertical_alignment = options.verticalAlignment;
    if (options.bold !== undefined) args.bold = options.bold;
    if (options.italic !== undefined) args.italic = options.italic;
    if (options.fontSize) args.font_size = options.fontSize;
    const result = await this.callTool("format_sheet_range", args);
    cache.invalidatePattern(sheetValuesCachePattern(spreadsheetId));
    return result;
  }

  async createSheet(spreadsheetId: string, sheetName: string): Promise<any> {
    const result = await this.callTool("create_sheet", { spreadsheet_id: spreadsheetId, sheet_name: sheetName });
    cache.invalidate(createCacheKey("spreadsheet_info", { id: spreadsheetId }));
    return result;
  }


  async getTask(taskListId: string, taskId: string): Promise<any> {
    const cacheKey = createCacheKey("task", { listId: taskListId, id: taskId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("get_task", { task_list_id: taskListId, task_id: taskId }),
      { ttl: TTL.FIVE_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async manageTask(
    action: "create" | "update" | "delete" | "move",
    taskListId: string,
    options: {
      taskId?: string; title?: string; notes?: string;
      status?: string; due?: string; parent?: string;
      previous?: string; destinationTaskList?: string;
    } = {}
  ): Promise<any> {
    const args: Record<string, any> = { action, task_list_id: taskListId };
    if (options.taskId) args.task_id = options.taskId;
    if (options.title) args.title = options.title;
    if (options.notes) args.notes = options.notes;
    if (options.status) args.status = options.status;
    if (options.due) args.due = options.due;
    if (options.parent) args.parent = options.parent;
    if (options.previous) args.previous = options.previous;
    if (options.destinationTaskList) args.destination_task_list = options.destinationTaskList;
    const result = await this.callTool("manage_task", args);
    cache.invalidatePattern(/^tasks:/);
    cache.invalidatePattern(/^task:/);
    return result;
  }


  async getPresentationComments(presentationId: string): Promise<any> {
    const cacheKey = createCacheKey("presentation_comments", { id: presentationId });
    return cache.getOrFetch(
      cacheKey,
      () => this.callTool("read_presentation_comments", { presentation_id: presentationId }),
      { ttl: TTL.FIFTEEN_MINUTES, bypassCache: this.cacheDisabled }
    );
  }

  async createPresentationComment(presentationId: string, slideId: string, text: string, location?: Record<string, any>): Promise<any> {
    const args: Record<string, any> = { presentation_id: presentationId, slide_id: slideId, text };
    if (location) args.location = location;
    const result = await this.callTool("create_presentation_comment", args);
    cache.invalidate(createCacheKey("presentation_comments", { id: presentationId }));
    return result;
  }

  async replyToPresentationComment(presentationId: string, commentId: string, text: string): Promise<any> {
    const result = await this.callTool("reply_to_presentation_comment", { presentation_id: presentationId, comment_id: commentId, text });
    cache.invalidate(createCacheKey("presentation_comments", { id: presentationId }));
    return result;
  }

  async resolvePresentationComment(presentationId: string, commentId: string): Promise<any> {
    const result = await this.callTool("resolve_presentation_comment", { presentation_id: presentationId, comment_id: commentId });
    cache.invalidate(createCacheKey("presentation_comments", { id: presentationId }));
    return result;
  }
}

export default GoogleWorkspaceMCPClient;

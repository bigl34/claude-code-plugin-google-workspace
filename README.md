<!-- AUTO-GENERATED README — DO NOT EDIT. Changes will be overwritten on next publish. -->
# claude-code-plugin-google-workspace

Business Google Workspace (Gmail, Calendar, Drive, Docs, Sheets, Tasks, Contacts, Chat, Filters, Comments)

![Version](https://img.shields.io/badge/version-1.12.0-blue) ![License: MIT](https://img.shields.io/badge/License-MIT-green) ![Node >= 18](https://img.shields.io/badge/node-%3E%3D18-brightgreen)

## Features

- Gmail
- **search-gmail** — READ
- **get-gmail-message** — READ
- **get-gmail-message-raw** — READ
- **get-gmail-thread** — READ
- **get-gmail-messages-batch** — READ
- **list-gmail-labels** — READ
- **list-gmail-filters** — READ
- **list-gmail-send-as** — READ
- **send-gmail** — WRITE
- **create-gmail-draft** — WRITE
- **create-gmail-filter** — WRITE
- **delete-gmail-filter** — WRITE
- **manage-gmail-label** — WRITE
- **modify-message-labels** — WRITE
- **update-gmail-send-as** — WRITE
- **delete-gmail-send-as** — WRITE
- Calendar
- **list-calendars** — READ
- **get-events** — READ
- **create-event** — WRITE
- **modify-event** — WRITE
- **delete-event** — WRITE
- **query-freebusy** — READ
- Drive
- **search-drive** — READ
- **get-drive-content** — READ
- **download-drive-file** — WRITE
- **list-drive-items** — READ
- **get-drive-share-link** — READ
- **get-drive-permissions** — READ
- **copy-drive-file** — WRITE
- **create-drive-folder** — WRITE
- **create-drive-file** — WRITE
- **share-drive-file** — WRITE
- Docs
- **search-docs** — READ
- **get-doc-content** — READ
- **get-doc-markdown** — READ
- **list-docs-in-folder** — READ
- **export-doc-pdf** — READ
- **create-doc** — WRITE
- **modify-doc-text** — WRITE
- **find-replace-doc** — WRITE
- Sheets
- **list-spreadsheets** — READ
- **get-spreadsheet-info** — READ
- **check-sheet-range** — READ
- **read-sheet** — READ
- **create-spreadsheet** — WRITE
- **write-sheet** — WRITE
- **write-rich-text** — WRITE
- **write-rich-text-batch** — WRITE
- **format-sheet-range** — WRITE
- **add-sheet** — WRITE
- Forms
- **get-form** — READ
- **get-form-response** — READ
- **list-form-responses** — READ
- Admin Directory
- **list-directory-users** — READ
- **get-directory-user** — READ
- **list-directory-user-aliases** — READ
- **preview-directory-user-alias-move** — READ
- **list-directory-groups** — READ
- **get-directory-group** — READ
- **list-directory-group-aliases** — READ
- **list-directory-group-members** — READ
- **get-directory-group-member** — READ
- **get-group-settings** — READ
- **create-directory-group** — WRITE
- **insert-directory-group-member** — WRITE
- **patch-group-settings** — WRITE
- **insert-directory-user-alias** — WRITE
- **delete-directory-user-alias** — WRITE
- Tasks
- **list-task-lists** — READ
- **list-tasks** — READ
- **get-task** — READ
- **create-task** — WRITE
- **complete-task** — WRITE
- **update-task** — WRITE
- **delete-task** — WRITE
- **move-task** — WRITE
- Contacts
- **list-contacts** — READ
- **get-contact** — READ
- **search-contacts** — READ
- **list-contact-groups** — READ
- **create-contact** — WRITE
- **update-contact** — WRITE
- **delete-contact** — WRITE
- Chat
- **list-chat-spaces** — READ
- **get-chat-messages** — READ
- **send-chat-message** — WRITE
- **search-chat-messages** — READ
- Comments (Docs, Sheets, Slides)
- **get-doc-comments** — READ
- **create-doc-comment** — WRITE
- **reply-doc-comment** — WRITE
- **resolve-doc-comment** — WRITE
- **get-sheet-comments** — READ
- **create-sheet-comment** — WRITE
- **reply-sheet-comment** — WRITE
- **resolve-sheet-comment** — WRITE
- **get-presentation-comments** — READ
- **create-presentation-comment** — WRITE
- **reply-presentation-comment** — WRITE
- **resolve-presentation-comment** — WRITE
- Utility
- **list-tools** — READ
- **cache-stats** — READ
- **cache-clear** — WRITE
- **cache-invalidate** — WRITE

## Prerequisites

- [Node.js](https://nodejs.org/) >= 18
- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) CLI
- MCP server binary for the target service (configured via `config.json`)

## Quick Start

```bash
git clone https://github.com/bigl34/claude-code-plugin-google-workspace.git
cd claude-code-plugin-google-workspace
cp config.template.json config.json  # fill in your credentials
npm --prefix scripts install
```

```bash
npm --prefix scripts run cli -- search-gmail
```

## Installation

1. Clone this repository
2. Copy `config.template.json` to `config.json` and fill in your credentials
3. Install dependencies:
   ```bash
   cd scripts && npm install
   ```
4. Ensure the MCP server binary is available on your system (see the service's documentation)

## Available Commands

### Gmail

| Command                    | Type  | Description                                                           | Key Options                                                                   |
| -------------------------- | ----- | --------------------------------------------------------------------- | ----------------------------------------------------------------------------- |
| `search-gmail`             | READ  | Search emails with pagination                                         | `--query`, `--limit`, `--pageToken`, `--account`                              |
| `get-gmail-message`        | READ  | Get full email                                                        | `--id`                                                                        |
| `get-gmail-message-raw`    | READ  | Get decoded raw RFC822/MIME email with headers for exact alias checks | `--id`, `--account`                                                           |
| `get-gmail-thread`         | READ  | Get full thread                                                       | `--id`                                                                        |
| `get-gmail-messages-batch` | READ  | Batch get emails                                                      | `--ids` (comma-separated, max 25), `--format`, `--bodyFormat`                 |
| `list-gmail-labels`        | READ  | List labels                                                           |                                                                               |
| `list-gmail-filters`       | READ  | List filters                                                          |                                                                               |
| `list-gmail-send-as`       | READ  | List send-as identities                                               | `--account`                                                                   |
| `send-gmail`               | WRITE | Send email                                                            | `--to`, `--subject`, `--body`, `--cc`, `--bcc`                                |
| `create-gmail-draft`       | WRITE | Create draft                                                          | `--to`, `--subject`, `--body`, `--threadId`                                   |
| `create-gmail-filter`      | WRITE | Create filter                                                         | `--criteria` (JSON), `--action` (JSON)                                        |
| `delete-gmail-filter`      | WRITE | Delete filter                                                         | `--id`                                                                        |
| `manage-gmail-label`       | WRITE | Create/update/delete label                                            | `--action`, `--name`, `--labelId`                                             |
| `modify-message-labels`    | WRITE | Add/remove labels                                                     | `--id`, `--add`, `--remove` (comma-separated label IDs)                       |
| `update-gmail-send-as`     | WRITE | Preview/update send-as settings                                       | `--email`, writable fields, `--dryRun`, exact `--confirmation UPDATE:<email>` |
| `delete-gmail-send-as`     | WRITE | Preview/delete a non-primary identity                                 | `--email`, `--dryRun`, exact `--confirmation DELETE:<email>`                  |

### Calendar

| Command          | Type  | Description          | Key Options                                                                                                                        |
| ---------------- | ----- | -------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| `list-calendars` | READ  | List all calendars   |                                                                                                                                    |
| `get-events`     | READ  | Get events           | `--calendarId`, `--timeMin`, `--timeMax`, `--eventId`, `--limit`                                                                   |
| `create-event`   | WRITE | Create event         | `--summary`, `--start`, `--end` (ISO 8601 with offset), `--description`, `--location`, `--attendees`, `--timezone`, `--calendarId` |
| `modify-event`   | WRITE | Create/update/delete | `--action`, `--eventId`, `--summary`, `--start`, `--end`, `--calendarId`, `--addGoogleMeet`, `--transparency`, `--visibility`      |
| `delete-event`   | WRITE | Delete event         | `--id`, `--calendarId`                                                                                                             |
| `query-freebusy` | READ  | Check free/busy      | `--timeMin`, `--timeMax`, `--calendarIds` (comma-separated)                                                                        |

### Drive

| Command                 | Type  | Description                                                              | Key Options                                                                                                                                                                                                              |
| ----------------------- | ----- | ------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `search-drive`          | READ  | Search files                                                             | `--query`, `--limit`                                                                                                                                                                                                     |
| `get-drive-content`     | READ  | Get file content                                                         | `--id`                                                                                                                                                                                                                   |
| `download-drive-file`   | WRITE | Preview/download bounded binary/export content to an explicit local path | `--id`, `--outputFile`, `--exportFormat`, `--maxBytes`, `--overwrite`, `--dryRun`, exact `--confirmation DOWNLOAD:<id>:<absolute-path>`                                                                                  |
| `list-drive-items`      | READ  | List folder items                                                        | `--folder-id`                                                                                                                                                                                                            |
| `get-drive-share-link`  | READ  | Get shareable link                                                       | `--id`                                                                                                                                                                                                                   |
| `get-drive-permissions` | READ  | Get permissions                                                          | `--id`                                                                                                                                                                                                                   |
| `copy-drive-file`       | WRITE | Copy a file                                                              | `--id`, `--name`, `--parentId`                                                                                                                                                                                           |
| `create-drive-folder`   | WRITE | Create folder                                                            | `--name`, `--parentId`                                                                                                                                                                                                   |
| `create-drive-file`     | WRITE | Create a Drive file from inline text or a source URL                     | `--name`, exactly one of `--content` / `--file-url`, `--folder-id`, `--mime-type`. Local `file://` sources must sit under the connector attachment root `~/.workspace-mcp/attachments` unless `ALLOWED_FILE_DIRS` is set |
| `share-drive-file`      | WRITE | Share/unshare                                                            | `--id`, `--action` (grant/revoke/update/transfer_owner), `--shareWith`, `--role`, `--shareType`                                                                                                                          |

### Docs

| Command               | Type  | Description           | Key Options                                                                                                                          |
| --------------------- | ----- | --------------------- | ------------------------------------------------------------------------------------------------------------------------------------ |
| `search-docs`         | READ  | Search documents      | `--query`                                                                                                                            |
| `get-doc-content`     | READ  | Get content           | `--id`, `--suggestionsMode`                                                                                                          |
| `get-doc-markdown`    | READ  | Get as Markdown       | `--id`, `--includeComments`, `--commentMode` (inline/appendix/none)                                                                  |
| `list-docs-in-folder` | READ  | List Docs in folder   | `--folder-id`, `--limit`                                                                                                             |
| `export-doc-pdf`      | READ  | Export to PDF         | `--id`, `--filename`, `--folder-id`                                                                                                  |
| `create-doc`          | WRITE | Create document       | `--title`, `--content`                                                                                                               |
| `modify-doc-text`     | WRITE | Insert/replace/delete | `--id`, `--operation` (insert/replace/delete), `--text`, `--index`, `--startIndex`, `--endIndex`, `--bold`, `--italic`, `--fontSize` |
| `find-replace-doc`    | WRITE | Find and replace      | `--id`, `--find`, `--replace`, `--replaceAll`                                                                                        |

### Sheets

| Command                 | Type  | Description                                                                     | Key Options                                                                                            |
| ----------------------- | ----- | ------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| `list-spreadsheets`     | READ  | List spreadsheets                                                               |                                                                                                        |
| `get-spreadsheet-info`  | READ  | Get metadata                                                                    | `--id`                                                                                                 |
| `check-sheet-range`     | READ  | Normalize A1 range, detect DATA_SOURCE sheets, and report grid expansion needed | `--id`, `--range`                                                                                      |
| `read-sheet`            | READ  | Read values                                                                     | `--id`, `--range` (e.g. A1:D10)                                                                        |
| `create-spreadsheet`    | WRITE | Create spreadsheet                                                              | `--title`, `--sheetNames` (JSON array)                                                                 |
| `write-sheet`           | WRITE | Write values after sheet-type/grid checks                                       | `--id`, `--range`, `--values` (JSON array), `--expandGrid`                                             |
| `write-rich-text`       | WRITE | Write rich text cell                                                            | `--id`, `--cell`, `--segments` (JSON), `--sheetName`                                                   |
| `write-rich-text-batch` | WRITE | Batch rich text                                                                 | `--id`, `--cells` (JSON), `--sheetName`                                                                |
| `format-sheet-range`    | WRITE | Format range                                                                    | `--id`, `--range`, `--backgroundColor`, `--textColor`, `--bold`, `--fontSize`, `--horizontalAlignment` |
| `add-sheet`             | WRITE | Add sheet tab                                                                   | `--id`, `--name`                                                                                       |

### Forms

| Command               | Type | Description                     | Key Options                                       |
| --------------------- | ---- | ------------------------------- | ------------------------------------------------- |
| `get-form`            | READ | Get form metadata and questions | `--id`, `--account`                               |
| `get-form-response`   | READ | Get one response                | `--formId`, `--responseId`, `--account`           |
| `list-form-responses` | READ | List responses with pagination  | `--formId`, `--limit`, `--pageToken`, `--account` |

### Admin Directory

| Command                             | Type  | Description                                                            | Key Options                                                                                                       |
| ----------------------------------- | ----- | ---------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------- |
| `list-directory-users`              | READ  | List users with pagination                                             | `--customer`, `--domain`, `--query`, `--limit`, `--pageToken`                                                     |
| `get-directory-user`                | READ  | Get a user by ID/email/alias                                           | `--userKey`                                                                                                       |
| `list-directory-user-aliases`       | READ  | List aliases                                                           | `--userKey`                                                                                                       |
| `preview-directory-user-alias-move` | READ  | Validate a proposed user-alias reassignment without changing Workspace | `--source-user-key`, `--target-user-key`, `--alias`                                                               |
| `list-directory-groups`             | READ  | List groups with pagination                                            | `--customer`, `--domain`, `--query`, `--limit`, `--pageToken`                                                     |
| `get-directory-group`               | READ  | Get a group by ID/email/alias                                          | `--groupKey`                                                                                                      |
| `list-directory-group-aliases`      | READ  | List group aliases                                                     | `--groupKey`                                                                                                      |
| `list-directory-group-members`      | READ  | List group members and roles                                           | `--groupKey`, `--limit`, `--pageToken`, `--includeDerivedMembership`, `--roles`                                   |
| `get-directory-group-member`        | READ  | Get a member and delivery subscription                                 | `--groupKey`, `--memberKey`                                                                                       |
| `get-group-settings`                | READ  | Get complete Google Groups settings                                    | `--groupEmail`                                                                                                    |
| `create-directory-group`            | WRITE | Preview/create a group                                                 | `--groupEmail`, `--name`, `--description`, `--dryRun`, exact `--confirmation CREATE_GROUP:<groupEmail>`           |
| `insert-directory-group-member`     | WRITE | Preview/add member with role and delivery                              | `--groupKey`, `--memberEmail`, `--role`, `--deliverySettings`, `--dryRun`, exact confirmation returned by preview |
| `patch-group-settings`              | WRITE | Preview/patch allowlisted group settings                               | `--groupEmail`, `--settings` JSON, `--dryRun`, exact payload-bound confirmation returned by preview               |
| `insert-directory-user-alias`       | WRITE | Preview/insert alias                                                   | `--userKey`, `--alias`, `--dryRun`, exact `--confirmation INSERT:<userKey>:<alias>`                               |
| `delete-directory-user-alias`       | WRITE | Preview/delete alias                                                   | `--userKey`, `--alias`, `--dryRun`, exact `--confirmation DELETE:<userKey>:<alias>`                               |

### Tasks

| Command           | Type  | Description      | Key Options                                                   |
| ----------------- | ----- | ---------------- | ------------------------------------------------------------- |
| `list-task-lists` | READ  | List task lists  |                                                               |
| `list-tasks`      | READ  | List tasks       | `--id` (task list ID)                                         |
| `get-task`        | READ  | Get task details | `--listId`, `--id`                                            |
| `create-task`     | WRITE | Create task      | `--listId`, `--title`, `--notes`, `--due`                     |
| `complete-task`   | WRITE | Mark complete    | `--listId`, `--id`                                            |
| `update-task`     | WRITE | Update task      | `--listId`, `--id`, `--title`, `--notes`, `--status`, `--due` |
| `delete-task`     | WRITE | Delete task      | `--listId`, `--id`                                            |
| `move-task`       | WRITE | Move task        | `--listId`, `--id`, `--destinationListId`, `--parent`         |

### Contacts

| Command               | Type  | Description    | Key Options                                                                         |
| --------------------- | ----- | -------------- | ----------------------------------------------------------------------------------- |
| `list-contacts`       | READ  | List contacts  | `--limit`, `--sortOrder`                                                            |
| `get-contact`         | READ  | Get details    | `--id`                                                                              |
| `search-contacts`     | READ  | Search         | `--query`, `--limit`                                                                |
| `list-contact-groups` | READ  | List groups    | `--limit`                                                                           |
| `create-contact`      | WRITE | Create contact | `--givenName`, `--familyName`, `--email`, `--phone`, `--organization`, `--jobTitle` |
| `update-contact`      | WRITE | Update contact | `--id`, `--givenName`, `--familyName`, `--email`, `--phone`                         |
| `delete-contact`      | WRITE | Delete contact | `--id`                                                                              |

### Chat

| Command                | Type  | Description     | Key Options                                          |
| ---------------------- | ----- | --------------- | ---------------------------------------------------- |
| `list-chat-spaces`     | READ  | List spaces     | `--type` (all/room/dm), `--limit`                    |
| `get-chat-messages`    | READ  | Get messages    | `--spaceId`, `--limit`, `--orderBy`                  |
| `send-chat-message`    | WRITE | Send message    | `--spaceId`, `--text`, `--threadName`, `--threadKey` |
| `search-chat-messages` | READ  | Search messages | `--query`, `--spaceId`, `--limit`                    |

### Comments (Docs, Sheets, Slides)

| Command                        | Type  | Description    | Key Options                                                  |
| ------------------------------ | ----- | -------------- | ------------------------------------------------------------ |
| `get-doc-comments`             | READ  | Doc comments   | `--id`                                                       |
| `create-doc-comment`           | WRITE | Add comment    | `--id`, `--text`, `--location` (JSON)                        |
| `reply-doc-comment`            | WRITE | Reply          | `--id`, `--commentId`, `--text`                              |
| `resolve-doc-comment`          | WRITE | Resolve        | `--id`, `--commentId`                                        |
| `get-sheet-comments`           | READ  | Sheet comments | `--id`                                                       |
| `create-sheet-comment`         | WRITE | Add comment    | `--id`, `--sheetId`, `--rowIndex`, `--columnIndex`, `--text` |
| `reply-sheet-comment`          | WRITE | Reply          | `--id`, `--commentId`, `--text`                              |
| `resolve-sheet-comment`        | WRITE | Resolve        | `--id`, `--commentId`                                        |
| `get-presentation-comments`    | READ  | Slide comments | `--id`                                                       |
| `create-presentation-comment`  | WRITE | Add comment    | `--id`, `--slideId`, `--text`                                |
| `reply-presentation-comment`   | WRITE | Reply          | `--id`, `--commentId`, `--text`                              |
| `resolve-presentation-comment` | WRITE | Resolve        | `--id`, `--commentId`                                        |

### Utility

| Command            | Type  | Description              |
| ------------------ | ----- | ------------------------ |
| `list-tools`       | READ  | List available MCP tools |
| `cache-stats`      | READ  | Cache statistics         |
| `cache-clear`      | WRITE | Clear cache              |
| `cache-invalidate` | WRITE | Invalidate pattern       |

### Global Options

| Option       | Description                |
| ------------ | -------------------------- |
| `--no-cache` | Bypass cache for this call |
| `--help`     | Show help                  |

## Usage Examples

```bash
# Search Gmail
npm --prefix "scripts" run cli -- search-gmail --query "from:supplier@example.com" --limit 10

# Get calendar events for this week
npm --prefix "scripts" run cli -- get-events --timeMin "2026-04-01T00:00:00Z" --timeMax "2026-04-07T23:59:59Z"

# Read spreadsheet data
npm --prefix "scripts" run cli -- read-sheet --id "SPREADSHEET_ID" --range "Sheet1!A1:D10"

# Search contacts
npm --prefix "scripts" run cli -- search-contacts --query "john"
```

## How It Works

This plugin wraps an MCP (Model Context Protocol) server, providing a CLI interface that communicates with the service's MCP binary. The CLI translates commands into MCP tool calls and returns structured JSON responses.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Authentication errors | Verify credentials in `config.json` |
| `ERR_MODULE_NOT_FOUND` | Run `cd scripts && npm install` |
| MCP connection timeout | Ensure the MCP server binary is installed and accessible |
| Rate limiting | The CLI handles retries automatically; wait and retry if persistent |
| Unexpected JSON output | Check API credentials haven't expired |

## Contributing

Issues and pull requests are welcome.

## License

MIT

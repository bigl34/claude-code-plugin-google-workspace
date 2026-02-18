<!-- AUTO-GENERATED README — DO NOT EDIT. Changes will be overwritten on next publish. -->
# claude-code-plugin-google-workspace

Business Google Workspace (Gmail, Calendar, Drive, Docs, Sheets, Rich Text, Comments, Suggestions)

![Version](https://img.shields.io/badge/version-1.7.1-blue) ![License: MIT](https://img.shields.io/badge/License-MIT-green) ![Node >= 18](https://img.shields.io/badge/node-%3E%3D18-brightgreen)

## Features

- **search-gmail** — READ
- **get-gmail-message** — READ
- **get-gmail-thread** — READ
- **list-gmail-labels** — READ
- **send-gmail** — ⚠️ WRITE
- **create-gmail-draft** — ⚠️ WRITE

## Prerequisites

- [Node.js](https://nodejs.org/) >= 18
- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) CLI
- MCP server binary for the target service (configured via `config.json`)

## Quick Start

```bash
git clone https://github.com/YOUR_GITHUB_USER/claude-code-plugin-google-workspace.git
cd claude-code-plugin-google-workspace
cp config.template.json config.json  # fill in your credentials
cd scripts && npm install
```

```bash
node scripts/dist/cli.js search-gmail
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

| Command              | Type     | Description           | Required Options        |
| -------------------- | -------- | --------------------- | ----------------------- |
| `search-gmail`       | READ     | Search emails         | `--query`               |
| `get-gmail-message`  | READ     | Get email details     | `--id`                  |
| `get-gmail-thread`   | READ     | Get full email thread | `--id`                  |
| `list-gmail-labels`  | READ     | List labels           | (none)                  |
| `send-gmail`         | ⚠️ WRITE | Send email            | `--to --subject --body` |
| `create-gmail-draft` | ⚠️ WRITE | Create draft          | `--to --subject --body` |

## Usage Examples

```bash
# Search emails from suppliers
node /Users/USER/node scripts/dist/cli.js search-gmail --query "from:supplier subject:invoice"

# Get calendar events for this week
node /Users/USER/node scripts/dist/cli.js get-events --time-min "2024-01-01T00:00:00Z" --time-max "2024-01-07T23:59:59Z"

# Search Drive for invoices
node /Users/USER/node scripts/dist/cli.js search-drive --query "invoice 2024"

# Read spreadsheet data
node /Users/USER/node scripts/dist/cli.js read-sheet --id "spreadsheetId" --range "Sheet1!A1:D10"

# Send an email
node /Users/USER/node scripts/dist/cli.js send-gmail --to "customer@example.com" --subject "Order Update" --body "Your order has shipped."

# Get comments from a Google Doc
node /Users/USER/node scripts/dist/cli.js get-doc-comments --id "YOUR_GOOGLE_DOC_ID"

# Add a comment to a document
node /Users/USER/node scripts/dist/cli.js create-doc-comment --id "documentId" --text "Please review this section"

# Reply to a comment
node /Users/USER/node scripts/dist/cli.js reply-doc-comment --id "documentId" --comment-id "commentId" --text "Thanks, I've updated it"

# Get document content with suggestions shown inline
node /Users/USER/node scripts/dist/cli.js get-doc-content --id "documentId" --suggestions-mode "SUGGESTIONS_INLINE"

# Get document preview with all suggestions accepted
node /Users/USER/node scripts/dist/cli.js get-doc-content --id "documentId" --suggestions-mode "PREVIEW_SUGGESTIONS_ACCEPTED"
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

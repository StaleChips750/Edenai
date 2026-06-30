# Google Docs MCP Server

FastMCP server with 94 tools for Google Docs, Sheets, Drive, Gmail, and Calendar.

## Tool Categories

| Category      | Count | Examples                                                                                                                                                                 |
| ------------- | ----- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Docs          | 5     | `readGoogleDoc`, `appendToGoogleDoc`, `insertText`, `deleteRange`, `listDocumentTabs`                                                                                    |
| Markdown      | 2     | `replaceDocumentWithMarkdown`, `appendMarkdownToGoogleDoc`                                                                                                               |
| Formatting    | 3     | `applyTextStyle`, `applyParagraphStyle`, `formatMatchingText`                                                                                                            |
| Structure     | 9     | `insertTable`, `insertPageBreak`, `insertSectionBreak`, `updateSectionStyle`, `insertImageFromUrl`, `insertLocalImage`, `editTableCell`\*, `findElement`\*, `fixListFormatting`\* |
| Comments      | 6     | `listComments`, `getComment`, `addComment`, `replyToComment`, `resolveComment`, `deleteComment`                                                                          |
| Sheets        | 31    | `readSpreadsheet`, `writeSpreadsheet`, `appendRows`, `clearRange`, `batchWrite`, `createSpreadsheet`, `listSpreadsheets`, `duplicateSheet`, `copySheetTo`, `renameSheet`, `deleteSheet`, `formatCells`, `setCellBorders`, `autoResizeColumns`, `autoResizeRows`, `setColumnWidths`, `setRowHeights`, `freezeRowsAndColumns`, `groupRows`, `protectRange`, `addConditionalFormatting`, `getConditionalFormatting`, `deleteConditionalFormatting`, `setDropdownValidation`, `insertChart`, `deleteChart` |
| Sheets Tables | 6     | `createTable`, `listTables`, `getTable`, `deleteTable`, `updateTableRange`, `appendTableRows`                                                                            |
| Drive         | 13    | `listGoogleDocs`, `searchGoogleDocs`, `getDocumentInfo`, `createFolder`, `moveFile`, `copyFile`, `createDocument`                                                        |
| Gmail         | 13    | `listMessages`, `getMessage`, `sendEmail`, `trashMessage`, `modifyMessageLabels`, `listLabels`, `createDraft`, `listDrafts`, `getDraft`, `updateDraft`, `sendDraft`, `deleteDraft`, `triageInbox` |
| Calendar      | 5     | `listEvents`, `createEvent`, `updateEvent`, `deleteEvent`, `quickAddEvent`                                                                                               |

\*Not fully implemented

## Setup

### 1. Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a new project
3. Enable these APIs: Google Docs, Google Sheets, Google Drive, Gmail, Google Calendar
4. Create OAuth 2.0 credentials (Desktop App)
5. Download `credentials.json` to the project root

### 2. Authenticate

```bash
npm install
npm run auth   # Opens browser for OAuth consent
```

This saves a `token.json` refresh token. Alternatively set environment variables:

```env
GOOGLE_CLIENT_ID=...
GOOGLE_CLIENT_SECRET=...
GOOGLE_REFRESH_TOKEN=...
```

### 3. Build & Run

```bash
npm run build
npm start
```

Or for development:

```bash
npm run dev
```

### 4. Configure Claude Code

Add to `~/.claude/settings.json`:

```json
{
  "mcpServers": {
    "google-docs": {
      "command": "node",
      "args": ["/path/to/edenai/dist/index.js"],
      "env": {
        "GOOGLE_CLIENT_ID": "...",
        "GOOGLE_CLIENT_SECRET": "...",
        "GOOGLE_REFRESH_TOKEN": "..."
      }
    }
  }
}
```

## Shared Drives Support

All Drive file operations use `supportsAllDrives: true` and `includeItemsFromAllDrives: true`.

## Parameter Patterns

- **Document ID:** Extract from URL: `docs.google.com/document/d/DOCUMENT_ID/edit`
- **Text targeting:** Use `textToFind` + `matchInstance` OR `startIndex`/`endIndex`
- **Colors:** Hex format `#RRGGBB` or `#RGB`
- **Alignment:** `START`, `END`, `CENTER`, `JUSTIFIED`
- **Indices:** 1-based, ranges are [start, end)
- **Tabs:** Optional `tabId` parameter (defaults to first tab)

## Source Files

| File                                         | Contains                                                                             |
| -------------------------------------------- | ------------------------------------------------------------------------------------ |
| `src/types.ts`                               | Zod schemas, hex color validation, style parameter definitions                       |
| `src/googleDocsApiHelpers.ts`                | `findTextRange`, `executeBatchUpdate`, style request builders                        |
| `src/googleSheetsApiHelpers.ts`              | A1 notation parsing, range/format operations, protected-range helpers, table helpers |
| `src/markdown-transformer/markdownToDocs.ts` | Markdown-to-Google-Docs conversion logic                                             |
| `src/markdown-transformer/docsToMarkdown.ts` | Google-Docs-to-Markdown conversion logic                                             |
| `src/markdown-transformer/index.ts`          | Markdown-it configuration and public API                                             |
| `src/index.ts`                               | Entry point, CLI handling, and MCP server startup                                    |

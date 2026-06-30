import { FastMCP } from "fastmcp";
import { google } from "googleapis";
import { z } from "zod";
import {
  textStyleParamsSchema,
  paragraphStyleParamsSchema,
  cellFormatSchema,
  borderStyleSchema,
  hexColorSchema,
} from "./types.js";
import {
  findTextRange,
  executeBatchUpdate,
  executeBatchUpdateWithSplitting,
  buildTextStyleRequest,
  buildParagraphStyleRequest,
  getDocumentEndIndex,
  extractPlainText,
} from "./googleDocsApiHelpers.js";
import {
  parseA1Notation,
  columnIndexToLetter,
  getSheetIdByName,
  getAllSheets,
  hexToSheetsColor,
  buildGridRange,
  buildCellFormat,
  findNamedRange,
} from "./googleSheetsApiHelpers.js";
import { markdownToDocs, appendMarkdownToDocs, docsToMarkdown } from "./markdown-transformer/index.js";

// --- OAuth2 setup ---
const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID,
  process.env.GOOGLE_CLIENT_SECRET,
);
oauth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });

const docsClient = google.docs({ version: "v1", auth: oauth2Client });
const sheetsClient = google.sheets({ version: "v4", auth: oauth2Client });
const driveClient = google.drive({ version: "v3", auth: oauth2Client });
const gmailClient = google.gmail({ version: "v1", auth: oauth2Client });
const calendarClient = google.calendar({ version: "v3", auth: oauth2Client });

const server = new FastMCP({
  name: "Google Docs MCP",
  version: "1.0.0",
});

// ============================================================
// DOCS TOOLS (5)
// ============================================================

server.addTool({
  name: "readGoogleDoc",
  description: "Read the full content of a Google Doc. Returns markdown by default.",
  parameters: z.object({
    documentId: z.string(),
    asMarkdown: z.boolean().optional().default(true),
  }),
  execute: async (args) => {
    const doc = await docsClient.documents.get({ documentId: args.documentId });
    if (args.asMarkdown) return docsToMarkdown(doc.data);
    return extractPlainText(doc.data);
  },
});

server.addTool({
  name: "appendToGoogleDoc",
  description: "Append plain text to the end of a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    text: z.string(),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    const doc = await docsClient.documents.get({ documentId: args.documentId });
    const endIndex = getDocumentEndIndex(doc.data);
    await executeBatchUpdate(docsClient, args.documentId, [{
      insertText: {
        location: { index: endIndex, ...(args.tabId ? { tabId: args.tabId } : {}) },
        text: args.text,
      },
    }]);
    return `Appended ${args.text.length} characters.`;
  },
});

server.addTool({
  name: "insertText",
  description: "Insert text at a specific index in a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    text: z.string(),
    index: z.number().int().positive(),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    await executeBatchUpdate(docsClient, args.documentId, [{
      insertText: {
        location: { index: args.index, ...(args.tabId ? { tabId: args.tabId } : {}) },
        text: args.text,
      },
    }]);
    return `Inserted text at index ${args.index}.`;
  },
});

server.addTool({
  name: "deleteRange",
  description: "Delete a range of content from a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    startIndex: z.number().int(),
    endIndex: z.number().int(),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    await executeBatchUpdate(docsClient, args.documentId, [{
      deleteContentRange: {
        range: {
          startIndex: args.startIndex,
          endIndex: args.endIndex,
          ...(args.tabId ? { tabId: args.tabId } : {}),
        },
      },
    }]);
    return `Deleted content from ${args.startIndex} to ${args.endIndex}.`;
  },
});

server.addTool({
  name: "listDocumentTabs",
  description: "List all tabs in a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
  }),
  execute: async (args) => {
    const doc = await docsClient.documents.get({ documentId: args.documentId });
    const tabs = doc.data.tabs ?? [];
    if (!tabs.length) return "No tabs (document does not use tabs).";
    return JSON.stringify(
      tabs.map((t) => ({
        tabId: t.tabProperties?.tabId,
        title: t.tabProperties?.title,
        nestingLevel: t.tabProperties?.nestingLevel,
      })),
      null,
      2,
    );
  },
});

// ============================================================
// MARKDOWN TOOLS (2)
// ============================================================

server.addTool({
  name: "replaceDocumentWithMarkdown",
  description: "Replace a Google Doc's entire content with rendered markdown.",
  parameters: z.object({
    documentId: z.string(),
    markdown: z.string(),
    firstHeadingAsTitle: z.boolean().optional().default(false),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    await markdownToDocs(docsClient, args.documentId, args.markdown, {
      firstHeadingAsTitle: args.firstHeadingAsTitle,
      tabId: args.tabId,
    });
    return "Document replaced with markdown content.";
  },
});

server.addTool({
  name: "appendMarkdownToGoogleDoc",
  description: "Append markdown content to the end of a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    markdown: z.string(),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    await appendMarkdownToDocs(docsClient, args.documentId, args.markdown, {
      tabId: args.tabId,
    });
    return "Markdown appended to document.";
  },
});

// ============================================================
// FORMATTING TOOLS (3)
// ============================================================

server.addTool({
  name: "applyTextStyle",
  description: "Apply text formatting (bold, italic, font size, color, etc.) to a range in a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    startIndex: z.number().int().optional(),
    endIndex: z.number().int().optional(),
    textToFind: z.string().optional(),
    matchInstance: z.number().int().positive().optional().default(1),
    style: textStyleParamsSchema,
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    let start = args.startIndex;
    let end = args.endIndex;
    if (args.textToFind) {
      const range = await findTextRange(docsClient, args.documentId, args.textToFind, args.matchInstance, args.tabId);
      start = range.startIndex;
      end = range.endIndex;
    }
    if (start === undefined || end === undefined) throw new Error("Provide startIndex/endIndex or textToFind");
    await executeBatchUpdate(docsClient, args.documentId, [
      buildTextStyleRequest(start, end, args.style, args.tabId),
    ]);
    return `Text style applied from ${start} to ${end}.`;
  },
});

server.addTool({
  name: "applyParagraphStyle",
  description: "Apply paragraph-level formatting (alignment, spacing, heading style) to a range in a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    startIndex: z.number().int().optional(),
    endIndex: z.number().int().optional(),
    textToFind: z.string().optional(),
    matchInstance: z.number().int().positive().optional().default(1),
    style: paragraphStyleParamsSchema,
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    let start = args.startIndex;
    let end = args.endIndex;
    if (args.textToFind) {
      const range = await findTextRange(docsClient, args.documentId, args.textToFind, args.matchInstance, args.tabId);
      start = range.startIndex;
      end = range.endIndex;
    }
    if (start === undefined || end === undefined) throw new Error("Provide startIndex/endIndex or textToFind");
    await executeBatchUpdate(docsClient, args.documentId, [
      buildParagraphStyleRequest(start, end, args.style, args.tabId),
    ]);
    return `Paragraph style applied from ${start} to ${end}.`;
  },
});

server.addTool({
  name: "formatMatchingText",
  description: "Find all instances of a string in a Google Doc and apply style to each match.",
  parameters: z.object({
    documentId: z.string(),
    textToFind: z.string(),
    textStyle: textStyleParamsSchema.optional(),
    paragraphStyle: paragraphStyleParamsSchema.optional(),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    const doc = await docsClient.documents.get({ documentId: args.documentId });
    const content = (doc.data.tabs?.[0]?.documentTab?.body?.content ?? doc.data.body?.content ?? []) as any[];
    const ranges: { start: number; end: number }[] = [];

    function scan(items: any[]) {
      for (const elem of items) {
        if (elem.paragraph) {
          for (const pe of elem.paragraph.elements ?? []) {
            const text: string = pe.textRun?.content ?? "";
            let from = 0, idx: number;
            while ((idx = text.indexOf(args.textToFind, from)) !== -1) {
              const base = pe.startIndex ?? 0;
              ranges.push({ start: base + idx, end: base + idx + args.textToFind.length });
              from = idx + 1;
            }
          }
        } else if (elem.table) {
          for (const row of elem.table.tableRows ?? []) {
            for (const cell of row.tableCells ?? []) scan(cell.content ?? []);
          }
        }
      }
    }
    scan(content);
    if (!ranges.length) return `No matches found for "${args.textToFind}".`;

    const requests: any[] = [];
    for (const r of ranges.reverse()) {
      if (args.textStyle) requests.push(buildTextStyleRequest(r.start, r.end, args.textStyle, args.tabId));
      if (args.paragraphStyle) requests.push(buildParagraphStyleRequest(r.start, r.end, args.paragraphStyle, args.tabId));
    }
    await executeBatchUpdateWithSplitting(docsClient, args.documentId, requests);
    return `Applied style to ${ranges.length} match(es).`;
  },
});

// ============================================================
// STRUCTURE TOOLS (9)
// ============================================================

server.addTool({
  name: "insertTable",
  description: "Insert a table into a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    rows: z.number().int().positive(),
    columns: z.number().int().positive(),
    index: z.number().int().optional(),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    let idx = args.index;
    if (idx === undefined) {
      const doc = await docsClient.documents.get({ documentId: args.documentId });
      idx = getDocumentEndIndex(doc.data);
    }
    await executeBatchUpdate(docsClient, args.documentId, [{
      insertTable: {
        rows: args.rows,
        columns: args.columns,
        location: { index: idx, ...(args.tabId ? { tabId: args.tabId } : {}) },
      },
    }]);
    return `Inserted ${args.rows}x${args.columns} table.`;
  },
});

server.addTool({
  name: "insertPageBreak",
  description: "Insert a page break in a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    index: z.number().int().optional(),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    let idx = args.index;
    if (idx === undefined) {
      const doc = await docsClient.documents.get({ documentId: args.documentId });
      idx = getDocumentEndIndex(doc.data);
    }
    await executeBatchUpdate(docsClient, args.documentId, [{
      insertPageBreak: {
        location: { index: idx, ...(args.tabId ? { tabId: args.tabId } : {}) },
      },
    }]);
    return "Page break inserted.";
  },
});

server.addTool({
  name: "insertSectionBreak",
  description: "Insert a section break in a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    sectionType: z.enum(["CONTINUOUS", "NEXT_PAGE", "EVEN_PAGE", "ODD_PAGE"]).optional().default("NEXT_PAGE"),
    index: z.number().int().optional(),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    let idx = args.index;
    if (idx === undefined) {
      const doc = await docsClient.documents.get({ documentId: args.documentId });
      idx = getDocumentEndIndex(doc.data);
    }
    await executeBatchUpdate(docsClient, args.documentId, [{
      insertSectionBreak: {
        location: { index: idx, ...(args.tabId ? { tabId: args.tabId } : {}) },
        sectionType: args.sectionType,
      },
    }]);
    return `Section break (${args.sectionType}) inserted.`;
  },
});

server.addTool({
  name: "updateSectionStyle",
  description: "Update page margins for a document section.",
  parameters: z.object({
    documentId: z.string(),
    marginTop: z.number().optional(),
    marginBottom: z.number().optional(),
    marginLeft: z.number().optional(),
    marginRight: z.number().optional(),
  }),
  execute: async (args) => {
    const doc = await docsClient.documents.get({ documentId: args.documentId });
    const content = (doc.data.tabs?.[0]?.documentTab?.body?.content ?? doc.data.body?.content ?? []) as any[];
    const sectionStyle: any = {};
    const fields: string[] = [];
    if (args.marginTop !== undefined) { sectionStyle.marginTop = { magnitude: args.marginTop, unit: "PT" }; fields.push("marginTop"); }
    if (args.marginBottom !== undefined) { sectionStyle.marginBottom = { magnitude: args.marginBottom, unit: "PT" }; fields.push("marginBottom"); }
    if (args.marginLeft !== undefined) { sectionStyle.marginLeft = { magnitude: args.marginLeft, unit: "PT" }; fields.push("marginLeft"); }
    if (args.marginRight !== undefined) { sectionStyle.marginRight = { magnitude: args.marginRight, unit: "PT" }; fields.push("marginRight"); }
    if (!fields.length) return "No margin values provided.";
    const endIndex = getDocumentEndIndex(doc.data);
    await executeBatchUpdate(docsClient, args.documentId, [{
      updateSectionStyle: {
        range: { startIndex: 1, endIndex },
        sectionStyle,
        fields: fields.join(","),
      },
    }]);
    return "Section style updated.";
  },
});

server.addTool({
  name: "insertImageFromUrl",
  description: "Insert an image from a URL into a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    imageUrl: z.string().url(),
    index: z.number().int().optional(),
    width: z.number().optional(),
    height: z.number().optional(),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    let idx = args.index;
    if (idx === undefined) {
      const doc = await docsClient.documents.get({ documentId: args.documentId });
      idx = getDocumentEndIndex(doc.data);
    }
    const objSize: any = {};
    if (args.width) objSize.width = { magnitude: args.width, unit: "PT" };
    if (args.height) objSize.height = { magnitude: args.height, unit: "PT" };
    await executeBatchUpdate(docsClient, args.documentId, [{
      insertInlineImage: {
        location: { index: idx, ...(args.tabId ? { tabId: args.tabId } : {}) },
        uri: args.imageUrl,
        ...(Object.keys(objSize).length ? { objectSize: objSize } : {}),
      },
    }]);
    return "Image inserted.";
  },
});

server.addTool({
  name: "insertLocalImage",
  description: "Insert an image from Google Drive into a Google Doc using a Drive file ID.",
  parameters: z.object({
    documentId: z.string(),
    driveFileId: z.string(),
    index: z.number().int().optional(),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    const fileRes = await driveClient.files.get({
      fileId: args.driveFileId,
      fields: "webContentLink",
      supportsAllDrives: true,
    });
    const url = fileRes.data.webContentLink;
    if (!url) throw new Error("Could not get content link for Drive file");
    let idx = args.index;
    if (idx === undefined) {
      const doc = await docsClient.documents.get({ documentId: args.documentId });
      idx = getDocumentEndIndex(doc.data);
    }
    await executeBatchUpdate(docsClient, args.documentId, [{
      insertInlineImage: {
        location: { index: idx, ...(args.tabId ? { tabId: args.tabId } : {}) },
        uri: url,
      },
    }]);
    return "Image from Drive inserted.";
  },
});

server.addTool({
  name: "editTableCell",
  description: "Replace the text content of a table cell in a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    tableIndex: z.number().int().nonnegative(),
    rowIndex: z.number().int().nonnegative(),
    columnIndex: z.number().int().nonnegative(),
    newText: z.string(),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    const doc = await docsClient.documents.get({ documentId: args.documentId });
    const content = (doc.data.tabs?.[0]?.documentTab?.body?.content ?? doc.data.body?.content ?? []) as any[];
    const tables = content.filter((e: any) => e.table);
    const table = tables[args.tableIndex];
    if (!table) throw new Error(`Table at index ${args.tableIndex} not found`);
    const row = table.table.tableRows?.[args.rowIndex];
    if (!row) throw new Error(`Row ${args.rowIndex} not found`);
    const cell = row.tableCells?.[args.columnIndex];
    if (!cell) throw new Error(`Column ${args.columnIndex} not found`);
    const cellContent = cell.content?.[0];
    if (!cellContent?.paragraph) throw new Error("Cell content not accessible");
    const startIndex = cellContent.startIndex;
    const endIndex = cellContent.endIndex - 1;
    const requests: any[] = [];
    if (endIndex > startIndex) {
      requests.push({ deleteContentRange: { range: { startIndex, endIndex, ...(args.tabId ? { tabId: args.tabId } : {}) } } });
    }
    requests.push({ insertText: { location: { index: startIndex, ...(args.tabId ? { tabId: args.tabId } : {}) }, text: args.newText } });
    await executeBatchUpdate(docsClient, args.documentId, requests);
    return `Cell [${args.rowIndex}, ${args.columnIndex}] updated.`;
  },
});

server.addTool({
  name: "findElement",
  description: "Find the start/end index of text in a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    textToFind: z.string(),
    matchInstance: z.number().int().positive().optional().default(1),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    const range = await findTextRange(docsClient, args.documentId, args.textToFind, args.matchInstance, args.tabId);
    return JSON.stringify(range);
  },
});

server.addTool({
  name: "fixListFormatting",
  description: "Apply or remove bullet/numbered list formatting from a paragraph range in a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    startIndex: z.number().int(),
    endIndex: z.number().int(),
    listPreset: z.enum(["BULLET_DISC_CIRCLE_SQUARE", "BULLET_ARROW_DIAMOND_DISC", "NUMBERED_DECIMAL_ALPHA_ROMAN"]).optional().default("BULLET_DISC_CIRCLE_SQUARE"),
    remove: z.boolean().optional().default(false),
    tabId: z.string().optional(),
  }),
  execute: async (args) => {
    const rangeObj = { startIndex: args.startIndex, endIndex: args.endIndex, ...(args.tabId ? { tabId: args.tabId } : {}) };
    const requests: any[] = args.remove
      ? [{ deleteParagraphBullets: { range: rangeObj } }]
      : [{ createParagraphBullets: { range: rangeObj, bulletPreset: args.listPreset } }];
    await executeBatchUpdate(docsClient, args.documentId, requests);
    return args.remove ? "List formatting removed." : `List formatted with ${args.listPreset}.`;
  },
});

// ============================================================
// COMMENTS TOOLS (6)
// ============================================================

server.addTool({
  name: "listComments",
  description: "List all comments on a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    includeDeleted: z.boolean().optional().default(false),
  }),
  execute: async (args) => {
    const res = await driveClient.comments.list({
      fileId: args.documentId,
      fields: "comments(id,content,author,createdTime,resolved,replies)",
      includeDeleted: args.includeDeleted,
    });
    return JSON.stringify(res.data.comments ?? [], null, 2);
  },
});

server.addTool({
  name: "getComment",
  description: "Get a specific comment from a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    commentId: z.string(),
  }),
  execute: async (args) => {
    const res = await driveClient.comments.get({
      fileId: args.documentId,
      commentId: args.commentId,
      fields: "*",
    });
    return JSON.stringify(res.data, null, 2);
  },
});

server.addTool({
  name: "addComment",
  description: "Add a comment to a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    content: z.string(),
  }),
  execute: async (args) => {
    const res = await driveClient.comments.create({
      fileId: args.documentId,
      fields: "id,content,createdTime",
      requestBody: { content: args.content },
    });
    return JSON.stringify(res.data, null, 2);
  },
});

server.addTool({
  name: "replyToComment",
  description: "Reply to an existing comment on a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    commentId: z.string(),
    content: z.string(),
  }),
  execute: async (args) => {
    const res = await driveClient.replies.create({
      fileId: args.documentId,
      commentId: args.commentId,
      fields: "id,content,createdTime",
      requestBody: { content: args.content },
    });
    return JSON.stringify(res.data, null, 2);
  },
});

server.addTool({
  name: "resolveComment",
  description: "Resolve or unresolve a comment on a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    commentId: z.string(),
    resolved: z.boolean().optional().default(true),
  }),
  execute: async (args) => {
    await driveClient.comments.update({
      fileId: args.documentId,
      commentId: args.commentId,
      fields: "id,resolved",
      requestBody: { resolved: args.resolved },
    });
    return `Comment ${args.resolved ? "resolved" : "unresolved"}.`;
  },
});

server.addTool({
  name: "deleteComment",
  description: "Delete a comment from a Google Doc.",
  parameters: z.object({
    documentId: z.string(),
    commentId: z.string(),
  }),
  execute: async (args) => {
    await driveClient.comments.delete({
      fileId: args.documentId,
      commentId: args.commentId,
    });
    return "Comment deleted.";
  },
});

// ============================================================
// SHEETS TOOLS (32)
// ============================================================

server.addTool({
  name: "readSpreadsheet",
  description: "Read values from a Google Sheets range.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string().describe("A1 notation, e.g. Sheet1!A1:D10"),
    valueRenderOption: z.enum(["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"]).optional().default("FORMATTED_VALUE"),
  }),
  execute: async (args) => {
    const res = await sheetsClient.spreadsheets.values.get({
      spreadsheetId: args.spreadsheetId,
      range: args.range,
      valueRenderOption: args.valueRenderOption,
    });
    return JSON.stringify(res.data.values ?? [], null, 2);
  },
});

server.addTool({
  name: "writeSpreadsheet",
  description: "Write values to a Google Sheets range.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string(),
    values: z.array(z.array(z.any())),
    valueInputOption: z.enum(["RAW", "USER_ENTERED"]).optional().default("USER_ENTERED"),
  }),
  execute: async (args) => {
    const res = await sheetsClient.spreadsheets.values.update({
      spreadsheetId: args.spreadsheetId,
      range: args.range,
      valueInputOption: args.valueInputOption,
      requestBody: { values: args.values },
    });
    return `Updated ${res.data.updatedCells} cells.`;
  },
});

server.addTool({
  name: "appendRows",
  description: "Append rows of data to a Google Sheet after existing content.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string(),
    values: z.array(z.array(z.any())),
    valueInputOption: z.enum(["RAW", "USER_ENTERED"]).optional().default("USER_ENTERED"),
  }),
  execute: async (args) => {
    await sheetsClient.spreadsheets.values.append({
      spreadsheetId: args.spreadsheetId,
      range: args.range,
      valueInputOption: args.valueInputOption,
      requestBody: { values: args.values },
    });
    return `Appended ${args.values.length} row(s).`;
  },
});

server.addTool({
  name: "clearRange",
  description: "Clear all values from a Google Sheets range.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string(),
  }),
  execute: async (args) => {
    await sheetsClient.spreadsheets.values.clear({
      spreadsheetId: args.spreadsheetId,
      range: args.range,
    });
    return `Range ${args.range} cleared.`;
  },
});

server.addTool({
  name: "batchWrite",
  description: "Write multiple ranges in a Google Spreadsheet in one request.",
  parameters: z.object({
    spreadsheetId: z.string(),
    data: z.array(z.object({ range: z.string(), values: z.array(z.array(z.any())) })),
    valueInputOption: z.enum(["RAW", "USER_ENTERED"]).optional().default("USER_ENTERED"),
  }),
  execute: async (args) => {
    const res = await sheetsClient.spreadsheets.values.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: { valueInputOption: args.valueInputOption, data: args.data },
    });
    return `Batch write complete. Updated ${res.data.totalUpdatedCells ?? 0} cells.`;
  },
});

server.addTool({
  name: "createSpreadsheet",
  description: "Create a new Google Spreadsheet.",
  parameters: z.object({
    title: z.string(),
    sheetTitles: z.array(z.string()).optional(),
  }),
  execute: async (args) => {
    const res = await sheetsClient.spreadsheets.create({
      requestBody: {
        properties: { title: args.title },
        sheets: args.sheetTitles?.map((t) => ({ properties: { title: t } })),
      },
    });
    return JSON.stringify({ spreadsheetId: res.data.spreadsheetId, url: res.data.spreadsheetUrl }, null, 2);
  },
});

server.addTool({
  name: "getSpreadsheetInfo",
  description: "Get metadata about a Google Spreadsheet (title, all sheet names and IDs).",
  parameters: z.object({
    spreadsheetId: z.string(),
  }),
  execute: async (args) => {
    const res = await sheetsClient.spreadsheets.get({
      spreadsheetId: args.spreadsheetId,
      fields: "spreadsheetId,properties,sheets.properties",
    });
    return JSON.stringify({
      spreadsheetId: res.data.spreadsheetId,
      title: res.data.properties?.title,
      sheets: res.data.sheets?.map((s) => ({
        sheetId: s.properties?.sheetId,
        title: s.properties?.title,
        index: s.properties?.index,
        rowCount: s.properties?.gridProperties?.rowCount,
        columnCount: s.properties?.gridProperties?.columnCount,
      })),
    }, null, 2);
  },
});

server.addTool({
  name: "listSpreadsheets",
  description: "List Google Spreadsheets in Drive.",
  parameters: z.object({
    query: z.string().optional(),
    maxResults: z.number().int().optional().default(20),
    folderId: z.string().optional(),
  }),
  execute: async (args) => {
    let q = "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false";
    if (args.query) q += ` and name contains '${args.query.replace(/'/g, "\\'")}' `;
    if (args.folderId) q += ` and '${args.folderId}' in parents`;
    const res = await driveClient.files.list({
      q,
      pageSize: args.maxResults,
      fields: "files(id,name,modifiedTime,webViewLink)",
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    return JSON.stringify(res.data.files ?? [], null, 2);
  },
});

server.addTool({
  name: "addSheet",
  description: "Add a new sheet tab to a Google Spreadsheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    title: z.string(),
    rowCount: z.number().int().optional().default(1000),
    columnCount: z.number().int().optional().default(26),
  }),
  execute: async (args) => {
    const res = await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          addSheet: {
            properties: {
              title: args.title,
              gridProperties: { rowCount: args.rowCount, columnCount: args.columnCount },
            },
          },
        }],
      },
    });
    const added = (res.data.replies?.[0] as any)?.addSheet?.properties;
    return JSON.stringify({ sheetId: added?.sheetId, title: added?.title }, null, 2);
  },
});

server.addTool({
  name: "duplicateSheet",
  description: "Duplicate a sheet within a Google Spreadsheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
    newSheetName: z.string().optional(),
    insertSheetIndex: z.number().int().optional(),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.sheetName);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          duplicateSheet: {
            sourceSheetId: sheetId,
            newSheetName: args.newSheetName,
            insertSheetIndex: args.insertSheetIndex,
          },
        }],
      },
    });
    return `Sheet "${args.sheetName}" duplicated.`;
  },
});

server.addTool({
  name: "copySheetTo",
  description: "Copy a sheet from one spreadsheet to another.",
  parameters: z.object({
    sourceSpreadsheetId: z.string(),
    sheetName: z.string(),
    destinationSpreadsheetId: z.string(),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.sourceSpreadsheetId, args.sheetName);
    const res = await sheetsClient.spreadsheets.sheets.copyTo({
      spreadsheetId: args.sourceSpreadsheetId,
      sheetId,
      requestBody: { destinationSpreadsheetId: args.destinationSpreadsheetId },
    });
    return JSON.stringify({ sheetId: res.data.sheetId, title: res.data.title }, null, 2);
  },
});

server.addTool({
  name: "renameSheet",
  description: "Rename a sheet tab in a Google Spreadsheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    currentName: z.string(),
    newName: z.string(),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.currentName);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          updateSheetProperties: {
            properties: { sheetId, title: args.newName },
            fields: "title",
          },
        }],
      },
    });
    return `Sheet renamed to "${args.newName}".`;
  },
});

server.addTool({
  name: "deleteSheet",
  description: "Delete a sheet from a Google Spreadsheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.sheetName);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: { requests: [{ deleteSheet: { sheetId } }] },
    });
    return `Sheet "${args.sheetName}" deleted.`;
  },
});

server.addTool({
  name: "formatCells",
  description: "Apply cell formatting to a Google Sheets range.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string(),
    format: cellFormatSchema,
  }),
  execute: async (args) => {
    const parsed = parseA1Notation(args.range);
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, parsed.sheetName ?? "Sheet1");
    const gridRange = buildGridRange(sheetId, parsed);
    const cellFormat = buildCellFormat(args.format);
    const fields = Object.keys(cellFormat).join(",");
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          repeatCell: {
            range: gridRange,
            cell: { userEnteredFormat: cellFormat },
            fields: `userEnteredFormat(${fields})`,
          },
        }],
      },
    });
    return `Format applied to ${args.range}.`;
  },
});

server.addTool({
  name: "setCellBorders",
  description: "Set borders on a Google Sheets range.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string(),
    top: borderStyleSchema.optional(),
    bottom: borderStyleSchema.optional(),
    left: borderStyleSchema.optional(),
    right: borderStyleSchema.optional(),
    innerHorizontal: borderStyleSchema.optional(),
    innerVertical: borderStyleSchema.optional(),
    color: hexColorSchema.optional().default("#000000"),
    width: z.number().int().optional().default(1),
  }),
  execute: async (args) => {
    const parsed = parseA1Notation(args.range);
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, parsed.sheetName ?? "Sheet1");
    const borderColor = hexToSheetsColor(args.color ?? "#000000");
    const makeBorder = (style?: string) => style ? { style, color: borderColor, width: args.width } : undefined;
    const borders: any = {};
    if (args.top) borders.top = makeBorder(args.top);
    if (args.bottom) borders.bottom = makeBorder(args.bottom);
    if (args.left) borders.left = makeBorder(args.left);
    if (args.right) borders.right = makeBorder(args.right);
    if (args.innerHorizontal) borders.innerHorizontal = makeBorder(args.innerHorizontal);
    if (args.innerVertical) borders.innerVertical = makeBorder(args.innerVertical);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: { requests: [{ updateBorders: { range: buildGridRange(sheetId, parsed), ...borders } }] },
    });
    return `Borders set on ${args.range}.`;
  },
});

server.addTool({
  name: "autoResizeColumns",
  description: "Auto-resize columns in a Google Sheet to fit content.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
    startColumnIndex: z.number().int().nonnegative().optional().default(0),
    endColumnIndex: z.number().int().optional(),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.sheetName);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          autoResizeDimensions: {
            dimensions: { sheetId, dimension: "COLUMNS", startIndex: args.startColumnIndex, endIndex: args.endColumnIndex },
          },
        }],
      },
    });
    return "Columns auto-resized.";
  },
});

server.addTool({
  name: "autoResizeRows",
  description: "Auto-resize rows in a Google Sheet to fit content.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
    startRowIndex: z.number().int().nonnegative().optional().default(0),
    endRowIndex: z.number().int().optional(),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.sheetName);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          autoResizeDimensions: {
            dimensions: { sheetId, dimension: "ROWS", startIndex: args.startRowIndex, endIndex: args.endRowIndex },
          },
        }],
      },
    });
    return "Rows auto-resized.";
  },
});

server.addTool({
  name: "setColumnWidths",
  description: "Set column widths in a Google Sheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
    startColumnIndex: z.number().int().nonnegative(),
    endColumnIndex: z.number().int(),
    pixelSize: z.number().int().positive(),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.sheetName);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          updateDimensionProperties: {
            range: { sheetId, dimension: "COLUMNS", startIndex: args.startColumnIndex, endIndex: args.endColumnIndex },
            properties: { pixelSize: args.pixelSize },
            fields: "pixelSize",
          },
        }],
      },
    });
    return `Column width set to ${args.pixelSize}px.`;
  },
});

server.addTool({
  name: "setRowHeights",
  description: "Set row heights in a Google Sheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
    startRowIndex: z.number().int().nonnegative(),
    endRowIndex: z.number().int(),
    pixelSize: z.number().int().positive(),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.sheetName);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          updateDimensionProperties: {
            range: { sheetId, dimension: "ROWS", startIndex: args.startRowIndex, endIndex: args.endRowIndex },
            properties: { pixelSize: args.pixelSize },
            fields: "pixelSize",
          },
        }],
      },
    });
    return `Row height set to ${args.pixelSize}px.`;
  },
});

server.addTool({
  name: "freezeRowsAndColumns",
  description: "Freeze rows and/or columns in a Google Sheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
    frozenRowCount: z.number().int().nonnegative().optional().default(0),
    frozenColumnCount: z.number().int().nonnegative().optional().default(0),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.sheetName);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          updateSheetProperties: {
            properties: {
              sheetId,
              gridProperties: { frozenRowCount: args.frozenRowCount, frozenColumnCount: args.frozenColumnCount },
            },
            fields: "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
          },
        }],
      },
    });
    return `Frozen ${args.frozenRowCount} row(s) and ${args.frozenColumnCount} column(s).`;
  },
});

server.addTool({
  name: "groupRows",
  description: "Group rows or columns in a Google Sheet for collapsible sections.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
    startIndex: z.number().int().nonnegative(),
    endIndex: z.number().int(),
    dimension: z.enum(["ROWS", "COLUMNS"]).optional().default("ROWS"),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.sheetName);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          addDimensionGroup: {
            range: { sheetId, dimension: args.dimension, startIndex: args.startIndex, endIndex: args.endIndex },
          },
        }],
      },
    });
    return `${args.dimension} ${args.startIndex}–${args.endIndex} grouped.`;
  },
});

server.addTool({
  name: "sortRange",
  description: "Sort a range of data in a Google Sheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string(),
    sortColumnIndex: z.number().int().nonnegative(),
    ascending: z.boolean().optional().default(true),
  }),
  execute: async (args) => {
    const parsed = parseA1Notation(args.range);
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, parsed.sheetName ?? "Sheet1");
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          sortRange: {
            range: buildGridRange(sheetId, parsed),
            sortSpecs: [{ dimensionIndex: args.sortColumnIndex, sortOrder: args.ascending ? "ASCENDING" : "DESCENDING" }],
          },
        }],
      },
    });
    return `Range sorted by column ${args.sortColumnIndex} ${args.ascending ? "ascending" : "descending"}.`;
  },
});

server.addTool({
  name: "mergeCells",
  description: "Merge cells in a Google Sheet range.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string(),
    mergeType: z.enum(["MERGE_ALL", "MERGE_COLUMNS", "MERGE_ROWS"]).optional().default("MERGE_ALL"),
  }),
  execute: async (args) => {
    const parsed = parseA1Notation(args.range);
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, parsed.sheetName ?? "Sheet1");
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: { requests: [{ mergeCells: { range: buildGridRange(sheetId, parsed), mergeType: args.mergeType } }] },
    });
    return `Cells merged in ${args.range}.`;
  },
});

server.addTool({
  name: "unmergeCells",
  description: "Unmerge previously merged cells in a Google Sheet range.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string(),
  }),
  execute: async (args) => {
    const parsed = parseA1Notation(args.range);
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, parsed.sheetName ?? "Sheet1");
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: { requests: [{ unmergeCells: { range: buildGridRange(sheetId, parsed) } }] },
    });
    return `Cells unmerged in ${args.range}.`;
  },
});

server.addTool({
  name: "protectRange",
  description: "Protect a range or sheet from editing in a Google Spreadsheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string().optional(),
    sheetName: z.string().optional(),
    description: z.string().optional(),
    warningOnly: z.boolean().optional().default(true),
  }),
  execute: async (args) => {
    let protectedRange: any = { warningOnly: args.warningOnly, description: args.description };
    if (args.range) {
      const parsed = parseA1Notation(args.range);
      const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, parsed.sheetName ?? args.sheetName ?? "Sheet1");
      protectedRange.range = buildGridRange(sheetId, parsed);
    } else if (args.sheetName) {
      const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.sheetName);
      protectedRange.sheetId = sheetId;
    }
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: { requests: [{ addProtectedRange: { protectedRange } }] },
    });
    return "Protection applied.";
  },
});

server.addTool({
  name: "addConditionalFormatting",
  description: "Add a conditional formatting rule to a Google Sheets range.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string(),
    conditionType: z.enum(["NUMBER_GREATER", "NUMBER_LESS", "NUMBER_EQUAL", "TEXT_CONTAINS", "TEXT_EQ", "BLANK", "NOT_BLANK", "CUSTOM_FORMULA"]),
    conditionValues: z.array(z.string()).optional(),
    backgroundColor: hexColorSchema.optional(),
    textColor: hexColorSchema.optional(),
    bold: z.boolean().optional(),
  }),
  execute: async (args) => {
    const parsed = parseA1Notation(args.range);
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, parsed.sheetName ?? "Sheet1");
    const format: any = {};
    if (args.backgroundColor) format.backgroundColor = hexToSheetsColor(args.backgroundColor);
    if (args.textColor || args.bold !== undefined) {
      format.textFormat = {};
      if (args.textColor) format.textFormat.foregroundColor = hexToSheetsColor(args.textColor);
      if (args.bold !== undefined) format.textFormat.bold = args.bold;
    }
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          addConditionalFormatRule: {
            rule: {
              ranges: [buildGridRange(sheetId, parsed)],
              booleanRule: {
                condition: {
                  type: args.conditionType,
                  values: args.conditionValues?.map((v) => ({ userEnteredValue: v })),
                },
                format,
              },
            },
            index: 0,
          },
        }],
      },
    });
    return "Conditional formatting rule added.";
  },
});

server.addTool({
  name: "getConditionalFormatting",
  description: "List all conditional formatting rules in a Google Sheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
  }),
  execute: async (args) => {
    const res = await sheetsClient.spreadsheets.get({
      spreadsheetId: args.spreadsheetId,
      fields: "sheets.properties,sheets.conditionalFormats",
    });
    const sheet = res.data.sheets?.find((s) => s.properties?.title === args.sheetName);
    if (!sheet) throw new Error(`Sheet "${args.sheetName}" not found`);
    return JSON.stringify(sheet.conditionalFormats ?? [], null, 2);
  },
});

server.addTool({
  name: "deleteConditionalFormatting",
  description: "Delete a conditional formatting rule by index from a Google Sheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
    ruleIndex: z.number().int().nonnegative(),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.sheetName);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: { requests: [{ deleteConditionalFormatRule: { sheetId, index: args.ruleIndex } }] },
    });
    return `Rule ${args.ruleIndex} deleted.`;
  },
});

server.addTool({
  name: "setDropdownValidation",
  description: "Add a dropdown list validation to a Google Sheets range.",
  parameters: z.object({
    spreadsheetId: z.string(),
    range: z.string(),
    options: z.array(z.string()),
    strict: z.boolean().optional().default(true),
    showCustomUi: z.boolean().optional().default(true),
  }),
  execute: async (args) => {
    const parsed = parseA1Notation(args.range);
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, parsed.sheetName ?? "Sheet1");
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          setDataValidation: {
            range: buildGridRange(sheetId, parsed),
            rule: {
              condition: {
                type: "ONE_OF_LIST",
                values: args.options.map((v) => ({ userEnteredValue: v })),
              },
              strict: args.strict,
              showCustomUi: args.showCustomUi,
            },
          },
        }],
      },
    });
    return `Dropdown validation set with ${args.options.length} option(s).`;
  },
});

server.addTool({
  name: "insertChart",
  description: "Insert a chart into a Google Sheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
    dataRange: z.string(),
    chartType: z.enum(["BAR", "LINE", "AREA", "COLUMN", "SCATTER", "PIE"]).optional().default("COLUMN"),
    title: z.string().optional(),
    anchorRow: z.number().int().nonnegative().optional().default(0),
    anchorColumn: z.number().int().nonnegative().optional().default(0),
  }),
  execute: async (args) => {
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, args.sheetName);
    const parsed = parseA1Notation(args.dataRange);
    const dataSheetId = parsed.sheetName
      ? await getSheetIdByName(sheetsClient, args.spreadsheetId, parsed.sheetName)
      : sheetId;
    const gridRange = buildGridRange(dataSheetId, parsed);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          addChart: {
            chart: {
              spec: {
                title: args.title,
                basicChart: {
                  chartType: args.chartType,
                  headerCount: 1,
                  domains: [{ domain: { sourceRange: { sources: [gridRange] } } }],
                  series: [{ series: { sourceRange: { sources: [gridRange] } } }],
                },
              },
              position: {
                overlayPosition: {
                  anchorCell: { sheetId, rowIndex: args.anchorRow, columnIndex: args.anchorColumn },
                },
              },
            },
          },
        }],
      },
    });
    return `${args.chartType} chart inserted.`;
  },
});

server.addTool({
  name: "deleteChart",
  description: "Delete a chart from a Google Sheet by chart ID.",
  parameters: z.object({
    spreadsheetId: z.string(),
    chartId: z.number().int(),
  }),
  execute: async (args) => {
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: { requests: [{ deleteEmbeddedObject: { objectId: args.chartId } }] },
    });
    return `Chart ${args.chartId} deleted.`;
  },
});

server.addTool({
  name: "findInSheet",
  description: "Search for a value in a Google Sheet and return matching cell references.",
  parameters: z.object({
    spreadsheetId: z.string(),
    sheetName: z.string(),
    searchValue: z.string(),
    searchByRegex: z.boolean().optional().default(false),
  }),
  execute: async (args) => {
    const res = await sheetsClient.spreadsheets.values.get({
      spreadsheetId: args.spreadsheetId,
      range: args.sheetName,
      valueRenderOption: "FORMATTED_VALUE",
    });
    const values = res.data.values ?? [];
    const regex = args.searchByRegex ? new RegExp(args.searchValue, "i") : null;
    const matches: { cell: string; row: number; col: number; value: string }[] = [];
    values.forEach((row, rowIdx) => {
      row.forEach((cell, colIdx) => {
        const cellStr = String(cell);
        const hit = regex ? regex.test(cellStr) : cellStr.includes(args.searchValue);
        if (hit) {
          matches.push({ cell: `${columnIndexToLetter(colIdx)}${rowIdx + 1}`, row: rowIdx + 1, col: colIdx + 1, value: cellStr });
        }
      });
    });
    if (!matches.length) return `No matches found for "${args.searchValue}".`;
    return JSON.stringify(matches, null, 2);
  },
});

// ============================================================
// SHEETS TABLES TOOLS (6)
// ============================================================

server.addTool({
  name: "createTable",
  description: "Create a named range (table) in a Google Spreadsheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    name: z.string(),
    range: z.string(),
  }),
  execute: async (args) => {
    const parsed = parseA1Notation(args.range);
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, parsed.sheetName ?? "Sheet1");
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          addNamedRange: {
            namedRange: { name: args.name, range: buildGridRange(sheetId, parsed) },
          },
        }],
      },
    });
    return `Named range "${args.name}" created.`;
  },
});

server.addTool({
  name: "listTables",
  description: "List all named ranges in a Google Spreadsheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
  }),
  execute: async (args) => {
    const res = await sheetsClient.spreadsheets.get({
      spreadsheetId: args.spreadsheetId,
      fields: "namedRanges",
    });
    return JSON.stringify(res.data.namedRanges ?? [], null, 2);
  },
});

server.addTool({
  name: "getTable",
  description: "Get a named range definition and read its data.",
  parameters: z.object({
    spreadsheetId: z.string(),
    name: z.string(),
  }),
  execute: async (args) => {
    const namedRange = await findNamedRange(sheetsClient, args.spreadsheetId, args.name);
    if (!namedRange) throw new Error(`Named range "${args.name}" not found`);
    const gr = namedRange.range!;
    const allSheets = await getAllSheets(sheetsClient, args.spreadsheetId);
    const sheet = allSheets.find((s) => s.properties?.sheetId === gr.sheetId);
    const sheetTitle = sheet?.properties?.title ?? "Sheet1";
    const startCol = columnIndexToLetter(gr.startColumnIndex ?? 0);
    const endCol = columnIndexToLetter((gr.endColumnIndex ?? 1) - 1);
    const a1 = `${sheetTitle}!${startCol}${(gr.startRowIndex ?? 0) + 1}:${endCol}${gr.endRowIndex ?? 1000}`;
    const values = await sheetsClient.spreadsheets.values.get({ spreadsheetId: args.spreadsheetId, range: a1 });
    return JSON.stringify({ definition: namedRange, values: values.data.values ?? [] }, null, 2);
  },
});

server.addTool({
  name: "deleteTable",
  description: "Delete a named range from a Google Spreadsheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    name: z.string(),
  }),
  execute: async (args) => {
    const namedRange = await findNamedRange(sheetsClient, args.spreadsheetId, args.name);
    if (!namedRange) throw new Error(`Named range "${args.name}" not found`);
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: { requests: [{ deleteNamedRange: { namedRangeId: namedRange.namedRangeId } }] },
    });
    return `Named range "${args.name}" deleted.`;
  },
});

server.addTool({
  name: "updateTableRange",
  description: "Update the range of a named range in a Google Spreadsheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    name: z.string(),
    newRange: z.string(),
  }),
  execute: async (args) => {
    const namedRange = await findNamedRange(sheetsClient, args.spreadsheetId, args.name);
    if (!namedRange) throw new Error(`Named range "${args.name}" not found`);
    const parsed = parseA1Notation(args.newRange);
    const sheetId = await getSheetIdByName(sheetsClient, args.spreadsheetId, parsed.sheetName ?? "Sheet1");
    await sheetsClient.spreadsheets.batchUpdate({
      spreadsheetId: args.spreadsheetId,
      requestBody: {
        requests: [{
          updateNamedRange: {
            namedRange: { namedRangeId: namedRange.namedRangeId, name: args.name, range: buildGridRange(sheetId, parsed) },
            fields: "range",
          },
        }],
      },
    });
    return `Named range "${args.name}" updated to ${args.newRange}.`;
  },
});

server.addTool({
  name: "appendTableRows",
  description: "Append rows to a named range (table) in a Google Spreadsheet.",
  parameters: z.object({
    spreadsheetId: z.string(),
    name: z.string(),
    values: z.array(z.array(z.any())),
    valueInputOption: z.enum(["RAW", "USER_ENTERED"]).optional().default("USER_ENTERED"),
  }),
  execute: async (args) => {
    const namedRange = await findNamedRange(sheetsClient, args.spreadsheetId, args.name);
    if (!namedRange) throw new Error(`Named range "${args.name}" not found`);
    const gr = namedRange.range!;
    const allSheets = await getAllSheets(sheetsClient, args.spreadsheetId);
    const sheet = allSheets.find((s) => s.properties?.sheetId === gr.sheetId);
    const sheetTitle = sheet?.properties?.title ?? "Sheet1";
    const startCol = columnIndexToLetter(gr.startColumnIndex ?? 0);
    const a1 = `${sheetTitle}!${startCol}`;
    await sheetsClient.spreadsheets.values.append({
      spreadsheetId: args.spreadsheetId,
      range: a1,
      valueInputOption: args.valueInputOption,
      requestBody: { values: args.values },
    });
    return `Appended ${args.values.length} row(s) to "${args.name}".`;
  },
});

// ============================================================
// DRIVE TOOLS (13)
// ============================================================

server.addTool({
  name: "listGoogleDocs",
  description: "List Google Docs files in Drive.",
  parameters: z.object({
    maxResults: z.number().int().optional().default(20),
    folderId: z.string().optional(),
    query: z.string().optional(),
  }),
  execute: async (args) => {
    let q = "mimeType='application/vnd.google-apps.document' and trashed=false";
    if (args.query) q += ` and name contains '${args.query.replace(/'/g, "\\'")}' `;
    if (args.folderId) q += ` and '${args.folderId}' in parents`;
    const res = await driveClient.files.list({
      q,
      pageSize: args.maxResults,
      fields: "files(id,name,modifiedTime,webViewLink,parents)",
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    return JSON.stringify(res.data.files ?? [], null, 2);
  },
});

server.addTool({
  name: "searchGoogleDocs",
  description: "Full-text search across Google Docs content in Drive.",
  parameters: z.object({
    query: z.string(),
    maxResults: z.number().int().optional().default(10),
  }),
  execute: async (args) => {
    const q = `mimeType='application/vnd.google-apps.document' and trashed=false and fullText contains '${args.query.replace(/'/g, "\\'")}' `;
    const res = await driveClient.files.list({
      q,
      pageSize: args.maxResults,
      fields: "files(id,name,modifiedTime,webViewLink)",
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    return JSON.stringify(res.data.files ?? [], null, 2);
  },
});

server.addTool({
  name: "getDocumentInfo",
  description: "Get metadata about a file in Google Drive.",
  parameters: z.object({
    fileId: z.string(),
  }),
  execute: async (args) => {
    const res = await driveClient.files.get({
      fileId: args.fileId,
      fields: "id,name,mimeType,createdTime,modifiedTime,size,webViewLink,parents,owners,shared",
      supportsAllDrives: true,
    });
    return JSON.stringify(res.data, null, 2);
  },
});

server.addTool({
  name: "createFolder",
  description: "Create a folder in Google Drive.",
  parameters: z.object({
    name: z.string(),
    parentFolderId: z.string().optional(),
  }),
  execute: async (args) => {
    const res = await driveClient.files.create({
      requestBody: {
        name: args.name,
        mimeType: "application/vnd.google-apps.folder",
        parents: args.parentFolderId ? [args.parentFolderId] : undefined,
      },
      fields: "id,name,webViewLink",
      supportsAllDrives: true,
    });
    return JSON.stringify(res.data, null, 2);
  },
});

server.addTool({
  name: "moveFile",
  description: "Move a file to a different folder in Google Drive.",
  parameters: z.object({
    fileId: z.string(),
    newParentFolderId: z.string(),
  }),
  execute: async (args) => {
    const file = await driveClient.files.get({ fileId: args.fileId, fields: "parents", supportsAllDrives: true });
    const oldParents = (file.data.parents ?? []).join(",");
    await driveClient.files.update({
      fileId: args.fileId,
      addParents: args.newParentFolderId,
      removeParents: oldParents,
      fields: "id,parents",
      supportsAllDrives: true,
    });
    return `File moved to folder ${args.newParentFolderId}.`;
  },
});

server.addTool({
  name: "copyFile",
  description: "Copy a file in Google Drive.",
  parameters: z.object({
    fileId: z.string(),
    newName: z.string().optional(),
    destinationFolderId: z.string().optional(),
  }),
  execute: async (args) => {
    const res = await driveClient.files.copy({
      fileId: args.fileId,
      requestBody: {
        name: args.newName,
        parents: args.destinationFolderId ? [args.destinationFolderId] : undefined,
      },
      fields: "id,name,webViewLink",
      supportsAllDrives: true,
    });
    return JSON.stringify(res.data, null, 2);
  },
});

server.addTool({
  name: "createDocument",
  description: "Create a new Google Doc in Drive, optionally with initial markdown content.",
  parameters: z.object({
    title: z.string(),
    folderId: z.string().optional(),
    initialContent: z.string().optional(),
  }),
  execute: async (args) => {
    const file = await driveClient.files.create({
      requestBody: {
        name: args.title,
        mimeType: "application/vnd.google-apps.document",
        parents: args.folderId ? [args.folderId] : undefined,
      },
      fields: "id,name,webViewLink",
      supportsAllDrives: true,
    });
    if (args.initialContent && file.data.id) {
      await markdownToDocs(docsClient, file.data.id, args.initialContent, {});
    }
    return JSON.stringify(file.data, null, 2);
  },
});

server.addTool({
  name: "renameFile",
  description: "Rename a file in Google Drive.",
  parameters: z.object({
    fileId: z.string(),
    newName: z.string(),
  }),
  execute: async (args) => {
    await driveClient.files.update({
      fileId: args.fileId,
      requestBody: { name: args.newName },
      supportsAllDrives: true,
    });
    return `File renamed to "${args.newName}".`;
  },
});

server.addTool({
  name: "trashFile",
  description: "Move a file to trash (or permanently delete) in Google Drive.",
  parameters: z.object({
    fileId: z.string(),
    permanently: z.boolean().optional().default(false),
  }),
  execute: async (args) => {
    if (args.permanently) {
      await driveClient.files.delete({ fileId: args.fileId, supportsAllDrives: true });
      return "File permanently deleted.";
    }
    await driveClient.files.update({ fileId: args.fileId, requestBody: { trashed: true }, supportsAllDrives: true });
    return "File moved to trash.";
  },
});

server.addTool({
  name: "shareFile",
  description: "Share a Google Drive file with a user or make it public.",
  parameters: z.object({
    fileId: z.string(),
    role: z.enum(["reader", "commenter", "writer", "organizer", "fileOrganizer"]).optional().default("reader"),
    type: z.enum(["user", "group", "domain", "anyone"]).optional().default("user"),
    emailAddress: z.string().email().optional(),
    domain: z.string().optional(),
    sendNotification: z.boolean().optional().default(false),
  }),
  execute: async (args) => {
    const permBody: any = { role: args.role, type: args.type };
    if (args.emailAddress) permBody.emailAddress = args.emailAddress;
    if (args.domain) permBody.domain = args.domain;
    const res = await driveClient.permissions.create({
      fileId: args.fileId,
      sendNotificationEmail: args.sendNotification,
      requestBody: permBody,
      supportsAllDrives: true,
    });
    return JSON.stringify(res.data, null, 2);
  },
});

server.addTool({
  name: "getFilePermissions",
  description: "Get all sharing permissions for a Google Drive file.",
  parameters: z.object({
    fileId: z.string(),
  }),
  execute: async (args) => {
    const res = await driveClient.permissions.list({
      fileId: args.fileId,
      fields: "permissions(id,type,role,emailAddress,domain,displayName)",
      supportsAllDrives: true,
    });
    return JSON.stringify(res.data.permissions ?? [], null, 2);
  },
});

server.addTool({
  name: "listFolderContents",
  description: "List the contents of a folder in Google Drive.",
  parameters: z.object({
    folderId: z.string(),
    maxResults: z.number().int().optional().default(50),
    mimeType: z.string().optional(),
  }),
  execute: async (args) => {
    let q = `'${args.folderId}' in parents and trashed=false`;
    if (args.mimeType) q += ` and mimeType='${args.mimeType}'`;
    const res = await driveClient.files.list({
      q,
      pageSize: args.maxResults,
      fields: "files(id,name,mimeType,modifiedTime,size,webViewLink)",
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    return JSON.stringify(res.data.files ?? [], null, 2);
  },
});

server.addTool({
  name: "searchDriveFiles",
  description: "Search for files in Google Drive by content or name.",
  parameters: z.object({
    query: z.string(),
    maxResults: z.number().int().optional().default(20),
    mimeType: z.string().optional(),
    folderId: z.string().optional(),
  }),
  execute: async (args) => {
    let q = `trashed=false and fullText contains '${args.query.replace(/'/g, "\\'")}' `;
    if (args.mimeType) q += ` and mimeType='${args.mimeType}'`;
    if (args.folderId) q += ` and '${args.folderId}' in parents`;
    const res = await driveClient.files.list({
      q,
      pageSize: args.maxResults,
      fields: "files(id,name,mimeType,modifiedTime,webViewLink,parents)",
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    return JSON.stringify(res.data.files ?? [], null, 2);
  },
});

// ============================================================
// GMAIL TOOLS (13)
// ============================================================

server.addTool({
  name: "listMessages",
  description: "List Gmail messages matching a search query.",
  parameters: z.object({
    query: z.string().optional(),
    maxResults: z.number().int().optional().default(10),
    labelIds: z.array(z.string()).optional(),
  }),
  execute: async (args) => {
    const res = await gmailClient.users.messages.list({
      userId: "me",
      q: args.query,
      maxResults: args.maxResults,
      labelIds: args.labelIds,
    });
    return JSON.stringify(res.data.messages ?? [], null, 2);
  },
});

server.addTool({
  name: "getMessage",
  description: "Get the full content of a Gmail message.",
  parameters: z.object({
    messageId: z.string(),
    format: z.enum(["full", "metadata", "minimal", "raw"]).optional().default("full"),
  }),
  execute: async (args) => {
    const res = await gmailClient.users.messages.get({ userId: "me", id: args.messageId, format: args.format });
    const msg = res.data;
    const headers = msg.payload?.headers ?? [];
    const get = (name: string) => headers.find((h) => h.name === name)?.value ?? "";
    function extractBody(part: any): string {
      if (part?.body?.data) return Buffer.from(part.body.data, "base64").toString("utf-8");
      if (part?.parts) for (const p of part.parts) { const t = extractBody(p); if (t) return t; }
      return "";
    }
    return JSON.stringify({
      id: msg.id,
      subject: get("Subject"),
      from: get("From"),
      to: get("To"),
      date: get("Date"),
      snippet: msg.snippet,
      body: extractBody(msg.payload),
    }, null, 2);
  },
});

server.addTool({
  name: "sendEmail",
  description: "Send an email via Gmail.",
  parameters: z.object({
    to: z.string(),
    subject: z.string(),
    body: z.string(),
    cc: z.string().optional(),
    bcc: z.string().optional(),
    isHtml: z.boolean().optional().default(false),
    threadId: z.string().optional(),
  }),
  execute: async (args) => {
    const contentType = args.isHtml ? "text/html" : "text/plain";
    const headerLines = [`To: ${args.to}`, `Subject: ${args.subject}`, `Content-Type: ${contentType}; charset=utf-8`];
    if (args.cc) headerLines.push(`Cc: ${args.cc}`);
    if (args.bcc) headerLines.push(`Bcc: ${args.bcc}`);
    const raw = Buffer.from(`${headerLines.join("\r\n")}\r\n\r\n${args.body}`).toString("base64url");
    const res = await gmailClient.users.messages.send({
      userId: "me",
      requestBody: { raw, threadId: args.threadId },
    });
    return `Email sent. Message ID: ${res.data.id}`;
  },
});

server.addTool({
  name: "trashMessage",
  description: "Move a Gmail message to trash.",
  parameters: z.object({ messageId: z.string() }),
  execute: async (args) => {
    await gmailClient.users.messages.trash({ userId: "me", id: args.messageId });
    return "Message moved to trash.";
  },
});

server.addTool({
  name: "modifyMessageLabels",
  description: "Add or remove labels on a Gmail message.",
  parameters: z.object({
    messageId: z.string(),
    addLabelIds: z.array(z.string()).optional(),
    removeLabelIds: z.array(z.string()).optional(),
  }),
  execute: async (args) => {
    await gmailClient.users.messages.modify({
      userId: "me",
      id: args.messageId,
      requestBody: { addLabelIds: args.addLabelIds, removeLabelIds: args.removeLabelIds },
    });
    return "Labels updated.";
  },
});

server.addTool({
  name: "listLabels",
  description: "List all Gmail labels.",
  parameters: z.object({}),
  execute: async (_args) => {
    const res = await gmailClient.users.labels.list({ userId: "me" });
    return JSON.stringify(res.data.labels ?? [], null, 2);
  },
});

server.addTool({
  name: "createDraft",
  description: "Create a Gmail draft.",
  parameters: z.object({
    to: z.string(),
    subject: z.string(),
    body: z.string(),
    cc: z.string().optional(),
    isHtml: z.boolean().optional().default(false),
  }),
  execute: async (args) => {
    const contentType = args.isHtml ? "text/html" : "text/plain";
    const headerLines = [`To: ${args.to}`, `Subject: ${args.subject}`, `Content-Type: ${contentType}; charset=utf-8`];
    if (args.cc) headerLines.push(`Cc: ${args.cc}`);
    const raw = Buffer.from(`${headerLines.join("\r\n")}\r\n\r\n${args.body}`).toString("base64url");
    const res = await gmailClient.users.drafts.create({ userId: "me", requestBody: { message: { raw } } });
    return JSON.stringify({ draftId: res.data.id }, null, 2);
  },
});

server.addTool({
  name: "listDrafts",
  description: "List Gmail drafts.",
  parameters: z.object({
    maxResults: z.number().int().optional().default(10),
  }),
  execute: async (args) => {
    const res = await gmailClient.users.drafts.list({ userId: "me", maxResults: args.maxResults });
    return JSON.stringify(res.data.drafts ?? [], null, 2);
  },
});

server.addTool({
  name: "getDraft",
  description: "Get the content of a Gmail draft.",
  parameters: z.object({ draftId: z.string() }),
  execute: async (args) => {
    const res = await gmailClient.users.drafts.get({ userId: "me", id: args.draftId, format: "full" });
    return JSON.stringify(res.data, null, 2);
  },
});

server.addTool({
  name: "updateDraft",
  description: "Update the content of a Gmail draft.",
  parameters: z.object({
    draftId: z.string(),
    to: z.string(),
    subject: z.string(),
    body: z.string(),
    cc: z.string().optional(),
    isHtml: z.boolean().optional().default(false),
  }),
  execute: async (args) => {
    const contentType = args.isHtml ? "text/html" : "text/plain";
    const headerLines = [`To: ${args.to}`, `Subject: ${args.subject}`, `Content-Type: ${contentType}; charset=utf-8`];
    if (args.cc) headerLines.push(`Cc: ${args.cc}`);
    const raw = Buffer.from(`${headerLines.join("\r\n")}\r\n\r\n${args.body}`).toString("base64url");
    await gmailClient.users.drafts.update({ userId: "me", id: args.draftId, requestBody: { message: { raw } } });
    return "Draft updated.";
  },
});

server.addTool({
  name: "sendDraft",
  description: "Send an existing Gmail draft.",
  parameters: z.object({ draftId: z.string() }),
  execute: async (args) => {
    const res = await gmailClient.users.drafts.send({ userId: "me", requestBody: { id: args.draftId } });
    return `Draft sent. Message ID: ${res.data.id}`;
  },
});

server.addTool({
  name: "deleteDraft",
  description: "Delete a Gmail draft.",
  parameters: z.object({ draftId: z.string() }),
  execute: async (args) => {
    await gmailClient.users.drafts.delete({ userId: "me", id: args.draftId });
    return "Draft deleted.";
  },
});

server.addTool({
  name: "triageInbox",
  description: "Get a summary of unread inbox messages for triage.",
  parameters: z.object({
    maxMessages: z.number().int().optional().default(20),
    includeSnippets: z.boolean().optional().default(true),
  }),
  execute: async (args) => {
    const listRes = await gmailClient.users.messages.list({ userId: "me", q: "is:unread in:inbox", maxResults: args.maxMessages });
    const messages = listRes.data.messages ?? [];
    if (!messages.length) return "Inbox is empty (no unread messages).";
    const details = await Promise.all(
      messages.map((m) =>
        gmailClient.users.messages.get({ userId: "me", id: m.id!, format: "metadata", metadataHeaders: ["Subject", "From", "Date"] }),
      ),
    );
    const summaries = details.map((d) => {
      const h = d.data.payload?.headers ?? [];
      const get = (name: string) => h.find((x) => x.name === name)?.value ?? "";
      return {
        id: d.data.id,
        subject: get("Subject") || "(no subject)",
        from: get("From"),
        date: get("Date"),
        ...(args.includeSnippets ? { snippet: d.data.snippet } : {}),
      };
    });
    return JSON.stringify({ unreadCount: listRes.data.resultSizeEstimate, messages: summaries }, null, 2);
  },
});

// ============================================================
// CALENDAR TOOLS (5)
// ============================================================

server.addTool({
  name: "listEvents",
  description: "List Google Calendar events.",
  parameters: z.object({
    calendarId: z.string().optional().default("primary"),
    timeMin: z.string().optional(),
    timeMax: z.string().optional(),
    maxResults: z.number().int().optional().default(10),
    query: z.string().optional(),
    singleEvents: z.boolean().optional().default(true),
    orderBy: z.enum(["startTime", "updated"]).optional().default("startTime"),
  }),
  execute: async (args) => {
    const res = await calendarClient.events.list({
      calendarId: args.calendarId,
      timeMin: args.timeMin,
      timeMax: args.timeMax,
      maxResults: args.maxResults,
      q: args.query,
      singleEvents: args.singleEvents,
      orderBy: args.orderBy,
    });
    return JSON.stringify(
      (res.data.items ?? []).map((e) => ({
        id: e.id,
        summary: e.summary,
        start: e.start,
        end: e.end,
        location: e.location,
        description: e.description,
        status: e.status,
      })),
      null,
      2,
    );
  },
});

server.addTool({
  name: "createEvent",
  description: "Create a new Google Calendar event.",
  parameters: z.object({
    calendarId: z.string().optional().default("primary"),
    summary: z.string(),
    description: z.string().optional(),
    location: z.string().optional(),
    startDateTime: z.string(),
    endDateTime: z.string(),
    timeZone: z.string().optional().default("UTC"),
    attendees: z.array(z.string().email()).optional(),
    sendNotifications: z.boolean().optional().default(false),
    colorId: z.string().optional(),
    recurrence: z.array(z.string()).optional(),
  }),
  execute: async (args) => {
    const res = await calendarClient.events.insert({
      calendarId: args.calendarId,
      sendNotifications: args.sendNotifications,
      requestBody: {
        summary: args.summary,
        description: args.description,
        location: args.location,
        start: { dateTime: args.startDateTime, timeZone: args.timeZone },
        end: { dateTime: args.endDateTime, timeZone: args.timeZone },
        attendees: args.attendees?.map((email) => ({ email })),
        colorId: args.colorId,
        recurrence: args.recurrence,
      },
    });
    return JSON.stringify({ id: res.data.id, htmlLink: res.data.htmlLink }, null, 2);
  },
});

server.addTool({
  name: "updateEvent",
  description: "Update an existing Google Calendar event.",
  parameters: z.object({
    calendarId: z.string().optional().default("primary"),
    eventId: z.string(),
    summary: z.string().optional(),
    description: z.string().optional(),
    location: z.string().optional(),
    startDateTime: z.string().optional(),
    endDateTime: z.string().optional(),
    timeZone: z.string().optional(),
    attendees: z.array(z.string().email()).optional(),
    colorId: z.string().optional(),
  }),
  execute: async (args) => {
    const existing = await calendarClient.events.get({ calendarId: args.calendarId ?? "primary", eventId: args.eventId });
    const updated: any = { ...existing.data };
    if (args.summary) updated.summary = args.summary;
    if (args.description !== undefined) updated.description = args.description;
    if (args.location !== undefined) updated.location = args.location;
    if (args.startDateTime) updated.start = { dateTime: args.startDateTime, timeZone: args.timeZone ?? updated.start?.timeZone };
    if (args.endDateTime) updated.end = { dateTime: args.endDateTime, timeZone: args.timeZone ?? updated.end?.timeZone };
    if (args.attendees) updated.attendees = args.attendees.map((email) => ({ email }));
    if (args.colorId) updated.colorId = args.colorId;
    const res = await calendarClient.events.update({ calendarId: args.calendarId ?? "primary", eventId: args.eventId, requestBody: updated });
    return JSON.stringify({ id: res.data.id, htmlLink: res.data.htmlLink }, null, 2);
  },
});

server.addTool({
  name: "deleteEvent",
  description: "Delete a Google Calendar event.",
  parameters: z.object({
    calendarId: z.string().optional().default("primary"),
    eventId: z.string(),
    sendNotifications: z.boolean().optional().default(false),
  }),
  execute: async (args) => {
    await calendarClient.events.delete({ calendarId: args.calendarId, eventId: args.eventId, sendNotifications: args.sendNotifications });
    return "Event deleted.";
  },
});

server.addTool({
  name: "quickAddEvent",
  description: "Create a Google Calendar event from natural language (e.g. 'Lunch with Alice tomorrow at noon').",
  parameters: z.object({
    calendarId: z.string().optional().default("primary"),
    text: z.string(),
    sendNotifications: z.boolean().optional().default(false),
  }),
  execute: async (args) => {
    const res = await calendarClient.events.quickAdd({ calendarId: args.calendarId, text: args.text, sendNotifications: args.sendNotifications });
    return JSON.stringify({ id: res.data.id, summary: res.data.summary, start: res.data.start, htmlLink: res.data.htmlLink }, null, 2);
  },
});

// ============================================================
// START
// ============================================================

server.start({ transportType: "stdio" });

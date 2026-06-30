import type { docs_v1 } from "googleapis";
import { md } from "./index.js";
import { hexToRgb } from "../types.js";

export interface MarkdownToDocsOptions {
  preserveTitle?: boolean;
  firstHeadingAsTitle?: boolean;
  tabId?: string;
}

interface Segment {
  text: string;
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  code?: boolean;
  link?: string;
  headingLevel?: number;
  isBulletItem?: boolean;
  isOrderedItem?: boolean;
  nestingLevel?: number;
  isCodeBlock?: boolean;
  codeBlockLang?: string;
}

function parseMarkdownToSegments(markdown: string): Segment[] {
  const tokens = md.parse(markdown, {});
  const segments: Segment[] = [];

  interface InlineState {
    bold: boolean;
    italic: boolean;
    strikethrough: boolean;
    code: boolean;
    link?: string;
  }

  function defaultState(): InlineState {
    return { bold: false, italic: false, strikethrough: false, code: false };
  }

  let state = defaultState();
  const stateStack: InlineState[] = [];
  let headingLevel = 0;
  let isBulletItem = false;
  let isOrderedItem = false;
  let nestingLevel = 0;
  let inCodeBlock = false;
  let codeBlockLang = "";
  let inListItem = false;

  for (const token of tokens) {
    switch (token.type) {
      case "heading_open":
        headingLevel = parseInt(token.tag.slice(1), 10);
        break;
      case "heading_close":
        segments.push({ text: "\n", headingLevel });
        headingLevel = 0;
        break;
      case "bullet_list_open":
        nestingLevel++;
        isBulletItem = true;
        break;
      case "bullet_list_close":
        nestingLevel--;
        if (nestingLevel === 0) isBulletItem = false;
        break;
      case "ordered_list_open":
        nestingLevel++;
        isOrderedItem = true;
        break;
      case "ordered_list_close":
        nestingLevel--;
        if (nestingLevel === 0) isOrderedItem = false;
        break;
      case "list_item_open":
        inListItem = true;
        break;
      case "list_item_close":
        inListItem = false;
        break;
      case "fence":
        segments.push({
          text: token.content,
          isCodeBlock: true,
          codeBlockLang: token.info?.trim() || undefined,
        });
        break;
      case "code_block":
        segments.push({ text: token.content, isCodeBlock: true });
        break;
      case "paragraph_open":
        break;
      case "paragraph_close":
        if (!inListItem) segments.push({ text: "\n" });
        break;
      case "inline":
        if (!token.children) break;
        for (const child of token.children) {
          switch (child.type) {
            case "strong_open":
              stateStack.push({ ...state });
              state = { ...state, bold: true };
              break;
            case "strong_close":
              state = stateStack.pop() ?? defaultState();
              break;
            case "em_open":
              stateStack.push({ ...state });
              state = { ...state, italic: true };
              break;
            case "em_close":
              state = stateStack.pop() ?? defaultState();
              break;
            case "s_open":
              stateStack.push({ ...state });
              state = { ...state, strikethrough: true };
              break;
            case "s_close":
              state = stateStack.pop() ?? defaultState();
              break;
            case "link_open":
              stateStack.push({ ...state });
              state = { ...state, link: child.attrGet("href") ?? undefined };
              break;
            case "link_close":
              state = stateStack.pop() ?? defaultState();
              break;
            case "code_inline":
              segments.push({
                text: child.content,
                code: true,
                bold: state.bold,
                italic: state.italic,
              });
              break;
            case "text":
            case "softbreak":
              segments.push({
                text: child.type === "softbreak" ? "\n" : child.content,
                bold: state.bold || undefined,
                italic: state.italic || undefined,
                strikethrough: state.strikethrough || undefined,
                link: state.link,
                headingLevel: headingLevel || undefined,
                isBulletItem: (isBulletItem && inListItem) || undefined,
                isOrderedItem: (isOrderedItem && inListItem) || undefined,
                nestingLevel: nestingLevel > 0 ? nestingLevel - 1 : undefined,
              });
              break;
            case "hardbreak":
              segments.push({ text: "\n" });
              break;
          }
        }
        if (inListItem) segments.push({ text: "\n", isBulletItem: isBulletItem || undefined, isOrderedItem: isOrderedItem || undefined, nestingLevel: nestingLevel > 0 ? nestingLevel - 1 : undefined });
        break;
    }
  }

  return segments;
}

function buildRequestsFromSegments(
  segments: Segment[],
  insertAt: number,
  firstHeadingAsTitle = false,
  tabId?: string,
): docs_v1.Schema$Request[] {
  const requests: docs_v1.Schema$Request[] = [];
  let cursor = insertAt;
  let firstHeading = true;

  // First pass: build full text and track positions
  const positions: Array<{ start: number; end: number; seg: Segment }> = [];

  for (const seg of segments) {
    if (!seg.text) continue;

    if (seg.isCodeBlock) {
      // Code blocks: insert as 1x1 table with monospace styling
      requests.push({
        insertTable: {
          rows: 1,
          columns: 1,
          location: { index: cursor, ...(tabId ? { tabId } : {}) },
        },
      });
      // Table insert adds ~5 chars; we skip tracking for simplicity
      // The text goes inside the table cell at cursor+4
      const textLen = seg.text.length;
      requests.push({
        insertText: {
          text: seg.text,
          location: { index: cursor + 4, ...(tabId ? { tabId } : {}) },
        },
      });
      requests.push({
        updateTextStyle: {
          range: { startIndex: cursor + 4, endIndex: cursor + 4 + textLen, ...(tabId ? { tabId } : {}) },
          textStyle: {
            weightedFontFamily: { fontFamily: "Courier New" },
            fontSize: { magnitude: 9, unit: "PT" },
          },
          fields: "weightedFontFamily,fontSize",
        },
      });
      cursor += 4 + textLen + 1; // approximate
      continue;
    }

    const start = cursor;
    const end = cursor + seg.text.length;
    positions.push({ start, end, seg });
    cursor = end;
  }

  // Build one big InsertText from all non-codeblock text
  const allText = segments
    .filter((s) => !s.isCodeBlock)
    .map((s) => s.text)
    .join("");

  if (allText) {
    requests.unshift({
      insertText: {
        text: allText,
        location: { index: insertAt, ...(tabId ? { tabId } : {}) },
      },
    });
  }

  // Second pass: apply styles
  for (const { start, end, seg } of positions) {
    if (start === end) continue;
    const range = { startIndex: start, endIndex: end, ...(tabId ? { tabId } : {}) };

    // Paragraph / heading styles
    if (seg.headingLevel) {
      let namedStyleType: string;
      if (firstHeadingAsTitle && firstHeading && seg.headingLevel === 1) {
        namedStyleType = "TITLE";
        firstHeading = false;
      } else {
        namedStyleType = `HEADING_${seg.headingLevel}`;
        firstHeading = false;
      }
      requests.push({
        updateParagraphStyle: {
          range,
          paragraphStyle: { namedStyleType },
          fields: "namedStyleType",
        },
      });
    }

    // List items
    if (seg.isBulletItem || seg.isOrderedItem) {
      requests.push({
        createParagraphBullets: {
          range,
          bulletPreset: seg.isOrderedItem ? "NUMBERED_DECIMAL_ALPHA_ROMAN" : "BULLET_DISC_CIRCLE_SQUARE",
        },
      });
    }

    // Text styles
    const textStyle: docs_v1.Schema$TextStyle = {};
    const fields: string[] = [];

    if (seg.bold) { textStyle.bold = true; fields.push("bold"); }
    if (seg.italic) { textStyle.italic = true; fields.push("italic"); }
    if (seg.strikethrough) { textStyle.strikethrough = true; fields.push("strikethrough"); }
    if (seg.link) { textStyle.link = { url: seg.link }; fields.push("link"); }
    if (seg.code) {
      textStyle.weightedFontFamily = { fontFamily: "Courier New" };
      textStyle.foregroundColor = { color: { rgbColor: hexToRgb("#16a34a") } }; // green
      textStyle.backgroundColor = { color: { rgbColor: hexToRgb("#f3f4f6") } }; // gray
      fields.push("weightedFontFamily", "foregroundColor", "backgroundColor");
    }

    if (fields.length) {
      requests.push({ updateTextStyle: { range, textStyle, fields: fields.join(",") } });
    }
  }

  return requests;
}

export async function markdownToDocs(
  docs: docs_v1.Docs,
  documentId: string,
  markdown: string,
  options: MarkdownToDocsOptions = {},
): Promise<void> {
  const { firstHeadingAsTitle = false, tabId } = options;

  // Get current doc to find end index
  const doc = await docs.documents.get({ documentId });
  const body = doc.data.tabs?.[0]?.documentTab?.body ?? doc.data.body;
  const content = body?.content ?? [];
  const endIndex = content.length
    ? (content[content.length - 1].endIndex ?? 2) - 1
    : 1;

  const requests: docs_v1.Schema$Request[] = [];

  // Delete existing content (leave the trailing newline at index 1)
  if (endIndex > 1) {
    requests.push({
      deleteContentRange: {
        range: { startIndex: 1, endIndex, ...(tabId ? { tabId } : {}) },
      },
    });
  }

  const segments = parseMarkdownToSegments(markdown);
  const insertRequests = buildRequestsFromSegments(segments, 1, firstHeadingAsTitle, tabId);
  requests.push(...insertRequests);

  if (requests.length) {
    await docs.documents.batchUpdate({
      documentId,
      requestBody: { requests },
    });
  }
}

export async function appendMarkdownToDocs(
  docs: docs_v1.Docs,
  documentId: string,
  markdown: string,
  options: MarkdownToDocsOptions & { addNewlineIfNeeded?: boolean } = {},
): Promise<void> {
  const { firstHeadingAsTitle = false, addNewlineIfNeeded = true, tabId } = options;

  const doc = await docs.documents.get({ documentId });
  const body = doc.data.tabs?.[0]?.documentTab?.body ?? doc.data.body;
  const content = body?.content ?? [];
  const endIndex = content.length
    ? (content[content.length - 1].endIndex ?? 2) - 1
    : 1;

  const requests: docs_v1.Schema$Request[] = [];
  let insertAt = endIndex;

  if (addNewlineIfNeeded && endIndex > 1) {
    requests.push({
      insertText: {
        text: "\n",
        location: { index: endIndex, ...(tabId ? { tabId } : {}) },
      },
    });
    insertAt = endIndex + 1;
  }

  const segments = parseMarkdownToSegments(markdown);
  const insertRequests = buildRequestsFromSegments(segments, insertAt, firstHeadingAsTitle, tabId);
  requests.push(...insertRequests);

  if (requests.length) {
    await docs.documents.batchUpdate({
      documentId,
      requestBody: { requests },
    });
  }
}

import type { docs_v1 } from "googleapis";
import { hexToRgb, type TextStyleParams, type ParagraphStyleParams } from "./types.js";

type DocsClient = docs_v1.Docs;

export async function findTextRange(
  docs: DocsClient,
  documentId: string,
  textToFind: string,
  matchInstance = 1,
  tabId?: string,
): Promise<{ startIndex: number; endIndex: number }> {
  const doc = await docs.documents.get({ documentId });
  const content =
    doc.data.tabs?.[0]?.documentTab?.body?.content ??
    doc.data.body?.content ??
    [];

  let found = 0;

  function search(
    items: docs_v1.Schema$StructuralElement[],
  ): { startIndex: number; endIndex: number } | null {
    for (const elem of items) {
      if (elem.paragraph) {
        for (const pe of elem.paragraph.elements ?? []) {
          const text = pe.textRun?.content ?? "";
          let from = 0;
          let idx: number;
          while ((idx = text.indexOf(textToFind, from)) !== -1) {
            found++;
            if (found === matchInstance) {
              const base = pe.startIndex ?? 0;
              return { startIndex: base + idx, endIndex: base + idx + textToFind.length };
            }
            from = idx + 1;
          }
        }
      } else if (elem.table) {
        for (const row of elem.table.tableRows ?? []) {
          for (const cell of row.tableCells ?? []) {
            const r = search(cell.content as docs_v1.Schema$StructuralElement[]);
            if (r) return r;
          }
        }
      }
    }
    return null;
  }

  const result = search(content as docs_v1.Schema$StructuralElement[]);
  if (!result) throw new Error(`"${textToFind}" not found (instance ${matchInstance})`);
  return result;
}

export async function executeBatchUpdate(
  docs: DocsClient,
  documentId: string,
  requests: docs_v1.Schema$Request[],
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
  const res = await docs.documents.batchUpdate({
    documentId,
    requestBody: { requests },
  });
  return res.data;
}

export async function executeBatchUpdateWithSplitting(
  docs: DocsClient,
  documentId: string,
  requests: docs_v1.Schema$Request[],
  chunkSize = 50,
): Promise<void> {
  for (let i = 0; i < requests.length; i += chunkSize) {
    await executeBatchUpdate(docs, documentId, requests.slice(i, i + chunkSize));
  }
}

export function buildTextStyleRequest(
  startIndex: number,
  endIndex: number,
  style: TextStyleParams,
  tabId?: string,
): docs_v1.Schema$Request {
  const textStyle: docs_v1.Schema$TextStyle = {};
  const fields: string[] = [];

  if (style.bold !== undefined) { textStyle.bold = style.bold; fields.push("bold"); }
  if (style.italic !== undefined) { textStyle.italic = style.italic; fields.push("italic"); }
  if (style.underline !== undefined) { textStyle.underline = style.underline; fields.push("underline"); }
  if (style.strikethrough !== undefined) { textStyle.strikethrough = style.strikethrough; fields.push("strikethrough"); }
  if (style.fontSize !== undefined) {
    textStyle.fontSize = { magnitude: style.fontSize, unit: "PT" };
    fields.push("fontSize");
  }
  if (style.foregroundColor) {
    textStyle.foregroundColor = { color: { rgbColor: hexToRgb(style.foregroundColor) } };
    fields.push("foregroundColor");
  }
  if (style.backgroundColor) {
    textStyle.backgroundColor = { color: { rgbColor: hexToRgb(style.backgroundColor) } };
    fields.push("backgroundColor");
  }
  if (style.fontFamily) {
    textStyle.weightedFontFamily = { fontFamily: style.fontFamily };
    fields.push("weightedFontFamily");
  }
  if (style.link) {
    textStyle.link = { url: style.link };
    fields.push("link");
  }

  return {
    updateTextStyle: {
      range: { startIndex, endIndex, ...(tabId ? { tabId } : {}) },
      textStyle,
      fields: fields.join(","),
    },
  };
}

export function buildParagraphStyleRequest(
  startIndex: number,
  endIndex: number,
  style: ParagraphStyleParams,
  tabId?: string,
): docs_v1.Schema$Request {
  const paragraphStyle: docs_v1.Schema$ParagraphStyle = {};
  const fields: string[] = [];

  if (style.alignment) { paragraphStyle.alignment = style.alignment; fields.push("alignment"); }
  if (style.lineSpacing !== undefined) { paragraphStyle.lineSpacing = style.lineSpacing; fields.push("lineSpacing"); }
  if (style.spaceAbove !== undefined) {
    paragraphStyle.spaceAbove = { magnitude: style.spaceAbove, unit: "PT" };
    fields.push("spaceAbove");
  }
  if (style.spaceBelow !== undefined) {
    paragraphStyle.spaceBelow = { magnitude: style.spaceBelow, unit: "PT" };
    fields.push("spaceBelow");
  }
  if (style.indentStart !== undefined) {
    paragraphStyle.indentStart = { magnitude: style.indentStart, unit: "PT" };
    fields.push("indentStart");
  }
  if (style.indentEnd !== undefined) {
    paragraphStyle.indentEnd = { magnitude: style.indentEnd, unit: "PT" };
    fields.push("indentEnd");
  }
  if (style.namedStyleType) { paragraphStyle.namedStyleType = style.namedStyleType; fields.push("namedStyleType"); }

  return {
    updateParagraphStyle: {
      range: { startIndex, endIndex, ...(tabId ? { tabId } : {}) },
      paragraphStyle,
      fields: fields.join(","),
    },
  };
}

export function getDocumentEndIndex(doc: docs_v1.Schema$Document): number {
  const content =
    doc.tabs?.[0]?.documentTab?.body?.content ?? doc.body?.content ?? [];
  if (!content.length) return 1;
  const last = content[content.length - 1];
  return (last.endIndex ?? 2) - 1;
}

export function extractPlainText(doc: docs_v1.Schema$Document): string {
  const content =
    doc.tabs?.[0]?.documentTab?.body?.content ?? doc.body?.content ?? [];
  const parts: string[] = [];

  function walk(items: docs_v1.Schema$StructuralElement[]) {
    for (const elem of items) {
      if (elem.paragraph) {
        for (const pe of elem.paragraph.elements ?? []) {
          if (pe.textRun?.content) parts.push(pe.textRun.content);
        }
      } else if (elem.table) {
        for (const row of elem.table.tableRows ?? []) {
          for (const cell of row.tableCells ?? []) {
            walk(cell.content as docs_v1.Schema$StructuralElement[]);
          }
        }
      }
    }
  }

  walk(content as docs_v1.Schema$StructuralElement[]);
  return parts.join("");
}

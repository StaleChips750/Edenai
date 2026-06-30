import type { sheets_v4 } from "googleapis";
import { hexToRgb, type CellFormat } from "./types.js";

type SheetsClient = sheets_v4.Sheets;

export function parseA1Notation(range: string): {
  sheetName?: string;
  startCol?: number;
  startRow?: number;
  endCol?: number;
  endRow?: number;
} {
  const bang = range.indexOf("!");
  const sheetName = bang !== -1 ? range.slice(0, bang).replace(/^'|'$/g, "") : undefined;
  const cellPart = bang !== -1 ? range.slice(bang + 1) : range;

  const match = cellPart.match(/^([A-Z]*)?(\d*)?(?::([A-Z]*)?(\d*)?)?$/i);
  if (!match) return { sheetName };

  const [, sc = "", sr = "", ec = "", er = ""] = match;

  function colToIdx(col: string): number | undefined {
    if (!col) return undefined;
    let n = 0;
    for (const c of col.toUpperCase()) n = n * 26 + c.charCodeAt(0) - 64;
    return n - 1;
  }

  return {
    sheetName,
    startCol: colToIdx(sc),
    startRow: sr ? parseInt(sr, 10) - 1 : undefined,
    endCol: colToIdx(ec),
    endRow: er ? parseInt(er, 10) - 1 : undefined,
  };
}

export function columnIndexToLetter(index: number): string {
  let result = "";
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

export async function getSheetIdByName(
  sheets: SheetsClient,
  spreadsheetId: string,
  sheetName: string,
): Promise<number> {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = res.data.sheets?.find((s) => s.properties?.title === sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);
  return sheet.properties!.sheetId!;
}

export async function getAllSheets(
  sheets: SheetsClient,
  spreadsheetId: string,
): Promise<sheets_v4.Schema$Sheet[]> {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  return res.data.sheets ?? [];
}

export function hexToSheetsColor(hex: string): sheets_v4.Schema$Color {
  return hexToRgb(hex);
}

export function buildGridRange(
  sheetId: number,
  parsed: { startRow?: number; startCol?: number; endRow?: number; endCol?: number },
): sheets_v4.Schema$GridRange {
  return {
    sheetId,
    startRowIndex: parsed.startRow,
    endRowIndex: parsed.endRow !== undefined ? parsed.endRow + 1 : undefined,
    startColumnIndex: parsed.startCol,
    endColumnIndex: parsed.endCol !== undefined ? parsed.endCol + 1 : undefined,
  };
}

export function buildCellFormat(fmt: CellFormat): sheets_v4.Schema$CellFormat {
  const cellFormat: sheets_v4.Schema$CellFormat = {};

  if (fmt.backgroundColor) {
    cellFormat.backgroundColor = hexToSheetsColor(fmt.backgroundColor);
  }
  if (fmt.horizontalAlignment) cellFormat.horizontalAlignment = fmt.horizontalAlignment;
  if (fmt.verticalAlignment) cellFormat.verticalAlignment = fmt.verticalAlignment;
  if (fmt.wrapStrategy) cellFormat.wrapStrategy = fmt.wrapStrategy;
  if (fmt.numberFormat) cellFormat.numberFormat = fmt.numberFormat;

  const textFmt: sheets_v4.Schema$TextFormat = {};
  if (fmt.bold !== undefined) textFmt.bold = fmt.bold;
  if (fmt.italic !== undefined) textFmt.italic = fmt.italic;
  if (fmt.strikethrough !== undefined) textFmt.strikethrough = fmt.strikethrough;
  if (fmt.underline !== undefined) textFmt.underline = fmt.underline;
  if (fmt.fontSize !== undefined) textFmt.fontSize = fmt.fontSize;
  if (fmt.foregroundColor) textFmt.foregroundColor = hexToSheetsColor(fmt.foregroundColor);
  if (fmt.fontFamily) textFmt.fontFamily = fmt.fontFamily;

  if (Object.keys(textFmt).length) cellFormat.textFormat = textFmt;
  return cellFormat;
}

export async function findNamedRange(
  sheets: SheetsClient,
  spreadsheetId: string,
  name: string,
): Promise<sheets_v4.Schema$NamedRange | null> {
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "namedRanges",
  });
  return res.data.namedRanges?.find((r) => r.name === name) ?? null;
}

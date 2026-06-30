import type { docs_v1 } from "googleapis";

function rgbToHex(r: number, g: number, b: number): string {
  const toHex = (n: number) => Math.round(n * 255).toString(16).padStart(2, "0");
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function processTextRun(
  run: docs_v1.Schema$TextRun,
): string {
  let text = run.content ?? "";
  const style = run.textStyle ?? {};

  // Inline code detection: monospace font
  const isCode = style.weightedFontFamily?.fontFamily?.toLowerCase().includes("mono") ||
    style.weightedFontFamily?.fontFamily?.toLowerCase().includes("courier");

  if (isCode) return `\`${text}\``;

  if (style.bold && style.italic) text = `***${text}***`;
  else if (style.bold) text = `**${text}**`;
  else if (style.italic) text = `*${text}*`;
  if (style.strikethrough) text = `~~${text}~~`;
  if (style.link?.url) text = `[${text}](${style.link.url})`;

  return text;
}

function paragraphToMarkdown(
  para: docs_v1.Schema$Paragraph,
): string {
  const style = para.paragraphStyle ?? {};
  const namedStyle = style.namedStyleType ?? "NORMAL_TEXT";
  const bullet = para.bullet;

  const textParts: string[] = [];
  for (const elem of para.elements ?? []) {
    if (elem.textRun) textParts.push(processTextRun(elem.textRun));
    else if (elem.inlineObjectElement) textParts.push("[image]");
    else if (elem.pageBreak) textParts.push("");
  }
  let text = textParts.join("").replace(/\n$/, "");

  // Heading
  if (namedStyle === "TITLE") return `# ${text}`;
  if (namedStyle === "SUBTITLE") return `## ${text}`;
  const headingMatch = namedStyle.match(/^HEADING_(\d)$/);
  if (headingMatch) {
    const level = parseInt(headingMatch[1], 10);
    return `${"#".repeat(level)} ${text}`;
  }

  // List
  if (bullet) {
    const indent = "  ".repeat(bullet.nestingLevel ?? 0);
    const listId = bullet.listId;
    // We don't have the full lists map here, so default to bullet
    return `${indent}- ${text}`;
  }

  return text;
}

function tableToMarkdown(table: docs_v1.Schema$Table): string {
  const rows = table.tableRows ?? [];
  if (!rows.length) return "";

  const lines: string[] = [];

  rows.forEach((row, rowIdx) => {
    const cells = row.tableCells ?? [];
    const cellTexts = cells.map((cell) => {
      const parts: string[] = [];
      for (const elem of cell.content ?? []) {
        if (elem.paragraph) parts.push(paragraphToMarkdown(elem.paragraph).trim());
      }
      return parts.join(" ");
    });
    lines.push(`| ${cellTexts.join(" | ")} |`);
    if (rowIdx === 0) {
      lines.push(`| ${cells.map(() => "---").join(" | ")} |`);
    }
  });

  return lines.join("\n");
}

export function docsToMarkdown(doc: docs_v1.Schema$Document): string {
  const content =
    doc.tabs?.[0]?.documentTab?.body?.content ?? doc.body?.content ?? [];

  const parts: string[] = [];

  for (const elem of content as docs_v1.Schema$StructuralElement[]) {
    if (elem.paragraph) {
      const line = paragraphToMarkdown(elem.paragraph);
      if (line !== undefined) parts.push(line);
    } else if (elem.table) {
      parts.push("");
      parts.push(tableToMarkdown(elem.table));
      parts.push("");
    } else if (elem.sectionBreak) {
      parts.push("");
      parts.push("---");
      parts.push("");
    }
  }

  return parts.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

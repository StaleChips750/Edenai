import { z } from "zod";

export const hexColorSchema = z
  .string()
  .regex(/^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/, "Must be #RGB or #RRGGBB");

export function hexToRgb(hex: string): { red: number; green: number; blue: number } {
  const h = hex.replace("#", "");
  const full = h.length === 3 ? h.split("").map((c) => c + c).join("") : h;
  return {
    red: parseInt(full.slice(0, 2), 16) / 255,
    green: parseInt(full.slice(2, 4), 16) / 255,
    blue: parseInt(full.slice(4, 6), 16) / 255,
  };
}

export const alignmentSchema = z.enum(["START", "END", "CENTER", "JUSTIFIED"]);

export const namedStyleTypeSchema = z.enum([
  "NORMAL_TEXT",
  "TITLE",
  "SUBTITLE",
  "HEADING_1",
  "HEADING_2",
  "HEADING_3",
  "HEADING_4",
  "HEADING_5",
  "HEADING_6",
]);

export const textStyleParamsSchema = z.object({
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  strikethrough: z.boolean().optional(),
  fontSize: z.number().positive().optional(),
  foregroundColor: hexColorSchema.optional(),
  backgroundColor: hexColorSchema.optional(),
  fontFamily: z.string().optional(),
  link: z.string().url().optional(),
});

export const paragraphStyleParamsSchema = z.object({
  alignment: alignmentSchema.optional(),
  lineSpacing: z.number().optional(),
  spaceAbove: z.number().optional(),
  spaceBelow: z.number().optional(),
  indentStart: z.number().optional(),
  indentEnd: z.number().optional(),
  namedStyleType: namedStyleTypeSchema.optional(),
});

export const borderStyleSchema = z.enum([
  "DOTTED",
  "DASHED",
  "SOLID",
  "SOLID_MEDIUM",
  "SOLID_THICK",
  "NONE",
  "DOUBLE",
]);

export const cellFormatSchema = z.object({
  backgroundColor: hexColorSchema.optional(),
  horizontalAlignment: z.enum(["LEFT", "CENTER", "RIGHT"]).optional(),
  verticalAlignment: z.enum(["TOP", "MIDDLE", "BOTTOM"]).optional(),
  wrapStrategy: z.enum(["OVERFLOW_CELL", "CLIP", "WRAP"]).optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  strikethrough: z.boolean().optional(),
  underline: z.boolean().optional(),
  fontSize: z.number().positive().optional(),
  foregroundColor: hexColorSchema.optional(),
  fontFamily: z.string().optional(),
  numberFormat: z
    .object({
      type: z.enum(["TEXT", "NUMBER", "PERCENT", "CURRENCY", "DATE", "TIME", "DATE_TIME", "SCIENTIFIC"]),
      pattern: z.string().optional(),
    })
    .optional(),
});

export type TextStyleParams = z.infer<typeof textStyleParamsSchema>;
export type ParagraphStyleParams = z.infer<typeof paragraphStyleParamsSchema>;
export type CellFormat = z.infer<typeof cellFormatSchema>;

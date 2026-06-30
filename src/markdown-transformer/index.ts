import MarkdownIt from "markdown-it";

export const md = new MarkdownIt({
  html: false,
  breaks: true,
  linkify: true,
  typographer: false,
});

export { markdownToDocs, type MarkdownToDocsOptions } from "./markdownToDocs.js";
export { docsToMarkdown } from "./docsToMarkdown.js";

#!/usr/bin/env node

/**
 * DOCX MCP Server — Local MCP server for reading, editing, formatting,
 * commenting, and highlighting Word documents.
 *
 * Transport: stdio (runs locally, no file uploads)
 * Usage with Claude Code:  Add to ~/.claude/settings.json under mcpServers
 * Usage with Cursor:       Add to MCP server configuration
 */

import { createRequire } from "node:module";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  readDocument,
  getDocumentInfo,
  searchText,
  replaceText,
  editParagraph,
  editParagraphs,
  insertParagraph,
  insertParagraphs,
  deleteParagraph,
  deleteParagraphs,
  formatText,
  setParagraphFormat,
  setParagraphFormats,
  addComment,
  addComments,
  readComments,
  replyToComment,
  deleteComment,
  createDocument,
  applyDocumentPreset,
  highlightText,
  insertTable,
  setHeading,
  setHeadings,
  getPageLayout,
  setPageLayout,
  acceptAllChanges,
  rejectAllChanges,
  readHeaderFooter,
  editTableCell,
  editTableCells,
  readFootnotes,
  listImages,
  EngineError,
  ErrorCode,
} from "./docx-engine.js";

const require = createRequire(import.meta.url);
const { version: VERSION } = require("../package.json") as { version: string };

function formatError(e: unknown): string {
  if (e instanceof EngineError) {
    return `[${e.code}] ${e.message}`;
  }
  if (e instanceof Error) {
    return `[INTERNAL_ERROR] ${e.message}`;
  }
  return `[INTERNAL_ERROR] ${String(e)}`;
}

// Server-side enforcement of F-004 (regulated-industry safety): the LLM cannot
// silently disable tracked changes by passing track_changes=false alone.
// Requires a second, independently named capability flag. Belt-and-braces
// against prompt injection or long-context drift.
function assertTrackChanges(
  track_changes: boolean | undefined,
  allow_untracked_edit: boolean | undefined,
): void {
  if (track_changes === false && allow_untracked_edit !== true) {
    throw new EngineError(
      ErrorCode.UNTRACKED_EDIT_NOT_ALLOWED,
      "track_changes=false requires allow_untracked_edit=true. " +
        "This is a safety check for regulated-industry use: silent edits to legal/regulated documents must be opted into with two independent flags.",
    );
  }
}

// ---------------------------------------------------------------------------
// Server setup
// ---------------------------------------------------------------------------

const server = new McpServer({
  name: "docx-editor",
  version: VERSION,
  description: [
    "Read, edit, format, comment, and manage Word (.docx) documents.",
    "",
    "Supported: paragraph read/write, tracked changes (w:ins/w:del), text formatting",
    "(bold/italic/underline/font/color/highlight), paragraph formatting (alignment/spacing/indent),",
    "headings, tables, comments with threading, page layout, headers/footers, footnotes, images.",
    "",
    "NOT supported (use python-docx or direct XML instead):",
    "- Embedded chart editing",
    "- Form fields / content controls editing",
    "- Macro execution (.docm)",
    "- Image insertion or modification (read-only via list_images)",
    "- Style definition creation (applies inline formatting, not named styles)",
  ].join("\n"),
});

// ---------------------------------------------------------------------------
// Tool: read_document
// ---------------------------------------------------------------------------

server.tool(
  "read_document",
  "Read the content of a DOCX file. Returns paragraphs with indices, styles, and formatting hints. Use start_paragraph/end_paragraph for large documents. Use show_revisions to see tracked changes annotations.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    start_paragraph: z
      .number()
      .optional()
      .describe("Start reading from this block index (inclusive)"),
    end_paragraph: z
      .number()
      .optional()
      .describe("Stop reading at this block index (exclusive)"),
    show_revisions: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        "Show tracked changes with annotations: [-deleted-] and [+inserted+]. Default false shows accepted text only.",
      ),
  },
  async ({ file_path, start_paragraph, end_paragraph, show_revisions }) => {
    try {
      const result = await readDocument(
        file_path,
        start_paragraph,
        end_paragraph,
        show_revisions,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: get_document_info
// ---------------------------------------------------------------------------

server.tool(
  "get_document_info",
  "Get metadata and structure overview of a DOCX file — paragraph count, headings outline, tables, comment count.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
  },
  async ({ file_path }) => {
    try {
      const result = await getDocumentInfo(file_path);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: search_text
// ---------------------------------------------------------------------------

server.tool(
  "search_text",
  "Search for text in a DOCX file. Returns matching blocks with context.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    query: z.string().describe("Text to search for"),
    case_sensitive: z
      .boolean()
      .optional()
      .default(false)
      .describe("Case-sensitive search"),
  },
  async ({ file_path, query, case_sensitive }) => {
    try {
      const result = await searchText(file_path, query, case_sensitive);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: replace_text
// ---------------------------------------------------------------------------

server.tool(
  "replace_text",
  "Find and replace text throughout a DOCX file. Handles text that spans multiple runs. Returns the number of replacements made.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    search: z.string().describe("Text to find"),
    replace: z.string().describe("Replacement text"),
    case_sensitive: z
      .boolean()
      .optional()
      .default(false)
      .describe("Case-sensitive matching"),
    track_changes: z
      .boolean()
      .optional()
      .default(true)
      .describe(
        "Record edits as tracked changes (w:del/w:ins) so they appear as revisions in Word. Default true.",
      ),
    author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Author name for tracked changes"),
    allow_untracked_edit: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        "Capability flag required to disable tracked changes. When track_changes is false, this must also be true or the call fails with UNTRACKED_EDIT_NOT_ALLOWED. Default false. This is a safety guard against prompt injection or long-context drift in regulated-industry use — silent edits to legal/regulated documents must be opted into with two independent flags.",
      ),
    include_headers_footers: z
      .boolean()
      .optional()
      .default(false)
      .describe("Also replace text in headers and footers. Default false."),
  },
  async ({ file_path, search, replace, case_sensitive, track_changes, author, include_headers_footers, allow_untracked_edit }) => {
    try {
      assertTrackChanges(track_changes, allow_untracked_edit);
      const result = await replaceText(
        file_path,
        search,
        replace,
        case_sensitive,
        track_changes,
        author,
        include_headers_footers,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: edit_paragraph
// ---------------------------------------------------------------------------

server.tool(
  "edit_paragraph",
  "Replace the entire text content of a specific paragraph (by index). Preserves the paragraph's style and formatting properties.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    paragraph_index: z
      .number()
      .describe("Index of the paragraph to edit (from read_document output)"),
    new_text: z.string().describe("New text content for the paragraph"),
    track_changes: z
      .boolean()
      .optional()
      .default(true)
      .describe(
        "Record edits as tracked changes (w:del/w:ins) so they appear as revisions in Word. Default true.",
      ),
    author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Author name for tracked changes"),
    allow_untracked_edit: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        "Capability flag required to disable tracked changes. When track_changes is false, this must also be true or the call fails with UNTRACKED_EDIT_NOT_ALLOWED. Default false. This is a safety guard against prompt injection or long-context drift in regulated-industry use — silent edits to legal/regulated documents must be opted into with two independent flags.",
      ),
  },
  async ({ file_path, paragraph_index, new_text, track_changes, author, allow_untracked_edit }) => {
    try {
      assertTrackChanges(track_changes, allow_untracked_edit);
      const result = await editParagraph(
        file_path,
        paragraph_index,
        new_text,
        track_changes,
        author,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: edit_paragraphs (bulk)
// ---------------------------------------------------------------------------

server.tool(
  "edit_paragraphs",
  "Replace the text content of multiple paragraphs in one operation. Opens and saves the file only once. Paragraph indices remain stable because edits don't change paragraph count.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    edits: z
      .array(
        z.object({
          paragraph_index: z
            .number()
            .describe("Index of the paragraph to edit"),
          new_text: z.string().describe("New text content for the paragraph"),
        }),
      )
      .describe("Array of paragraph edits"),
    track_changes: z
      .boolean()
      .optional()
      .default(true)
      .describe(
        "Record edits as tracked changes (w:del/w:ins). Default true.",
      ),
    author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Author name for tracked changes"),
    allow_untracked_edit: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        "Capability flag required to disable tracked changes. When track_changes is false, this must also be true or the call fails with UNTRACKED_EDIT_NOT_ALLOWED. Default false. This is a safety guard against prompt injection or long-context drift in regulated-industry use — silent edits to legal/regulated documents must be opted into with two independent flags.",
      ),
  },
  async ({ file_path, edits, track_changes, author, allow_untracked_edit }) => {
    try {
      assertTrackChanges(track_changes, allow_untracked_edit);
      const engineEdits = edits.map((e) => ({
        paragraphIndex: e.paragraph_index,
        newText: e.new_text,
      }));
      const result = await editParagraphs(
        file_path,
        engineEdits,
        track_changes,
        author,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: insert_paragraph
// ---------------------------------------------------------------------------

server.tool(
  "insert_paragraph",
  "Insert a new paragraph at a specific position. Use style names like 'Heading1', 'Heading2', 'Normal', etc. Use position=-1 to append at the end. To reproduce Word's list-based numbering (e.g. 第1条, 第2条…), use num_id/num_level or copy_format_from.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    text: z.string().describe("Text content of the new paragraph"),
    position: z
      .number()
      .describe("Block index to insert before (-1 for end of document)"),
    style: z
      .string()
      .optional()
      .describe("Paragraph style (e.g., 'Heading1', 'Heading2', 'Normal')"),
    num_id: z
      .number()
      .optional()
      .describe("Numbering definition ID (w:numId). Produces <w:numPr> in the paragraph properties. Use with num_level. Ignored if copy_format_from is set."),
    num_level: z
      .number()
      .optional()
      .default(0)
      .describe("Numbering indentation level (w:ilvl), 0-based. Default 0."),
    copy_format_from: z
      .number()
      .optional()
      .describe("Block index of an existing paragraph whose w:pPr to deep-copy (numbering, indentation, spacing, borders, etc.). When set, style/num_id/num_level are ignored."),
    track_changes: z
      .boolean()
      .optional()
      .default(true)
      .describe(
        "Record insertion as a tracked change so it appears as a revision in Word. Default true.",
      ),
    author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Author name for tracked changes"),
    allow_untracked_edit: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        "Capability flag required to disable tracked changes. When track_changes is false, this must also be true or the call fails with UNTRACKED_EDIT_NOT_ALLOWED. Default false. This is a safety guard against prompt injection or long-context drift in regulated-industry use — silent edits to legal/regulated documents must be opted into with two independent flags.",
      ),
  },
  async ({ file_path, text, position, style, track_changes, author, num_id, num_level, copy_format_from, allow_untracked_edit }) => {
    try {
      assertTrackChanges(track_changes, allow_untracked_edit);
      const result = await insertParagraph(
        file_path,
        text,
        position,
        style,
        track_changes,
        author,
        num_id,
        num_level,
        copy_format_from,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: insert_paragraphs (bulk)
// ---------------------------------------------------------------------------

server.tool(
  "insert_paragraphs",
  "Insert multiple paragraphs in one operation. Handles index shifting internally by processing in reverse order. Opens and saves the file only once. Supports numbering (num_id/num_level) and format copying (copy_format_from).",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    paragraphs: z
      .array(
        z.object({
          text: z.string().describe("Text content of the new paragraph"),
          position: z
            .number()
            .describe("Block index to insert before (-1 for end of document)"),
          style: z
            .string()
            .optional()
            .describe("Paragraph style (e.g., 'Heading1', 'Normal')"),
          num_id: z
            .number()
            .optional()
            .describe("Numbering definition ID (w:numId). Ignored if copy_format_from is set."),
          num_level: z
            .number()
            .optional()
            .default(0)
            .describe("Numbering indentation level (w:ilvl), 0-based. Default 0."),
          copy_format_from: z
            .number()
            .optional()
            .describe("Block index of an existing paragraph whose w:pPr to deep-copy. When set, style/num_id/num_level are ignored."),
        }),
      )
      .describe("Array of paragraphs to insert"),
    track_changes: z
      .boolean()
      .optional()
      .default(true)
      .describe("Record insertions as tracked changes. Default true."),
    author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Author name for tracked changes"),
    allow_untracked_edit: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        "Capability flag required to disable tracked changes. When track_changes is false, this must also be true or the call fails with UNTRACKED_EDIT_NOT_ALLOWED. Default false. This is a safety guard against prompt injection or long-context drift in regulated-industry use — silent edits to legal/regulated documents must be opted into with two independent flags.",
      ),
  },
  async ({ file_path, paragraphs, track_changes, author, allow_untracked_edit }) => {
    try {
      assertTrackChanges(track_changes, allow_untracked_edit);
      const result = await insertParagraphs(
        file_path,
        paragraphs.map(p => ({
          text: p.text,
          position: p.position,
          style: p.style,
          numId: p.num_id,
          numLevel: p.num_level,
          copyFormatFrom: p.copy_format_from,
        })),
        track_changes,
        author,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: delete_paragraph
// ---------------------------------------------------------------------------

server.tool(
  "delete_paragraph",
  "Delete a paragraph or table block by its index.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    paragraph_index: z
      .number()
      .describe("Index of the block to delete"),
    track_changes: z
      .boolean()
      .optional()
      .default(true)
      .describe(
        "Record deletion as a tracked change instead of removing the paragraph. Default true.",
      ),
    author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Author name for tracked changes"),
    allow_untracked_edit: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        "Capability flag required to disable tracked changes. When track_changes is false, this must also be true or the call fails with UNTRACKED_EDIT_NOT_ALLOWED. Default false. This is a safety guard against prompt injection or long-context drift in regulated-industry use — silent edits to legal/regulated documents must be opted into with two independent flags.",
      ),
  },
  async ({ file_path, paragraph_index, track_changes, author, allow_untracked_edit }) => {
    try {
      assertTrackChanges(track_changes, allow_untracked_edit);
      const result = await deleteParagraph(
        file_path,
        paragraph_index,
        track_changes,
        author,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: delete_paragraphs (bulk)
// ---------------------------------------------------------------------------

server.tool(
  "delete_paragraphs",
  "Delete multiple paragraphs or table blocks by their indices in one operation. Handles index reordering internally.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    paragraph_indices: z
      .array(z.number())
      .describe("Array of block indices to delete"),
    track_changes: z
      .boolean()
      .optional()
      .default(true)
      .describe(
        "Record deletions as tracked changes instead of removing the paragraphs. Default true.",
      ),
    author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Author name for tracked changes"),
    allow_untracked_edit: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        "Capability flag required to disable tracked changes. When track_changes is false, this must also be true or the call fails with UNTRACKED_EDIT_NOT_ALLOWED. Default false. This is a safety guard against prompt injection or long-context drift in regulated-industry use — silent edits to legal/regulated documents must be opted into with two independent flags.",
      ),
  },
  async ({ file_path, paragraph_indices, track_changes, author, allow_untracked_edit }) => {
    try {
      assertTrackChanges(track_changes, allow_untracked_edit);
      const result = await deleteParagraphs(
        file_path,
        paragraph_indices,
        track_changes,
        author,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: format_text
// ---------------------------------------------------------------------------

server.tool(
  "format_text",
  "Apply character formatting (bold, italic, underline, highlight, font, size, color) to all runs matching the search text.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    search: z.string().describe("Text to find and format"),
    bold: z.boolean().optional().describe("Set bold (true/false)"),
    italic: z.boolean().optional().describe("Set italic (true/false)"),
    underline: z.boolean().optional().describe("Set underline (true/false)"),
    strikethrough: z
      .boolean()
      .optional()
      .describe("Set strikethrough (true/false)"),
    highlight_color: z
      .string()
      .optional()
      .describe("Highlight color: yellow, green, cyan, magenta, blue, red, etc."),
    font_name: z.string().optional().describe("Font family name"),
    font_size: z.number().optional().describe("Font size in points (e.g. 12)"),
    font_color: z
      .string()
      .optional()
      .describe("Font color as hex (e.g. 'FF0000' for red)"),
    case_sensitive: z
      .boolean()
      .optional()
      .default(false)
      .describe("Case-sensitive text matching"),
  },
  async ({
    file_path,
    search,
    bold,
    italic,
    underline,
    strikethrough,
    highlight_color,
    font_name,
    font_size,
    font_color,
    case_sensitive,
  }) => {
    try {
      const result = await formatText(
        file_path,
        search,
        {
          bold,
          italic,
          underline,
          strikethrough,
          highlightColor: highlight_color,
          fontName: font_name,
          fontSize: font_size,
          fontColor: font_color,
        },
        case_sensitive,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: set_paragraph_format
// ---------------------------------------------------------------------------

server.tool(
  "set_paragraph_format",
  "Set paragraph-level formatting: alignment, spacing, and indentation.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    paragraph_index: z
      .number()
      .describe("Index of the paragraph to format"),
    alignment: z
      .enum(["left", "center", "right", "justify"])
      .optional()
      .describe("Text alignment"),
    space_before: z
      .number()
      .optional()
      .describe("Space before paragraph in points"),
    space_after: z
      .number()
      .optional()
      .describe("Space after paragraph in points"),
    line_spacing: z
      .number()
      .optional()
      .describe("Line spacing in points"),
    indent_left: z
      .number()
      .optional()
      .describe("Left indentation in twips (1440 twips = 1 inch)"),
    indent_right: z
      .number()
      .optional()
      .describe("Right indentation in twips"),
    first_line_indent: z
      .number()
      .optional()
      .describe("First line indent in twips"),
    hanging_indent: z
      .number()
      .optional()
      .describe("Hanging indent in twips"),
  },
  async ({
    file_path,
    paragraph_index,
    alignment,
    space_before,
    space_after,
    line_spacing,
    indent_left,
    indent_right,
    first_line_indent,
    hanging_indent,
  }) => {
    try {
      const result = await setParagraphFormat(file_path, paragraph_index, {
        alignment,
        spaceBefore: space_before,
        spaceAfter: space_after,
        lineSpacing: line_spacing,
        indentLeft: indent_left,
        indentRight: indent_right,
        firstLineIndent: first_line_indent,
        hangingIndent: hanging_indent,
      });
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: set_paragraph_formats
// ---------------------------------------------------------------------------

server.tool(
  "set_paragraph_formats",
  "Apply paragraph formatting to multiple paragraphs in one operation. Much faster than calling set_paragraph_format repeatedly.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    groups: z
      .array(
        z.object({
          indices: z
            .array(z.number())
            .describe("Paragraph indices to apply this format to"),
          alignment: z
            .enum(["left", "center", "right", "justify"])
            .optional()
            .describe("Text alignment"),
          space_before: z
            .number()
            .optional()
            .describe("Space before paragraph in points"),
          space_after: z
            .number()
            .optional()
            .describe("Space after paragraph in points"),
          line_spacing: z
            .number()
            .optional()
            .describe("Line spacing in points"),
          indent_left: z
            .number()
            .optional()
            .describe("Left indentation in twips (1440 twips = 1 inch)"),
          indent_right: z
            .number()
            .optional()
            .describe("Right indentation in twips"),
          first_line_indent: z
            .number()
            .optional()
            .describe("First line indent in twips"),
          hanging_indent: z
            .number()
            .optional()
            .describe("Hanging indent in twips"),
        }),
      )
      .describe("Array of formatting groups, each with indices and format options"),
  },
  async ({ file_path, groups }) => {
    try {
      const engineGroups = groups.map((g) => ({
        indices: g.indices,
        format: {
          alignment: g.alignment,
          spaceBefore: g.space_before,
          spaceAfter: g.space_after,
          lineSpacing: g.line_spacing,
          indentLeft: g.indent_left,
          indentRight: g.indent_right,
          firstLineIndent: g.first_line_indent,
          hangingIndent: g.hanging_indent,
        },
      }));
      const result = await setParagraphFormats(file_path, engineGroups);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: add_comment
// ---------------------------------------------------------------------------

server.tool(
  "add_comment",
  "Add a comment to specific text in the document. The comment is anchored to the first occurrence of the anchor text.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    anchor_text: z
      .string()
      .describe("Text to attach the comment to (must exist in the document)"),
    comment_text: z.string().describe("The comment content"),
    author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Comment author name"),
  },
  async ({ file_path, anchor_text, comment_text, author }) => {
    try {
      const result = await addComment(
        file_path,
        anchor_text,
        comment_text,
        author,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: add_comments
// ---------------------------------------------------------------------------

server.tool(
  "add_comments",
  "Add multiple comments to a document in a single operation. Opens and saves the file only once. Supports partial success: comments with unfound anchors are reported as failures without blocking the rest.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    comments: z
      .array(
        z.object({
          anchor_text: z.string().describe("Text to attach the comment to"),
          comment_text: z.string().describe("The comment content"),
          author: z.string().optional().describe("Comment author (overrides default_author)"),
        }),
      )
      .describe("Array of comments to add"),
    default_author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Default author name for comments without an explicit author"),
  },
  async ({ file_path, comments, default_author }) => {
    try {
      const result = await addComments(file_path, comments, default_author);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: read_comments
// ---------------------------------------------------------------------------

server.tool(
  "read_comments",
  "Read all comments in a DOCX file. Shows threaded replies indented under parent comments when available.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
  },
  async ({ file_path }) => {
    try {
      const result = await readComments(file_path);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: reply_to_comment
// ---------------------------------------------------------------------------

server.tool(
  "reply_to_comment",
  "Reply to an existing comment, creating a threaded conversation. The reply appears under the parent comment in Word's comment pane.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    parent_comment_id: z.number().describe("ID of the parent comment to reply to (from read_comments)"),
    comment_text: z.string().describe("The reply content"),
    author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Reply author name"),
  },
  async ({ file_path, parent_comment_id, comment_text, author }) => {
    try {
      const result = await replyToComment(file_path, parent_comment_id, comment_text, author);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: delete_comment
// ---------------------------------------------------------------------------

server.tool(
  "delete_comment",
  "Delete a comment by its ID. Also removes range markers from the document.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    comment_id: z.number().describe("Comment ID to delete (from read_comments)"),
  },
  async ({ file_path, comment_id }) => {
    try {
      const result = await deleteComment(file_path, comment_id);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: create_document
// ---------------------------------------------------------------------------

server.tool(
  "create_document",
  "Create a new DOCX file. Optionally provide initial content and title.",
  {
    file_path: z
      .string()
      .describe("Absolute path where the new .docx file will be created"),
    title: z
      .string()
      .optional()
      .describe("Document title (added as Heading 1)"),
    content: z
      .string()
      .optional()
      .describe(
        "Initial text content. Use newlines to separate paragraphs.",
      ),
    preset: z
      .enum(["ja-business"])
      .optional()
      .describe(
        "Optional style preset. 'ja-business' applies Japanese business-document defaults such as Yu Gothic body text and roomier paragraph spacing.",
      ),
  },
  async ({ file_path, title, content, preset }) => {
    try {
      const result = await createDocument(file_path, title, content, preset);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: apply_document_preset
// ---------------------------------------------------------------------------

server.tool(
  "apply_document_preset",
  "Apply a named document style preset in one pass by updating styles.xml. Use this instead of repeated paragraph-level formatting when you want a document-wide baseline.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    preset: z
      .enum(["ja-business"])
      .describe(
        "Preset name. 'ja-business' applies Yu Gothic body defaults plus roomier paragraph and heading spacing.",
      ),
  },
  async ({ file_path, preset }) => {
    try {
      const result = await applyDocumentPreset(file_path, preset);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: highlight_text
// ---------------------------------------------------------------------------

server.tool(
  "highlight_text",
  "Highlight all occurrences of text with a specified color.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    search: z.string().describe("Text to highlight"),
    color: z
      .string()
      .optional()
      .default("yellow")
      .describe("Highlight color: yellow, green, cyan, magenta, blue, red, etc."),
    case_sensitive: z
      .boolean()
      .optional()
      .default(false)
      .describe("Case-sensitive matching"),
  },
  async ({ file_path, search, color, case_sensitive }) => {
    try {
      const result = await highlightText(
        file_path,
        search,
        color,
        case_sensitive,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: insert_table
// ---------------------------------------------------------------------------

server.tool(
  "insert_table",
  "Insert a table at a specific position in the document.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    position: z
      .number()
      .describe("Block index to insert before (-1 for end)"),
    rows: z.number().describe("Number of rows"),
    cols: z.number().describe("Number of columns"),
    data: z
      .array(z.array(z.string()))
      .optional()
      .describe("Optional 2D array of cell values, e.g. [['A1','B1'],['A2','B2']]"),
  },
  async ({ file_path, position, rows, cols, data }) => {
    try {
      const result = await insertTable(file_path, position, rows, cols, data);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: set_heading
// ---------------------------------------------------------------------------

server.tool(
  "set_heading",
  "Convert a paragraph to a heading with the specified level (1–9).",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    paragraph_index: z
      .number()
      .describe("Index of the paragraph to convert"),
    level: z.number().min(1).max(9).describe("Heading level (1–9)"),
  },
  async ({ file_path, paragraph_index, level }) => {
    try {
      const result = await setHeading(file_path, paragraph_index, level);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: set_headings
// ---------------------------------------------------------------------------

server.tool(
  "set_headings",
  "Convert multiple paragraphs to headings in one operation. Opens and saves the file only once.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    headings: z
      .array(
        z.object({
          paragraph_index: z
            .number()
            .describe("Index of the paragraph to convert"),
          level: z.number().min(1).max(9).describe("Heading level (1–9)"),
        }),
      )
      .describe("Array of heading assignments"),
  },
  async ({ file_path, headings }) => {
    try {
      const engineItems = headings.map((h) => ({
        paragraphIndex: h.paragraph_index,
        level: h.level,
      }));
      const result = await setHeadings(file_path, engineItems);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: accept_all_changes
// ---------------------------------------------------------------------------

server.tool(
  "accept_all_changes",
  "Accept all tracked changes in the document. Insertions become permanent text, deletions are removed.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
  },
  async ({ file_path }) => {
    try {
      const result = await acceptAllChanges(file_path);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: reject_all_changes
// ---------------------------------------------------------------------------

server.tool(
  "reject_all_changes",
  "Reject all tracked changes in the document. Insertions are removed, deletions are restored to normal text.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
  },
  async ({ file_path }) => {
    try {
      const result = await rejectAllChanges(file_path);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: get_page_layout
// ---------------------------------------------------------------------------

server.tool(
  "get_page_layout",
  "Get page size, margins, and orientation of a DOCX file.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
  },
  async ({ file_path }) => {
    try {
      const result = await getPageLayout(file_path);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: set_page_layout
// ---------------------------------------------------------------------------

server.tool(
  "set_page_layout",
  "Set page size, margins, and orientation. Use presets (A4, LETTER, NARROW, etc.) or custom values in millimeters.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    page_size_preset: z
      .string()
      .optional()
      .describe(
        "Page size preset: A3, A4, A5, B4, B5, LETTER, LEGAL",
      ),
    orientation: z
      .enum(["portrait", "landscape"])
      .optional()
      .describe("Page orientation"),
    width_mm: z
      .number()
      .optional()
      .describe("Custom page width in millimeters (overrides preset)"),
    height_mm: z
      .number()
      .optional()
      .describe("Custom page height in millimeters (overrides preset)"),
    margin_preset: z
      .string()
      .optional()
      .describe(
        "Margin preset: NORMAL, NARROW, WIDE, JP_COURT_25, JP_COURT_30_20",
      ),
    top_mm: z.number().optional().describe("Top margin in mm"),
    right_mm: z.number().optional().describe("Right margin in mm"),
    bottom_mm: z.number().optional().describe("Bottom margin in mm"),
    left_mm: z.number().optional().describe("Left margin in mm"),
    header_mm: z.number().optional().describe("Header distance in mm"),
    footer_mm: z.number().optional().describe("Footer distance in mm"),
    gutter_mm: z.number().optional().describe("Gutter margin in mm"),
  },
  async ({
    file_path,
    page_size_preset,
    orientation,
    width_mm,
    height_mm,
    margin_preset,
    top_mm,
    right_mm,
    bottom_mm,
    left_mm,
    header_mm,
    footer_mm,
    gutter_mm,
  }) => {
    try {
      const result = await setPageLayout(file_path, {
        pageSizePreset: page_size_preset,
        orientation,
        widthMm: width_mm,
        heightMm: height_mm,
        marginPreset: margin_preset,
        topMm: top_mm,
        rightMm: right_mm,
        bottomMm: bottom_mm,
        leftMm: left_mm,
        headerMm: header_mm,
        footerMm: footer_mm,
        gutterMm: gutter_mm,
      });
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: read_header_footer
// ---------------------------------------------------------------------------

server.tool(
  "read_header_footer",
  "Read the content of headers and footers in a DOCX file.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
  },
  async ({ file_path }) => {
    try {
      const result = await readHeaderFooter(file_path);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: edit_table_cell
// ---------------------------------------------------------------------------

server.tool(
  "edit_table_cell",
  "Replace the text content of a specific table cell identified by block index, row, and column.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    block_index: z
      .number()
      .describe("Index of the table block (from read_document output)"),
    row_index: z.number().describe("Zero-based row index"),
    col_index: z.number().describe("Zero-based column index"),
    new_text: z.string().describe("New text content for the cell"),
    track_changes: z
      .boolean()
      .optional()
      .default(true)
      .describe("Record edits as tracked changes. Default true."),
    author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Author name for tracked changes"),
    allow_untracked_edit: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        "Capability flag required to disable tracked changes. When track_changes is false, this must also be true or the call fails with UNTRACKED_EDIT_NOT_ALLOWED. Default false. This is a safety guard against prompt injection or long-context drift in regulated-industry use — silent edits to legal/regulated documents must be opted into with two independent flags.",
      ),
  },
  async ({ file_path, block_index, row_index, col_index, new_text, track_changes, author, allow_untracked_edit }) => {
    try {
      assertTrackChanges(track_changes, allow_untracked_edit);
      const result = await editTableCell(
        file_path,
        block_index,
        row_index,
        col_index,
        new_text,
        track_changes,
        author,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: edit_table_cells (bulk)
// ---------------------------------------------------------------------------

server.tool(
  "edit_table_cells",
  "Replace the text content of multiple table cells in one operation. Cells can span different tables. Opens and saves the file only once.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
    edits: z
      .array(
        z.object({
          block_index: z
            .number()
            .describe("Index of the table block"),
          row_index: z.number().describe("Zero-based row index"),
          col_index: z.number().describe("Zero-based column index"),
          new_text: z.string().describe("New text content for the cell"),
        }),
      )
      .describe("Array of cell edits"),
    track_changes: z
      .boolean()
      .optional()
      .default(true)
      .describe("Record edits as tracked changes. Default true."),
    author: z
      .string()
      .optional()
      .default("Claude")
      .describe("Author name for tracked changes"),
    allow_untracked_edit: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        "Capability flag required to disable tracked changes. When track_changes is false, this must also be true or the call fails with UNTRACKED_EDIT_NOT_ALLOWED. Default false. This is a safety guard against prompt injection or long-context drift in regulated-industry use — silent edits to legal/regulated documents must be opted into with two independent flags.",
      ),
  },
  async ({ file_path, edits, track_changes, author, allow_untracked_edit }) => {
    try {
      assertTrackChanges(track_changes, allow_untracked_edit);
      const engineEdits = edits.map((e) => ({
        blockIndex: e.block_index,
        rowIndex: e.row_index,
        colIndex: e.col_index,
        newText: e.new_text,
      }));
      const result = await editTableCells(
        file_path,
        engineEdits,
        track_changes,
        author,
      );
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: read_footnotes
// ---------------------------------------------------------------------------

server.tool(
  "read_footnotes",
  "Read all footnotes in a DOCX file.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
  },
  async ({ file_path }) => {
    try {
      const result = await readFootnotes(file_path);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Tool: list_images
// ---------------------------------------------------------------------------

server.tool(
  "list_images",
  "List all images embedded in a DOCX file. Returns filename, dimensions, alt text, and block index for each image.",
  {
    file_path: z.string().describe("Absolute path to the .docx file"),
  },
  async ({ file_path }) => {
    try {
      const result = await listImages(file_path);
      return { content: [{ type: "text", text: result }] };
    } catch (e: unknown) {
      return {
        content: [{ type: "text", text: formatError(e) }],
        isError: true,
      };
    }
  },
);

// ---------------------------------------------------------------------------
// Start server
// ---------------------------------------------------------------------------

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error("Fatal:", err);
  process.exit(1);
});

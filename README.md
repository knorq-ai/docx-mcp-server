# docx-mcp-server

A local [MCP](https://modelcontextprotocol.io/) server for reading and editing Word (.docx) documents. Works with Claude Code, Cursor, and any MCP-compatible client.

**23 tools** for document content, formatting, comments, page layout, and track changes — all running locally via stdio with no file uploads.

## Features

| Category | Tools |
|---|---|
| **Read** | `read_document`, `get_document_info`, `search_text` |
| **Edit** | `replace_text`, `edit_paragraph`, `insert_paragraph`, `delete_paragraph` |
| **Format** | `format_text`, `set_paragraph_format`, `highlight_text`, `set_heading` |
| **Structure** | `insert_table`, `create_document` |
| **Review** | `add_comment`, `read_comments`, `delete_comment` |
| **Track changes** | `accept_all_changes`, `reject_all_changes` |
| **Page layout** | `get_page_layout`, `set_page_layout` |
| **Headers/footers** | `read_header_footer` |
| **Tables** | `edit_table_cell` |
| **Footnotes** | `read_footnotes` |

### Track changes

The editing tools (`replace_text`, `edit_paragraph`, `insert_paragraph`, `delete_paragraph`) support **tracked changes** — edits are recorded as Word revisions (`w:ins`/`w:del`) with author and timestamp, so reviewers can accept or reject them in Word.

Track changes is **on by default**. Pass `track_changes: false` to make direct edits.

Use `read_document` with `show_revisions: true` to see tracked changes annotated as `[-deleted-]` and `[+inserted+]`. The default view shows accepted text only.

Use `accept_all_changes` / `reject_all_changes` to finalize or revert all pending revisions.

### Page layout

`get_page_layout` / `set_page_layout` support:

- **Page size presets**: A3, A4, A5, B4, B5, Letter, Legal
- **Margin presets**: Normal, Narrow, Wide, JP Court 25mm, JP Court 30/20mm
- **Custom values** in millimeters for page size and individual margins
- **Orientation** (portrait / landscape)

## Quick start

### Option 1: Install from npm

```bash
npm install -g docx-mcp-server
```

Then add to your MCP config (see [Configuration](#configuration) below).

### Option 2: Use npx (no install)

Just add the config — `npx` downloads and runs it automatically:

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "npx",
      "args": ["-y", "docx-mcp-server"]
    }
  }
}
```

### Option 3: Build from source

```bash
git clone <repo-url>
cd mcp-server
npm install
npm run build
npm link        # makes `docx-mcp-server` available globally
```

## Configuration

### Claude Code

Add to your project's `.mcp.json` (per-project) or `~/.claude/settings.json` (global):

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "npx",
      "args": ["-y", "docx-mcp-server"]
    }
  }
}
```

### Cursor

Add to your MCP server configuration in Cursor settings:

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "npx",
      "args": ["-y", "docx-mcp-server"]
    }
  }
}
```

### Using a local build (without npm)

If you built from source and ran `npm link`:

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "docx-mcp-server"
    }
  }
}
```

Or reference the built file directly:

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "node",
      "args": ["/absolute/path/to/mcp-server/dist/index.js"]
    }
  }
}
```

## Distributing to others

### Via npm (recommended)

```bash
cd mcp-server
npm publish
```

Recipients install with:

```bash
npm install -g docx-mcp-server
```

Or skip the install entirely — just share the `.mcp.json` config with the `npx` setup above and it works out of the box.

### Via zip / git

Share the `mcp-server/` directory. Recipients run:

```bash
cd mcp-server
npm install
npm run build
npm link
```

Then add the config above.

## Tool reference

### Reading

**`read_document`** — Read content with block indices, styles, and formatting hints. Use `show_revisions` to see tracked changes.
```
file_path, start_paragraph?, end_paragraph?, show_revisions?
```

**`get_document_info`** — Paragraph count, heading outline, table count, comment status.
```
file_path
```

**`search_text`** — Search with context snippets.
```
file_path, query, case_sensitive?
```

### Editing

All editing tools accept `track_changes` (default `true`) and `author` (default `"Claude"`).

**`replace_text`** — Find and replace across the entire document. Handles text spanning multiple runs.
```
file_path, search, replace, case_sensitive?, track_changes?, author?, include_headers_footers?
```

**`edit_paragraph`** — Replace a paragraph's text content by index.
```
file_path, paragraph_index, new_text, track_changes?, author?
```

**`insert_paragraph`** — Insert a new paragraph at a position.
```
file_path, text, position, style?, track_changes?, author?
```

**`delete_paragraph`** — Delete a paragraph by index.
```
file_path, paragraph_index, track_changes?, author?
```

### Formatting

**`format_text`** — Apply bold, italic, underline, font, size, color, highlight to matching text.
```
file_path, search, bold?, italic?, underline?, strikethrough?, highlight_color?, font_name?, font_size?, font_color?, case_sensitive?
```

**`set_paragraph_format`** — Set alignment, spacing, indentation on a paragraph.
```
file_path, paragraph_index, alignment?, space_before?, space_after?, line_spacing?, indent_left?, indent_right?, first_line_indent?, hanging_indent?
```

**`highlight_text`** — Highlight matching text with a color.
```
file_path, search, color?, case_sensitive?
```

**`set_heading`** — Convert a paragraph to a heading (level 1-9).
```
file_path, paragraph_index, level
```

### Structure

**`insert_table`** — Insert a table with optional cell data.
```
file_path, position, rows, cols, data?
```

**`create_document`** — Create a new .docx file with optional title and content.
```
file_path, title?, content?
```

### Review

**`add_comment`** — Anchor a comment to specific text.
```
file_path, anchor_text, comment_text, author?
```

**`read_comments`** — List all comments with IDs, authors, and text.
```
file_path
```

**`delete_comment`** — Remove a comment by ID.
```
file_path, comment_id
```

### Track changes

**`accept_all_changes`** — Accept all tracked changes. Insertions become permanent, deletions are removed.
```
file_path
```

**`reject_all_changes`** — Reject all tracked changes. Insertions are removed, deleted text is restored.
```
file_path
```

### Page layout

**`get_page_layout`** — Read page size, margins, orientation.
```
file_path
```

**`set_page_layout`** — Set page size, margins, orientation by preset or custom mm values.
```
file_path, page_size_preset?, orientation?, width_mm?, height_mm?, margin_preset?, top_mm?, right_mm?, bottom_mm?, left_mm?, header_mm?, footer_mm?, gutter_mm?
```

### Headers and footers

**`read_header_footer`** — Read the text content of all headers and footers.
```
file_path
```

### Tables

**`edit_table_cell`** — Replace the text in a specific table cell by block, row, and column index.
```
file_path, block_index, row_index, col_index, new_text, track_changes?, author?
```

### Footnotes

**`read_footnotes`** — Read all footnotes with their IDs and text content.
```
file_path
```

## Requirements

- Node.js 18+

## License

MIT

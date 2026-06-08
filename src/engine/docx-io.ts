/**
 * DOCX file I/O — open, save, parse, serialize, body helpers.
 */

import * as fs from "fs/promises";
import JSZip from "jszip";
import {
  type XNode,
  parser,
  builder,
  findAll,
  findOne,
  sanitizeXmlText,
} from "./xml-helpers.js";

// ---------------------------------------------------------------------------
// Structured error codes
// ---------------------------------------------------------------------------

export const ErrorCode = {
  FILE_NOT_FOUND: "FILE_NOT_FOUND",
  INVALID_DOCX: "INVALID_DOCX",
  INDEX_OUT_OF_RANGE: "INDEX_OUT_OF_RANGE",
  NOT_A_PARAGRAPH: "NOT_A_PARAGRAPH",
  NOT_A_TABLE: "NOT_A_TABLE",
  ANCHOR_NOT_FOUND: "ANCHOR_NOT_FOUND",
  AMBIGUOUS_ANCHOR: "AMBIGUOUS_ANCHOR",
  INVALID_LOCATOR: "INVALID_LOCATOR",
  INVALID_PARAMETER: "INVALID_PARAMETER",
  UNTRACKED_EDIT_NOT_ALLOWED: "UNTRACKED_EDIT_NOT_ALLOWED",
  PENDING_REVISIONS: "PENDING_REVISIONS",
} as const;

export type ErrorCodeType = (typeof ErrorCode)[keyof typeof ErrorCode];

export class EngineError extends Error {
  code: ErrorCodeType;
  constructor(code: ErrorCodeType, message: string) {
    super(message);
    this.name = "EngineError";
    this.code = code;
  }
}

// ---------------------------------------------------------------------------
// DocxHandle
// ---------------------------------------------------------------------------

export interface DocxHandle {
  zip: JSZip;
  filePath: string;
}

export async function openDocx(filePath: string): Promise<DocxHandle> {
  let data: Buffer;
  try {
    data = await fs.readFile(filePath);
  } catch (err: unknown) {
    if ((err as NodeJS.ErrnoException).code === "ENOENT") {
      throw new EngineError(ErrorCode.FILE_NOT_FOUND, `File not found: ${filePath}`);
    }
    throw err;
  }
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    throw new EngineError(ErrorCode.INVALID_DOCX, `Not a valid DOCX/ZIP file: ${filePath}`);
  }
  return { zip, filePath };
}

export async function saveDocx(
  handle: DocxHandle,
  outPath?: string,
): Promise<void> {
  const buf = await handle.zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  await fs.writeFile(outPath ?? handle.filePath, buf);
}

export async function parseDocXml(handle: DocxHandle): Promise<XNode[]> {
  const xml = await handle.zip.file("word/document.xml")?.async("string");
  if (!xml) throw new EngineError(ErrorCode.INVALID_DOCX, "word/document.xml not found in DOCX");
  return parser.parse(xml);
}

export function serializeDocXml(handle: DocxHandle, parsed: XNode[]): void {
  const xml = builder.build(parsed);
  handle.zip.file("word/document.xml", xml);
}

// ---------------------------------------------------------------------------
// Body helpers
// ---------------------------------------------------------------------------

export function getBody(parsed: XNode[]): XNode[] {
  const docEl = parsed.find((n: XNode) => n["w:document"]);
  if (!docEl) return [];
  const bodyEl = docEl["w:document"].find((n: XNode) => n["w:body"]);
  if (!bodyEl) return [];
  return bodyEl["w:body"];
}

/**
 * Recursively iterate all paragraphs inside a table (including nested tables).
 * Calls `callback` with each paragraph's children array.
 */
export function forEachParagraphInTable(
  tblChildren: XNode[],
  callback: (pChildren: XNode[]) => void,
): void {
  const rows = findAll(tblChildren, "w:tr");
  for (const row of rows) {
    const cells = findAll(row["w:tr"], "w:tc");
    for (const cell of cells) {
      const cellChildren = cell["w:tc"] as XNode[];
      for (const cellChild of cellChildren) {
        if (cellChild["w:p"]) {
          callback(cellChild["w:p"]);
        } else if (cellChild["w:tbl"]) {
          forEachParagraphInTable(cellChild["w:tbl"], callback);
        }
      }
    }
  }
}

/** Returns list of header/footer file paths from the zip */
export function getHeaderFooterFiles(handle: DocxHandle): string[] {
  const files: string[] = [];
  handle.zip.forEach((relativePath) => {
    if (
      relativePath.startsWith("word/header") ||
      relativePath.startsWith("word/footer")
    ) {
      files.push(relativePath);
    }
  });
  return files;
}

// ---------------------------------------------------------------------------
// XML escaping
// ---------------------------------------------------------------------------

export function escapeXml(text: string): string {
  // Strip XML-1.0-illegal control chars first (same chokepoint as textNode),
  // then escape the markup-significant characters. Covers the string-building
  // path (createDocument) just as textNode covers the builder path.
  return sanitizeXmlText(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Test helpers for the MCP docx-engine tests.
 */

import * as os from "os";
import * as path from "path";
import * as fs from "fs/promises";
import * as crypto from "crypto";
import JSZip from "jszip";
import { createDocument, EngineError, ErrorCode, escapeXml } from "../docx-engine.js";

export { EngineError, ErrorCode };

/** Generate a unique tmp file path for a .docx file */
export function tmpDocxPath(): string {
  return path.join(os.tmpdir(), `mcp-test-${crypto.randomUUID()}.docx`);
}

/** List of tmp file paths to clean up after each test */
const tmpFiles: string[] = [];

/** Register a tmp path for cleanup */
export function trackTmpFile(p: string): string {
  tmpFiles.push(p);
  return p;
}

/** Remove all tracked tmp files */
export async function cleanupTmpFiles(): Promise<void> {
  for (const p of tmpFiles) {
    try {
      await fs.unlink(p);
    } catch {
      // ignore — file may not exist
    }
  }
  tmpFiles.length = 0;
}

/** Create a tmp docx and return its path (auto-tracked for cleanup) */
export async function createTmpDoc(
  content?: string,
  title?: string,
): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  await createDocument(p, content, title);
  return p;
}

/** Read raw word/document.xml from a .docx file */
export async function readRawDocXml(filePath: string): Promise<string> {
  const data = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(data);
  const xml = await zip.file("word/document.xml")?.async("string");
  if (!xml) throw new Error("word/document.xml not found");
  return xml;
}

/** Read raw word/comments.xml from a .docx file (returns empty string if not present) */
export async function readRawCommentsXml(filePath: string): Promise<string> {
  const data = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(data);
  const xml = await zip.file("word/comments.xml")?.async("string");
  return xml ?? "";
}

/** Read raw [Content_Types].xml */
export async function readRawContentTypes(filePath: string): Promise<string> {
  const data = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(data);
  const xml = await zip.file("[Content_Types].xml")?.async("string");
  return xml ?? "";
}

/** Read raw word/_rels/document.xml.rels */
export async function readRawDocRels(filePath: string): Promise<string> {
  const data = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(data);
  const xml = await zip
    .file("word/_rels/document.xml.rels")
    ?.async("string");
  return xml ?? "";
}

/**
 * Create a docx fixture where text is split across multiple w:r elements.
 * This is essential for testing cross-run text replacement.
 */
export async function createCrossRunDoc(
  parts: string[],
): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);

  // Build runs from parts
  const runs = parts
    .map(
      (text) =>
        `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`,
    )
    .join("");

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>${runs}</w:p>
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>`;

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", relsXml);
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/_rels/document.xml.rels", docRelsXml);

  const buf = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
  await fs.writeFile(p, buf);

  return p;
}

// escapeXml is imported from docx-engine.ts (canonical implementation in engine/docx-io.ts)

export async function createDocWithBookmark(
  text: string,
  bookmarkName: string = "testBookmark",
): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:bookmarkStart w:id="0" w:name="${escapeXml(bookmarkName)}"/>
  <w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`
  await writeMinimalDocx(p, documentXml);
  return p;
}

export async function createDocWithInsertedRun(
  before: string,
  inserted: string,
  after: string,
): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:r><w:t xml:space="preserve">${escapeXml(before)}</w:t></w:r>
  <w:ins w:id="1" w:author="TestAuthor" w:date="2024-01-01T00:00:00Z">
    <w:r><w:t xml:space="preserve">${escapeXml(inserted)}</w:t></w:r>
  </w:ins>
  <w:r><w:t xml:space="preserve">${escapeXml(after)}</w:t></w:r>
</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`
  await writeMinimalDocx(p, documentXml);
  return p;
}

export async function createDocWithCommentRange(
  text: string,
): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>
  <w:commentRangeEnd w:id="0"/>
</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
  await writeMinimalDocx(p, documentXml);
  return p;
}

export async function createDocWithDrawing(): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:r><w:t xml:space="preserve">Before image </w:t></w:r>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="1000000" cy="1000000"/>
        <wp:docPr id="1" name="Image1"/>
      </wp:inline>
    </w:drawing>
  </w:r>
  <w:r><w:t xml:space="preserve"> after image</w:t></w:r>
</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
  await writeMinimalDocx(p, documentXml);
  return p;
}

export async function createDocWithSdt(
  sdtText: string,
  tagName: string = "MyField",
): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p><w:r><w:t>Normal paragraph</w:t></w:r></w:p>
<w:sdt>
  <w:sdtPr>
    <w:tag w:val="${escapeXml(tagName)}"/>
  </w:sdtPr>
  <w:sdtContent>
    <w:p><w:r><w:t xml:space="preserve">${escapeXml(sdtText)}</w:t></w:r></w:p>
  </w:sdtContent>
</w:sdt>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
  await writeMinimalDocx(p, documentXml);
  return p;
}

export async function createDocWithHeaderFooter(
  bodyText: string,
  headerText: string,
  footerText: string,
): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p><w:r><w:t xml:space="preserve">${escapeXml(bodyText)}</w:t></w:r></w:p>
<w:sectPr>
  <w:headerReference w:type="default" r:id="rId2"/>
  <w:footerReference w:type="default" r:id="rId3"/>
  <w:pgSz w:w="11906" w:h="16838"/>
</w:sectPr>
</w:body>
</w:document>`;
  const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:r><w:t xml:space="preserve">${escapeXml(headerText)}</w:t></w:r></w:p>
</w:hdr>`;
  const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:r><w:t xml:space="preserve">${escapeXml(footerText)}</w:t></w:r></w:p>
</w:ftr>`;
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`;
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`;
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", relsXml);
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/header1.xml", headerXml);
  zip.file("word/footer1.xml", footerXml);
  zip.file("word/_rels/document.xml.rels", docRelsXml);
  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await fs.writeFile(p, buf);
  return p;
}


export async function createDocWithFootnote(
  bodyText: string,
  footnoteText: string,
): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:r><w:t xml:space="preserve">${escapeXml(bodyText)}</w:t></w:r>
  <w:r>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
    <w:footnoteReference w:id="1"/>
  </w:r>
</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
  const footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
<w:footnote w:id="1">
  <w:p><w:r><w:t xml:space="preserve">${escapeXml(footnoteText)}</w:t></w:r></w:p>
</w:footnote>
</w:footnotes>`;
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`;
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>`;
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>`;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", relsXml);
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/footnotes.xml", footnotesXml);
  zip.file("word/_rels/document.xml.rels", docRelsXml);
  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await fs.writeFile(p, buf);
  return p;
}


export async function createDocWithNestedTable(): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:tbl>
  <w:tr>
    <w:tc>
      <w:p><w:r><w:t>outer cell</w:t></w:r></w:p>
      <w:tbl>
        <w:tr>
          <w:tc><w:p><w:r><w:t>nested text</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    </w:tc>
  </w:tr>
</w:tbl>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
  await writeMinimalDocx(p, documentXml);
  return p;
}

/**
 * 1x1 透明 PNG (67 bytes) — テスト用最小画像。
 */
const TINY_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVQI12NgAAIABQAB" +
    "Nl7BcQAAAABJRU5ErkJggg==",
  "base64",
);

export async function createDocWithEmbeddedImage(
  altText: string = "Test image",
): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p><w:r><w:t>Paragraph before image</w:t></w:r></w:p>
<w:p>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="914400" cy="914400"/>
        <wp:docPr id="1" name="Picture 1" descr="${escapeXml(altText)}"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:blipFill>
                <a:blip r:embed="rId2"/>
              </pic:blipFill>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>
<w:p><w:r><w:t>Paragraph after image</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`;

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`;

  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", relsXml);
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/_rels/document.xml.rels", docRelsXml);
  zip.file("word/media/image1.png", TINY_PNG);

  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await fs.writeFile(p, buf);
  return p;
}

/**
 * Create a doc where text inside a paragraph is wrapped in w:sdt > w:sdtContent > w:r
 * (Google Docs export pattern: inline SDT within paragraph children).
 */
export async function createDocWithInlineSdt(
  sdtText: string,
  tagName: string = "goog_rdk_0",
): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:sdt>
    <w:sdtPr><w:tag w:val="${escapeXml(tagName)}"/></w:sdtPr>
    <w:sdtContent>
      <w:r><w:t xml:space="preserve">${escapeXml(sdtText)}</w:t></w:r>
    </w:sdtContent>
  </w:sdt>
</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
  await writeMinimalDocx(p, documentXml);
  return p;
}

/**
 * Create a doc with a paragraph that has numPr (numbering properties) in its pPr.
 * Optionally includes style, alignment, and indentation for testing copy_format_from.
 */
export async function createDocWithNumberedParagraph(
  text: string,
  numId: number,
  ilvl: number = 0,
  opts?: { style?: string; alignment?: string; indentLeft?: number },
): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);

  const pPrParts: string[] = [];
  if (opts?.style) pPrParts.push(`<w:pStyle w:val="${escapeXml(opts.style)}"/>`);
  pPrParts.push(`<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`);
  if (opts?.alignment) pPrParts.push(`<w:jc w:val="${escapeXml(opts.alignment)}"/>`);
  if (opts?.indentLeft !== undefined) pPrParts.push(`<w:ind w:left="${opts.indentLeft}"/>`);

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:pPr>${pPrParts.join("")}</w:pPr>
  <w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>
</w:p>
<w:p><w:r><w:t>Plain paragraph</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
  await writeMinimalDocx(p, documentXml);
  return p;
}

/**
 * Create a doc with a paragraph that has tracked change markers in its pPr > rPr.
 */
export async function createDocWithTrackedPPr(text: string): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>
    <w:rPr>
      <w:ins w:id="99" w:author="OldAuthor" w:date="2024-01-01T00:00:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>
</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
  await writeMinimalDocx(p, documentXml);
  return p;
}

export async function writeMinimalDocx(filePath: string, documentXml: string): Promise<void> {
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>`;
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", relsXml);
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/_rels/document.xml.rels", docRelsXml);
  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await fs.writeFile(filePath, buf);
}

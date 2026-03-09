/**
 * Image listing helpers: scan w:drawing elements and resolve relationships.
 */

import {
  type XNode,
  attr,
  findOne,
} from "./xml-helpers.js";
import type { DocxHandle } from "./docx-io.js";
import { getBody, parseDocXml } from "./docx-io.js";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface ImageInfo {
  /** Relationship ID (e.g. "rId5") */
  relationshipId: string;
  /** Filename inside word/ (e.g. "media/image1.png") */
  filename: string;
  /** MIME content type if determinable */
  contentType: string;
  /** File size in bytes (0 if unavailable) */
  sizeBytes: number;
  /** Block index where the image appears */
  blockIndex: number;
  /** Alt text from wp:docPr */
  altText: string;
  /** Name from wp:docPr */
  name: string;
  /** Width in EMU (English Metric Units, 1 inch = 914400 EMU) */
  widthEmu: number;
  /** Height in EMU */
  heightEmu: number;
}

export interface ListImagesResult {
  file: string;
  totalImages: number;
  images: ImageInfo[];
}

// ---------------------------------------------------------------------------
// Relationship parsing
// ---------------------------------------------------------------------------

/**
 * document.xml.rels から rId → Target のマップを構築する。
 * image 関係のみフィルタする。
 */
export function parseImageRelationships(
  relsXml: string,
): Map<string, string> {
  const map = new Map<string, string>();
  // Match <Relationship> elements with image type
  const re =
    /<Relationship\s[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?\/>/g;
  const imageType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  let m: RegExpExecArray | null;
  while ((m = re.exec(relsXml)) !== null) {
    const full = m[0];
    if (full.includes(imageType)) {
      map.set(m[1], m[2]);
    }
  }
  return map;
}

// ---------------------------------------------------------------------------
// Image scanning in body
// ---------------------------------------------------------------------------

/** MIME type from file extension */
function contentTypeFromExt(filename: string): string {
  const ext = filename.split(".").pop()?.toLowerCase() ?? "";
  const map: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    tif: "image/tiff",
    tiff: "image/tiff",
    svg: "image/svg+xml",
    emf: "image/x-emf",
    wmf: "image/x-wmf",
  };
  return map[ext] ?? "application/octet-stream";
}

/** 単一の XNode ツリーから w:drawing 要素を再帰的に収集する */
function collectDrawings(nodes: XNode[]): XNode[] {
  const result: XNode[] = [];
  for (const node of nodes) {
    if (typeof node !== "object" || node === null) continue;
    for (const key of Object.keys(node)) {
      if (key === ":@" || key === "#text") continue;
      if (key === "w:drawing") {
        result.push(node);
      }
      const children = node[key];
      if (Array.isArray(children)) {
        result.push(...collectDrawings(children));
      }
    }
  }
  return result;
}

/** w:drawing ノードから画像情報を抽出する */
function extractImageFromDrawing(
  drawingNode: XNode,
  rIdMap: Map<string, string>,
): Omit<ImageInfo, "blockIndex" | "sizeBytes"> | null {
  const drawingChildren = drawingNode["w:drawing"];
  if (!Array.isArray(drawingChildren)) return null;

  // wp:inline or wp:anchor
  for (const child of drawingChildren) {
    const containerKey = child["wp:inline"]
      ? "wp:inline"
      : child["wp:anchor"]
        ? "wp:anchor"
        : null;
    if (!containerKey) continue;

    const container = child[containerKey] as XNode[];

    // Extent (width/height)
    let widthEmu = 0;
    let heightEmu = 0;
    for (const c of container) {
      if (c["wp:extent"]) {
        widthEmu = parseInt(attr(c, "cx") ?? "0");
        heightEmu = parseInt(attr(c, "cy") ?? "0");
      }
    }

    // docPr (name, alt text)
    let name = "";
    let altText = "";
    for (const c of container) {
      if (c["wp:docPr"]) {
        name = attr(c, "name") ?? "";
        altText = attr(c, "descr") ?? "";
      }
    }

    // Find a:blip with r:embed
    const rId = findBlipRId(container);
    if (!rId) continue;

    const target = rIdMap.get(rId);
    if (!target) continue;

    return {
      relationshipId: rId,
      filename: target,
      contentType: contentTypeFromExt(target),
      altText,
      name,
      widthEmu,
      heightEmu,
    };
  }

  return null;
}

/** Recursively find a:blip r:embed attribute */
function findBlipRId(nodes: XNode[]): string | null {
  for (const node of nodes) {
    if (typeof node !== "object" || node === null) continue;
    for (const key of Object.keys(node)) {
      if (key === ":@" || key === "#text") continue;
      if (key === "a:blip") {
        const rEmbed = attr(node, "r:embed");
        if (rEmbed) return rEmbed;
      }
      const children = node[key];
      if (Array.isArray(children)) {
        const found = findBlipRId(children);
        if (found) return found;
      }
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * ドキュメント内の全画像を列挙する。
 */
export async function scanImages(handle: DocxHandle): Promise<ImageInfo[]> {
  const relsXml = await handle.zip
    .file("word/_rels/document.xml.rels")
    ?.async("string");
  if (!relsXml) return [];

  const rIdMap = parseImageRelationships(relsXml);
  if (rIdMap.size === 0) return [];

  const parsed = await parseDocXml(handle);
  const body = getBody(parsed);

  const images: ImageInfo[] = [];
  let blockIndex = 0;

  /** Scan nodes for drawings, assigning each found image the given block index. */
  const scanNodesForImages = async (nodes: XNode[], bi: number) => {
    const drawings = collectDrawings(nodes);
    for (const d of drawings) {
      const info = extractImageFromDrawing(d, rIdMap);
      if (info) {
        let sizeBytes = 0;
        const zipFile = handle.zip.file(`word/${info.filename}`);
        if (zipFile) {
          const data = await zipFile.async("nodebuffer");
          sizeBytes = data.length;
        }
        images.push({ ...info, sizeBytes, blockIndex: bi });
      }
    }
  };

  for (const child of body) {
    if (child["w:p"]) {
      await scanNodesForImages(child["w:p"] as XNode[], blockIndex);
      blockIndex++;
    } else if (child["w:tbl"]) {
      await scanNodesForImages(child["w:tbl"] as XNode[], blockIndex);
      blockIndex++;
    } else if (child["w:sdt"]) {
      // Content controls — match enumerateBlocks logic
      const sdtChildren = child["w:sdt"] as XNode[];
      const sdtContent = findOne(sdtChildren, "w:sdtContent");
      if (sdtContent) {
        const contentChildren = sdtContent["w:sdtContent"] as XNode[];
        for (const contentChild of contentChildren) {
          if (contentChild["w:p"]) {
            await scanNodesForImages(contentChild["w:p"] as XNode[], blockIndex);
            blockIndex++;
          }
        }
      }
    }
    // Skip w:sectPr and other non-content elements
  }

  return images;
}

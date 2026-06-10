import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  cleanupTmpFiles,
  readRawDocXml,
  writeMinimalDocx,
  tmpDocxPath,
  trackTmpFile,
  createDocWithSdt,
} from "./helpers.js";
import {
  readDocument,
  getDocumentInfo,
  editParagraphs,
  insertParagraphs,
  insertParagraphsStructured,
  deleteParagraphs,
  setHeadings,
  setParagraphFormats,
  searchTextStructured,
  ensureAnchors,
  ensureAnchorsStructured,
} from "../docx-engine.js";
import {
  generateDocParaId,
  collectAllParaIds,
  ensureW14Namespace,
  getDocumentRoot,
} from "../engine/anchors.js";
import { parser } from "../engine/xml-helpers.js";

afterEach(cleanupTmpFiles);

/** Build a doc with N simple paragraphs via untracked inserts; returns its path. */
async function makeParagraphs(texts: string[]): Promise<string> {
  const p = await createTmpDoc(texts[0]);
  for (let i = 1; i < texts.length; i++) {
    await insertParagraphs(p, [{ text: texts[i], position: -1 }], false);
  }
  return p;
}

/** Get the index→anchor map for a doc (seeds if needed). */
async function anchorsOf(p: string): Promise<(string | null)[]> {
  const r = await ensureAnchorsStructured(p);
  return r.blocks.map((b) => b.anchor);
}

// =========================================================================
// generateDocParaId
// =========================================================================

describe("generateDocParaId", () => {
  it("produces 8 uppercase hex chars in 00000001–7FFFFFFF and is unique", () => {
    const used = new Set<string>();
    for (let i = 0; i < 5000; i++) {
      const id = generateDocParaId(used);
      expect(id).toMatch(/^[0-9A-F]{8}$/);
      const n = parseInt(id, 16);
      expect(n).toBeGreaterThan(0);
      expect(n).toBeLessThan(0x80000000);
    }
    expect(used.size).toBe(5000); // all unique
  });

  it("never returns an id already in the used set", () => {
    const used = new Set<string>(["00000001", "7FFFFFFF"]);
    const id = generateDocParaId(used);
    expect(id).not.toBe("00000001");
    expect(id).not.toBe("7FFFFFFF");
  });
});

// =========================================================================
// ensure_anchors
// =========================================================================

describe("ensure_anchors", () => {
  it("seeds every body paragraph and declares the namespaces", async () => {
    // Raw fixture with three un-anchored paragraphs (inserts would auto-seed).
    const path = trackTmpFile(tmpDocxPath());
    await writeMinimalDocx(
      path,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>one</w:t></w:r></w:p>
<w:p><w:r><w:t>two</w:t></w:r></w:p>
<w:p><w:r><w:t>three</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`,
    );
    const r = await ensureAnchorsStructured(path);
    expect(r.seeded).toBe(3);
    expect(r.blocks.every((b) => b.anchor && /^[0-9A-F]{8}$/.test(b.anchor))).toBe(true);

    const xml = await readRawDocXml(path);
    expect(xml).toContain('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
    expect(xml).toContain('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"');
    expect(xml).toMatch(/mc:Ignorable="[^"]*\bw14\b[^"]*"/);
  });

  it("is idempotent: a second call seeds nothing and adds no duplicate namespace", async () => {
    const p = await makeParagraphs(["a", "b"]);
    await ensureAnchors(p);
    const first = await readRawDocXml(p);
    const r2 = await ensureAnchorsStructured(p);
    expect(r2.seeded).toBe(0);
    expect(r2.repaired).toBe(0);
    const second = await readRawDocXml(p);
    expect(second).toBe(first); // byte-stable
    expect((second.match(/xmlns:w14=/g) ?? []).length).toBe(1);
  });

  it("preserves existing valid anchors", async () => {
    const p = await makeParagraphs(["x", "y"]);
    const before = await anchorsOf(p);
    const after = await anchorsOf(p);
    expect(after).toEqual(before);
  });

  it("repairs duplicate paraIds across body paragraphs", async () => {
    const path = trackTmpFile(tmpDocxPath());
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p w14:paraId="AAAAAAAA"><w:r><w:t>first</w:t></w:r></w:p>
<w:p w14:paraId="AAAAAAAA"><w:r><w:t>second</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
    await writeMinimalDocx(path, documentXml);

    const r = await ensureAnchorsStructured(path);
    expect(r.repaired).toBeGreaterThanOrEqual(1);
    const anchors = r.blocks.map((b) => b.anchor);
    expect(new Set(anchors).size).toBe(anchors.length); // all unique now
  });

  it("reseeds an out-of-range / invalid existing paraId", async () => {
    const path = trackTmpFile(tmpDocxPath());
    // 00000000 (zero) and 80000000 (high bit set) are both invalid per MS-DOCX.
    await writeMinimalDocx(
      path,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p w14:paraId="00000000"><w:r><w:t>zero</w:t></w:r></w:p>
<w:p w14:paraId="80000000"><w:r><w:t>highbit</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`,
    );
    const r = await ensureAnchorsStructured(path);
    expect(r.repaired).toBe(2);
    expect(r.blocks.every((b) => b.anchor && /^[0-9A-F]{8}$/.test(b.anchor) && parseInt(b.anchor, 16) > 0 && parseInt(b.anchor, 16) < 0x80000000)).toBe(true);
  });

  it("treats a top-level w:sdt as one non-anchorable block (not expanded, not anchored)", async () => {
    const p = await createDocWithSdt("inside the content control");
    const r = await ensureAnchorsStructured(p);
    // Doc is [0] paragraph "Normal paragraph", [1] the content control.
    expect(r.blocks).toHaveLength(2);
    expect(r.blocks[0].type).toBe("paragraph");
    expect(r.blocks[0].anchor).toMatch(/^[0-9A-F]{8}$/);
    // The sdt is exactly one opaque block: type "sdt", no anchor, content shown.
    const sdtBlock = r.blocks[1];
    expect(sdtBlock.type).toBe("sdt");
    expect(sdtBlock.anchor).toBeNull();
    expect(sdtBlock.textPreview).toContain("inside the content control");
    // …and it never receives a paraId.
    const xml = await readRawDocXml(p);
    const sdtPart = xml.slice(xml.indexOf("<w:sdt"));
    expect(sdtPart).not.toContain("w14:paraId");
  });
});

// =========================================================================
// w:sdt block-index consistency: a block index from search_text/read_document
// resolves to the same block in the edit tools, even past a content control.
// =========================================================================

describe("w:sdt block-index consistency (read ↔ edit)", () => {
  /** [0] "Normal paragraph", [1] w:sdt, [2] the appended target paragraph. */
  async function sdtDocWithTrailingParagraph(target: string): Promise<string> {
    const p = await createDocWithSdt("inside the content control");
    await insertParagraphs(p, [{ text: target, position: -1 }], false);
    return p;
  }

  it("counts a content control as exactly one block in get_document_info", async () => {
    const p = await createDocWithSdt("field value");
    const info = await getDocumentInfo(p);
    // [0] Normal paragraph + [1] sdt = 2 blocks (sdt not expanded into its inner paragraph).
    expect(info).toContain("Total blocks: 2");
  });

  it("search_text index for a paragraph after an sdt round-trips through edit_paragraphs", async () => {
    const p = await sdtDocWithTrailingParagraph("TARGET_PARAGRAPH");

    const found = await searchTextStructured(p, "TARGET_PARAGRAPH");
    expect(found.totalMatches).toBe(1);
    const blockIndex = found.matches[0].blockIndex;
    expect(blockIndex).toBe(2); // [0] para, [1] sdt, [2] target

    // The crux: editing by that exact index must hit the target paragraph,
    // not a different block (the old enumerate/blockBodyIndices mismatch).
    await editParagraphs(p, [{ paragraphIndex: blockIndex, newText: "EDITED_OK" }], false);

    const after = await searchTextStructured(p, "EDITED_OK");
    expect(after.totalMatches).toBe(1);
    expect(after.matches[0].blockIndex).toBe(2);
    // The content control and the first paragraph are untouched.
    expect((await searchTextStructured(p, "inside the content control")).totalMatches).toBe(1);
    expect((await searchTextStructured(p, "Normal paragraph")).totalMatches).toBe(1);
    expect((await searchTextStructured(p, "TARGET_PARAGRAPH")).totalMatches).toBe(0);
  });

  it("rejects editing a content-control block by index with NOT_A_PARAGRAPH", async () => {
    const p = await createDocWithSdt("field value");
    await expect(
      editParagraphs(p, [{ paragraphIndex: 1, newText: "nope" }], false),
    ).rejects.toMatchObject({ code: "NOT_A_PARAGRAPH" });
  });

  it("rejects deleting a content-control block by index with a clear error", async () => {
    const p = await createDocWithSdt("field value");
    await expect(deleteParagraphs(p, [1], false)).rejects.toMatchObject({
      code: "NOT_A_PARAGRAPH",
    });
  });
});

// =========================================================================
// Anchor-based editing survives index shifts
// =========================================================================

describe("anchor-based editing", () => {
  it("edits the right paragraph by anchor after an earlier delete shifts indices", async () => {
    const p = await makeParagraphs(["alpha", "beta", "gamma"]);
    const anchors = await anchorsOf(p);
    const betaAnchor = anchors[1]!;

    // Delete paragraph 0 (untracked) — beta is now at index 0.
    await deleteParagraphs(p, [0], false);

    // Edit by beta's anchor: must still hit "beta", not whatever sits at index 1.
    await editParagraphs(p, [{ anchor: betaAnchor, newText: "BETA!" }], false);
    const doc = await readDocument(p);
    expect(doc).toContain("BETA!");
    expect(doc).toContain("gamma");
    expect(doc).not.toContain("beta");
  });

  it("set_headings by anchor", async () => {
    const p = await makeParagraphs(["title", "body"]);
    const anchors = await anchorsOf(p);
    await setHeadings(p, [{ anchor: anchors[0]!, level: 1 }]);
    expect(await readDocument(p)).toContain("(H1)");
  });

  it("set_paragraph_formats by anchor", async () => {
    const p = await makeParagraphs(["one", "two"]);
    const anchors = await anchorsOf(p);
    await setParagraphFormats(p, [{ anchors: [anchors[1]!], format: { alignment: "center" } }]);
    expect(await readDocument(p)).toContain("[align:center]");
  });

  it("delete_paragraphs by anchor", async () => {
    const p = await makeParagraphs(["keep", "drop", "keep2"]);
    const anchors = await anchorsOf(p);
    await deleteParagraphs(p, [], false, "Claude", [anchors[1]!]);
    const doc = await readDocument(p);
    expect(doc).not.toContain("drop");
    expect(doc).toContain("keep");
    expect(doc).toContain("keep2");
  });

  it("auto-seeds the touched paragraph so a fresh doc becomes anchor-addressable after one edit", async () => {
    const p = await makeParagraphs(["solo"]);
    // No ensure_anchors yet.
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "solo edited" }], false);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w14:paraId");
    expect(xml).toContain("xmlns:w14");
  });
});

// =========================================================================
// Insert with anchor placement
// =========================================================================

describe("insert_paragraphs with anchors", () => {
  it("inserts before / after an anchor and returns new anchors that resolve", async () => {
    const p = await makeParagraphs(["A", "C"]);
    const anchors = await anchorsOf(p);
    const cAnchor = anchors[1]!;

    const r = await insertParagraphsStructured(
      p,
      [{ text: "B", anchor: cAnchor, placement: "before" }],
      false,
    );
    expect(r.newParagraphs).toHaveLength(1);
    expect(r.newParagraphs[0].anchor).toMatch(/^[0-9A-F]{8}$/);

    const doc = await readDocument(p);
    expect(doc.indexOf("A")).toBeLessThan(doc.indexOf("B"));
    expect(doc.indexOf("B")).toBeLessThan(doc.indexOf("C"));

    // The returned anchor resolves on a follow-up edit.
    await editParagraphs(p, [{ anchor: r.newParagraphs[0].anchor, newText: "B-edited" }], false);
    expect(await readDocument(p)).toContain("B-edited");
  });

  it("multiple 'after' inserts against the same anchor keep array order", async () => {
    const p = await makeParagraphs(["p_aaa", "p_zzz"]);
    const anchors = await anchorsOf(p);
    await insertParagraphs(
      p,
      [
        { text: "p_bbb", anchor: anchors[0]!, placement: "after" },
        { text: "p_ccc", anchor: anchors[0]!, placement: "after" },
        { text: "p_ddd", anchor: anchors[0]!, placement: "after" },
      ],
      false,
    );
    const doc = await readDocument(p);
    expect(doc.indexOf("p_aaa")).toBeLessThan(doc.indexOf("p_bbb"));
    expect(doc.indexOf("p_bbb")).toBeLessThan(doc.indexOf("p_ccc"));
    expect(doc.indexOf("p_ccc")).toBeLessThan(doc.indexOf("p_ddd"));
    expect(doc.indexOf("p_ddd")).toBeLessThan(doc.indexOf("p_zzz"));
  });

  it("copy_format_from_anchor copies the source paragraph's numbering", async () => {
    const p = await makeParagraphs(["first"]);
    await insertParagraphs(p, [{ text: "numbered", position: -1, numId: 7 }], false);
    const anchors = await anchorsOf(p);
    const numberedAnchor = anchors[1]!;
    await insertParagraphs(
      p,
      [{ text: "copied", position: -1, copyFormatFromAnchor: numberedAnchor }],
      false,
    );
    const xml = await readRawDocXml(p);
    expect((xml.match(/w:numId w:val="7"/g) ?? []).length).toBe(2);
  });
});

// =========================================================================
// Locator validation & error paths
// =========================================================================

describe("anchor locator validation", () => {
  it("rejects an item with both paragraph_index and anchor (INVALID_LOCATOR)", async () => {
    const p = await makeParagraphs(["x"]);
    const anchors = await anchorsOf(p);
    await expect(
      editParagraphs(p, [{ paragraphIndex: 0, anchor: anchors[0]!, newText: "y" }], false),
    ).rejects.toMatchObject({ code: "INVALID_LOCATOR" });
  });

  it("rejects an item with neither paragraph_index nor anchor (INVALID_LOCATOR)", async () => {
    const p = await makeParagraphs(["x"]);
    await expect(
      editParagraphs(p, [{ newText: "y" }], false),
    ).rejects.toMatchObject({ code: "INVALID_LOCATOR" });
  });

  it("throws ANCHOR_NOT_FOUND for an unknown anchor", async () => {
    const p = await makeParagraphs(["x"]);
    await ensureAnchors(p);
    await expect(
      editParagraphs(p, [{ anchor: "DEADBEEF", newText: "y" }], false),
    ).rejects.toMatchObject({ code: "ANCHOR_NOT_FOUND" });
  });

  it("throws ANCHOR_NOT_FOUND after an untracked delete removed the paragraph", async () => {
    const p = await makeParagraphs(["alpha", "beta"]);
    const anchors = await anchorsOf(p);
    await deleteParagraphs(p, [1], false); // hard-delete beta
    await expect(
      editParagraphs(p, [{ anchor: anchors[1]!, newText: "z" }], false),
    ).rejects.toMatchObject({ code: "ANCHOR_NOT_FOUND" });
  });

  it("a tracked-deleted paragraph still resolves by anchor", async () => {
    const p = await makeParagraphs(["alpha", "beta"]);
    const anchors = await anchorsOf(p);
    await deleteParagraphs(p, [1], true); // tracked delete keeps the node
    // Resolving the anchor still works (no throw) — set a heading on it.
    await expect(
      setHeadings(p, [{ anchor: anchors[1]!, level: 2 }]),
    ).resolves.toBeTruthy();
  });
});

// =========================================================================
// Read surfaces
// =========================================================================

describe("anchors in read surfaces", () => {
  it("search_text reports the matched paragraph's anchor", async () => {
    const p = await makeParagraphs(["the magic word here"]);
    const anchors = await anchorsOf(p);
    const r = await searchTextStructured(p, "magic");
    expect(r.matches[0].anchor).toBe(anchors[0]);
  });

  it("read_document show_anchors annotates paragraphs", async () => {
    const p = await makeParagraphs(["hello"]);
    const anchors = await anchorsOf(p);
    const withAnchors = await readDocument(p, undefined, undefined, false, true);
    expect(withAnchors).toContain(`@${anchors[0]}`);
    const without = await readDocument(p);
    expect(without).not.toContain("@");
  });
});

// =========================================================================
// ensureW14Namespace URI validation
// =========================================================================

describe("ensureW14Namespace", () => {
  it("refuses a root that binds w14 to a different URI", () => {
    const parsed = parser.parse(
      `<w:document xmlns:w="x" xmlns:w14="http://wrong/uri"><w:body/></w:document>`,
    );
    const root = getDocumentRoot(parsed);
    expect(() => ensureW14Namespace(root)).toThrowError(/unexpected namespace/i);
  });

  it("collectAllParaIds counts ids inside tables and sdt", () => {
    const parsed = parser.parse(
      `<w:body>` +
        `<w:p w14:paraId="00000001"/>` +
        `<w:tbl><w:tr><w:tc><w:p w14:paraId="00000002"/></w:tc></w:tr></w:tbl>` +
        `<w:sdt><w:sdtContent><w:p w14:paraId="00000003"/></w:sdtContent></w:sdt>` +
        `</w:body>`,
    );
    const body = parsed[0]["w:body"];
    const { all } = collectAllParaIds(body);
    expect(all).toEqual(new Set(["00000001", "00000002", "00000003"]));
  });
});

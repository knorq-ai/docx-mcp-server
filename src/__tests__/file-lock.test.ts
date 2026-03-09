import { describe, it, expect } from "vitest";
import { withFileLock } from "../engine/file-lock.js";

describe("withFileLock", () => {
  it("serializes concurrent writes to the same file", async () => {
    const order: number[] = [];

    const task = (id: number, delayMs: number) =>
      withFileLock("/tmp/same-file.docx", async () => {
        order.push(id);
        await new Promise((r) => setTimeout(r, delayMs));
        order.push(id * 10);
      });

    // Launch concurrently — task 1 takes longer but starts first
    await Promise.all([task(1, 50), task(2, 10)]);

    // If serialized, task 1 must fully complete before task 2 starts
    expect(order).toEqual([1, 10, 2, 20]);
  });

  it("allows concurrent writes to different files", async () => {
    const order: string[] = [];

    const task = (file: string, id: string, delayMs: number) =>
      withFileLock(file, async () => {
        order.push(`${id}-start`);
        await new Promise((r) => setTimeout(r, delayMs));
        order.push(`${id}-end`);
      });

    await Promise.all([
      task("/tmp/file-a.docx", "A", 50),
      task("/tmp/file-b.docx", "B", 10),
    ]);

    // Both should start before either ends (parallel execution)
    const aStart = order.indexOf("A-start");
    const bStart = order.indexOf("B-start");
    const bEnd = order.indexOf("B-end");
    // B starts before A ends (both started concurrently)
    expect(bStart).toBeLessThan(order.indexOf("A-end"));
    // B finishes before A because it's faster
    expect(bEnd).toBeLessThan(order.indexOf("A-end"));
  });

  it("releases lock on exception", async () => {
    // First call throws
    await expect(
      withFileLock("/tmp/error-file.docx", async () => {
        throw new Error("intentional error");
      }),
    ).rejects.toThrow("intentional error");

    // Second call should still acquire the lock (not deadlocked)
    const result = await withFileLock("/tmp/error-file.docx", async () => {
      return "recovered";
    });
    expect(result).toBe("recovered");
  });
});

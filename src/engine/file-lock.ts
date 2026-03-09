/**
 * Per-file promise chain lock for serializing write operations.
 *
 * 読み取り専用関数はロック不要。書き込み関数のみ withFileLock でラップする。
 */

import * as path from "path";

const locks = new Map<string, Promise<void>>();

/**
 * 同一ファイルパスへの書き込み操作を直列化する。
 * 異なるファイルへの操作は並列に実行される。
 * 例外発生時もロックを正しく解放する。
 */
export async function withFileLock<T>(
  filePath: string,
  fn: () => Promise<T>,
): Promise<T> {
  const key = path.resolve(filePath);
  const prev = locks.get(key) ?? Promise.resolve();

  let resolve!: () => void;
  const next = new Promise<void>((r) => {
    resolve = r;
  });
  locks.set(key, next);

  await prev;
  try {
    return await fn();
  } finally {
    resolve();
    // GC: clean up if this is the last pending operation
    if (locks.get(key) === next) {
      locks.delete(key);
    }
  }
}

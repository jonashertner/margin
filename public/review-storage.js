export class ReviewStorageConflict extends Error {
  constructor() {
    super('Another tab saved a review. Automatic saving in this tab is paused. Save this review file before reloading or opening another document.');
    this.name = 'ReviewStorageConflict';
    this.code = 'REVIEW_STORAGE_CONFLICT';
  }
}

// Each tab owns only the exact saved value it last read or wrote. A mismatch
// pauses writes for the rest of the tab's life; refreshing cannot silently
// re-authorize this in-memory version to replace another tab's work.
export function createReviewStorage(storage, { activeKey, historyKey }) {
  let known = false, expected = null, conflicted = false;
  function readActive() {
    const raw = storage.getItem(activeKey);
    if (!known) { expected = raw; known = true; }
    return raw;
  }
  function assertCurrent() {
    if (conflicted) throw new ReviewStorageConflict();
    const actual = storage.getItem(activeKey);
    if (!known) { expected = actual; known = true; }
    if (actual !== expected) { conflicted = true; throw new ReviewStorageConflict(); }
  }
  function readHistory() {
    const raw = storage.getItem(historyKey);
    if (!raw) return [];
    const value = JSON.parse(raw);
    if (!Array.isArray(value)) throw new Error('Recent reviews could not be read. Save this review file before opening another document.');
    return value;
  }
  function save(session) {
    assertCurrent();
    const raw = JSON.stringify(session);
    storage.setItem(activeKey, raw);
    expected = raw;
  }
  function archive(session, at = new Date().toISOString()) {
    // This check happens after asynchronous import/conversion and before any
    // history or active-review write. The newest saved review remains intact.
    assertCurrent();
    const entries = readHistory().filter(item => item.workspaceId !== session.workspaceId);
    entries.unshift({ ...JSON.parse(JSON.stringify(session)), savedAt: at });
    storage.setItem(historyKey, JSON.stringify(entries));
  }
  return { readActive, readHistory, assertCurrent, save, archive, get conflicted() { return conflicted; } };
}

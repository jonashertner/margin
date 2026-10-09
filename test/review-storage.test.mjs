import test from 'node:test';
import assert from 'node:assert/strict';
import {createReviewStorage, ReviewStorageConflict} from '../public/review-storage.js';

const keys = {activeKey: 'current', historyKey: 'recent'};
function storage() {
  const values = new Map();
  return {getItem: key => values.get(key) ?? null, setItem: (key, value) => values.set(key, String(value))};
}
function review(workspaceId, text = 'Original wording') {
  return {workspaceId, before: text, after: text, decisions: {}, comments: {}, context: {openQuestions: 'Confirm the date.'}};
}
function tab(shared) {
  const manager = createReviewStorage(shared, keys);
  const memory = JSON.parse(manager.readActive());
  return {manager, memory};
}

test('a tab can save its own successive changes and archive before opening another review', () => {
  const shared = storage(), a = review('a');
  shared.setItem(keys.activeKey, JSON.stringify(a));
  const current = tab(shared);
  current.memory.comments.clause = [{text: 'Discuss the payment date.'}];
  current.manager.save(current.memory);
  current.memory.decisions.clause = 'accepted';
  current.manager.save(current.memory);
  current.manager.archive(current.memory, '2026-10-09');
  current.manager.save(review('b'));
  assert.equal(JSON.parse(shared.getItem(keys.activeKey)).workspaceId, 'b');
  assert.deepEqual(current.manager.readHistory()[0], {...current.memory, savedAt: '2026-10-09'});
  assert.equal(current.manager.conflicted, false);
});

test('a stale tab cannot overwrite a handoff review or its review history', () => {
  const shared = storage();
  shared.setItem(keys.activeKey, JSON.stringify(review('a')));
  const first = tab(shared), handoff = tab(shared);
  handoff.manager.archive(handoff.memory);
  const incoming = review('word-snapshot', 'Word snapshot wording');
  incoming.decisions.payment = 'rejected';
  incoming.comments.payment = [{text: 'Keep the earlier payment deadline.'}];
  handoff.manager.save(incoming);
  const expectedActive = shared.getItem(keys.activeKey), expectedHistory = shared.getItem(keys.historyKey);
  first.memory.comments.timing = [{text: 'My unsaved response remains here.'}];
  assert.throws(() => first.manager.save(first.memory), ReviewStorageConflict);
  assert.throws(() => first.manager.archive(first.memory), ReviewStorageConflict);
  assert.equal(shared.getItem(keys.activeKey), expectedActive);
  assert.equal(shared.getItem(keys.historyKey), expectedHistory);
  assert.equal(first.memory.comments.timing[0].text, 'My unsaved response remains here.');
  // A portable copy can still be made directly from this tab's memory.
  assert.deepEqual(JSON.parse(JSON.stringify(first.memory)), first.memory);
});

test('two tabs editing the same workspace cannot silently replace decisions or questions', () => {
  const shared = storage();
  shared.setItem(keys.activeKey, JSON.stringify(review('same-id')));
  const first = tab(shared), second = tab(shared);
  second.memory.decisions.clause = 'rejected';
  second.memory.context.openQuestions = 'Does the client agree with the revised scope?';
  second.manager.save(second.memory);
  first.memory.decisions.clause = 'accepted';
  assert.throws(() => first.manager.save(first.memory), {code: 'REVIEW_STORAGE_CONFLICT'});
  assert.deepEqual(JSON.parse(shared.getItem(keys.activeKey)), second.memory);
});

test('the post-conversion archive guard catches another tab saving while an import is in flight', async () => {
  const shared = storage();
  shared.setItem(keys.activeKey, JSON.stringify(review('existing')));
  const importer = tab(shared), editor = tab(shared);
  importer.manager.assertCurrent();
  // Import/conversion yields to other work before it can replace the review.
  await Promise.resolve();
  editor.memory.decisions.deadline = 'accepted';
  editor.manager.save(editor.memory);
  assert.throws(() => importer.manager.archive(importer.memory), ReviewStorageConflict);
  assert.equal(shared.getItem(keys.historyKey), null);
  assert.deepEqual(JSON.parse(shared.getItem(keys.activeKey)), editor.memory);
});

test('observing a conflict pauses writes even if storage later returns to its earlier bytes', () => {
  const shared = storage(), original = JSON.stringify(review('a'));
  shared.setItem(keys.activeKey, original);
  const current = tab(shared);
  shared.setItem(keys.activeKey, JSON.stringify(review('b')));
  assert.throws(() => current.manager.assertCurrent(), ReviewStorageConflict);
  shared.setItem(keys.activeKey, original);
  current.manager.readActive();
  assert.equal(current.manager.conflicted, true);
  assert.throws(() => current.manager.save(current.memory), ReviewStorageConflict);
});

test('a failed browser write does not advance the tab’s last saved value', () => {
  const shared = storage(), originalSet = shared.setItem;
  shared.setItem(keys.activeKey, JSON.stringify(review('a')));
  const current = tab(shared);
  current.memory.decisions.clause = 'accepted';
  shared.setItem = () => { throw new Error('Quota exceeded'); };
  assert.throws(() => current.manager.save(current.memory), /Quota exceeded/);
  assert.equal(current.manager.conflicted, false);
  shared.setItem = originalSet;
  current.manager.save(current.memory);
  assert.deepEqual(JSON.parse(shared.getItem(keys.activeKey)), current.memory);
});

test('unreadable history is preserved rather than replaced during a new document open', () => {
  const shared = storage();
  shared.setItem(keys.activeKey, JSON.stringify(review('a')));
  shared.setItem(keys.historyKey, '{"not":"an array"}');
  const current = tab(shared), before = shared.getItem(keys.historyKey);
  assert.throws(() => current.manager.archive(current.memory), /could not be read/);
  assert.equal(shared.getItem(keys.historyKey), before);
  assert.equal(JSON.parse(shared.getItem(keys.activeKey)).workspaceId, 'a');
});

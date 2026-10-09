import assert from 'node:assert/strict';
import { mkdtemp, readFile, readdir, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { installWordManifest } from '../scripts/word-setup.mjs';

const manifest = await readFile(new URL('../office/manifest.xml', import.meta.url), 'utf8');

async function temporaryDirectory(t) {
  const directory = await mkdtemp(join(tmpdir(), 'margin-word-setup-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  return directory;
}

test('Word sideload is idempotent and preserves the preceding own manifest', async t => {
  const directory = await temporaryDirectory(t);
  const destination = join(directory, 'wef', 'margin-local.xml');
  assert.deepEqual(await installWordManifest(destination, manifest), { changed: true, backup: null });
  assert.deepEqual(await installWordManifest(destination, manifest), { changed: false, backup: null });
  const updated = manifest.replaceAll('localhost:4318', 'localhost:4418');
  const result = await installWordManifest(destination, updated);
  assert.equal(result.changed, true);
  assert.equal(await readFile(result.backup, 'utf8'), manifest);
  assert.equal(await readFile(destination, 'utf8'), updated);
  assert.equal((await readdir(join(directory, 'wef'))).length, 2);
});

test('Word sideload refuses to replace a foreign manifest', async t => {
  const directory = await temporaryDirectory(t);
  const destination = join(directory, 'margin-local.xml');
  const foreign = '<OfficeApp><Id>unrelated-add-in</Id></OfficeApp>';
  await writeFile(destination, foreign);
  await assert.rejects(installWordManifest(destination, manifest), /different add-in/);
  assert.equal(await readFile(destination, 'utf8'), foreign);
  assert.deepEqual(await readdir(directory), ['margin-local.xml']);
});

test('Word sideload refuses a symlink even when its target has Margin’s ID', async t => {
  const directory = await temporaryDirectory(t);
  const target = join(directory, 'important.xml');
  const destination = join(directory, 'margin-local.xml');
  await writeFile(target, manifest);
  await symlink(target, destination);
  await assert.rejects(installWordManifest(destination, manifest), /non-file/);
  assert.equal(await readFile(target, 'utf8'), manifest);
});

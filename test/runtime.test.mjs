import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp, writeFile, symlink, rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {findExecutable, sandboxProfile} from '../lib/runtime.mjs';

test('explicit tool selection preserves invocation symlinks and rejects a missing override', async () => {
  const root = await mkdtemp(path.join(tmpdir(), 'margin-runtime-'));
  const previous = process.env.MARGIN_TEST_TOOL;
  try {
    const binary = path.join(root, 'engine');
    const alias = path.join(root, 'xelatex');
    await writeFile(binary, '#!/bin/sh\nexit 0\n', {mode:0o700});
    await symlink(binary, alias);
    process.env.MARGIN_TEST_TOOL = alias;
    assert.equal(findExecutable('node', ['MARGIN_TEST_TOOL']), alias);
    process.env.MARGIN_TEST_TOOL = path.join(root, 'absent');
    assert.equal(findExecutable('node', ['MARGIN_TEST_TOOL']), null);
  } finally {
    if (previous === undefined) delete process.env.MARGIN_TEST_TOOL;
    else process.env.MARGIN_TEST_TOOL = previous;
    await rm(root, {recursive:true, force:true});
  }
});

test('sandbox paths remain quoted data and writes are limited to the job', () => {
  const profile = sandboxProfile('/private/tmp/job', ['/private/tmp/runtime "name"']);
  assert.ok(profile.includes('(subpath "/private/tmp/runtime \\"name\\"")'));
  assert.ok(profile.includes('(deny default)'));
  assert.ok(!profile.includes('(allow network'));
  const writes = profile.slice(profile.indexOf('(allow file-write*'));
  assert.ok(writes.includes('(subpath "/private/tmp/job")'));
  assert.ok(!writes.includes('runtime'));
});

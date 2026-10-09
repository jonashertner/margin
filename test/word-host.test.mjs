import test from 'node:test';
import assert from 'node:assert/strict';
import { createWordHost, MAX_WORD_BYTES } from '../public/word-host.js';

const succeeded = value => ({ status: 'succeeded', value });
const failed = message => ({ status: 'failed', error: { message } });
const errorCode = code => error => error.code === code;
const revision = (nativeId, text = 'CHF 52,000', type = 'Added') => ({ nativeId, text, type, author: 'Mara Keller', date: new Date('2026-10-09T08:30:00Z') });

function hostFixture(options = {}) {
  const state = { changes: [revision('native-a'), revision('native-b', 'without notice', 'Deleted')],
    surrounding: 'A separate unchanged clause.', decisions: [], selections: [], calls: [], mode: 'TrackAll', ...options.state };
  const word = {
    async run(callback) {
      const jobs = [];
      const context = {
        async sync() { while (jobs.length) jobs.shift()(); },
        document: {
          load(property) { state.calls.push(['load-document', property]); },
          get changeTrackingMode() { return state.mode; },
          set changeTrackingMode(value) { jobs.push(() => { state.mode = value; }); },
          body: {
            getOoxml() {
              const result = {};
              jobs.push(() => { result.value = state.xml ?? JSON.stringify({ body: state.surrounding, changes: state.changes }); });
              return result;
            },
            getTrackedChanges() {
              const collection = {
                items: [],
                load(fields) {
                  state.calls.push(['load-changes', fields]);
                  jobs.push(() => { collection.items = state.changes.map(record => ({
                    ...record,
                    accept() { jobs.push(() => decide(record, 'accept')); },
                    reject() { jobs.push(() => decide(record, 'reject')); },
                    getRange() { return {
                      select(mode) { jobs.push(() => state.selections.push([record.nativeId, mode])); },
                      getReviewedText(version) {
                        (state.rangeQueries ??= []).push([record.nativeId, version]);
                        const result = {};
                        jobs.push(() => {
                          if (state.reviewedTextFails) throw new Error('Word could not read this range');
                          result.value = state.reviewedText?.[record.nativeId]?.[version] ?? '';
                        });
                        return result;
                      },
                    }; },
                  })); });
                },
              };
              return collection;
            },
          },
        },
      };
      function decide(record, action) {
        const index = state.changes.findIndex(item => item === record);
        if (index < 0) throw Object.assign(new Error('Native object no longer exists'), { code: 'ItemNotFound' });
        state.decisions.push([record.nativeId, action]);
        state.changes.splice(index, 1);
      }
      return callback(context);
    },
  };
  const office = {
    onReady: async () => ({ host: options.host ?? 'Word', platform: options.platform ?? 'Mac' }),
    AsyncResultStatus: { Succeeded: 'succeeded' },
    FileType: { Compressed: 'compressed' },
    context: {
      requirements: { isSetSupported: (set, version) => options.supports?.(set, version) ?? true },
      document: { url: 'https://example.invalid/drafts/Settlement%20agreement.docx', getFileAsync() {} },
    },
  };
  return { state, word, office, bridge: createWordHost({ office, word, timeoutMs: 100 }) };
}

function captureFixture(options = {}) {
  const fixture = hostFixture(options);
  const bytes = options.bytes ?? Uint8Array.from([0x50, 0x4b, 3, 4, ...Array.from({ length: 100_000 }, (_, index) => index % 256)]);
  const state = { closes: 0, indices: [], type: null, options: null };
  const file = {
    size: options.size ?? bytes.length,
    sliceCount: options.sliceCount ?? Math.ceil(bytes.length / 65_536),
    getSliceAsync(index, callback) {
      state.indices.push(index);
      if (options.sliceThrows) throw new Error('Host threw');
      if (options.sliceTimeout) return;
      if (options.failSlice === index) return callback(failed('Cannot read slice'));
      const data = Array.from(bytes.subarray(index * 65_536, (index + 1) * 65_536));
      const slice = { index, data, size: data.length };
      options.mutateSlice?.(slice);
      callback(succeeded(slice));
    },
    closeAsync(callback) {
      state.closes++;
      if (options.closeThrows) throw new Error('Close threw');
      callback(options.closeFails ? failed('Cannot close') : succeeded());
    },
  };
  fixture.office.context.document.getFileAsync = (type, settings, callback) => {
    state.type = type;
    state.options = settings;
    if (options.fileFails) callback(failed('Cannot get document'));
    else if (options.fileDelay) setTimeout(() => callback(succeeded(file)), options.fileDelay);
    else callback(succeeded(file));
  };
  if (Object.hasOwn(options, 'url')) fixture.office.context.document.url = options.url;
  fixture.bridge = createWordHost({ office: fixture.office, word: fixture.word, timeoutMs: options.timeoutMs ?? 100 });
  return { ...fixture, file, captureState: state, bytes };
}

test('normal browser and a non-Word host expose unavailable capabilities without writes', async () => {
  const outside = createWordHost({ office: null, word: null });
  assert.equal((await outside.waitForWord()).available, false);
  await assert.rejects(outside.listChanges(), errorCode('UNSUPPORTED_HOST'));
  const other = hostFixture({ host: 'Excel' });
  assert.equal((await other.bridge.waitForWord()).available, false);
  await assert.rejects(other.bridge.setTrackingMode('Off'), errorCode('UNSUPPORTED_HOST'));
  assert.deepEqual(other.state.calls, []);
});

test('runtime requirements independently gate review, tracking and document capture', async () => {
  const fixture = hostFixture({ supports: (set, version) => set === 'WordApi' && version === '1.4' });
  const capabilities = await fixture.bridge.waitForWord();
  assert.equal(capabilities.available, true);
  assert.equal(capabilities.trackingMode, true);
  assert.equal(capabilities.trackedChanges, false);
  assert.equal(capabilities.captureDocument, false);
  await assert.rejects(fixture.bridge.listChanges(), errorCode('UNSUPPORTED_API'));
  await assert.rejects(fixture.bridge.captureDocument(), errorCode('UNSUPPORTED_API'));
  assert.equal(await fixture.bridge.getTrackingMode(), 'TrackAll');
});

test('Word on the web permits native review but does not offer unsupported compressed capture', async () => {
  const fixture = hostFixture({ platform: 'OfficeOnline' });
  const capabilities = await fixture.bridge.waitForWord();
  assert.equal(capabilities.trackedChanges, true);
  assert.equal(capabilities.captureDocument, false);
  await assert.rejects(fixture.bridge.captureDocument(), errorCode('UNSUPPORTED_API'));
});

test('Office initialization timeout is bounded and safely unavailable', async () => {
  const bridge = createWordHost({ office: { onReady: () => new Promise(() => {}) }, timeoutMs: 5 });
  const capabilities = await bridge.waitForWord();
  assert.equal(capabilities.available, false);
  assert.match(capabilities.reason, /did not become ready/);
});

test('native revision snapshots display author, exact text, date and type with opaque ids', async () => {
  const { bridge } = hostFixture();
  const changes = await bridge.listChanges();
  assert.equal(changes.length, 2);
  assert.deepEqual({ ...changes[0], id: undefined }, {
    author: 'Mara Keller', date: '2026-10-09T08:30:00.000Z', text: 'CHF 52,000', type: 'Added', id: undefined,
    textSource: 'tracked-change', requiresWordReview: false,
  });
  assert.notEqual(changes[0].id, changes[1].id);
  assert.equal('nativeId' in changes[0], false);
});

test('select, accept and reject address the validated native object only', async () => {
  const { bridge, state } = hostFixture();
  let changes = await bridge.listChanges();
  await bridge.selectChange(changes[1].id);
  assert.deepEqual(state.selections, [['native-b', 'Select']]);
  await bridge.decideChange(changes[1].id, 'reject');
  assert.deepEqual(state.decisions, [['native-b', 'reject']]);
  changes = await bridge.listChanges();
  await bridge.decideChange(changes[0].id, 'accept');
  assert.deepEqual(state.decisions, [['native-b', 'reject'], ['native-a', 'accept']]);
  assert.deepEqual(await bridge.listChanges(), []);
});

test('empty deletion text uses the original text of only its native change range', async () => {
  const { bridge, state } = hostFixture({ state: {
    changes: [revision('deleted', '', 'Deleted')],
    reviewedText: { deleted: { Original: '48,000', Current: '' } },
  } });
  const [change] = await bridge.listChanges();
  assert.equal(change.text, '48,000');
  assert.equal(change.textSource, 'reviewed-range');
  assert.equal(change.requiresWordReview, false);
  assert.deepEqual(state.rangeQueries, [['deleted', 'Original']]);
  // Display fallback must not change the raw native metadata signature.
  await bridge.decideChange(change.id, 'reject');
  assert.deepEqual(state.decisions, [['deleted', 'reject']]);
});

test('empty insertion text uses the current version, never the original wording', async () => {
  const { bridge, state } = hostFixture({ state: {
    changes: [revision('added', '', 'Added')],
    reviewedText: { added: { Original: '', Current: '52,000' } },
  } });
  const [change] = await bridge.listChanges();
  assert.equal(change.text, '52,000');
  assert.deepEqual(state.rangeQueries, [['added', 'Current']]);
  await bridge.decideChange(change.id, 'accept');
});

test('unavailable deletion text gates decisions until that specific change is shown in Word', async () => {
  const { bridge, state } = hostFixture({ state: { changes: [revision('a', '', 'Deleted'), revision('b', '', 'Deleted')] } });
  const changes = await bridge.listChanges();
  assert.equal(changes[0].requiresWordReview, true);
  assert.equal(changes[0].textSource, 'unavailable');
  await assert.rejects(bridge.decideChange(changes[0].id, 'accept'), errorCode('WORD_REVIEW_REQUIRED'));
  assert.deepEqual(state.decisions, []);
  await bridge.selectChange(changes[0].id);
  await assert.rejects(bridge.decideChange(changes[1].id, 'accept'), errorCode('WORD_REVIEW_REQUIRED'));
  await bridge.decideChange(changes[0].id, 'reject');
  assert.deepEqual(state.decisions, [['a', 'reject']]);
});

test('a failed reviewed-range read preserves the native list but still requires Word inspection', async () => {
  const { bridge, state } = hostFixture({ state: { changes: [revision('a', '', 'Deleted')], reviewedTextFails: true } });
  const [change] = await bridge.listChanges();
  assert.equal(change.requiresWordReview, true);
  await assert.rejects(bridge.decideChange(change.id, 'reject'), errorCode('WORD_REVIEW_REQUIRED'));
  assert.deepEqual(state.decisions, []);
});

test('a refresh does not carry over Word inspection from an older snapshot', async () => {
  const { bridge } = hostFixture({ state: { changes: [revision('a', '\r', 'Deleted')] } });
  const [first] = await bridge.listChanges();
  await bridge.selectChange(first.id);
  const [second] = await bridge.listChanges();
  await assert.rejects(bridge.decideChange(second.id, 'accept'), errorCode('WORD_REVIEW_REQUIRED'));
});

test('accepting a change invalidates every snapshot id and a second click cannot shift to another change', async () => {
  const { bridge, state } = hostFixture();
  const changes = await bridge.listChanges();
  const first = bridge.decideChange(changes[0].id, 'accept');
  const doubleClick = bridge.decideChange(changes[0].id, 'accept');
  await first;
  await assert.rejects(doubleClick, errorCode('STALE_CHANGE'));
  await assert.rejects(bridge.decideChange(changes[1].id, 'accept'), errorCode('STALE_CHANGE'));
  assert.deepEqual(state.decisions, [['native-a', 'accept']]);
});

test('refresh invalidates previous ids, and returned metadata cannot mutate the safety snapshot', async () => {
  const { bridge, state } = hostFixture();
  const old = await bridge.listChanges();
  const current = await bridge.listChanges();
  old[0].text = 'Tampered display';
  current[1].author = 'Tampered author';
  await assert.rejects(bridge.selectChange(old[0].id), errorCode('STALE_CHANGE'));
  await bridge.decideChange(current[1].id, 'accept');
  assert.deepEqual(state.decisions, [['native-b', 'accept']]);
});

for (const [name, alter] of [
  ['an earlier revision was removed', state => state.changes.shift()],
  ['a revision was edited', state => { state.changes[1].text = 'new wording'; }],
  ['unchanged surrounding wording was edited', state => { state.surrounding = 'Changed clause.'; }],
  ['a revision was added', state => state.changes.push(revision('native-c'))],
]) {
  test(`stale list refuses selection and all writes when ${name}`, async () => {
    const { bridge, state } = hostFixture();
    const changes = await bridge.listChanges();
    alter(state);
    await assert.rejects(bridge.decideChange(changes[1].id, 'accept'), errorCode('STALE_CHANGE'));
    await assert.rejects(bridge.selectChange(changes[0].id), errorCode('STALE_CHANGE'));
    assert.deepEqual(state.decisions, []);
    assert.deepEqual(state.selections, []);
  });
}

test('identical repeated edits cannot masquerade as the removed revision at a reused index', async () => {
  const { bridge, state } = hostFixture({ state: { changes: [revision('native-a'), revision('native-b')] } });
  const changes = await bridge.listChanges();
  // Every displayed field and the collection length remain identical. Only the
  // actual OOXML revision identity has changed at the same location.
  state.changes[0] = revision('native-replacement');
  await assert.rejects(bridge.decideChange(changes[0].id, 'reject'), errorCode('STALE_CHANGE'));
  assert.deepEqual(state.decisions, []);
});

const wordXml = (rsid, extra = '') => `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p w:rsidR="${rsid}" w:rsidRDefault="${rsid}" w:rsidP="${rsid}"><w:ins w:id="4" w:author="Mara Keller" w:date="2026-10-09T08:30:00Z"><w:r w:rsidRPr="${rsid}"><w:t>CHF 52,000</w:t></w:r></w:ins>${extra}</w:p></w:body></w:document>`;

test('Word for Mac serialization RSIDs do not falsely stale an unchanged native revision', async () => {
  const { bridge, state } = hostFixture({ state: { xml: wordXml('00953E7F') } });
  const [change] = await bridge.listChanges();
  state.xml = wordXml('005A7702');
  await bridge.selectChange(change.id);
  await bridge.decideChange(change.id, 'accept');
  assert.deepEqual(state.decisions, [['native-a', 'accept']]);
});

const terminalXml = id => wordXml('00953E7F').replace('</w:body>', `<w:p w14:paraId="${id}" w14:textId="77777777" w:rsidR="00953E7F" w:rsidRDefault="00953E7F"/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body>`);

test('Word for Mac empty terminal paragraph serialization ID does not falsely stale a revision', async () => {
  const { bridge, state } = hostFixture({ state: { xml: terminalXml('44A05565') } });
  const [change] = await bridge.listChanges();
  state.xml = terminalXml('69C5B741');
  await bridge.selectChange(change.id);
  await bridge.decideChange(change.id, 'accept');
  assert.deepEqual(state.decisions, [['native-a', 'accept']]);
});

for (const [name, transform] of [
  ['a content-bearing paragraph', xml => xml.replace('w:rsidRDefault="00953E7F"/><w:sectPr>', 'w:rsidRDefault="00953E7F"><w:r><w:t>Substantive text.</w:t></w:r></w:p><w:sectPr>')],
  ['a non-terminal empty paragraph', xml => xml.replace('/><w:sectPr>', '/><w:p><w:r><w:t>Another paragraph.</w:t></w:r></w:p><w:sectPr>')],
  ['an empty paragraph with an unknown attribute', xml => xml.replace('w14:textId="77777777"', 'w14:textId="77777777" w:custom="value"')],
  ['a paragraph without the empty-text sentinel', xml => xml.replace('w14:textId="77777777"', 'w14:textId="12345678"')],
  ['a literal paragraph in a CDATA section', xml => `<![CDATA[${xml}]]>`],
]) {
  test(`terminal paragraph normalization preserves identity for ${name}`, async () => {
    const { bridge, state } = hostFixture({ state: { xml: transform(terminalXml('44A05565')) } });
    const [change] = await bridge.listChanges();
    state.xml = transform(terminalXml('69C5B741'));
    await assert.rejects(bridge.decideChange(change.id, 'accept'), errorCode('STALE_CHANGE'));
    assert.deepEqual(state.decisions, []);
  });
}

const separatorXml = (id, type = 'separator', note = 'footnote') => `${wordXml('00953E7F')}<w:${note}s><w:${note} w:type="${type}" w:id="-1"><w:p w14:paraId="${id}" w14:textId="77777777"><w:r><w:separator/></w:r></w:p></w:${note}></w:${note}s>`;

for (const note of ['footnote', 'endnote']) {
  for (const type of ['separator', 'continuationSeparator']) {
    test(`${note} ${type} synthetic paragraph IDs do not falsely stale a revision`, async () => {
      const { bridge, state } = hostFixture({ state: { xml: separatorXml('429A17F3', type, note) } });
      const [change] = await bridge.listChanges();
      state.xml = separatorXml('752053D5', type, note);
      await bridge.selectChange(change.id);
      await bridge.decideChange(change.id, 'accept');
      assert.deepEqual(state.decisions, [['native-a', 'accept']]);
    });
  }
}

for (const [name, original, changed] of [
  ['ordinary footnote paragraph identity', separatorXml('429A17F3', 'normal'), separatorXml('752053D5', 'normal')],
  ['separator content', separatorXml('429A17F3'), separatorXml('752053D5').replace('<w:separator/>', '<w:t>Different wording.</w:t>')],
  ['separator formatting', separatorXml('429A17F3'), separatorXml('752053D5').replace('<w:r>', '<w:r><w:rPr><w:b/></w:rPr>')],
  ['separator identity', separatorXml('429A17F3'), separatorXml('752053D5').replace('w:id="-1"', 'w:id="-2"')],
  ['revision identity inside separator', separatorXml('429A17F3').replace('<w:separator/>', '<w:ins w:id="8"><w:t>Text</w:t></w:ins>'), separatorXml('752053D5').replace('<w:separator/>', '<w:ins w:id="9"><w:t>Text</w:t></w:ins>')],
  ['literal XML in CDATA', `<![CDATA[${separatorXml('429A17F3')}]]>`, `<![CDATA[${separatorXml('752053D5')}]]>`],
]) {
  test(`separator normalization preserves ${name}`, async () => {
    const { bridge, state } = hostFixture({ state: { xml: original } });
    const [change] = await bridge.listChanges();
    state.xml = changed;
    await assert.rejects(bridge.decideChange(change.id, 'accept'), errorCode('STALE_CHANGE'));
    assert.deepEqual(state.decisions, []);
  });
}

const catalogue = entries => `${wordXml('00953E7F')}<w:settings><w:rsids><w:rsidRoot w:val="00000000"/>${entries}</w:rsids></w:settings>`;

test('Word serialization may extend the metadata-only editing-session catalogue', async () => {
  const { bridge, state } = hostFixture({ state: { xml: catalogue('<w:rsid w:val="00034616"/><w:rsid w:val="0015074B"/>') } });
  const [change] = await bridge.listChanges();
  state.xml = catalogue('<w:rsid w:val="00034616"/><w:rsid w:val="0015074B"/><w:rsid w:val="00172D67"/>');
  await bridge.selectChange(change.id);
  await bridge.decideChange(change.id, 'accept');
  assert.deepEqual(state.decisions, [['native-a', 'accept']]);
});

for (const [name, original, changed] of [
  ['unknown children', '<w:custom value="old"/>', '<w:custom value="new"/>'],
  ['unknown attributes', '<w:rsid w:val="00034616" w:custom="old"/>', '<w:rsid w:val="00034616" w:custom="new"/>'],
  ['unexpected text', 'old wording', 'new wording'],
  ['unexpected values', '<w:rsid w:val="unexpected-old"/>', '<w:rsid w:val="unexpected-new"/>'],
  ['tracked revisions', '<w:ins w:id="4"><w:t>Text</w:t></w:ins>', '<w:ins w:id="5"><w:t>Text</w:t></w:ins>'],
]) {
  test(`editing-session catalogue normalization preserves ${name}`, async () => {
    const { bridge, state } = hostFixture({ state: { xml: catalogue(original) } });
    const [change] = await bridge.listChanges();
    state.xml = catalogue(changed);
    await assert.rejects(bridge.decideChange(change.id, 'accept'), errorCode('STALE_CHANGE'));
    assert.deepEqual(state.decisions, []);
  });
}

for (const [name, before, after] of [
  ['text', wordXml('00953E7F'), wordXml('005A7702').replace('52,000', '53,000')],
  ['revision identity', wordXml('00953E7F'), wordXml('005A7702').replace('w:id="4"', 'w:id="5"')],
  ['revision author', wordXml('00953E7F'), wordXml('005A7702').replace('Mara Keller', 'Other author')],
  ['revision date', wordXml('00953E7F'), wordXml('005A7702').replace('08:30:00Z', '09:30:00Z')],
  ['RSID-looking literal wording', wordXml('00953E7F', '<w:r><w:t>The literal w:rsidR="11111111" is quoted.</w:t></w:r>'), wordXml('005A7702', '<w:r><w:t>The literal w:rsidR="22222222" is quoted.</w:t></w:r>')],
  ['RSID-looking text inside another attribute', wordXml('00953E7F', '<w:r data-note=\'literal w:rsidR="11111111"\'/>'), wordXml('005A7702', '<w:r data-note=\'literal w:rsidR="22222222"\'/>')],
  ['CDATA', wordXml('00953E7F', '<![CDATA[<w:r w:rsidR="11111111"/>]]>'), wordXml('005A7702', '<![CDATA[<w:r w:rsidR="22222222"/>]]>')],
  ['XML comment', wordXml('00953E7F', '<!-- <w:r w:rsidR="11111111"/> -->'), wordXml('005A7702', '<!-- <w:r w:rsidR="22222222"/> -->')],
]) {
  test(`RSID normalization still refuses changed ${name}`, async () => {
    const { bridge, state } = hostFixture({ state: { xml: before } });
    const [change] = await bridge.listChanges();
    state.xml = after;
    await assert.rejects(bridge.decideChange(change.id, 'accept'), error => error.code === 'STALE_CHANGE' && error.reason === 'body_snapshot_changed' && !('debugDetails' in error));
    assert.deepEqual(state.decisions, []);
  });
}

test('invalid ids, decisions and tracking modes never invoke native writes', async () => {
  const { bridge, state } = hostFixture();
  await bridge.listChanges();
  await assert.rejects(bridge.decideChange('arbitrary', 'accept'), errorCode('STALE_CHANGE'));
  await assert.rejects(bridge.decideChange('arbitrary', 'acceptAll'), errorCode('INVALID_DECISION'));
  await assert.rejects(bridge.setTrackingMode('disable'), errorCode('INVALID_MODE'));
  assert.deepEqual(state.decisions, []);
  assert.equal(state.mode, 'TrackAll');
});

test('tracking mode changes are read back and preserve existing revisions', async () => {
  const { bridge, state } = hostFixture();
  assert.equal(await bridge.getTrackingMode(), 'TrackAll');
  assert.equal(await bridge.setTrackingMode('TrackMineOnly'), 'TrackMineOnly');
  assert.equal(await bridge.setTrackingMode('Off'), 'Off');
  assert.equal(state.changes.length, 2);
  assert.deepEqual(state.decisions, []);
});

test('capture reads sequential 64 KB slices, preserves exact bytes and closes the file', async () => {
  const { bridge, bytes, captureState } = captureFixture();
  const result = await bridge.captureDocument();
  assert.equal(result.name, 'Settlement agreement.docx');
  assert.deepEqual(Buffer.from(result.data, 'base64'), Buffer.from(bytes));
  assert.equal(captureState.type, 'compressed');
  assert.deepEqual(captureState.options, { sliceSize: 65_536 });
  assert.deepEqual(captureState.indices, [0, 1]);
  assert.equal(captureState.closes, 1);
});

test('capture safely names an unsaved document and rejects path injection in the filename', async () => {
  assert.equal((await captureFixture({ url: null }).bridge.captureDocument()).name, 'Word document.docx');
  assert.equal((await captureFixture({ url: 'https://example.invalid/a%2Fb%00.docx' }).bridge.captureDocument()).name, 'ab.docx');
});

test('a document over the application 4 MB limit is closed without reading slices', async () => {
  const { bridge, captureState } = captureFixture({ size: MAX_WORD_BYTES + 1 });
  await assert.rejects(bridge.captureDocument(), errorCode('FILE_TOO_LARGE'));
  assert.deepEqual(captureState.indices, []);
  assert.equal(captureState.closes, 1);
});

test('capture permits the exact 4 MB boundary and preserves all bytes', async () => {
  const bytes = new Uint8Array(MAX_WORD_BYTES);
  bytes.set([0x50, 0x4b, 3, 4]);
  bytes[MAX_WORD_BYTES - 1] = 254;
  const { bridge, captureState } = captureFixture({ bytes });
  const result = await bridge.captureDocument();
  assert.deepEqual(Buffer.from(result.data, 'base64'), Buffer.from(bytes));
  assert.equal(captureState.closes, 1);
});

for (const [name, options, code] of [
  ['file acquisition failed', { fileFails: true }, 'FILE_READ_FAILED'],
  ['a later slice failed', { failSlice: 1 }, 'SLICE_READ_FAILED'],
  ['a slice call threw', { sliceThrows: true }, 'SLICE_READ_FAILED'],
  ['a slice timed out', { sliceTimeout: true, timeoutMs: 5 }, 'HOST_TIMEOUT'],
  ['size was invalid', { size: -1 }, 'INVALID_FILE'],
  ['slice count was inconsistent', { sliceCount: 3 }, 'INVALID_FILE'],
  ['slice index was wrong', { mutateSlice: slice => { slice.index++; } }, 'INVALID_FILE'],
  ['slice was truncated', { mutateSlice: slice => { slice.data.pop(); } }, 'INVALID_FILE'],
  ['slice contained an invalid byte', { mutateSlice: slice => { slice.data[0] = 256; } }, 'INVALID_FILE'],
  ['slice contained a missing byte', { mutateSlice: slice => { delete slice.data[12]; } }, 'INVALID_FILE'],
  ['content was not a DOCX', { bytes: new Uint8Array([1, 2, 3, 4]) }, 'INVALID_FILE'],
  ['file release failed', { closeFails: true }, 'FILE_CLOSE_FAILED'],
]) {
  test(`capture fails safely and releases the file when ${name}`, async () => {
    const { bridge, captureState } = captureFixture(options);
    await assert.rejects(bridge.captureDocument(), errorCode(code));
    assert.equal(captureState.closes, options.fileFails ? 0 : 1);
  });
}

test('a late file acquisition after timeout releases the acquired Word handle', async () => {
  const { bridge, captureState } = captureFixture({ fileDelay: 15, timeoutMs: 5 });
  await assert.rejects(bridge.captureDocument(), errorCode('HOST_TIMEOUT'));
  await new Promise(resolve => setTimeout(resolve, 30));
  assert.equal(captureState.closes, 1);
  assert.deepEqual(captureState.indices, []);
});

test('release failure does not conceal the primary slice failure', async () => {
  const { bridge, captureState } = captureFixture({ failSlice: 0, closeThrows: true });
  await assert.rejects(bridge.captureDocument(), error => error.code === 'SLICE_READ_FAILED' && error.closeError.code === 'FILE_CLOSE_FAILED');
  assert.equal(captureState.closes, 1);
});

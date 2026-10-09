// The only module that talks to Office.js. Document bytes never leave the caller.
// Native changes: https://learn.microsoft.com/office/dev/add-ins/word/manage-tracked-changes
// File capture: https://learn.microsoft.com/javascript/api/office/office.document
export const MAX_WORD_BYTES = 4_000_000;
const SLICE_BYTES = 65_536;
const MODES = new Set(['TrackAll', 'TrackMineOnly', 'Off']);
const STALE_MESSAGE = 'The document changed since this list was loaded. Refresh the changes before continuing.';

export class WordHostError extends Error {
  constructor(code, message, cause) {
    super(message, cause ? { cause } : undefined);
    this.name = 'WordHostError';
    this.code = code;
  }
}

const failure = (code, message, cause) => new WordHostError(code, message, cause);
const stale = reason => Object.assign(failure('STALE_CHANGE', STALE_MESSAGE), reason ? { reason } : {});

function snapshotXml(xml) {
  // Word for Mac regenerates editing-session RSIDs while serializing getOoxml.
  // They are save-session metadata, not the w:id identity of an ins/del revision.
  // Normalize only their eight-hex-digit attribute values inside element tags;
  // leave actual text, tracked revision IDs, authors, dates and all other XML.
  const normalized = xml.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<[A-Za-z_][\w:.-]*(?:[^"'<>]|"[^"]*"|'[^']*')*>/g, tag => {
    if (tag.startsWith('<!')) return tag;
    return tag.replace(/(\s+)([A-Za-z_:][\w:.-]*)(\s*=\s*)("[^"]*"|'[^']*')/g, (attribute, space, name, equals, quoted) => {
      if (!/^w:rsid(?:RDefault|RPr|Sect|Del|R|P)$/.test(name) || !/^[0-9A-Fa-f]{8}$/.test(quoted.slice(1, -1))) return attribute;
      return `${space}${name}${equals}${quoted[0]}00000000${quoted.at(-1)}`;
    });
  });
  // This same Mac serializer emits an empty final paragraph with a fresh ID.
  // Limit this exception to a self-closing paragraph with metadata attributes,
  // the observed empty-text sentinel, and section properties ending the body.
  const terminalNormalized = normalized.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<w:p(?:\s+[A-Za-z_:][\w:.-]*\s*=\s*(?:"[^"]*"|'[^']*'))*\s*\/>(?=\s*<w:sectPr\b)/g, (tag, offset) => {
    if (tag.startsWith('<!')) return tag;
    const tail = normalized.slice(offset + tag.length);
    const bodyEnd = tail.indexOf('</w:body>');
    if (bodyEnd < 0) return tag;
    const section = tail.slice(0, bodyEnd).trim();
    if (!/^<w:sectPr\b/.test(section) || !(section.endsWith('</w:sectPr>') || /^<w:sectPr(?:\s+[A-Za-z_:][\w:.-]*\s*=\s*(?:"[^"]*"|'[^']*'))*\s*\/>$/.test(section)) || /<w:p(?=[\s/>])/.test(section)) return tag;
    const attributes = [...tag.matchAll(/\s+([A-Za-z_:][\w:.-]*)\s*=\s*("[^"]*"|'[^']*')/g)];
    if (attributes.some(([, name]) => !['w14:paraId', 'w14:textId', 'w:rsidR', 'w:rsidRDefault'].includes(name))) return tag;
    if (!attributes.some(([, name, value]) => name === 'w14:textId' && value.slice(1, -1) === '77777777')) return tag;
    return tag.replace(/(\s+w14:paraId\s*=\s*)(["'])[0-9A-Fa-f]{8}\2/, (_, attribute, quote) => `${attribute}${quote}00000000${quote}`);
  });
  // Footnote/endnote separator definitions also receive fresh paragraph IDs.
  // Keep the separator's identity, content and formatting, and every ordinary
  // footnote paragraph ID. Only these explicitly typed synthetic notes qualify.
  const separatorsNormalized = terminalNormalized.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<w:(footnote|endnote)(?=[\s>])(?:[^"'<>]|"[^"]*"|'[^']*')*>[\s\S]*?<\/w:\1>/g, note => {
    if (note.startsWith('<!')) return note;
    const opening = note.match(/^<(?:[^"'<>]|"[^"]*"|'[^']*')*>/)[0];
    const attributes = [...opening.matchAll(/\s+([A-Za-z_:][\w:.-]*)\s*=\s*("[^"]*"|'[^']*')/g)];
    const type = attributes.find(([, name]) => name === 'w:type')?.[2].slice(1, -1);
    if (!['separator', 'continuationSeparator'].includes(type)) return note;
    return note.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<w:p(?=[\s/>])(?:[^"'<>]|"[^"]*"|'[^']*')*>/g, tag => {
      if (tag.startsWith('<!')) return tag;
      return tag.replace(/(\s+)([A-Za-z_:][\w:.-]*)(\s*=\s*)("[^"]*"|'[^']*')/g, (attribute, space, name, equals, quoted) => {
        if (name !== 'w14:paraId' || !/^[0-9A-Fa-f]{8}$/.test(quoted.slice(1, -1))) return attribute;
        return `${space}${name}${equals}${quoted[0]}00000000${quoted.at(-1)}`;
      });
    });
  });
  // The settings catalogue lists those same editing-session RSIDs and grows
  // during serialization. Match the entire known metadata-only grammar; an
  // unexpected attribute, child, comment or non-hex value stays byte-sensitive.
  return separatorsNormalized.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<w:rsids>(?:\s*<w:(?:rsidRoot|rsid)\s+w:val\s*=\s*(?:"[0-9A-Fa-f]{8}"|'[0-9A-Fa-f]{8}')\s*\/>)*\s*<\/w:rsids>/g,
    catalogue => catalogue.startsWith('<w:rsids>') ? '<w:rsids/>' : catalogue);
}

function metadata(change) {
  const date = change.date instanceof Date ? change.date : new Date(change.date ?? NaN);
  return {
    author: String(change.author ?? ''),
    date: Number.isFinite(date.getTime()) ? date.toISOString() : '',
    text: String(change.text ?? ''),
    type: String(change.type ?? 'None'),
  };
}

function documentName(url) {
  let name = String(url || '').split(/[?#]/, 1)[0].split(/[\\/]/).pop();
  try { name = decodeURIComponent(name); } catch { /* Keep a literal filename. */ }
  name = name.replace(/[\u0000-\u001f\u007f/\\]/g, '').trim();
  if (!name || !/\.docx$/i.test(name)) return 'Word document.docx';
  return name;
}

function base64(bytes) {
  // Chunk the conversion to avoid argument/stack limits on large documents.
  let binary = '';
  for (let offset = 0; offset < bytes.length; offset += 8192) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  }
  return globalThis.btoa(binary);
}

/** Injectable Office/Word objects let host behaviour be tested without Word. */
export function createWordHost(dependencies = {}) {
  const office = () => Object.hasOwn(dependencies, 'office') ? dependencies.office : globalThis.Office;
  const word = () => Object.hasOwn(dependencies, 'word') ? dependencies.word : globalThis.Word;
  const timeoutMs = dependencies.timeoutMs ?? 20_000;
  let readiness;
  let snapshot = null;
  let sequence = 0;
  let queue = Promise.resolve();
  const session = globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`;

  // All host actions are serialized: double clicks cannot act on an old list.
  const serial = action => {
    const result = queue.then(action);
    queue = result.catch(() => {});
    return result;
  };

  function asyncCall(invoke, code, message, onLateSuccess) {
    const api = office();
    return new Promise((resolve, reject) => {
      let settled = false;
      const timer = setTimeout(() => {
        settled = true;
        reject(failure('HOST_TIMEOUT', `${message} Word did not respond in time.`));
      }, timeoutMs);
      const callback = result => {
        const succeeded = result?.status === (api?.AsyncResultStatus?.Succeeded ?? 'succeeded');
        if (settled) {
          if (succeeded && onLateSuccess) onLateSuccess(result.value);
          return;
        }
        settled = true;
        clearTimeout(timer);
        if (succeeded) resolve(result.value);
        else reject(failure(code, message, result?.error));
      };
      try { invoke(callback); }
      catch (error) {
        if (!settled) {
          settled = true;
          clearTimeout(timer);
          reject(failure(code, message, error));
        }
      }
    });
  }

  async function waitForWord() {
    const api = office();
    const absent = reason => ({ available: false, host: null, platform: null,
      trackedChanges: false, trackingMode: false, captureDocument: false, reason });
    if (!api || typeof api.onReady !== 'function') return absent('Open Margin from Microsoft Word to use these controls.');
    if (!readiness) readiness = (async () => {
      let timer;
      try {
        const info = await Promise.race([
          api.onReady(),
          new Promise((_, reject) => { timer = setTimeout(() => reject(failure('HOST_TIMEOUT', 'Word did not become ready.')), timeoutMs); }),
        ]);
        const host = info?.host ?? api.context?.host ?? null;
        const platform = info?.platform ?? api.context?.platform ?? null;
        if (String(host).toLowerCase() !== 'word' || typeof word()?.run !== 'function') {
          return absent('Open this add-in in Microsoft Word.');
        }
        const supports = (set, version) => {
          try { return api.context?.requirements?.isSetSupported(set, version) === true; }
          catch { return false; }
        };
        const web = /officeonline|web/i.test(String(platform));
        return Object.freeze({
          available: true, host, platform,
          trackedChanges: supports('WordApi', '1.6'),
          trackingMode: supports('WordApi', '1.4'),
          captureDocument: !web && supports('CompressedFile', '1.1') && typeof api.context?.document?.getFileAsync === 'function',
          reason: '',
        });
      } catch (error) {
        readiness = null;
        return absent(error.code === 'HOST_TIMEOUT' ? error.message : 'Word could not initialize the add-in. Reopen Margin to try again.');
      } finally { clearTimeout(timer); }
    })();
    return readiness;
  }

  async function requireCapability(key) {
    const capabilities = await waitForWord();
    if (!capabilities.available) throw failure('UNSUPPORTED_HOST', capabilities.reason);
    if (!capabilities[key]) throw failure('UNSUPPORTED_API', key === 'captureDocument'
      ? 'This version of Word cannot send a document copy. Save a .docx and open it in Margin.'
      : 'This version of Word does not support this control. Update Word to use it.');
  }

  async function readBody(context) {
    const body = context.document.body;
    const collection = body.getTrackedChanges();
    collection.load('items/author,items/date,items/text,items/type');
    const xml = body.getOoxml();
    await context.sync();
    // OOXML includes revision identities and unchanged surrounding text.
    // Metadata alone cannot distinguish identical edits in repeated clauses.
    if (typeof xml.value !== 'string' || !xml.value) throw failure('HOST_READ_FAILED', 'Word could not provide a complete document snapshot.');
    const details = collection.items.map(metadata);
    return { items: collection.items, details, xml: snapshotXml(xml.value), signature: JSON.stringify(details) };
  }

  function listChanges() {
    return serial(async () => {
      await requireCapability('trackedChanges');
      snapshot = null;
      return word().run(async context => {
        const current = await readBody(context);
        const prefix = `${session}-${++sequence}`;
        const changes = current.details.map((detail, index) => ({ ...detail, id: `${prefix}-${index}`, textSource: detail.text.trim() ? 'tracked-change' : 'unavailable' }));
        // Some Word builds omit deleted text from TrackedChange.text. Read only
        // that change's own range, in its original version; never substitute the
        // enclosing paragraph, which could conceal the scope of the deletion.
        const missing = [];
        for (const [index, detail] of current.details.entries()) {
          if (detail.text.trim() || !['Added', 'Deleted'].includes(detail.type)) continue;
          try {
            const range = current.items[index].getRange();
            if (typeof range.getReviewedText === 'function') {
              missing.push({ index, result: range.getReviewedText(detail.type === 'Deleted' ? 'Original' : 'Current') });
            }
          } catch { /* Remain explicit about text Word could not expose. */ }
        }
        if (missing.length) {
          try {
            await context.sync();
            for (const { index, result } of missing) {
              if (typeof result.value === 'string' && result.value.trim()) {
                changes[index].text = result.value;
                changes[index].textSource = 'reviewed-range';
              }
            }
          } catch { /* Missing text requires review in Word before any decision. */ }
        }
        for (const change of changes) change.requiresWordReview = !change.text.trim();
        snapshot = { xml: current.xml, signature: current.signature, ids: new Map(changes.map((item, index) => [item.id, index])),
          requiresWordReview: new Set(changes.filter(item => item.requiresWordReview).map(item => item.id)), shownInWord: new Set() };
        return changes;
      });
    });
  }

  function withChange(id, operation, { selecting = false, deciding = false } = {}) {
    return serial(async () => {
      await requireCapability('trackedChanges');
      const expected = snapshot;
      if (typeof id !== 'string' || !expected?.ids.has(id)) throw stale();
      try {
        return await word().run(async context => {
          const current = await readBody(context);
          if (current.xml !== expected.xml || current.signature !== expected.signature) {
            snapshot = null;
            throw stale(current.signature !== expected.signature ? 'revision_list_changed' : 'body_snapshot_changed');
          }
          if (deciding && expected.requiresWordReview.has(id) && !expected.shownInWord.has(id)) {
            throw failure('WORD_REVIEW_REQUIRED', 'Word did not expose this change’s text. Use Show in Word to inspect it before deciding.');
          }
          const change = current.items[expected.ids.get(id)];
          if (!change) throw stale();
          // The freshly loaded native object is used inside this same Word.run.
          // No object path or collection index survives across Word.run calls.
          operation(change);
          await context.sync();
          if (selecting) expected.shownInWord.add(id);
        });
      } catch (error) {
        if (error.code === 'WORD_REVIEW_REQUIRED') throw error;
        snapshot = null;
        if (['InvalidObjectPath', 'ItemNotFound', 'InvalidObject', 'ObjectNotFound'].includes(error.code)) throw stale();
        throw error;
      }
    });
  }

  function selectChange(id) {
    return withChange(id, change => change.getRange().select('Select'), { selecting: true });
  }

  function decideChange(id, decision) {
    if (!['accept', 'reject'].includes(decision)) return Promise.reject(failure('INVALID_DECISION', 'Choose accept or reject.'));
    return withChange(id, change => {
      // Invalidated before the write: even an ambiguous host failure needs refresh.
      snapshot = null;
      change[decision]();
    }, { deciding: true });
  }

  function getTrackingMode() {
    return serial(async () => {
      await requireCapability('trackingMode');
      return word().run(async context => {
        context.document.load('changeTrackingMode');
        await context.sync();
        return context.document.changeTrackingMode;
      });
    });
  }

  function setTrackingMode(mode) {
    if (!MODES.has(mode)) return Promise.reject(failure('INVALID_MODE', 'Choose a supported Track Changes mode.'));
    return serial(async () => {
      await requireCapability('trackingMode');
      return word().run(async context => {
        context.document.changeTrackingMode = mode;
        snapshot = null;
        await context.sync();
        context.document.load('changeTrackingMode');
        await context.sync();
        return context.document.changeTrackingMode;
      });
    });
  }

  function captureDocument() {
    return serial(async () => {
      await requireCapability('captureDocument');
      const api = office();
      const doc = api.context.document;
      const file = await asyncCall(callback => doc.getFileAsync(api.FileType?.Compressed ?? 'compressed', { sliceSize: SLICE_BYTES }, callback),
        'FILE_READ_FAILED', 'Word could not create a document copy.',
        lateFile => { try { lateFile.closeAsync(() => {}); } catch { /* Best-effort release after timeout. */ } });
      let result;
      let problem;
      try {
        const { size, sliceCount } = file;
        if (!Number.isSafeInteger(size) || size <= 0 || !Number.isSafeInteger(sliceCount) || sliceCount <= 0) {
          throw failure('INVALID_FILE', 'Word returned an incomplete document copy.');
        }
        if (size > MAX_WORD_BYTES) throw failure('FILE_TOO_LARGE', 'Margin accepts Word documents up to 4 MB.');
        if (sliceCount !== Math.ceil(size / SLICE_BYTES)) throw failure('INVALID_FILE', 'Word returned inconsistent document slices.');
        const bytes = new Uint8Array(size);
        let offset = 0;
        for (let index = 0; index < sliceCount; index++) {
          const slice = await asyncCall(callback => file.getSliceAsync(index, callback), 'SLICE_READ_FAILED', 'Word could not read the complete document copy.');
          const data = slice?.data;
          const expectedSize = Math.min(SLICE_BYTES, size - offset);
          if (slice?.index !== index || (!Array.isArray(data) && !(data instanceof Uint8Array)) || data.length !== expectedSize ||
            (slice.size !== undefined && slice.size !== data.length) ||
            !Array.from(data).every(value => Number.isInteger(value) && value >= 0 && value <= 255)) {
            throw failure('INVALID_FILE', 'Word returned an invalid document slice.');
          }
          bytes.set(data, offset);
          offset += data.length;
        }
        if (bytes[0] !== 0x50 || bytes[1] !== 0x4b || bytes[2] !== 0x03 || bytes[3] !== 0x04) {
          throw failure('INVALID_FILE', 'Word did not return a .docx document.');
        }
        result = { data: base64(bytes), name: documentName(doc.url) };
      } catch (error) { problem = error; }
      try {
        await asyncCall(callback => file.closeAsync(callback), 'FILE_CLOSE_FAILED', 'Word could not release the document copy. Try reopening Margin.');
      } catch (error) {
        if (!problem) problem = error;
        else problem.closeError = error;
      }
      if (problem) throw problem;
      return result;
    });
  }

  return { waitForWord, listChanges, selectChange, decideChange, getTrackingMode, setTrackingMode, captureDocument };
}

const defaultHost = createWordHost();
export const waitForWord = (...args) => defaultHost.waitForWord(...args);
export const listChanges = (...args) => defaultHost.listChanges(...args);
export const selectChange = (...args) => defaultHost.selectChange(...args);
export const decideChange = (...args) => defaultHost.decideChange(...args);
export const getTrackingMode = (...args) => defaultHost.getTrackingMode(...args);
export const setTrackingMode = (...args) => defaultHost.setTrackingMode(...args);
export const captureDocument = (...args) => defaultHost.captureDocument(...args);

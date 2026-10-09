import {
  waitForWord, listChanges, selectChange, decideChange,
  getTrackingMode, setTrackingMode, captureDocument,
} from './word-host.js';

const $ = selector => document.querySelector(selector);
const state = {
  capabilities: null, changes: [], selectedId: null, tracking: null,
  busy: false, stale: false, loaded: false, config: null, configPromise: null, handoffUrl: null,
  inspected: new Set(),
};
const typeLabels = { Added: 'Addition', Deleted: 'Deletion', Formatted: 'Formatting', None: 'Change' };
const typeClasses = { Added: 'added', Deleted: 'deleted', Formatted: 'formatted', None: 'other' };
const dateFormatter = new Intl.DateTimeFormat(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
const timeFormatter = new Intl.DateTimeFormat(undefined, { hour: 'numeric', minute: '2-digit' });

function element(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}
function message(selector, text) {
  const node = $(selector);
  node.textContent = text || '';
  node.hidden = !text;
}
function announce(text) { $('#announcement').textContent = text; }
function readableError(error) {
  return error instanceof Error && error.message ? error.message : 'This action could not be completed. Please try again.';
}
function setBusy(value) {
  state.busy = value;
  $('#pane').classList.toggle('busy', value);
  $('#review-list').setAttribute('aria-busy', String(value));
  syncControls();
}
function syncControls() {
  const caps = state.capabilities || {};
  $('#refresh').disabled = state.busy || !caps.available;
  $('#tracking-toggle').disabled = state.busy || !caps.trackingMode || !state.tracking;
  $('#tracking-who').disabled = state.busy || !caps.trackingMode;
  $('#handoff').disabled = state.busy || !caps.captureDocument;
  $('#handoff-again').disabled = state.busy || !caps.captureDocument;
  $('#connect-again').disabled = state.busy;
  for (const button of document.querySelectorAll('[data-change-action]')) {
    const change = state.changes.find(item => item.id === button.dataset.changeId);
    const needsInspection = ['accept', 'reject'].includes(button.dataset.changeAction) && change?.requiresWordReview && !state.inspected.has(change.id);
    button.disabled = state.busy || state.stale || needsInspection;
  }
}
function renderTracking() {
  const mode = state.tracking;
  const on = mode === 'TrackAll' || mode === 'TrackMineOnly';
  $('#tracking-toggle').setAttribute('aria-checked', String(on));
  $('#tracking-toggle').hidden = !state.capabilities?.trackingMode || !mode;
  $('#tracking-description').textContent = !state.capabilities?.trackingMode
    ? 'Use Word’s Review tab to change this setting.'
    : mode === 'TrackAll' ? 'Recording everyone’s edits.'
      : mode === 'TrackMineOnly' ? 'Recording your edits.'
        : mode === 'Off' ? 'New edits are not being tracked.'
          : 'Could not read Word’s tracking setting.';
  $('#tracking-scope').hidden = !on;
  if (on) $('#tracking-who').value = mode;
  syncControls();
}
function emptyState(title, description) {
  const section = element('div', 'empty-state');
  section.append(element('h2', '', title), element('p', '', description));
  return section;
}
function revisionButton(text, action, id, className = 'button') {
  const button = element('button', className, text);
  button.type = 'button';
  button.dataset.changeAction = action;
  button.dataset.changeId = id;
  return button;
}
function renderChanges() {
  const list = $('#review-list');
  list.replaceChildren();
  $('#change-count').textContent = state.loaded ? String(state.changes.length) : '';
  $('#scope-note').hidden = !state.capabilities?.trackedChanges;
  if (!state.capabilities?.trackedChanges) {
    list.append(emptyState('Review in Word.', 'This version of Word does not expose individual changes to add-ins. You can keep using Word’s own Review tab.'));
    return;
  }
  if (state.stale) {
    list.append(emptyState('The document has changed.', 'Refresh to load its current changes before continuing.'));
    return;
  }
  if (!state.loaded) {
    list.append(emptyState('Reading the changes…', 'Margin is checking the main text of your open document.'));
    return;
  }
  if (!state.changes.length) {
    list.append(emptyState('No tracked changes\nin the main text.', state.tracking === 'Off'
      ? 'Turn on Track changes above to record your next edits.'
      : 'Keep writing in Word. Refresh here when you are ready to review.'));
    return;
  }
  state.changes.forEach((change, index) => {
    const selected = change.id === state.selectedId;
    const typeClass = typeClasses[change.type] || 'other';
    const article = element('article', `revision${selected ? ' is-selected' : ''}`);
    article.dataset.revisionId = change.id;
    const summary = revisionButton('', 'expand', change.id, 'revision-summary');
    summary.setAttribute('aria-expanded', String(selected));
    summary.setAttribute('aria-describedby', `revision-author-${index}`);
    const topline = element('span', 'revision-topline');
    const number = element('span', 'revision-number', String(index + 1).padStart(2, '0'));
    number.setAttribute('aria-hidden', 'true');
    topline.append(element('span', `revision-type ${typeClass}`, typeLabels[change.type] || 'Change'), element('span', 'sr-only', `Change ${index + 1} of ${state.changes.length}.`), number);
    const text = change.text || (change.requiresWordReview ? 'Review this change in Word.' : change.type === 'Formatted' ? 'A formatting change.' : 'A change without visible text.');
    const excerpt = element('span', `revision-text ${typeClass}${change.text ? '' : ' empty'}`, text);
    summary.append(topline, excerpt);
    article.append(summary);
    const byline = element('p', 'revision-byline');
    byline.id = `revision-author-${index}`;
    byline.append(element('span', '', change.author || 'Author not recorded'));
    if (change.date) {
      const date = new Date(change.date);
      if (!Number.isNaN(date.getTime())) {
        const time = element('time', '', dateFormatter.format(date));
        time.dateTime = date.toISOString();
        time.title = date.toLocaleString();
        byline.append(time);
      }
    }
    article.append(byline);
    if (selected) {
      const details = element('div', 'revision-details');
      details.append(revisionButton('Show in Word', 'select', change.id, 'text-button'));
      if (change.requiresWordReview) {
        const inspection = element('p', 'inspection-guide', state.inspected.has(change.id)
          ? 'Selected in Word. Check the passage there before deciding.'
          : 'Word has not provided a readable preview. Choose Show in Word and inspect the change before deciding.');
        inspection.id = `inspection-guide-${index}`;
        details.append(inspection);
      }
      const guidance = { Added: 'Accept keeps this addition. Reject removes it.', Deleted: 'Accept removes this text. Reject keeps it.', Formatted: 'Accept keeps the new formatting. Reject restores the previous formatting.' };
      if (guidance[change.type]) details.append(element('p', 'decision-guide', guidance[change.type]));
      const actions = element('div', 'revision-actions');
      const accept = revisionButton('Accept', 'accept', change.id);
      accept.setAttribute('aria-label', `Accept change ${index + 1}`);
      const reject = revisionButton('Reject', 'reject', change.id);
      reject.setAttribute('aria-label', `Reject change ${index + 1}`);
      if (change.requiresWordReview) {
        accept.setAttribute('aria-describedby', `inspection-guide-${index}`);
        reject.setAttribute('aria-describedby', `inspection-guide-${index}`);
      }
      actions.append(accept, reject);
      details.append(actions);
      article.append(details);
    }
    list.append(article);
  });
  syncControls();
}
function selectedCard() {
  return [...document.querySelectorAll('[data-revision-id]')].find(node => node.dataset.revisionId === state.selectedId);
}
function focusCurrent() {
  const summary = selectedCard()?.querySelector('.revision-summary');
  if (summary) summary.focus({ preventScroll: true });
  else $('#refresh').focus({ preventScroll: true });
}
function markStale(error) {
  state.stale = true;
  state.inspected.clear();
  state.loaded = false;
  state.changes = [];
  state.selectedId = null;
  $('#refresh-note').textContent = 'Refresh needed before your next decision.';
  $('#refresh-note').classList.add('is-stale');
  renderChanges();
  message('#error', error?.code === 'STALE_CHANGE'
    ? 'Your document changed since this list was loaded. Nothing was decided. Refresh to review the current version.'
    : readableError(error));
}
async function readCurrent() {
  state.inspected.clear();
  let trackingError = null;
  if (state.capabilities.trackingMode) {
    try { state.tracking = await getTrackingMode(); }
    catch (error) { state.tracking = null; trackingError = error; }
  }
  renderTracking();
  if (state.capabilities.trackedChanges) {
    const changes = await listChanges();
    if (!Array.isArray(changes)) throw new Error('Word returned an unreadable change list. Please refresh.');
    state.changes = changes;
    state.loaded = true;
    state.stale = false;
    if (!changes.some(change => change.id === state.selectedId)) state.selectedId = changes[0]?.id ?? null;
  }
  $('#refresh-note').textContent = `Updated ${timeFormatter.format(new Date())}. Refresh after editing in Word.`;
  $('#refresh-note').classList.remove('is-stale');
  renderChanges();
  if (trackingError) message('#error', `The changes are available, but the tracking setting could not be read. ${readableError(trackingError)}`);
}
async function refresh() {
  if (state.busy || !state.capabilities?.available) return;
  setBusy(true);
  message('#error', '');
  message('#notice', '');
  try {
    await readCurrent();
    announce(`${state.changes.length} tracked ${state.changes.length === 1 ? 'change' : 'changes'} in the main text.`);
  } catch (error) { markStale(error); }
  finally { setBusy(false); }
}
async function waitForOfficeScript() {
  if (globalThis.Office) return;
  const script = $('#office-runtime');
  if (!script) return;
  await new Promise(resolve => {
    let timer;
    const finish = () => {
      clearTimeout(timer);
      script.removeEventListener('load', finish);
      script.removeEventListener('error', finish);
      resolve();
    };
    script.addEventListener('load', finish, { once: true });
    script.addEventListener('error', finish, { once: true });
    timer = setTimeout(finish, 10_000);
    if (globalThis.Office) finish();
  });
}
async function connect() {
  if (state.busy) return;
  setBusy(true);
  $('#connection-description').textContent = 'Connecting to the document open in Word…';
  $('#connect-again').hidden = true;
  try {
    await waitForOfficeScript();
    state.capabilities = await waitForWord();
    if (!state.capabilities?.available) {
      $('#connection-description').textContent = state.capabilities?.reason || 'Open this pane inside Microsoft Word to review your document’s tracked changes.';
      $('#connection-links').hidden = false;
      $('#connect-again').hidden = false;
      return;
    }
    $('#connection-state').hidden = true;
    $('#word-workspace').hidden = false;
    $('#handoff-section').hidden = false;
    $('#host-label').textContent = 'IN WORD';
    if (!state.capabilities.captureDocument) {
      $('#handoff').hidden = true;
      $('#handoff-note').textContent = 'To continue locally, save this document as a .docx file and open it in the Margin web app.';
      const link = element('a', 'button primary', 'Open local web app ↗');
      link.dataset.webApp = '';
      configureWebAppLink(link);
      link.target = '_blank';
      link.rel = 'noopener';
      $('#handoff').after(link);
    }
    renderTracking();
    renderChanges();
    try { await readCurrent(); }
    catch (error) { markStale(error); }
  } catch (error) {
    $('#connection-description').textContent = readableError(error);
    $('#connection-links').hidden = false;
    $('#connect-again').hidden = false;
  } finally { setBusy(false); }
}
async function changeTracking(mode) {
  if (state.busy || !state.capabilities?.trackingMode) return;
  setBusy(true);
  message('#error', '');
  message('#notice', '');
  try {
    state.tracking = await setTrackingMode(mode);
    renderTracking();
    announce(state.tracking === 'Off' ? 'Track changes is off.' : state.tracking === 'TrackMineOnly' ? 'Tracking your changes.' : 'Tracking everyone’s changes.');
    try { await readCurrent(); }
    catch (error) { markStale(new Error(`The tracking setting was updated. Refresh before reviewing changes. ${readableError(error)}`)); }
  } catch (error) {
    // An unsuccessful read after a write leaves the current setting unknown.
    state.tracking = null;
    renderTracking();
    markStale(new Error(`${readableError(error)} Refresh to check Word’s current setting.`));
  } finally { setBusy(false); }
}
async function handleRevision(action, id) {
  if (state.busy || state.stale || !state.changes.some(change => change.id === id)) return;
  if (action === 'expand') {
    state.selectedId = id;
    renderChanges();
    focusCurrent();
    return;
  }
  setBusy(true);
  message('#error', '');
  message('#notice', '');
  try {
    if (action === 'select') {
      await selectChange(id);
      state.inspected.add(id);
      renderChanges();
      announce('Passage selected in Word. Review it there before accepting or rejecting.');
    } else if (action === 'accept' || action === 'reject') {
      const change = state.changes.find(item => item.id === id);
      if (change.requiresWordReview && !state.inspected.has(id)) {
        message('#notice', 'Choose Show in Word and inspect this change before deciding.');
        return;
      }
      const index = state.changes.findIndex(change => change.id === id);
      await decideChange(id, action);
      state.selectedId = null;
      const decision = action === 'accept' ? 'Change accepted in Word.' : 'Change rejected in Word.';
      message('#notice', `${decision} Use Word’s Undo to reverse it.`);
      announce(decision);
      try {
        await readCurrent();
        state.selectedId = state.changes[Math.min(index, state.changes.length - 1)]?.id ?? null;
        renderChanges();
      } catch (error) {
        markStale(new Error(`${decision} The list could not be refreshed. Refresh before continuing. ${readableError(error)}`));
      }
    }
  } catch (error) {
    if (error?.code === 'STALE_CHANGE') markStale(error);
    else if (error?.code === 'WORD_REVIEW_REQUIRED') {
      state.inspected.delete(id);
      renderChanges();
      message('#notice', 'Choose Show in Word and inspect this change before deciding.');
    } else if (action === 'accept' || action === 'reject') {
      markStale(new Error(`Word did not confirm the decision. Refresh to check the document before trying again. ${readableError(error)}`));
    } else markStale(error);
  } finally {
    setBusy(false);
    if (action === 'accept' || action === 'reject') focusCurrent();
  }
}
function validateWebOrigin(value) {
  const url = new URL(value);
  if (url.protocol !== 'http:' || url.hostname !== '127.0.0.1' || url.pathname !== '/' ||
      url.search || url.hash || url.username || url.password) {
    throw new Error('Margin returned an invalid local workspace address.');
  }
  return url.origin;
}
function configureWebAppLink(link) {
  if (state.config) {
    link.href = `${state.config.webOrigin}/`;
    link.removeAttribute('aria-disabled');
    link.removeAttribute('title');
  } else {
    link.removeAttribute('href');
    link.setAttribute('aria-disabled', 'true');
    link.title = 'The local workspace link is unavailable. Reload this pane to reconnect.';
  }
}
async function loadConfig() {
  if (state.config) return state.config;
  if (state.configPromise) return state.configPromise;
  const pending = (async () => {
    const response = await fetch('/api/config', { cache: 'no-store' });
    if (!response.ok) throw new Error('The local Margin workspace is unavailable. Keep Margin running and try again.');
    const config = await response.json();
    config.webOrigin = validateWebOrigin(config.webOrigin);
    state.config = config;
    document.querySelectorAll('[data-web-app]').forEach(configureWebAppLink);
    return config;
  })();
  state.configPromise = pending;
  try { return await pending; }
  finally { if (state.configPromise === pending) state.configPromise = null; }
}
async function handoffRequest(payload, retry = true) {
  await loadConfig();
  const response = await fetch('/api/word/handoff', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'X-Folio-Token': state.config.token },
    body: JSON.stringify(payload),
  });
  const result = await response.json();
  if (response.status === 403 && retry) {
    state.config = null;
    return handoffRequest(payload, false);
  }
  if (!response.ok) throw new Error(result.error || 'The local workspace could not prepare this copy.');
  return result;
}
async function handoff() {
  if (state.busy || !state.capabilities?.captureDocument) return;
  setBusy(true);
  message('#handoff-error', '');
  $('#handoff').textContent = 'Preparing your copy…';
  try {
    const documentCopy = await captureDocument();
    const result = await handoffRequest(documentCopy);
    const url = new URL(result.url, location.origin);
    const expectedOrigin = validateWebOrigin(state.config.webOrigin);
    if (url.origin !== expectedOrigin || url.pathname !== '/' || url.search || url.username || url.password || !/^#word=[a-f0-9]{48}$/.test(url.hash)) {
      throw new Error('Margin returned an invalid local workspace link. Please try again.');
    }
    state.handoffUrl = url.href;
    $('#handoff-link').href = url.href;
    $('#handoff').hidden = true;
    $('#handoff-result').hidden = false;
    const expires = new Date(result.expiresAt);
    $('#handoff-expiry').textContent = Number.isNaN(expires.getTime())
      ? 'Your copy is ready in the local workspace.'
      : `Your copy is ready. Open it by ${timeFormatter.format(expires)}.`;
    $('#handoff-note').textContent = 'A separate copy on this computer. Later edits in Word or Margin do not sync between them.';
    announce('Your document copy is ready. Continue in the browser to open it in Margin.');
    $('#handoff-link').focus();
  } catch (error) { message('#handoff-error', readableError(error)); }
  finally {
    $('#handoff').textContent = 'Open in Margin ↗';
    setBusy(false);
  }
}

$('#refresh').addEventListener('click', refresh);
$('#connect-again').addEventListener('click', connect);
$('#tracking-toggle').addEventListener('click', () => changeTracking(state.tracking === 'Off' ? 'TrackAll' : 'Off'));
$('#tracking-who').addEventListener('change', event => changeTracking(event.target.value));
$('#review-list').addEventListener('click', event => {
  const button = event.target.closest('[data-change-action]');
  if (button) handleRevision(button.dataset.changeAction, button.dataset.changeId);
});
$('#handoff').addEventListener('click', handoff);
$('#handoff-again').addEventListener('click', handoff);
window.addEventListener('focus', () => {
  if (state.loaded && !state.busy && !state.stale) {
    $('#refresh-note').textContent = 'Word may have changed. Refresh to update this list.';
  }
});
loadConfig().catch(() => document.querySelectorAll('[data-web-app]').forEach(configureWebAppLink));
connect();

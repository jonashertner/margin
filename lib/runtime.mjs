import {accessSync, constants, realpathSync} from 'node:fs';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const run = promisify(execFile);
export const projectRoot = fileURLToPath(new URL('../', import.meta.url));
const standardBins = ['/opt/homebrew/bin', '/usr/local/bin', '/Library/TeX/texbin', '/usr/bin', '/bin'];
const configured = names => names.map(name => process.env[name]).find(Boolean);

export function findExecutable(name, overrides = [], extra = []) {
  const selected = configured(overrides);
  const candidates = selected
    ? (selected.includes('/') ? [path.resolve(selected)] : (process.env.PATH || '').split(path.delimiter).filter(Boolean).map(dir => path.join(dir, selected)))
    : [...extra, ...[...(process.env.PATH || '').split(path.delimiter), ...standardBins].filter(Boolean).map(dir => path.join(dir, name))];
  for (const candidate of candidates) {
    try { accessSync(candidate, constants.X_OK); return path.resolve(candidate); } catch {}
  }
  return null;
}

// Keep the invoked path: resolving a Python venv symlink loses its environment,
// and resolving xelatex to xetex loses the selected TeX format.
export const wordPython = findExecutable('python3', ['MARGIN_WORD_PYTHON', 'FOLIO_WORD_PYTHON', 'MARGIN_PYTHON', 'FOLIO_PYTHON'], [path.join(projectRoot, '.venv/bin/python3')]);
export const pdfPython = findExecutable('python3', ['MARGIN_PYTHON', 'FOLIO_PYTHON'], [path.join(projectRoot, '.venv/bin/python3')]);
export const pandoc = findExecutable('pandoc', ['MARGIN_PANDOC', 'FOLIO_PANDOC']);
const selectedTexRoot = configured(['MARGIN_TEX_ROOT', 'FOLIO_TEX_ROOT']);
const texCandidates = selectedTexRoot ? ['universal-darwin', 'arm64-darwin', 'x86_64-darwin'].map(arch => path.join(selectedTexRoot, 'bin', arch, 'xelatex')) : [];
export const xelatex = findExecutable('xelatex', ['MARGIN_XELATEX', 'FOLIO_XELATEX'], texCandidates);
export const pdftoppm = findExecutable('pdftoppm', ['MARGIN_PDFTOPPM', 'FOLIO_PDFTOPPM']);
export const dataDir = configured(['MARGIN_DATA_DIR', 'FOLIO_DATA_DIR']) || path.join(projectRoot, '.margin-data');
export const sandbox = process.platform === 'darwin' ? findExecutable('sandbox-exec', [], ['/usr/bin/sandbox-exec']) : null;
export const texRoot = selectedTexRoot || (() => {
  if (!xelatex) return null;
  const real = realpathSync(xelatex), marker = `${path.sep}bin${path.sep}`;
  return real.includes(marker) ? real.slice(0, real.lastIndexOf(marker)) : path.dirname(real);
})();

export const systemReadRoots = ['/System', '/usr/lib', '/usr/share', '/Library/Apple', '/Library/Fonts', '/private/var/db/timezone', '/opt/homebrew/Cellar', '/opt/homebrew/opt', '/opt/homebrew/lib', '/opt/homebrew/bin', '/usr/local/Cellar', '/usr/local/opt', '/usr/local/lib', '/usr/local/bin'];
const pythonCache = new Map();
export async function pythonRuntime(executable) {
  if (!executable) return {roots:[], modules:{}};
  if (!pythonCache.has(executable)) pythonCache.set(executable, (async () => {
    try {
      const script = "import json,sys,sysconfig,importlib.util; print(json.dumps({'roots':list({sys.prefix,sys.base_prefix,sys.exec_prefix,*sysconfig.get_paths().values()}),'modules':{name:importlib.util.find_spec(name) is not None for name in ['lxml','pypdf']}}))";
      const result = await run(executable, ['-I', '-B', '-c', script], {timeout:10000, maxBuffer:50000});
      const value = JSON.parse(result.stdout);
      // Never derive a filesystem-wide or whole-home allowance from a runtime.
      value.roots = [...new Set(value.roots.filter(p => typeof p === 'string' && path.isAbsolute(p) && p !== '/' && p !== process.env.HOME))];
      return value;
    } catch { return {roots:[], modules:{}}; }
  })());
  return pythonCache.get(executable);
}

export function dependencyRoots(executable) {
  if (!executable) return [];
  return [path.dirname(executable), path.dirname(realpathSync(executable))];
}

export function sandboxProfile(job, roots) {
  const readable = [...new Set([job, ...systemReadRoots, ...roots].filter(Boolean))];
  return `(version 1)(deny default)(allow process*)(allow sysctl-read)(allow mach-lookup)(allow file-read-metadata)(allow file-read* (literal "/") (literal "/dev/urandom") (literal "/dev/null") ${readable.map(p => `(subpath ${JSON.stringify(p)})`).join(' ')})(allow file-write* (subpath ${JSON.stringify(job)}) (literal "/dev/null"))`;
}

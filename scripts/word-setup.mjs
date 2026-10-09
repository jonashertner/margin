#!/usr/bin/env node
import { X509Certificate, createPrivateKey, randomUUID } from 'node:crypto';
import { execFile, spawn } from 'node:child_process';
import { constants } from 'node:fs';
import { copyFile, lstat, mkdir, readFile, rename, unlink, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const manifestId = '47320ed7-0921-4fa2-987b-c2d78a7dbd91';
const certificateDirectory = join(homedir(), '.office-addin-dev-certs');
const certificatePath = process.env.MARGIN_TLS_CERT ? resolve(process.env.MARGIN_TLS_CERT) : join(certificateDirectory, 'localhost.crt');
const privateKeyPath = process.env.MARGIN_TLS_KEY ? resolve(process.env.MARGIN_TLS_KEY) : join(certificateDirectory, 'localhost.key');
const portText = process.env.MARGIN_HTTPS_PORT || '4318';
const port = Number(portText);
const destination = join(homedir(), 'Library/Containers/com.microsoft.Word/Data/Documents/wef/margin-local.xml');

function checkPort() {
  if (!/^\d+$/.test(portText) || !Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error('MARGIN_HTTPS_PORT must be an integer from 1 to 65535.');
  }
}

async function certificateStatus() {
  try {
    const certificate = new X509Certificate(await readFile(certificatePath));
    const privateKey = createPrivateKey(await readFile(privateKeyPath));
    if (!certificate.checkPrivateKey(privateKey)) throw new Error('Certificate and private key do not match.');
    if (!certificate.checkHost('localhost')) throw new Error('Certificate does not cover localhost.');
    const now = Date.now();
    if (now < Date.parse(certificate.validFrom)) throw new Error('Certificate is not valid yet.');
    if (now >= Date.parse(certificate.validTo)) throw new Error('Certificate has expired.');
    return { ready: true, expires: certificate.validTo };
  } catch (error) {
    return { ready: false, reason: error.code === 'ENOENT' ? 'Certificate or private key is missing.' : error.message };
  }
}

async function run(command, args, env = process.env) {
  await new Promise((resolveRun, reject) => {
    const child = spawn(command, args, { cwd: root, env, stdio: 'inherit' });
    child.once('error', reject);
    child.once('exit', (code, signal) => {
      if (code === 0) resolveRun();
      else reject(new Error(`${command} exited with ${code ?? signal}.`));
    });
  });
}

async function officialCertificates(command) {
  // This optional tool is deliberately absent from Margin's dependency tree.
  // It runs only for an explicit certificate command, never during startup.
  const args = ['exec', '--yes', '--ignore-scripts', '--package=office-addin-dev-certs@3.0.1', '--', 'office-addin-dev-certs', command];
  console.log('Running Microsoft’s office-addin-dev-certs@3.0.1 using npm’s isolated package cache.');
  if (process.env.npm_execpath) {
    await run(process.execPath, [process.env.npm_execpath, ...args]);
  } else if (process.platform !== 'win32') {
    await run('npm', args);
  } else {
    throw new Error('Run this certificate command through npm run word:setup so the npm executable is available.');
  }
}

async function verifyCertificate() {
  const certificate = await certificateStatus();
  if (!certificate.ready) throw new Error(certificate.reason);
  if (process.platform === 'darwin') {
    // Verify the actual certificate with macOS trust, SSL policy and hostname.
    // No npm package download or trust-store mutation is needed.
    await execFileAsync('/usr/bin/security', ['verify-cert', '-q', '-L', '-c', certificatePath, '-p', 'ssl', '-n', 'localhost'], { timeout: 15_000, maxBuffer: 64 * 1024 });
    console.log(`macOS trusts this localhost certificate; certificate and key match. Expires ${certificate.expires}.`);
  } else {
    if (process.env.MARGIN_TLS_CERT || process.env.MARGIN_TLS_KEY) {
      throw new Error('Certificate and key match, but custom certificate system trust must be checked with your operating system’s certificate tools.');
    }
    await officialCertificates('verify');
  }
}

async function readSideload(manifestPath = destination) {
  try {
    if (!(await lstat(manifestPath)).isFile()) throw new Error(`Refusing to change a non-file at ${manifestPath}`);
    const contents = await readFile(manifestPath, 'utf8');
    if (!new RegExp(`<Id>\\s*${manifestId}\\s*</Id>`, 'i').test(contents)) {
      throw new Error(`A different add-in uses ${manifestPath}. It has been left untouched.`);
    }
    return contents;
  } catch (error) {
    if (error.code === 'ENOENT') return null;
    throw error;
  }
}

async function status() {
  const certificate = await certificateStatus();
  console.log('Margin for Word · local setup\n');
  console.log(`Task pane: https://localhost:${port}/word`);
  console.log(`Certificate: ${certificatePath}`);
  console.log(certificate.ready ? `Certificate and key match; valid until ${certificate.expires}.` : certificate.reason);
  console.log('System trust is separate. Run npm run word:setup -- verify-certificate to check it.');
  if (process.platform === 'darwin') {
    const existing = await readSideload();
    console.log(`Word manifest: ${existing ? 'installed' : 'not installed'} (${destination})`);
  } else {
    console.log('Automatic sideloading is available on macOS. See docs/word-add-in.md for other hosts.');
  }
  console.log('\nCommands:');
  console.log('  npm run word:setup -- install-certificate  Create/trust the development certificate');
  console.log('  npm run word:setup -- verify-certificate   Check system trust for localhost');
  console.log('  npm run word:setup -- sideload             Add Margin to Word on this Mac');
  console.log('  npm run word:start                         Start the local web app and HTTPS task pane');
  console.log('  npm run word:setup -- unsideload           Remove this Margin manifest only');
}

async function sideload() {
  if (process.platform !== 'darwin') throw new Error('Automatic sideloading is macOS-only. Follow docs/word-add-in.md for Windows.');
  const source = (await readFile(join(root, 'office/manifest.xml'), 'utf8'))
    .replaceAll('https://localhost:4318', `https://localhost:${port}`);
  const result = await installWordManifest(destination, source);
  if (result.backup) console.log(`Previous Margin manifest preserved at ${result.backup}.`);
  console.log(`Margin manifest ${result.changed ? 'installed' : 'already installed'}: ${destination}`);
  console.log('Start Margin with npm run word:start, then open or restart Word and select Home → Add-ins → Margin.');
}

// Exported for isolated filesystem tests; normal CLI use always targets Word's
// documented macOS sideload directory and the checked-in Margin manifest.
export async function installWordManifest(manifestPath, source) {
  if (!new RegExp(`<Id>\\s*${manifestId}\\s*</Id>`, 'i').test(source)) {
    throw new Error('The source is not the Margin manifest.');
  }
  const existing = await readSideload(manifestPath);
  if (source === existing) {
    return { changed: false, backup: null };
  }
  await mkdir(dirname(manifestPath), { recursive: true });
  let backup = null;
  if (existing !== null) {
    backup = `${manifestPath}.backup-${Date.now()}-${randomUUID().slice(0, 8)}`;
    await copyFile(manifestPath, backup, constants.COPYFILE_EXCL);
    const temporary = `${manifestPath}.new-${randomUUID()}`;
    try {
      await writeFile(temporary, source, { flag: 'wx', mode: 0o600 });
      await rename(temporary, manifestPath);
    } finally {
      await unlink(temporary).catch(error => { if (error.code !== 'ENOENT') throw error; });
    }
  } else {
    await writeFile(manifestPath, source, { flag: 'wx', mode: 0o600 });
  }
  return { changed: true, backup };
}

async function unsideload() {
  if (process.platform !== 'darwin') throw new Error('Automatic sideload removal is macOS-only.');
  if (await readSideload() === null) {
    console.log('The Margin manifest is not installed.');
    return;
  }
  await unlink(destination);
  console.log('The Margin manifest was removed. Other add-ins and certificates were left in place. Restart Word.');
}

async function start() {
  const certificate = await certificateStatus();
  if (!certificate.ready) throw new Error(`${certificate.reason} Run npm run word:setup -- install-certificate first, or set MARGIN_TLS_CERT and MARGIN_TLS_KEY.`);
  console.log(`Using the existing certificate; no trust settings are changed. Expires ${certificate.expires}.`);
  await run(process.execPath, [join(root, 'server.mjs')], {
    ...process.env,
    MARGIN_TLS_CERT: certificatePath,
    MARGIN_TLS_KEY: privateKeyPath,
    MARGIN_HTTPS_PORT: String(port),
  });
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) try {
  checkPort();
  const command = process.argv[2] || 'status';
  if (process.argv.length > 3) throw new Error('Pass one setup command. Run npm run word:setup for help.');
  switch (command) {
    case 'status': case '--help': case '-h': await status(); break;
    case 'install-certificate':
      console.log('Microsoft’s certificate utility will create or renew its development certificate and add its CA to your user trust store. Your operating system may ask for confirmation.');
      await officialCertificates('install');
      break;
    case 'verify-certificate': await verifyCertificate(); break;
    case 'sideload': await sideload(); break;
    case 'unsideload': await unsideload(); break;
    case 'start': await start(); break;
    default: throw new Error(`Unknown setup command: ${command}. Run npm run word:setup for help.`);
  }
} catch (error) {
  console.error(`Margin: ${error.message}`);
  process.exitCode = 1;
}

# Margin in Word

Margin has two entrances: a local web app and a task pane inside Microsoft Word. Both run from the same checkout and local server. The Word pane uses Word’s document APIs; the full workspace retains Margin’s comparison, conversion and export tools.

This is a development add-in installed on your own computer, not a Microsoft Marketplace listing. The manifest has passed Microsoft’s validator. That validates its packaging, not every operation in every Word version.

## Set up on a Mac

Install Node.js 22 or later, then run these commands in your Margin checkout:

```sh
npm ci
npm run word:setup
```

The second command reports your existing certificate and sideload status without changing either. Word needs a trusted HTTPS origin, even though Margin runs locally. If you do not already have a trusted Office development certificate:

```sh
npm run word:setup -- install-certificate
```

This explicitly downloads and runs Microsoft’s `office-addin-dev-certs@3.0.1` utility through npm’s package cache. It is not installed as a Margin dependency. It creates or renews a certificate in `~/.office-addin-dev-certs` and installs its development certificate authority in your user trust store. macOS may ask for confirmation. The certificate is shared with other Office add-in development projects and normally expires after 30 days. Ordinary Margin startup never downloads this tool, installs a certificate or changes trust. Check existing system trust with:

```sh
npm run word:setup -- verify-certificate
```

On macOS this verifies the certificate with the operating system’s `security` tool, including the SSL policy and `localhost` hostname; it also checks expiry and the matching private key. It does not download a package or change trust. Other platforms use Microsoft’s optional certificate utility.

Install the manifest, then start Margin:

```sh
npm run word:setup -- sideload
npm run word:start
```

The web app is at [http://127.0.0.1:4317](http://127.0.0.1:4317). The Word pane is served at [https://localhost:4318/word](https://localhost:4318/word). Keep the terminal running. Stop an existing Margin server before starting this one if it already occupies port 4317.

After this one-time setup, double-click **Start Margin in Word.command** in Finder to start both the Word pane and the local web app. It uses your existing certificate and explains any missing setup; it never installs a certificate or changes trust automatically. The original **Start Margin.command** remains available for the web app alone.

Open or restart Word, open a disposable test document, and select **Home → Add-ins → Margin**. Once loaded, **Open Margin** is also available in the Home ribbon. Word’s menus can vary by version and organisation policy.

The sideload command writes only `margin-local.xml` in Word’s documented `wef` directory. It refuses to replace a foreign manifest or a symlink. Updating Margin’s own manifest preserves a backup of the preceding version.

## Review in Word

Word remains the editor. In the Margin pane:

1. Turn on Track changes for everyone or just yourself when drafting.
2. Read each listed revision with its author, date and type. Expand a revision and use **Show in Word** to see its location.
3. **Accept** or **Reject** applies that decision to the open Word document. Use Word’s Undo to reverse it.
4. Use **Refresh** after editing in Word. If the document has changed since the pane read it, Margin requires a refresh before applying a decision.

Word may expose a replacement as separate deletion and addition revisions. Decide both in context; rejecting only the deletion can temporarily leave both wordings present. The list covers the main document body. Review headers, footnotes and text boxes through Word’s own Review tools. Native review needs no document conversion.

For Margin’s larger workspace or LaTeX/PDF tools, choose **Open in Margin**, then **Continue in browser**. This captures a separate DOCX snapshot, up to 4 MB. The two copies are not live-synchronised: later edits in either workspace do not change the other automatically. Export a reviewed document from the web workspace when needed.

## Local documents and network use

The standalone web app needs no externally hosted JavaScript. The Word pane loads Microsoft’s Office.js runtime from its official CDN, so the add-in itself is not a fully offline application. Office, the certificate tools and the manifest validator may also make their own network requests.

Margin sends document data only to its server on this computer. Conversion and PDF jobs keep their existing filesystem and network sandbox. There is no Margin account, document upload service or model API. Word can independently sync a document through its own OneDrive or SharePoint settings; the add-in does not change those settings. A snapshot obtained from Word represents the current in-memory document, which may include unsaved edits, rather than a byte-for-byte copy of the last saved file.

The Word API permission is `ReadWriteDocument` because review actions can change the open document. Save a separate working copy when evaluating the add-in. See [security and document handling](../SECURITY.md) for local storage and converter boundaries.

## Compatibility

| Capability | Current boundary |
| --- | --- |
| Local web workspace | Modern desktop and phone browsers; Node.js server on the same computer. |
| Word pane | Office Add-ins host; each feature checks the relevant Word API at runtime. |
| Native tracked-change review | Requires WordApi 1.6; unsupported hosts receive an explanation. |
| Whole-document snapshot | Word’s `getFileAsync` with `CompressedFile`; Word on the web does not support this API. Save a `.docx` and open it in the local workspace there. |
| Word import/export and LaTeX/PDF conversion | Currently macOS only, with the dependencies in the main README. |
| Automatic sideload setup | macOS desktop Word. |

The manifest intentionally does not make WordApi 1.6 an installation prerequisite: a host can still display the pane and explain which actions it supports. A manifest validation pass is not a claim of Windows, web or iPad integration testing. A phone’s `localhost` points to that phone, not to your desktop; this development setup does not expose Margin to your network.

For Windows, Microsoft documents [sideloading through a network-share catalogue](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins). Use the provided manifest with a locally trusted server; conversion remains unavailable until a Windows isolation boundary is implemented. No automatic Windows or iPad installer is provided.

## Custom certificate or port

Use a certificate and key for `localhost` that your operating system already trusts:

```sh
MARGIN_TLS_CERT=/absolute/path/localhost.crt \
MARGIN_TLS_KEY=/absolute/path/localhost.key \
npm run word:start
```

To use a different HTTPS port, use the same value for sideloading and startup:

```sh
MARGIN_HTTPS_PORT=4418 npm run word:setup -- sideload
MARGIN_HTTPS_PORT=4418 npm run word:start
```

The sideload command adjusts URLs in the installed manifest. The source manifest stays on the default port, 4318. The `install-certificate` command always manages Microsoft’s standard certificate directory; it does not alter a custom certificate.

## Troubleshooting and removal

- **Word cannot load Margin:** open the HTTPS pane address in your browser, check that the server is running, and verify certificate trust. An expired certificate requires the explicit certificate installation command again.
- **Margin is missing from Add-ins:** restart Word after sideloading. Confirm your organisation permits custom Office Add-ins. Follow Microsoft’s [Mac sideloading instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac) if your menu differs.
- **Conversion is unavailable:** the Word pane and native review do not install Pandoc or a TeX distribution. Follow the [main setup instructions](../README.md).
- **Port already in use:** stop the earlier Margin server, or configure matching HTTP and HTTPS ports as described in the main README and above.
- **Remove the local add-in:** run `npm run word:setup -- unsideload`, then restart Word. Only Margin’s manifest is removed; no other add-in or shared certificate is touched.

Validate packaging after editing the manifest with `npm run word:validate`. Development certificates belong outside this repository; never commit a private key.

## Optional certificate tool dependency

As of 9 October 2026, Microsoft’s certificate utility depends on `mkcert` and `node-forge`, whose current release has an [unpatched signature-verification advisory](https://github.com/advisories/GHSA-86w9-cpqp-85rv). Margin therefore keeps that utility out of its installed dependency tree and loads it only for an explicit certificate installation command, or verification on non-Mac platforms. This packaging choice does not fix the upstream issue.

The inspected certificate-generation path creates and signs new local keys; its validation uses Node’s crypto and operating-system trust tools. No call to the affected `node-forge` signature-verification API was identified in that workflow. Margin’s server uses Node’s TLS implementation and never loads `node-forge`. An existing trusted localhost certificate can be supplied instead, as described above.

For contributors: `fs-extra` is an explicit development dependency because the manifest validator’s `@microsoft/app-manifest@1.1.2` dependency imports it at runtime but declares it only as a development dependency. The explicit pin keeps `npm run word:validate` working after a clean installation; the application server does not use it.

## API references

- [Microsoft: Office development certificates](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-dev-certs)
- [Microsoft: Word tracked changes](https://learn.microsoft.com/en-us/office/dev/add-ins/word/manage-tracked-changes)
- [Microsoft: Document.getFileAsync and platform support](https://learn.microsoft.com/en-us/javascript/api/office/office.document?view=office-js#office-office-document-getfileasync-member(1))

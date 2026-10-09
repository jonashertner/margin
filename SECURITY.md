# Security and document handling

Margin is an experimental, single-user local application. Do not expose its HTTP port to a network, run it behind a public reverse proxy, or treat its review record as an authenticated legal audit trail.

## Current boundary

The server binds to `127.0.0.1`. Requests must match the local Host and Origin, and state-changing API calls require a per-process token. These controls reduce unwanted browser requests; they do not defend against another process or user that already has access to the same computer. The token is not an account system.

Word and PDF conversion run only on macOS, under `sandbox-exec`. Jobs have temporary writable directories, limited read access to their selected toolchains and system dependencies, no network permission, and time limits. Python dependency discovery examines the configured, trusted interpreter. Executable paths and runtime installations must be trusted. TeX shell escape is disabled. The converters reject several unsafe archive, XML and external-resource forms, but this prototype has not received a comprehensive hostile-document security audit.

Do not remove the sandbox check to enable conversion on another platform. Linux and Windows currently support the source-review interface and pure source operations; Word/PDF conversion needs a separately designed and tested isolation boundary there. The Python engine tests call the converter directly with fictional fixtures and are not a secure route for processing untrusted files.

The application uses no cloud service, analytics or model API. Original Word files are copied to the selected data directory, with restrictive creation permissions and SHA-256 integrity checks. Active reviews are also stored in browser local storage. Neither location is encrypted by the application. Use appropriate device and account protection. Portable review files contain document text, comments, author information and original Word bytes; handle them as you would the underlying documents.

Imports, conversions, hashes and exports do not imply legal approval, a signature, sending or filing. A clean Word export applies revision decisions but can still contain comments. Review the actual outgoing file before sharing it.

## Reporting a vulnerability

Do not attach a confidential document, credential or exploitable payload to a public issue. Prefer a synthetic reproduction. Use GitHub's private vulnerability reporting if it is enabled for this repository; otherwise open a minimal issue requesting a private reporting channel without disclosing the exploit or private material. No response-time commitment or formal security support policy is offered for this prototype.

# Security Policy

## Reporting a vulnerability

Report security issues privately through the repository's GitHub security advisory page. Do not publish sensitive reports in a regular issue.

## Current security boundary

The shipped product is a browser renderer packaged as an MCP App. It contains no Electron main/preload runtime, IPC bridge, product AI provider, account login, telemetry client, or stored model credentials.

- The editor resource has a restrictive Content Security Policy and is packaged as self-contained HTML.
- Host commands use typed operation schemas and a monotonically increasing document revision.
- Mutations fail while the live editor is offline; there is no hidden second document authority.
- File access uses browser capabilities and explicit user permission, with download fallback when a writable handle is unavailable.
- The MCP server exposes only the documented resources and typed live-session interfaces; it does not provide an arbitrary shell or a second editable document store.
- All five packaged editors use the same live-session boundary. A format is not described as release-ready until it has passed the repository's current evidence gate.

## Out of scope

- Vulnerabilities requiring an already compromised machine or a deliberately modified build.
- Security properties of disabled migration inputs that cannot be reached from the installed TandemFolio plugin.

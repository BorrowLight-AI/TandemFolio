# Getting started

For prebuilt Windows/macOS release downloads, installation scripts, and updates, see
[Install a release](distribution.md). The release workflow publishes those bundles only
after the required evidence and native smoke checks pass.

The source-checkout workflow below needs Node.js 22.12+ and npm 10+. A compatible MCP Apps host is required for the embedded visual experience.

## Start the local server

From the repository root, install dependencies and start the MCP server:

```bash
npm install --ignore-scripts && npm run dev
```

This is the fastest way to run the local development server. It serves the packaged editor resources already present in the checkout.

## Build the plugin

Build all five web editors, the MCP server, and the self-contained plugin package:

```bash
npm install --ignore-scripts && npm run build
```

The built local plugin lives in `plugins/tandemfolio/`.

## Open TandemFolio in Codex

After building, register this checkout as a local marketplace and install the plugin:

```bash
codex plugin marketplace add .
codex plugin add tandemfolio@personal
```

Start a new Codex task after installation so it discovers the plugin's Skill and MCP server. If your Codex setup already uses a marketplace named `personal`, resolve that marketplace identity before installing rather than editing its configuration by hand.

## Work on one editor in a browser

For focused DOCX renderer work, run:

```bash
npm run dev:editor
```

The full development, test, package, and release workflow is documented in [Development guide](development.md).

## Before sharing a build

TandemFolio is pre-release. Run the relevant checks for your change, review the [project facts and attribution](project-facts.md), and do not advertise a release as ready until the evidence gate passes.

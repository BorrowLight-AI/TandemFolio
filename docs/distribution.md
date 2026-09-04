# Install a TandemFolio release / 安装发布包

Download from [GitHub Releases](https://github.com/BorrowLight-AI/TandemFolio/releases).
Choose the attached **`tandemfolio-X.Y.Z.zip`** (Windows or macOS) or
**`tandemfolio-X.Y.Z.tar.gz`** (macOS), plus **`SHA256SUMS`**. GitHub's automatically
generated “Source code” archives are source checkouts, not these installation bundles.
If no release assets are listed yet, the first release has not passed the required gate.

从 [GitHub Releases](https://github.com/BorrowLight-AI/TandemFolio/releases) 下载上述附件。
Windows 推荐 ZIP；macOS 可选 ZIP 或 tar.gz。不要选择自动生成的 “Source code”。
如果还没有发布附件，表示首个版本尚未完成发布门禁；可参考源码仓库的开发文档。

## Requirements / 环境要求

- Windows or macOS with a compatible Codex desktop host and a working Codex CLI that
  supports `codex plugin marketplace add` and `codex plugin add`.
- [Node.js](https://nodejs.org/) 22.12 or later on PATH. Use the native Windows/macOS
  installer for your CPU architecture. The same plugin files work with either runtime.
- No source build, Python, Git, or `npm install` is required for the downloaded bundle.

需要支持插件命令的 Codex CLI、支持 MCP Apps 的 Codex 桌面宿主，以及 PATH 中可用的
Node.js 22.12+。Windows 使用原生 PowerShell；不要把 WSL 的安装路径交给原生 Windows
Codex。安装或更新 Node.js 后，重新打开终端和 Codex。

The plugin bundle does not include Codex or Node.js. A marketplace is a distribution
mechanism; it does not change the host's account or regional availability requirements.

## macOS

1. Download the archive and `SHA256SUMS` into the same folder. Check the archive hash
   against the matching line in `SHA256SUMS`:

   ```bash
   shasum -a 256 tandemfolio-X.Y.Z.tar.gz
   ```

2. Extract it with Finder or `tar -xzf tandemfolio-X.Y.Z.tar.gz`. Move the extracted
   folder to a permanent location, for example `~/Documents/TandemFolio`.
   Rename the extracted top-level folder to `TandemFolio` so updates can reuse this path.
3. In Terminal, run:

   ```bash
   cd "$HOME/Documents/TandemFolio"
   bash ./install.sh
   ```

The script verifies the bundle, registers `tandemfolio-releases`, and installs
`tandemfolio@tandemfolio-releases`. It stops on any failure. For verification without
changing Codex, use `bash ./install.sh --check`.

下载后先校验，再解压并将目录固定为 `~/Documents/TandemFolio`，执行上面的安装命令。
脚本会校验包内文件、注册 marketplace 并安装插件；不需要管理员权限。

## Windows

1. Download the ZIP and `SHA256SUMS`. In PowerShell, compare this hash to the matching
   entry in `SHA256SUMS` (letter case does not matter):

   ```powershell
   Get-FileHash .\tandemfolio-X.Y.Z.zip -Algorithm SHA256
   ```

2. Use **Extract All** in Explorer. Move/rename the extracted top-level folder to a
   permanent location, for example `C:\Users\YOUR_NAME\Documents\TandemFolio`.
3. Open PowerShell in that folder and run:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
   ```

`-ExecutionPolicy Bypass` applies only to this new process; it does not change the
machine/user policy. No administrator access is required. To check the bundle without
changing Codex, append `-Check`.

Windows 下载 ZIP，校验后用资源管理器“全部解压”，将目录固定在 Documents 下。
在解压目录打开 PowerShell，执行上述命令。执行策略仅对这次进程生效，不修改系统设置。

## Start using it / 开始使用

Keep the installed folder in place: it is the registered marketplace source. Start a
**new Codex task** so the Skill and MCP server are discovered, then ask:

> Use TandemFolio to open and edit this document in the live editor.

> 用 TandemFolio 打开这份文档，在可视化编辑器中修改。

If you prefer manual installation after `node ./verify.mjs`, run from the extracted folder:

```text
codex plugin marketplace add .
codex plugin add tandemfolio@tandemfolio-releases
```

The release catalog is independent of an existing `personal` marketplace. It does not
edit that marketplace's files. Keep the editor visible while the agent changes a document.

## Updates and rollback / 更新与回退

1. Save documents and close tasks using the plugin.
2. Download and verify the new archive. Move the old installed folder aside as a backup.
3. Put the newly extracted folder at the **same absolute path** as the previous installation.
   Replace the whole folder; do not merge new files into old files.
4. Run `bash ./install.sh --update` on macOS, or
   `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Update` on Windows.
   This reinstalls the plugin using the already registered marketplace.
5. Start a new Codex task. To roll back, restore the backed-up folder to that same path,
   run the update command again, and start another new task.

更新前保存文档并关闭相关任务，保留旧目录备份。新包必须放回原来的绝对路径，整体替换，
不要覆盖合并文件。执行对应平台的 `--update` / `-Update` 命令，再开启新任务。
回退时恢复备份并重复安装。第一次安装不要使用 Update 参数。

## Troubleshooting / 排障

- **`node` or `codex` not found:** install the prerequisite and reopen the terminal.
  Check `node --version` and `codex --version`. If `codex plugin` is unavailable, update
  the CLI to a version supporting plugins. This package does not install those prerequisites.
- **Codex cannot start the MCP server:** check that the desktop app can find the same
  Node.js on PATH; fully restart Codex after installing Node.js. Terminal-only version
  managers may not expose Node.js to desktop apps.
- **Checksum failure:** download a fresh archive and extract into a new empty folder.
  Do not edit `checksums.json` to silence the error. An extra file inside the extracted
  bundle also fails verification (Finder/Explorer metadata is ignored); keep documents
  outside the installation folder.
- **Marketplace already registered:** if it refers to this same fixed installation path,
  rerun with `--update` / `-Update`. For an unrelated catalog with the same name, resolve
  that identity in Codex first; do not overwrite another catalog's configuration.
- **Old tools after updating:** finish the old task and create a new one.

## Maintainers / 维护者发布

The repository's `Release plugin` workflow runs on a `vX.Y.Z` tag (including SemVer
prerelease tags). Its tag must match the root `package.json` version and the plugin
manifest version before any `+codex...` local cache suffix. The release archive uses
that clean version without changing the developer's source manifest.

The workflow first fetches the configured read-only upstream history because the pinned community
commit is intentionally not part of the product repository's `origin` history. It then runs the
repository checks, approved source-current evidence gate, build, visual checks, and archive validation.
The source-current gate fingerprints cross-platform-stable source and packaging inputs; generated
editor bundles are validated after each platform build by visual, asset-budget, smoke, and archive
checksum checks rather than being treated as portable source bytes.
Windows and macOS then verify/extract the archive
and run the real packaged MCP smoke, including local document saves. Only after all jobs
pass does it create the GitHub Release with ZIP, tar.gz, and `SHA256SUMS`. SemVer prerelease
tags produce GitHub prereleases. Missing or stale evidence blocks both stable and
prerelease publication; the workflow never auto-approves a capture.

Locally, `npm run release:distribute` builds after the release gate and then validates
the gate again before creating archives in a fresh `out/releases/vX.Y.Z-*/` directory.
Archive creation requires macOS/Linux with `tar` and `zip`; the resulting package installs
on Windows/macOS. `npm run test:distribution` checks packaging/installer behavior on both
platforms independently of release readiness. It does not authorize publication.

Push a reviewed version commit and matching tag to trigger the workflow; this repository
does not push or create tags as part of the build. To rerun manually, select the same tag
in the workflow's **Run workflow** ref selector. A branch run skips publication. Existing
release assets are not overwritten by a rerun. Review a failed/partial release before retrying.

当前版本仍需按仓库规范补齐并审核发布证据。配置流水线不代表已有可下载的正式版本。
发布前同步版本号并提交，推送匹配的 `vX.Y.Z` 标签即可触发；预发布版本同样必须通过门禁。
README 使用固定的 Releases 页面地址，后续版本无需逐次修改下载链接。

param([switch]$Check, [switch]$Update)
$ErrorActionPreference = 'Stop'

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    throw 'Install Node.js 22.12+ from https://nodejs.org/, then reopen PowerShell.'
}
& node (Join-Path $PSScriptRoot 'verify.mjs')
if ($LASTEXITCODE -ne 0) { throw 'Bundle verification failed.' }
if ($Check) { exit 0 }

# Prefer the npm .cmd shim so npm's codex.ps1 does not require a policy change.
$codex = Get-Command codex.cmd -ErrorAction SilentlyContinue
if (-not $codex) { $codex = Get-Command codex -ErrorAction SilentlyContinue }
if (-not $codex) { throw 'Codex CLI is required. Install it, then run this script again.' }
& $codex.Source --version
if ($LASTEXITCODE -ne 0) { throw 'Codex CLI could not start.' }
if (-not $Update) {
    & $codex.Source plugin marketplace add $PSScriptRoot
    if ($LASTEXITCODE -ne 0) { throw 'Could not register the TandemFolio marketplace.' }
}
& $codex.Source plugin add 'tandemfolio@tandemfolio-releases'
if ($LASTEXITCODE -ne 0) { throw 'Could not install the TandemFolio plugin.' }
Write-Host 'TandemFolio installed. Keep this folder in place and start a new Codex task.'

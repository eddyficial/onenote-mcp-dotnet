# One-command setup: build the server, then point Codex, Claude Desktop, and
# Claude Code at the built exe. Mirrors the Bun repo's scripts/setup.ts.
#
# Usage: .\setup.ps1 [-DryRun] [-Client codex|claude-desktop|claude-code|claude|all]
param(
    [switch]$DryRun,
    [ValidateSet("codex", "claude-desktop", "claude-code", "claude", "all")]
    [string]$Client = "all"
)
$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
$exe = Join-Path $root "OneNoteMcp\bin\Release\net8.0-windows\OneNoteMcp.exe"

Write-Host "Building (dotnet build -c Release)..."
dotnet build (Join-Path $root "OneNoteMcp\OneNoteMcp.csproj") -c Release --nologo -v quiet
if ($LASTEXITCODE -ne 0) { throw "dotnet build failed ($LASTEXITCODE)." }
if (-not (Test-Path $exe)) { throw "Build output not found: $exe" }

function Wants([string]$name) {
    if ($Client -eq "all") { return $true }
    if ($Client -eq "claude") { return $name.StartsWith("claude-") }
    return $Client -eq $name
}

function Write-Config([string]$path, [string]$content) {
    $current = if (Test-Path $path) { Get-Content $path -Raw } else { "" }
    if ($current -eq $content) { Write-Host "Already configured: $path"; return }
    if ($DryRun) { Write-Host "Would update: $path"; return }
    New-Item -ItemType Directory -Force (Split-Path $path) | Out-Null
    [System.IO.File]::WriteAllText($path, $content)
    Write-Host "Configured: $path"
}

if (Wants "codex") {
    $codexHome = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $HOME ".codex" }
    $path = Join-Path $codexHome "config.toml"
    $source = if (Test-Path $path) { (Get-Content $path -Raw) -replace "`r`n", "`n" } else { "" }
    $cleaned = ([regex]::Replace($source, "(?m)^\[mcp_servers\.onenote\]\n[\s\S]*?(?=^\[|\z)", "")).TrimEnd()
    $exeToml = "'" + ($exe -replace "'", "''") + "'"
    $block = "[mcp_servers.onenote]`ncommand = $exeToml`nargs = []`nstartup_timeout_sec = 30`n"
    $merged = if ($cleaned) { "$cleaned`n`n$block" } else { $block }
    Write-Config $path $merged
}

if (Wants "claude-desktop") {
    if (-not $env:APPDATA) { throw "APPDATA is unavailable." }
    $path = Join-Path $env:APPDATA "Claude\claude_desktop_config.json"
    $parsed = if ((Test-Path $path) -and (Get-Content $path -Raw).Trim()) {
        Get-Content $path -Raw | ConvertFrom-Json
    } else { [pscustomobject]@{} }
    if (-not $parsed.PSObject.Properties["mcpServers"]) {
        $parsed | Add-Member -NotePropertyName mcpServers -NotePropertyValue ([pscustomobject]@{})
    }
    $entry = [pscustomobject]@{ command = $exe; args = @() }
    if ($parsed.mcpServers.PSObject.Properties["onenote"]) { $parsed.mcpServers.onenote = $entry }
    else { $parsed.mcpServers | Add-Member -NotePropertyName onenote -NotePropertyValue $entry }
    Write-Config $path (($parsed | ConvertTo-Json -Depth 10) + "`n")
}

if (Wants "claude-code") {
    $claude = Get-Command claude -ErrorAction SilentlyContinue
    if (-not $claude) { throw "Claude Code CLI ('claude') not found on PATH." }
    # cmd owns the 2>&1 so PowerShell 5.1 doesn't wrap stderr in ErrorRecords
    $out = (& cmd /c "`"$($claude.Source)`" mcp get onenote 2>&1") | Out-String
    $found = $LASTEXITCODE -eq 0
    $same = $found -and ($out.Replace("\", "/").ToLower().Contains($exe.Replace("\", "/").ToLower()))
    if ($same) { Write-Host "Already configured: Claude Code user MCP 'onenote'" }
    elseif ($found) { Write-Warning "Claude Code has a different 'onenote' MCP. Left unchanged." }
    elseif ($DryRun) { Write-Host "Would add: Claude Code user MCP 'onenote'" }
    else {
        & $claude.Source mcp add --scope user onenote -- $exe
        if ($LASTEXITCODE -ne 0) { throw "Claude Code setup failed ($LASTEXITCODE)." }
    }
}

Write-Host "Setup complete. Reload configured clients to discover the onenote_* tools."
exit 0

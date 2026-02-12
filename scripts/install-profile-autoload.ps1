param(
    [string]$ScriptPath
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
    $ScriptPath = Join-Path $repoRoot 'mail_mcp.ps1'
}
$ScriptPath = [System.IO.Path]::GetFullPath($ScriptPath)
if (-not (Test-Path -LiteralPath $ScriptPath)) {
    throw "Script introuvable: $ScriptPath"
}

$profilePath = $PROFILE.CurrentUserCurrentHost
$profileDir = Split-Path -Parent $profilePath

if (-not (Test-Path -LiteralPath $profileDir)) {
    New-Item -ItemType Directory -Path $profileDir | Out-Null
}
if (-not (Test-Path -LiteralPath $profilePath)) {
    New-Item -ItemType File -Path $profilePath | Out-Null
}

$block = @"
# Mail MCP auto-load
`$mailMcpPath = '$ScriptPath'
if (Test-Path -LiteralPath `$mailMcpPath) {
    . `$mailMcpPath
}
"@

$current = Get-Content -LiteralPath $profilePath -Raw
if ($null -eq $current) { $current = '' }

if ($current -notmatch '(?m)^# Mail MCP auto-load\s*$') {
    if ($current.Length -gt 0 -and -not $current.EndsWith("`r`n")) {
        $current += "`r`n"
    }
    $current += $block
    [System.IO.File]::WriteAllText($profilePath, $current, $enc)
}

Write-Host "Profile mis a jour:" $profilePath
Write-Host "Script charge automatiquement:" $ScriptPath
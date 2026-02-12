$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'mail_mcp.ps1')

try {
    $account = Invoke-MailMcp -Action account
    $folders = Invoke-MailMcp -Action folders
    $list = Invoke-MailMcp -Action list -Top 3

    [pscustomobject]@{
        account = $account
        foldersCount = @($folders).Count
        listCount = @($list).Count
    } | Format-List
}
finally {
    Disconnect-MailMcp
}

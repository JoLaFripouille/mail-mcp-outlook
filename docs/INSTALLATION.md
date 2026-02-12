# Installation

## 1. Prérequis

- Outlook Desktop configuré avec ton compte (ex: Gmail via Outlook)
- Une session Windows interactive (le COM Outlook ne fonctionne pas en mode service sans session)
- PowerShell 5.1 ou plus

## 2. Préparer PowerShell

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Optionnel si fichier téléchargé depuis internet:

```powershell
Unblock-File .\mail_mcp.ps1
```

## 3. Charger le script

```powershell
Set-Location <chemin-du-repo>
. .\mail_mcp.ps1
```

## 4. Vérifier que tout fonctionne

```powershell
Invoke-MailMcp -Action account
Invoke-MailMcp -Action folders
Invoke-MailMcp -Action list -Top 3
```

## 5. Chargement automatique à chaque ouverture PowerShell (optionnel)

```powershell
$profilePath = $PROFILE.CurrentUserCurrentHost
$profileDir = Split-Path -Parent $profilePath
if (-not (Test-Path -LiteralPath $profileDir)) { New-Item -ItemType Directory -Path $profileDir | Out-Null }
if (-not (Test-Path -LiteralPath $profilePath)) { New-Item -ItemType File -Path $profilePath | Out-Null }

Add-Content -LiteralPath $profilePath -Value "`n# Mail MCP auto-load`n`$mailMcpPath = 'C:\\Users\\Salon\\Bureau\\mail-mcp-outlook\\mail_mcp.ps1'`nif (Test-Path -LiteralPath `$mailMcpPath) {`n    . `$mailMcpPath`n}`n"
```

## 6. Test de fumée

```powershell
Invoke-MailMcp -Action search -Query "Google" -Top 5 -ScanLimit 500
Invoke-MailMcp -Action unread -Top 5 -ScanLimit 500
```

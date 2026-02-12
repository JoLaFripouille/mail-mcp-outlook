# Mail MCP Outlook (PowerShell)

MCP local pour piloter Outlook (lecture, recherche, envoi, pièces jointes, organisation, libellés) via COM, sans dépendances externes.

## Ce que fait le projet

- Connexion au profil Outlook par défaut
- Lecture et recherche de mails
- Envoi et réponses
- Gestion des pièces jointes (liste + téléchargement)
- Organisation (déplacement, archivage, suppression)
- Libellés via catégories Outlook (manuel + auto par sujet)

## Prérequis

- Windows + Outlook Desktop installé
- Au moins un compte mail configuré dans Outlook
- PowerShell 5.1+

## Installation rapide

1. Ouvrir PowerShell
2. Aller dans le dossier du repo
3. Autoriser l'exécution locale si nécessaire
4. Charger le script
5. Tester

```powershell
Set-Location .\mail-mcp-outlook
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
Unblock-File .\mail_mcp.ps1
. .\mail_mcp.ps1
Invoke-MailMcp -Action account
```

## Démarrage rapide

```powershell
Invoke-MailMcp -Action list -Top 5
Invoke-MailMcp -Action search -Query "facture" -Top 10 -ScanLimit 800
Invoke-MailMcp -Action attachments_list -Query "Avancé Adoption" -ScanLimit 2000
Invoke-MailMcp -Action attachments_download -Query "Avancé Adoption" -OnlyImagesAndPdf -OpenDir
```

## Documentation

- Installation: `docs/INSTALLATION.md`
- Dépannage: `docs/TROUBLESHOOTING.md`
- Référence des actions: `docs/ACTIONS_REFERENCE.md`
- Publication GitHub: `docs/GITHUB_PUBLISH.md`

## Sécurité

Actions sensibles disponibles: `move`, `archive`, `delete`.

Utilise toujours `-DryRun` avant un envoi réel ou une modification massive.

## Licence

MIT (voir `LICENSE`).

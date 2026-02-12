# Publication sur GitHub

## 1. Initialiser le dépôt local

```powershell
Set-Location <chemin-du-repo>
git init
git add .
git commit -m "Initial commit: Mail MCP Outlook"
```

## 2. Créer le repo distant GitHub

Option A: via interface web
- Créer un repo vide (sans README)

Option B: via GitHub CLI

```powershell
gh repo create mail-mcp-outlook --public --source . --remote origin --push
```

## 3. Lier et pousser (si repo créé via le web)

```powershell
git remote add origin https://github.com/<USER>/<REPO>.git
git branch -M main
git push -u origin main
```

## 4. Bonnes pratiques

- Ne pas versionner de données privées ni exports mails
- Vérifier `.gitignore`
- Utiliser `-DryRun` dans les exemples de doc

## 5. Releases (optionnel)

```powershell
git tag v1.0.0
git push origin v1.0.0
```

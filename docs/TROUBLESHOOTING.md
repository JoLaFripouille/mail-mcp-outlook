# Dépannage

## 1. Fenêtre "Choix d'un profil" qui bloque

Symptôme:
- Les commandes semblent bloquées ou timeout.

Cause:
- Outlook attend une sélection manuelle du profil.

Correction:
- Ouvrir Outlook une fois, définir le profil par défaut.
- Vérifier dans le registre `HKCU\Software\Microsoft\Office\16.0\Outlook` valeur `DefaultProfile`.

## 2. "Aucun profil Outlook par defaut trouve"

Symptôme:
- Erreur au démarrage de `Invoke-MailMcp`.

Correction:
- Ouvrir Outlook et terminer la configuration du compte.
- Fermer/réouvrir Outlook puis retester.

## 3. Erreur d'exécution de script

Symptôme:
- `running scripts is disabled`.

Correction:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

## 4. Le compte ne reçoit/pas les nouveaux mails

Symptôme:
- `search` ne trouve pas un mail visible sur webmail.

Cause probable:
- Synchronisation Outlook en retard.

Correction:
- Forcer un envoi/réception dans Outlook.
- Réduire ou augmenter `-ScanLimit` selon besoin.

## 5. "Dossier d archive introuvable"

Symptôme:
- `archive` échoue ou en dry-run indique qu'il manque un dossier.

Correction:
- Faire un essai réel (sans `-DryRun`) pour créer `Archive` automatiquement.
- Ou utiliser `move` vers un dossier explicite:

```powershell
Invoke-MailMcp -Action move -Query "newsletter" -FolderName "saved" -ScanLimit 800
```

## 6. Dossier ambigu pour `move`

Symptôme:
- Plusieurs dossiers portent le même nom.

Correction:
- Utiliser le chemin complet renvoyé par `Invoke-MailMcp -Action folders`.

## 7. Libellés et Gmail

Important:
- Les "libellés" gérés ici sont des catégories Outlook (`Categories`).
- Sur un compte Gmail IMAP, ce n'est pas forcément identique aux labels Gmail natifs côté web.

## 8. Performances lentes

Causes:
- Boîte volumineuse + `ScanLimit` élevé.

Bonnes pratiques:
- Commencer par `-ScanLimit 200` puis augmenter.
- Utiliser une `Query` précise.
- Utiliser `-Top` bas pour limiter le résultat.

## 9. Actions sensibles

Toujours tester d'abord avec:
- `-DryRun` sur `move`, `archive`, `delete`, `label_auto_subject`.

Exemple:

```powershell
Invoke-MailMcp -Action delete -Query "test" -Top 5 -ScanLimit 500 -DryRun
```

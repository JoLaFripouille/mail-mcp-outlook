# Référence des actions `Invoke-MailMcp`

## Signature générale

```powershell
Invoke-MailMcp -Action <action> [options]
```

Paramètres communs:
- `-Top` nombre de résultats (défaut 10)
- `-Query` texte de recherche
- `-ScanLimit` volume scanné (défaut 300)
- `-MessageId` ID mail exact (prioritaire sur `-Query`)
- `-DryRun` simulation sans modification

## Actions de base

### `account`
Retourne profil Outlook et adresse SMTP principale.

```powershell
Invoke-MailMcp -Action account
```

### `folders`
Liste les dossiers principaux et leurs compteurs.

```powershell
Invoke-MailMcp -Action folders
```

### `list`
Liste les derniers mails de la boîte de réception.

```powershell
Invoke-MailMcp -Action list -Top 20
```

### `unread`
Liste les mails non lus.

```powershell
Invoke-MailMcp -Action unread -Top 20 -ScanLimit 800
```

### `search`
Recherche plein texte (sujet, expéditeur, preview).

```powershell
Invoke-MailMcp -Action search -Query "facture" -Top 20 -ScanLimit 1200
```

### `read`
Lit un mail complet (métadonnées, corps, PJ, labels).

```powershell
Invoke-MailMcp -Action read -Query "Avancé Adoption" -ScanLimit 2000
```

Options:
- `-MaxBodyChars` taille max du corps renvoyé (défaut 6000)

## Envoi et réponse

### `send`
Envoie un mail.

```powershell
Invoke-MailMcp -Action send -To "destinataire@mail.com" -Subject "Test" -Body "Bonjour"
```

Options:
- `-Html` pour envoyer un corps HTML
- `-DryRun` pour sauvegarder en brouillon

### `reply`
Répond à un mail (brouillon par défaut).

```powershell
Invoke-MailMcp -Action reply -Query "Avancé Adoption" -ReplyBody "Merci, je regarde." -ScanLimit 2000
```

Options:
- `-ReplyAll`
- `-SendNow` pour envoyer directement
- `-Html`

## Lecture et téléchargement de pièces jointes

### `attachments_list`
Liste les pièces jointes d'un mail.

```powershell
Invoke-MailMcp -Action attachments_list -Query "Avancé Adoption" -ScanLimit 2000
```

### `attachments_download`
Télécharge les pièces jointes dans un dossier.

```powershell
Invoke-MailMcp -Action attachments_download -Query "Avancé Adoption" -OnlyImagesAndPdf -OpenDir
```

Options:
- `-OutputDir` dossier cible
- `-OnlyImagesAndPdf` filtre image/PDF
- `-Overwrite` écrase au lieu de suffixer
- `-OpenDir` ouvre le dossier après téléchargement

### `search_attachments`
Recherche les mails avec pièces jointes.

```powershell
Invoke-MailMcp -Action search_attachments -Query "devis" -AttachmentExtension pdf -Top 20 -ScanLimit 1500
```

Options:
- `-AttachmentName`
- `-AttachmentExtension` (ex: `pdf` ou `.pdf`)
- `-OnlyImagesAndPdf`

## Organisation des mails

### `move`
Déplace un mail vers un dossier.

```powershell
Invoke-MailMcp -Action move -Query "NordVPN" -FolderName "PROMO" -CreateFolder -DryRun
```

Options:
- `-FolderName` obligatoire
- `-CreateFolder` crée le dossier si absent
- `-DryRun` recommandé avant exécution réelle

### `archive`
Archive un mail vers dossier d'archive détecté.

```powershell
Invoke-MailMcp -Action archive -Query "newsletter" -DryRun
```

Notes:
- Si aucun dossier archive n'existe, le script peut créer `Archive` lors d'un run réel.

### `delete`
Supprime un mail.

```powershell
Invoke-MailMcp -Action delete -Query "test" -DryRun
```

## Libellés (catégories Outlook)

### `label_list`
Affiche les labels/catégories d'un mail.

```powershell
Invoke-MailMcp -Action label_list -Query "Avancé Adoption" -ScanLimit 2000
```

### `label_add`
Ajoute un ou plusieurs labels.

```powershell
Invoke-MailMcp -Action label_add -Query "adoption" -Label "Sujet/Adoption" -Top 20 -ScanLimit 1500
```

Options:
- `-Label` (unique)
- `-Labels` (liste)
- `-Top` pour appliquer sur plusieurs mails trouvés
- `-DryRun`

### `label_remove`
Retire un ou plusieurs labels.

```powershell
Invoke-MailMcp -Action label_remove -Query "adoption" -Label "Sujet/Adoption" -Top 20 -ScanLimit 1500
```

### `label_clear`
Supprime tous les labels d'un mail (ou lot de mails).

```powershell
Invoke-MailMcp -Action label_clear -Query "newsletter" -Top 10 -ScanLimit 800 -DryRun
```

### `label_auto_subject`
Crée/applique automatiquement un label basé sur le sujet (`Sujet/<topic>`).

```powershell
Invoke-MailMcp -Action label_auto_subject -Top 100 -ScanLimit 300 -LabelPrefix "Sujet" -UnreadOnly -DryRun
```

Options:
- `-LabelPrefix` préfixe du label
- `-UnreadOnly`
- `-DryRun` recommandé

## Fin de session

### `Disconnect-MailMcp`
Libère les objets COM Outlook.

```powershell
Disconnect-MailMcp
```

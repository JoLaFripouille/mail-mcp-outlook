Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:MailMcpContext = $null

function Get-MailMcpDefaultProfile {
    [CmdletBinding()]
    param()

    $paths = @(
        'HKCU:\Software\Microsoft\Office\16.0\Outlook',
        'HKCU:\Software\Microsoft\Office\15.0\Outlook',
        'HKCU:\Software\Microsoft\Office\14.0\Outlook'
    )

    foreach ($path in $paths) {
        if (-not (Test-Path $path)) {
            continue
        }

        try {
            $item = Get-ItemProperty -Path $path -Name 'DefaultProfile' -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($item.DefaultProfile)) {
                return $item.DefaultProfile
            }
        }
        catch {
        }
    }

    return $null
}

function New-MailMcpContext {
    [CmdletBinding()]
    param()

    $outlook = New-Object -ComObject Outlook.Application
    $session = $outlook.GetNamespace('MAPI')

    $profileName = Get-MailMcpDefaultProfile
    if ([string]::IsNullOrWhiteSpace($profileName)) {
        throw 'Aucun profil Outlook par defaut trouve. Ouvre Outlook une fois et choisis un profil par defaut.'
    }

    # Silent logon on the default profile to avoid profile picker UI.
    $session.Logon($profileName, [Type]::Missing, $false, $false)

    $inbox = $session.GetDefaultFolder(6)   # olFolderInbox
    $drafts = $session.GetDefaultFolder(16) # olFolderDrafts
    $sent = $session.GetDefaultFolder(5)    # olFolderSentMail

    $primaryAccount = $null
    if ($session.Accounts.Count -gt 0) {
        $primaryAccount = $session.Accounts.Item(1)
    }

    [pscustomobject]@{
        Outlook     = $outlook
        Session     = $session
        Inbox       = $inbox
        Drafts      = $drafts
        Sent        = $sent
        ProfileName = $profileName
        AccountSmtp = if ($primaryAccount) { $primaryAccount.SmtpAddress } else { $null }
    }
}

function Get-MailMcpMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Folder,

        [int]$Top = 10
    )

    if ($Top -lt 1) {
        throw 'Top doit etre >= 1.'
    }

    $items = $Folder.Items
    $items.Sort('[ReceivedTime]', $true)

    $result = New-Object System.Collections.Generic.List[object]

    for ($i = 1; $i -le $items.Count -and $result.Count -lt $Top; $i++) {
        $mail = $items.Item($i)
        if ($null -eq $mail -or $mail.Class -ne 43) {
            continue
        }

        $preview = $null
        if ($mail.Body) {
            $cleanBody = ($mail.Body -replace '\s+', ' ').Trim()
            if ($cleanBody.Length -gt 160) {
                $preview = $cleanBody.Substring(0, 160)
            }
            else {
                $preview = $cleanBody
            }
        }

        $result.Add([pscustomobject]@{
            id           = $mail.EntryID
            receivedTime = $mail.ReceivedTime
            fromName     = $mail.SenderName
            fromAddress  = $mail.SenderEmailAddress
            subject      = $mail.Subject
            isUnread     = [bool]$mail.UnRead
            preview      = $preview
        })
    }

    return $result
}

function Find-MailMcpMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Query,

        [int]$Top = 10,
        [int]$ScanLimit = 300
    )

    if ([string]::IsNullOrWhiteSpace($Query)) {
        throw 'Query est requis.'
    }
    if ($Top -lt 1) {
        throw 'Top doit etre >= 1.'
    }
    if ($ScanLimit -lt 1) {
        throw 'ScanLimit doit etre >= 1.'
    }

    $messages = Get-MailMcpMessages -Folder $script:MailMcpContext.Inbox -Top $ScanLimit
    $needle = $Query.Trim().ToLowerInvariant()

    $filtered = $messages | Where-Object {
        $subject = if ($_.subject) { $_.subject.ToLowerInvariant() } else { '' }
        $fromName = if ($_.fromName) { $_.fromName.ToLowerInvariant() } else { '' }
        $fromAddress = if ($_.fromAddress) { $_.fromAddress.ToLowerInvariant() } else { '' }
        $preview = if ($_.preview) { $_.preview.ToLowerInvariant() } else { '' }

        $subject.Contains($needle) -or
        $fromName.Contains($needle) -or
        $fromAddress.Contains($needle) -or
        $preview.Contains($needle)
    }

    return $filtered | Select-Object -First $Top
}

function Resolve-MailMcpMessage {
    [CmdletBinding()]
    param(
        [string]$MessageId,
        [string]$Query,
        [int]$ScanLimit = 300
    )

    $resolvedId = $null

    if (-not [string]::IsNullOrWhiteSpace($MessageId)) {
        $resolvedId = $MessageId
    }
    elseif (-not [string]::IsNullOrWhiteSpace($Query)) {
        $match = Find-MailMcpMessages -Query $Query -Top 1 -ScanLimit $ScanLimit
        if (-not $match) {
            throw "Aucun mail trouve pour la recherche: $Query"
        }
        $resolvedId = $match[0].id
    }
    else {
        throw 'Fournis MessageId ou Query.'
    }

    try {
        $mail = $script:MailMcpContext.Session.GetItemFromID($resolvedId)
    }
    catch {
        throw "Impossible de charger le mail via MessageId: $resolvedId"
    }

    if ($null -eq $mail -or $mail.Class -ne 43) {
        throw 'L element trouve n est pas un mail Outlook valide.'
    }

    return $mail
}

function Resolve-MailMcpMessages {
    [CmdletBinding()]
    param(
        [string]$MessageId,
        [string]$Query,
        [int]$ScanLimit = 300,
        [int]$Top = 10
    )

    if ($Top -lt 1) {
        throw 'Top doit etre >= 1.'
    }

    if (-not [string]::IsNullOrWhiteSpace($MessageId)) {
        $single = Resolve-MailMcpMessage -MessageId $MessageId -Query $null -ScanLimit $ScanLimit
        return @($single)
    }

    if ([string]::IsNullOrWhiteSpace($Query)) {
        throw 'Fournis MessageId ou Query.'
    }

    $matches = Find-MailMcpMessages -Query $Query -Top $Top -ScanLimit $ScanLimit
    if (-not $matches) {
        throw "Aucun mail trouve pour la recherche: $Query"
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($match in $matches) {
        try {
            $mail = $script:MailMcpContext.Session.GetItemFromID($match.id)
            if ($null -ne $mail -and $mail.Class -eq 43) {
                $result.Add($mail)
            }
        }
        catch {
        }
    }

    if ($result.Count -eq 0) {
        throw 'Aucun mail valide n a pu etre charge.'
    }

    return $result
}

function Get-MailMcpAllFolders {
    [CmdletBinding()]
    param(
        [int]$MaxDepth = 10
    )

    if ($MaxDepth -lt 0) {
        throw 'MaxDepth doit etre >= 0.'
    }

    $root = $script:MailMcpContext.Inbox.Parent
    $result = New-Object System.Collections.Generic.List[object]
    $queue = New-Object System.Collections.Generic.List[object]
    $queue.Add([pscustomobject]@{ Folder = $root; Depth = 0 })

    for ($i = 0; $i -lt $queue.Count; $i++) {
        $node = $queue[$i]
        $folder = $node.Folder
        $depth = [int]$node.Depth

        if ($depth -gt 0) {
            $result.Add([pscustomobject]@{
                name   = $folder.Name
                path   = $folder.FolderPath
                depth  = $depth
                folder = $folder
            })
        }

        if ($depth -ge $MaxDepth) {
            continue
        }

        try {
            foreach ($child in $folder.Folders) {
                $queue.Add([pscustomobject]@{ Folder = $child; Depth = ($depth + 1) })
            }
        }
        catch {
        }
    }

    return $result
}

function Resolve-MailMcpFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderName,

        [switch]$CreateIfMissing
    )

    if ([string]::IsNullOrWhiteSpace($FolderName)) {
        throw 'FolderName est requis.'
    }

    $candidate = $FolderName.Trim()
    $key = $candidate.ToLowerInvariant()

    $defaultMap = @{
        'inbox' = 6
        'boite de reception' = 6
        'boite de réception' = 6
        'sent' = 5
        'sent items' = 5
        'messages envoyes' = 5
        'messages envoyés' = 5
        'drafts' = 16
        'brouillons' = 16
        'outbox' = 4
        'boite d envoi' = 4
        'boite d''envoi' = 4
        'deleted' = 3
        'deleted items' = 3
        'corbeille' = 3
        'junk' = 23
        'spam' = 23
    }

    if ($defaultMap.ContainsKey($key)) {
        return $script:MailMcpContext.Session.GetDefaultFolder([int]$defaultMap[$key])
    }

    $all = Get-MailMcpAllFolders -MaxDepth 12

    $exactPath = @($all | Where-Object {
        [string]::Equals([string]$_.path, $candidate, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($exactPath.Count -eq 1) {
        return $exactPath[0].folder
    }
    if ($exactPath.Count -gt 1) {
        $choices = ($exactPath | Select-Object -First 5 | ForEach-Object { $_.path }) -join '; '
        throw "Dossier ambigu (chemin exact): $choices"
    }

    $exactName = @($all | Where-Object {
        [string]::Equals([string]$_.name, $candidate, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($exactName.Count -eq 1) {
        return $exactName[0].folder
    }
    if ($exactName.Count -gt 1) {
        $choices = ($exactName | Select-Object -First 5 | ForEach-Object { $_.path }) -join '; '
        throw "Nom de dossier ambigu, utilise le chemin complet. Exemples: $choices"
    }

    $needle = $candidate.ToLowerInvariant()
    $partial = @($all | Where-Object {
        ([string]$_.path).ToLowerInvariant().Contains($needle) -or ([string]$_.name).ToLowerInvariant().Contains($needle)
    })
    if ($partial.Count -eq 1) {
        return $partial[0].folder
    }
    if ($partial.Count -gt 1) {
        $choices = ($partial | Select-Object -First 5 | ForEach-Object { $_.path }) -join '; '
        throw "Plusieurs dossiers correspondent, precise davantage. Exemples: $choices"
    }

    if ($CreateIfMissing) {
        $parent = $script:MailMcpContext.Inbox.Parent
        try {
            return $parent.Folders.Add($candidate)
        }
        catch {
            throw "Impossible de creer le dossier: $candidate"
        }
    }

    throw "Dossier introuvable: $candidate"
}

function Resolve-MailMcpArchiveFolder {
    [CmdletBinding()]
    param(
        [switch]$CreateIfMissing
    )

    $all = Get-MailMcpAllFolders -MaxDepth 12
    $scored = New-Object System.Collections.Generic.List[object]

    foreach ($entry in $all) {
        $name = ([string]$entry.name).ToLowerInvariant()
        $path = ([string]$entry.path).ToLowerInvariant()
        $score = 0

        if ($name -eq 'tous les messages' -or $name -eq 'all mail') {
            $score = 100
        }
        elseif (($path.Contains('[gmail]')) -and ($path.Contains('tous les messages') -or $path.Contains('all mail'))) {
            $score = 95
        }
        elseif ($name.Contains('archive') -or $name.Contains('archiv')) {
            $score = 80
        }
        elseif ($path.Contains('archive') -or $path.Contains('archiv')) {
            $score = 70
        }

        if ($score -gt 0) {
            $scored.Add([pscustomobject]@{
                score  = $score
                path   = $entry.path
                folder = $entry.folder
            })
        }
    }

    if ($scored.Count -gt 0) {
        return ($scored | Sort-Object score -Descending | Select-Object -First 1).folder
    }

    if ($CreateIfMissing) {
        try {
            $root = $script:MailMcpContext.Inbox.Parent
            return $root.Folders.Add('Archive')
        }
        catch {
        }
    }

    throw 'Dossier d archive introuvable. Utilise -Action folders puis -Action move avec -FolderName.'
}

function Move-MailMcpMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem,

        [Parameter(Mandatory)]
        $TargetFolder,

        [switch]$DryRun
    )

    $fromPath = $null
    try {
        $fromPath = $MailItem.Parent.FolderPath
    }
    catch {
    }

    if ($DryRun) {
        return [pscustomobject]@{
            status   = 'dryrun_move'
            id       = $MailItem.EntryID
            subject  = $MailItem.Subject
            fromPath = $fromPath
            toPath   = $TargetFolder.FolderPath
        }
    }

    $moved = $MailItem.Move($TargetFolder)
    return [pscustomobject]@{
        status   = 'moved'
        id       = $moved.EntryID
        subject  = $moved.Subject
        fromPath = $fromPath
        toPath   = $TargetFolder.FolderPath
    }
}

function Move-MailMcpMessageToArchive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem,

        [switch]$DryRun
    )

    try {
        $archiveFolder = Resolve-MailMcpArchiveFolder -CreateIfMissing:(-not $DryRun)
    }
    catch {
        if ($DryRun) {
            $fromPath = $null
            try {
                $fromPath = $MailItem.Parent.FolderPath
            }
            catch {
            }

            return [pscustomobject]@{
                status   = 'dryrun_archive'
                id       = $MailItem.EntryID
                subject  = $MailItem.Subject
                fromPath = $fromPath
                toPath   = 'Archive (a creer)'
                note     = 'Aucun dossier archive detecte, un dossier Archive sera cree lors d un envoi reel.'
            }
        }

        throw
    }

    $result = Move-MailMcpMessage -MailItem $MailItem -TargetFolder $archiveFolder -DryRun:$DryRun
    $result | Add-Member -NotePropertyName archivePath -NotePropertyValue $archiveFolder.FolderPath -Force
    return $result
}

function Remove-MailMcpMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem,

        [switch]$DryRun
    )

    $fromPath = $null
    try {
        $fromPath = $MailItem.Parent.FolderPath
    }
    catch {
    }

    if ($DryRun) {
        return [pscustomobject]@{
            status   = 'dryrun_delete'
            id       = $MailItem.EntryID
            subject  = $MailItem.Subject
            fromPath = $fromPath
        }
    }

    $id = $MailItem.EntryID
    $subject = $MailItem.Subject
    $MailItem.Delete()

    return [pscustomobject]@{
        status   = 'deleted'
        id       = $id
        subject  = $subject
        fromPath = $fromPath
    }
}

function Normalize-MailMcpLabel {
    [CmdletBinding()]
    param(
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Label)) {
        return $null
    }

    $value = $Label.Trim()
    $value = $value -replace '[,;]+', ' '
    $value = $value -replace '\s+', ' '
    $value = $value.Trim()

    if ($value.Length -gt 60) {
        $value = $value.Substring(0, 60).Trim()
    }

    if ([string]::IsNullOrWhiteSpace($value)) {
        return $null
    }

    return $value
}

function Resolve-MailMcpInputLabels {
    [CmdletBinding()]
    param(
        [string]$Label,
        [string[]]$Labels
    )

    $raw = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($Label)) {
        $raw.Add($Label)
    }

    foreach ($entry in @($Labels)) {
        if ([string]::IsNullOrWhiteSpace($entry)) {
            continue
        }

        foreach ($part in ($entry -split ',')) {
            $raw.Add($part)
        }
    }

    $result = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($candidate in $raw) {
        $normalized = Normalize-MailMcpLabel -Label $candidate
        if ([string]::IsNullOrWhiteSpace($normalized)) {
            continue
        }

        $k = $normalized.ToLowerInvariant()
        if (-not $seen.ContainsKey($k)) {
            $seen[$k] = $true
            $result.Add($normalized)
        }
    }

    return $result.ToArray()
}

function Get-MailMcpLabelsFromMailItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem
    )

    return @(Resolve-MailMcpInputLabels -Label ([string]$MailItem.Categories) -Labels @())
}

function Add-MailMcpLabelsToMailItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem,

        [Parameter(Mandatory)]
        [string[]]$Labels,

        [switch]$DryRun
    )

    $toAdd = @(Resolve-MailMcpInputLabels -Label $null -Labels $Labels)
    if ($toAdd.Count -eq 0) {
        throw 'Aucun label valide a ajouter.'
    }

    $current = @(Get-MailMcpLabelsFromMailItem -MailItem $MailItem)
    $merged = New-Object System.Collections.Generic.List[string]
    $seen = @{}

    foreach ($label in $current) {
        $key = $label.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $merged.Add($label)
        }
    }

    $changed = $false
    foreach ($label in $toAdd) {
        $key = $label.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $merged.Add($label)
            $changed = $true
        }
    }

    if ($changed -and -not $DryRun) {
        $MailItem.Categories = ($merged -join ', ')
        $MailItem.Save()
    }

    return [pscustomobject]@{
        id      = $MailItem.EntryID
        subject = $MailItem.Subject
        status  = if ($changed) { if ($DryRun) { 'dryrun_label_add' } else { 'label_added' } } else { 'no_change' }
        labels  = $merged
    }
}

function Remove-MailMcpLabelsFromMailItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem,

        [Parameter(Mandatory)]
        [string[]]$Labels,

        [switch]$DryRun
    )

    $toRemove = @(Resolve-MailMcpInputLabels -Label $null -Labels $Labels)
    if ($toRemove.Count -eq 0) {
        throw 'Aucun label valide a retirer.'
    }

    $removeMap = @{}
    foreach ($label in $toRemove) {
        $removeMap[$label.ToLowerInvariant()] = $true
    }

    $current = @(Get-MailMcpLabelsFromMailItem -MailItem $MailItem)
    $remaining = New-Object System.Collections.Generic.List[string]
    $changed = $false

    foreach ($label in $current) {
        $key = $label.ToLowerInvariant()
        if ($removeMap.ContainsKey($key)) {
            $changed = $true
            continue
        }
        $remaining.Add($label)
    }

    if ($changed -and -not $DryRun) {
        $MailItem.Categories = ($remaining -join ', ')
        $MailItem.Save()
    }

    return [pscustomobject]@{
        id      = $MailItem.EntryID
        subject = $MailItem.Subject
        status  = if ($changed) { if ($DryRun) { 'dryrun_label_remove' } else { 'label_removed' } } else { 'no_change' }
        labels  = $remaining
    }
}

function Clear-MailMcpLabelsFromMailItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem,

        [switch]$DryRun
    )

    $current = @(Get-MailMcpLabelsFromMailItem -MailItem $MailItem)
    $changed = $current.Count -gt 0

    if ($changed -and -not $DryRun) {
        $MailItem.Categories = ''
        $MailItem.Save()
    }

    return [pscustomobject]@{
        id      = $MailItem.EntryID
        subject = $MailItem.Subject
        status  = if ($changed) { if ($DryRun) { 'dryrun_label_clear' } else { 'label_cleared' } } else { 'no_change' }
        labels  = @()
    }
}

function Get-MailMcpTopicFromSubject {
    [CmdletBinding()]
    param(
        [string]$Subject
    )

    $topic = if ([string]::IsNullOrWhiteSpace($Subject)) { 'SansSujet' } else { [string]$Subject }

    $prev = $null
    while ($topic -ne $prev) {
        $prev = $topic
        $topic = $topic -replace '^(?i)\s*(re|fw|fwd|tr|rv)\s*:\s*', ''
        $topic = $topic -replace '^\s*\[[^\]]+\]\s*', ''
        $topic = $topic.Trim()
    }

    $parts = [regex]::Split($topic, '\s*[\|\-:]\s*', 2)
    if ($parts.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($parts[0])) {
        $topic = $parts[0].Trim()
    }

    $topic = $topic -replace '\s+', ' '
    $topic = $topic.Trim()

    if ([string]::IsNullOrWhiteSpace($topic)) {
        $topic = 'SansSujet'
    }

    if ($topic.Length -gt 45) {
        $topic = $topic.Substring(0, 45).Trim()
    }

    return $topic
}

function Invoke-MailMcpAutoSubjectLabeling {
    [CmdletBinding()]
    param(
        [string]$Query,
        [int]$Top = 10,
        [int]$ScanLimit = 300,
        [string]$LabelPrefix = 'Sujet',
        [switch]$UnreadOnly,
        [switch]$DryRun
    )

    if ($Top -lt 1) {
        throw 'Top doit etre >= 1.'
    }
    if ($ScanLimit -lt 1) {
        throw 'ScanLimit doit etre >= 1.'
    }

    $prefix = Normalize-MailMcpLabel -Label $LabelPrefix
    if ([string]::IsNullOrWhiteSpace($prefix)) {
        $prefix = 'Sujet'
    }

    $candidates = if ([string]::IsNullOrWhiteSpace($Query)) {
        Get-MailMcpMessages -Folder $script:MailMcpContext.Inbox -Top $ScanLimit
    }
    else {
        Find-MailMcpMessages -Query $Query -Top $ScanLimit -ScanLimit $ScanLimit
    }

    if ($UnreadOnly) {
        $candidates = $candidates | Where-Object { $_.isUnread }
    }

    $candidates = $candidates | Select-Object -First $Top

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($candidate in @($candidates)) {
        try {
            $mail = $script:MailMcpContext.Session.GetItemFromID($candidate.id)
            if ($null -eq $mail -or $mail.Class -ne 43) {
                continue
            }

            $topic = Get-MailMcpTopicFromSubject -Subject $mail.Subject
            $label = Normalize-MailMcpLabel -Label ("$prefix/$topic")
            if ([string]::IsNullOrWhiteSpace($label)) {
                continue
            }

            $applied = Add-MailMcpLabelsToMailItem -MailItem $mail -Labels @($label) -DryRun:$DryRun
            $result.Add([pscustomobject]@{
                id      = $applied.id
                subject = $applied.subject
                topic   = $topic
                label   = $label
                status  = $applied.status
                labels  = $applied.labels
            })
        }
        catch {
        }
    }

    return $result
}

function Search-MailMcpAttachments {
    [CmdletBinding()]
    param(
        [string]$Query,
        [int]$Top = 10,
        [int]$ScanLimit = 300,
        [string]$AttachmentName,
        [string]$AttachmentExtension,
        [switch]$OnlyImagesAndPdf
    )

    if ($Top -lt 1) {
        throw 'Top doit etre >= 1.'
    }
    if ($ScanLimit -lt 1) {
        throw 'ScanLimit doit etre >= 1.'
    }

    $extension = $null
    if (-not [string]::IsNullOrWhiteSpace($AttachmentExtension)) {
        $extension = $AttachmentExtension.Trim().ToLowerInvariant()
        if (-not $extension.StartsWith('.')) {
            $extension = ".${extension}"
        }
    }

    $nameNeedle = if ([string]::IsNullOrWhiteSpace($AttachmentName)) { $null } else { $AttachmentName.Trim().ToLowerInvariant() }

    $candidates = if ([string]::IsNullOrWhiteSpace($Query)) {
        Get-MailMcpMessages -Folder $script:MailMcpContext.Inbox -Top $ScanLimit
    }
    else {
        Find-MailMcpMessages -Query $Query -Top $ScanLimit -ScanLimit $ScanLimit
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($candidate in @($candidates)) {
        if ($result.Count -ge $Top) {
            break
        }

        try {
            $mail = $script:MailMcpContext.Session.GetItemFromID($candidate.id)
            if ($null -eq $mail -or $mail.Class -ne 43) {
                continue
            }

            if ($mail.Attachments.Count -eq 0) {
                continue
            }

            $attachments = Get-MailMcpAttachmentsInfo -MailItem $mail

            if ($OnlyImagesAndPdf) {
                $attachments = @($attachments | Where-Object { $_.kind -in @('image', 'pdf') })
            }

            if (-not [string]::IsNullOrWhiteSpace($nameNeedle)) {
                $attachments = @($attachments | Where-Object {
                    ([string]$_.fileName).ToLowerInvariant().Contains($nameNeedle)
                })
            }

            if (-not [string]::IsNullOrWhiteSpace($extension)) {
                $attachments = @($attachments | Where-Object {
                    ([System.IO.Path]::GetExtension([string]$_.fileName).ToLowerInvariant() -eq $extension)
                })
            }

            if (@($attachments).Count -eq 0) {
                continue
            }

            $result.Add([pscustomobject]@{
                id              = $mail.EntryID
                subject         = $mail.Subject
                fromName        = $mail.SenderName
                fromAddress     = $mail.SenderEmailAddress
                receivedTime    = $mail.ReceivedTime
                attachmentCount = @($attachments).Count
                attachments     = $attachments
            })
        }
        catch {
        }
    }

    return $result
}

function Get-MailMcpAttachmentKind {
    [CmdletBinding()]
    param(
        [string]$FileName
    )

    $ext = [System.IO.Path]::GetExtension([string]$FileName).ToLowerInvariant()

    if (@('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tif', '.tiff').Contains($ext)) {
        return 'image'
    }
    if ($ext -eq '.pdf') {
        return 'pdf'
    }

    return 'other'
}

function Get-MailMcpAttachmentsInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem
    )

    $result = New-Object System.Collections.Generic.List[object]

    for ($i = 1; $i -le $MailItem.Attachments.Count; $i++) {
        $att = $MailItem.Attachments.Item($i)

        $result.Add([pscustomobject]@{
            index = $i
            fileName = $att.FileName
            sizeBytes = $att.Size
            type = $att.Type
            kind = Get-MailMcpAttachmentKind -FileName $att.FileName
        })
    }

    return $result
}

function New-MailMcpSafeFileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $safe = $Name
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace($ch, '_')
    }

    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = 'attachment'
    }

    return $safe
}

function Get-MailMcpDownloadFolder {
    [CmdletBinding()]
    param(
        [string]$OutputDir,
        [string]$Subject
    )

    if (-not [string]::IsNullOrWhiteSpace($OutputDir)) {
        if (-not (Test-Path -LiteralPath $OutputDir)) {
            New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
        }
        return (Resolve-Path -LiteralPath $OutputDir).Path
    }

    $baseDir = Join-Path $env:TEMP 'MailMcpDownloads'
    if (-not (Test-Path -LiteralPath $baseDir)) {
        New-Item -Path $baseDir -ItemType Directory -Force | Out-Null
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $subjectSafe = New-MailMcpSafeFileName -Name ([string]$Subject)
    if ($subjectSafe.Length -gt 60) {
        $subjectSafe = $subjectSafe.Substring(0, 60)
    }

    $folderName = "$stamp-$subjectSafe"
    $final = Join-Path $baseDir $folderName
    New-Item -Path $final -ItemType Directory -Force | Out-Null

    return $final
}

function Save-MailMcpAttachments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem,

        [string]$OutputDir,
        [switch]$OnlyImagesAndPdf,
        [switch]$Overwrite
    )

    $targetDir = Get-MailMcpDownloadFolder -OutputDir $OutputDir -Subject $MailItem.Subject
    $files = New-Object System.Collections.Generic.List[object]

    for ($i = 1; $i -le $MailItem.Attachments.Count; $i++) {
        $att = $MailItem.Attachments.Item($i)
        $kind = Get-MailMcpAttachmentKind -FileName $att.FileName

        if ($OnlyImagesAndPdf -and @('image', 'pdf') -notcontains $kind) {
            continue
        }

        $safeName = New-MailMcpSafeFileName -Name ([string]$att.FileName)
        $dest = Join-Path $targetDir $safeName

        if ((Test-Path -LiteralPath $dest) -and -not $Overwrite) {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($safeName)
            $ext = [System.IO.Path]::GetExtension($safeName)
            $idx = 2
            do {
                $candidate = "{0}_{1}{2}" -f $base, $idx, $ext
                $dest = Join-Path $targetDir $candidate
                $idx++
            } while (Test-Path -LiteralPath $dest)
        }

        $att.SaveAsFile($dest)

        $files.Add([pscustomobject]@{
            index = $i
            fileName = [System.IO.Path]::GetFileName($dest)
            fullPath = $dest
            sizeBytes = $att.Size
            kind = $kind
        })
    }

    return [pscustomobject]@{
        outputDir = $targetDir
        downloadedCount = $files.Count
        files = $files
    }
}

function Get-MailMcpFolders {
    [CmdletBinding()]
    param()

    $map = @(
        @{ Id = 6; Name = 'Inbox' },
        @{ Id = 5; Name = 'Sent' },
        @{ Id = 16; Name = 'Drafts' },
        @{ Id = 4; Name = 'Outbox' },
        @{ Id = 3; Name = 'Deleted' },
        @{ Id = 2; Name = 'DeletedItemsLegacy' },
        @{ Id = 23; Name = 'Junk' }
    )

    $result = New-Object System.Collections.Generic.List[object]

    foreach ($entry in $map) {
        try {
            $folder = $script:MailMcpContext.Session.GetDefaultFolder($entry.Id)
            if ($null -ne $folder) {
                $result.Add([pscustomobject]@{
                    name       = $entry.Name
                    path       = $folder.FolderPath
                    itemCount  = $folder.Items.Count
                    unread     = $folder.UnReadItemCount
                })
            }
        }
        catch {
        }
    }

    return $result
}

function Get-MailMcpMessageDetails {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem,

        [int]$MaxBodyChars = 6000
    )

    if ($MaxBodyChars -lt 1) {
        throw 'MaxBodyChars doit etre >= 1.'
    }

    $body = [string]$MailItem.Body
    $wasTruncated = $false
    if ($body.Length -gt $MaxBodyChars) {
        $body = $body.Substring(0, $MaxBodyChars)
        $wasTruncated = $true
    }

    $attachments = Get-MailMcpAttachmentsInfo -MailItem $MailItem
    $labels = @(Get-MailMcpLabelsFromMailItem -MailItem $MailItem)

    return [pscustomobject]@{
        id              = $MailItem.EntryID
        subject         = $MailItem.Subject
        fromName        = $MailItem.SenderName
        fromAddress     = $MailItem.SenderEmailAddress
        to              = $MailItem.To
        cc              = $MailItem.CC
        receivedTime    = $MailItem.ReceivedTime
        isUnread        = [bool]$MailItem.UnRead
        labels          = $labels
        attachmentCount = $attachments.Count
        bodyText        = $body
        bodyTruncated   = $wasTruncated
        attachments     = $attachments
    }
}

function Set-MailMcpReadState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem,

        [Parameter(Mandatory)]
        [bool]$IsRead
    )

    $MailItem.UnRead = -not $IsRead
    $MailItem.Save()

    return [pscustomobject]@{
        id       = $MailItem.EntryID
        subject  = $MailItem.Subject
        isUnread = [bool]$MailItem.UnRead
        status   = if ($IsRead) { 'marked_read' } else { 'marked_unread' }
    }
}

function New-MailMcpReply {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailItem,

        [Parameter(Mandatory)]
        [string]$ReplyBody,

        [switch]$ReplyAll,
        [switch]$Html,
        [switch]$SendNow
    )

    if ([string]::IsNullOrWhiteSpace($ReplyBody)) {
        throw 'ReplyBody est requis pour reply.'
    }

    $reply = if ($ReplyAll) { $MailItem.ReplyAll() } else { $MailItem.Reply() }

    if ($Html) {
        $reply.HTMLBody = "$ReplyBody<br><br>" + $reply.HTMLBody
    }
    else {
        $reply.Body = "$ReplyBody`r`n`r`n" + $reply.Body
    }

    if ($SendNow) {
        $reply.Send()
        return [pscustomobject]@{
            status  = 'sent'
            to      = $reply.To
            cc      = $reply.CC
            subject = $reply.Subject
        }
    }

    $reply.Save()
    return [pscustomobject]@{
        status  = 'draft_saved'
        to      = $reply.To
        cc      = $reply.CC
        subject = $reply.Subject
        id      = $reply.EntryID
    }
}

function Send-MailMcpMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$To,

        [Parameter(Mandatory)]
        [string]$Subject,

        [Parameter(Mandatory)]
        [string]$Body,

        [switch]$Html,
        [switch]$DryRun
    )

    if (-not $script:MailMcpContext) {
        $script:MailMcpContext = New-MailMcpContext
    }

    $mail = $script:MailMcpContext.Outlook.CreateItem(0) # olMailItem
    $mail.To = $To
    $mail.Subject = $Subject

    if ($Html) {
        $mail.HTMLBody = $Body
    }
    else {
        $mail.Body = $Body
    }

    if ($DryRun) {
        $mail.Save()
        return [pscustomobject]@{
            status  = 'draft_saved'
            to      = $To
            subject = $Subject
            id      = $mail.EntryID
        }
    }

    $mail.Send()

    return [pscustomobject]@{
        status  = 'sent'
        to      = $To
        subject = $Subject
    }
}

function Invoke-MailMcp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('account', 'folders', 'list', 'unread', 'search', 'search_attachments', 'read', 'send', 'reply', 'mark_read', 'mark_unread', 'move', 'archive', 'delete', 'label_list', 'label_add', 'label_remove', 'label_clear', 'label_auto_subject', 'attachments_list', 'attachments_download')]
        [string]$Action,

        [int]$Top = 10,
        [string]$Query,
        [string]$To,
        [string]$Subject,
        [string]$Body,
        [string]$ReplyBody,
        [string]$FolderName,
        [string]$Label,
        [string[]]$Labels,
        [string]$LabelPrefix = 'Sujet',
        [string]$AttachmentName,
        [string]$AttachmentExtension,
        [int]$ScanLimit = 300,
        [int]$MaxBodyChars = 6000,
        [string]$MessageId,
        [string]$OutputDir,
        [switch]$OnlyImagesAndPdf,
        [switch]$Overwrite,
        [switch]$OpenDir,
        [switch]$ReplyAll,
        [switch]$UnreadOnly,
        [switch]$CreateFolder,
        [switch]$SendNow,
        [switch]$Html,
        [switch]$DryRun
    )

    if (-not $script:MailMcpContext) {
        $script:MailMcpContext = New-MailMcpContext
    }

    switch ($Action) {
        'account' {
            return [pscustomobject]@{
                profile     = $script:MailMcpContext.ProfileName
                smtpAddress = $script:MailMcpContext.AccountSmtp
            }
        }

        'folders' {
            return Get-MailMcpFolders
        }

        'list' {
            return Get-MailMcpMessages -Folder $script:MailMcpContext.Inbox -Top $Top
        }

        'unread' {
            $messages = Get-MailMcpMessages -Folder $script:MailMcpContext.Inbox -Top $ScanLimit
            return $messages | Where-Object { $_.isUnread } | Select-Object -First $Top
        }

        'search' {
            return Find-MailMcpMessages -Query $Query -Top $Top -ScanLimit $ScanLimit
        }

        'search_attachments' {
            return Search-MailMcpAttachments -Query $Query -Top $Top -ScanLimit $ScanLimit -AttachmentName $AttachmentName -AttachmentExtension $AttachmentExtension -OnlyImagesAndPdf:$OnlyImagesAndPdf
        }

        'read' {
            $mail = Resolve-MailMcpMessage -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit
            return Get-MailMcpMessageDetails -MailItem $mail -MaxBodyChars $MaxBodyChars
        }

        'send' {
            if ([string]::IsNullOrWhiteSpace($To)) {
                throw 'To est requis pour Action=send.'
            }
            if ([string]::IsNullOrWhiteSpace($Subject)) {
                throw 'Subject est requis pour Action=send.'
            }
            if ([string]::IsNullOrWhiteSpace($Body)) {
                throw 'Body est requis pour Action=send.'
            }

            return Send-MailMcpMessage -To $To -Subject $Subject -Body $Body -Html:$Html -DryRun:$DryRun
        }

        'reply' {
            $mail = Resolve-MailMcpMessage -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit
            return New-MailMcpReply -MailItem $mail -ReplyBody $ReplyBody -ReplyAll:$ReplyAll -Html:$Html -SendNow:$SendNow
        }

        'mark_read' {
            $mail = Resolve-MailMcpMessage -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit
            return Set-MailMcpReadState -MailItem $mail -IsRead $true
        }

        'mark_unread' {
            $mail = Resolve-MailMcpMessage -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit
            return Set-MailMcpReadState -MailItem $mail -IsRead $false
        }

        'move' {
            if ([string]::IsNullOrWhiteSpace($FolderName)) {
                throw 'FolderName est requis pour Action=move.'
            }

            $mail = Resolve-MailMcpMessage -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit
            $target = Resolve-MailMcpFolder -FolderName $FolderName -CreateIfMissing:$CreateFolder
            return Move-MailMcpMessage -MailItem $mail -TargetFolder $target -DryRun:$DryRun
        }

        'archive' {
            $mail = Resolve-MailMcpMessage -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit
            return Move-MailMcpMessageToArchive -MailItem $mail -DryRun:$DryRun
        }

        'delete' {
            $mail = Resolve-MailMcpMessage -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit
            return Remove-MailMcpMessage -MailItem $mail -DryRun:$DryRun
        }

        'label_list' {
            $mail = Resolve-MailMcpMessage -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit
            $currentLabels = @(Get-MailMcpLabelsFromMailItem -MailItem $mail)
            return [pscustomobject]@{
                id = $mail.EntryID
                subject = $mail.Subject
                labelCount = $currentLabels.Count
                labels = $currentLabels
            }
        }

        'label_add' {
            $targetLabels = @(Resolve-MailMcpInputLabels -Label $Label -Labels $Labels)
            if ($targetLabels.Count -eq 0) {
                throw 'Fournis Label ou Labels pour Action=label_add.'
            }

            $mails = Resolve-MailMcpMessages -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit -Top $Top
            $results = foreach ($mail in $mails) {
                Add-MailMcpLabelsToMailItem -MailItem $mail -Labels $targetLabels -DryRun:$DryRun
            }
            return $results
        }

        'label_remove' {
            $targetLabels = @(Resolve-MailMcpInputLabels -Label $Label -Labels $Labels)
            if ($targetLabels.Count -eq 0) {
                throw 'Fournis Label ou Labels pour Action=label_remove.'
            }

            $mails = Resolve-MailMcpMessages -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit -Top $Top
            $results = foreach ($mail in $mails) {
                Remove-MailMcpLabelsFromMailItem -MailItem $mail -Labels $targetLabels -DryRun:$DryRun
            }
            return $results
        }

        'label_clear' {
            $mails = Resolve-MailMcpMessages -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit -Top $Top
            $results = foreach ($mail in $mails) {
                Clear-MailMcpLabelsFromMailItem -MailItem $mail -DryRun:$DryRun
            }
            return $results
        }

        'label_auto_subject' {
            return Invoke-MailMcpAutoSubjectLabeling -Query $Query -Top $Top -ScanLimit $ScanLimit -LabelPrefix $LabelPrefix -UnreadOnly:$UnreadOnly -DryRun:$DryRun
        }

        'attachments_list' {
            $mail = Resolve-MailMcpMessage -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit
            $attachments = Get-MailMcpAttachmentsInfo -MailItem $mail

            return [pscustomobject]@{
                id = $mail.EntryID
                subject = $mail.Subject
                fromName = $mail.SenderName
                fromAddress = $mail.SenderEmailAddress
                receivedTime = $mail.ReceivedTime
                attachmentCount = $attachments.Count
                attachments = $attachments
            }
        }

        'attachments_download' {
            $mail = Resolve-MailMcpMessage -MessageId $MessageId -Query $Query -ScanLimit $ScanLimit
            $attachments = Get-MailMcpAttachmentsInfo -MailItem $mail
            $download = Save-MailMcpAttachments -MailItem $mail -OutputDir $OutputDir -OnlyImagesAndPdf:$OnlyImagesAndPdf -Overwrite:$Overwrite

            if ($OpenDir -and (Test-Path -LiteralPath $download.outputDir)) {
                Start-Process explorer.exe $download.outputDir | Out-Null
            }

            return [pscustomobject]@{
                id = $mail.EntryID
                subject = $mail.Subject
                fromName = $mail.SenderName
                fromAddress = $mail.SenderEmailAddress
                receivedTime = $mail.ReceivedTime
                totalAttachments = $attachments.Count
                downloadedCount = $download.downloadedCount
                outputDir = $download.outputDir
                files = $download.files
            }
        }
    }
}

function Disconnect-MailMcp {
    [CmdletBinding()]
    param()

    if (-not $script:MailMcpContext) {
        return
    }

    foreach ($objName in @('Inbox', 'Drafts', 'Sent', 'Session', 'Outlook')) {
        $obj = $script:MailMcpContext.$objName
        if ($null -ne $obj) {
            try {
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj)
            }
            catch {
            }
        }
    }

    $script:MailMcpContext = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

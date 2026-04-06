Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# =========================
# Konfiguration
# =========================
$AppRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}
else {
    (Get-Location).Path
}

$SshConfigPath = Join-Path $HOME ".ssh\config"
$HostMetaPath = Join-Path $AppRoot "ssh-host-meta.json"
$UiStatePath = Join-Path $AppRoot "ssh-host-ui.json"
$AppName = "SSH Vault"
$AppAuthor = "kanuracer"
$AppVersion = "0.9.0"
$GitHubRepo = "kanuracer/ssh-vault"
$GitHubRepoUrl = "https://github.com/$GitHubRepo"
$GitHubBranch = "main"
$AppLogoPath = Join-Path $AppRoot "logo.png"
$WindowTitle = "$AppName v$AppVersion"

$Translations = @{
    de = @{
        Hosts = "Hosts"
        Info = "Info"
        Settings = "Einstellungen"
        RestartRequiredNote = "Hinweis: Nach Sprachwechsel App bitte neu starten."
        Tags = "Tags"
        Filter = "Filter"
        NewHost = "+ Host"
        Refresh = "Neu laden"
        SearchPlaceholder = ""
        SearchPrefix = "Suche"
        FilterAll = "Filter: Alle"
        FilterPrefix = "Filter"
        SortByHost = "Nach Hostname"
        SortByTag = "Nach Tag, dann Hostname"
        LanguageGerman = "Deutsch"
        LanguageEnglish = "Englisch"
        SettingsTitle = "Einstellungen"
        LanguageLabel = "Sprache"
        SortLabel = "Sortierung"
        AppInfoTitle = "App-Informationen"
        VersionLabel = "Version"
        AuthorLabel = "Autor"
        RepoLabel = "Repository"
        UpdateSection = "Updates"
        Untagged = "Ohne Tag"
        CheckUpdate = "Update pruefen"
        Update = "Update"
        UpdateChecking = "Pruefe Update..."
        VersionOnly = "Version v{0}"
        VersionWithUpdate = "Version v{0} | Update verfuegbar: v{1}"
        StatusLoadingHosts = "Lade Hosts..."
        StatusHosts = "{0} / {1} Host(s)"
        SortModeHost = "host"
        SortModeTag = "tag"
        TagDialogTitle = "Tags bearbeiten"
        NoHostsAvailable = "Keine Hosts verfuegbar."
        Notice = "Hinweis"
        HostLabel = "Host"
        AliasLabel = "Alias (Host)"
        HostNameLabel = "HostName"
        UserOptionalLabel = "User (optional)"
        PortOptionalLabel = "Port (optional)"
        ExistingTags = "Vorhandene Tags"
        NewTags = "Neue Tags"
        TagDialogInfo = "Mehrere Tags moeglich. Neue Tags komma-getrennt eingeben."
        NewHostDialogTitle = "Neuen SSH Host anlegen"
        InputIncomplete = "Alias und HostName sind Pflichtfelder."
        InputIncompleteTitle = "Eingabe unvollstaendig"
        InvalidAlias = "Der Alias darf keine Leerzeichen enthalten."
        InvalidAliasTitle = "Ungueltiger Alias"
        InvalidPort = "Port muss numerisch sein."
        InvalidPortTitle = "Ungueltiger Port"
        Save = "Speichern"
        Cancel = "Abbrechen"
        FilterDialogTitle = "Filter bearbeiten"
        FilterIncludeLabel = "Zeigen, wenn Host einen dieser Tags hat"
        FilterExcludeLabel = "Ausblenden, wenn Host einen dieser Tags hat"
        Apply = "Anwenden"
        CreateHostSaveError = "Host konnte nicht gespeichert werden.`n`n{0}"
        SaveErrorTitle = "Speicherfehler"
        MissingConfig = "Die Datei '{0}' wurde nicht gefunden."
        MissingConfigTitle = "SSH-Config fehlt"
        ReadError = "Fehler beim Einlesen der SSH-Config.`n`n{0}"
        ReadErrorTitle = "Lesefehler"
        NoUpdateInstalled = "Kein Update installiert"
        UpdateConfirm = "Soll SSH Vault wirklich von v{0} auf v{1} aktualisiert werden?"
        UpdateConfirmTitle = "Update bestaetigen"
        UpdateSuccess = "Update erfolgreich installiert.`nBackup: {0}`nBitte App neu starten."
        UpdateSuccessTitle = "Update erfolgreich"
        UpdateFailed = "Update fehlgeschlagen.`n`n{0}"
        UpdateFailedTitle = "Updatefehler"
    }
    en = @{
        Hosts = "Hosts"
        Info = "Info"
        Settings = "Settings"
        RestartRequiredNote = "Note: Restart the app after changing the language."
        Tags = "Tags"
        Filter = "Filter"
        NewHost = "+ Host"
        Refresh = "Refresh"
        SearchPlaceholder = ""
        SearchPrefix = "Search"
        FilterAll = "Filter: All"
        FilterPrefix = "Filter"
        SortByHost = "By host name"
        SortByTag = "By tag, then host name"
        LanguageGerman = "German"
        LanguageEnglish = "English"
        SettingsTitle = "Settings"
        LanguageLabel = "Language"
        SortLabel = "Sorting"
        AppInfoTitle = "App Information"
        VersionLabel = "Version"
        AuthorLabel = "Author"
        RepoLabel = "Repository"
        UpdateSection = "Updates"
        Untagged = "Untagged"
        CheckUpdate = "Check Update"
        Update = "Update"
        UpdateChecking = "Checking for updates..."
        VersionOnly = "Version v{0}"
        VersionWithUpdate = "Version v{0} | Update available: v{1}"
        StatusLoadingHosts = "Loading hosts..."
        StatusHosts = "{0} / {1} Host(s)"
        SortModeHost = "host"
        SortModeTag = "tag"
        TagDialogTitle = "Edit Tags"
        NoHostsAvailable = "No hosts available."
        Notice = "Notice"
        HostLabel = "Host"
        AliasLabel = "Alias (Host)"
        HostNameLabel = "HostName"
        UserOptionalLabel = "User (optional)"
        PortOptionalLabel = "Port (optional)"
        ExistingTags = "Existing tags"
        NewTags = "New tags"
        TagDialogInfo = "Multiple tags supported. Enter new tags separated by commas."
        NewHostDialogTitle = "Create New SSH Host"
        InputIncomplete = "Alias and HostName are required."
        InputIncompleteTitle = "Incomplete input"
        InvalidAlias = "Alias must not contain spaces."
        InvalidAliasTitle = "Invalid alias"
        InvalidPort = "Port must be numeric."
        InvalidPortTitle = "Invalid port"
        Save = "Save"
        Cancel = "Cancel"
        FilterDialogTitle = "Edit Filters"
        FilterIncludeLabel = "Show when host has one of these tags"
        FilterExcludeLabel = "Hide when host has one of these tags"
        Apply = "Apply"
        CreateHostSaveError = "Could not save host.`n`n{0}"
        SaveErrorTitle = "Save Error"
        MissingConfig = "The file '{0}' was not found."
        MissingConfigTitle = "SSH config missing"
        ReadError = "Error while reading SSH config.`n`n{0}"
        ReadErrorTitle = "Read Error"
        NoUpdateInstalled = "No update installed"
        UpdateConfirm = "Update SSH Vault from v{0} to v{1}?"
        UpdateConfirmTitle = "Confirm update"
        UpdateSuccess = "Update installed successfully.`nBackup: {0}`nPlease restart the app."
        UpdateSuccessTitle = "Update successful"
        UpdateFailed = "Update failed.`n`n{0}"
        UpdateFailedTitle = "Update error"
    }
}

# Dark theme colors
$BgMain = [System.Drawing.Color]::FromArgb(18, 22, 28)
$BgPanel = [System.Drawing.Color]::FromArgb(24, 30, 38)
$BgInput = [System.Drawing.Color]::FromArgb(34, 42, 52)
$BgCard = [System.Drawing.Color]::FromArgb(26, 34, 44)
$BgCardHover = [System.Drawing.Color]::FromArgb(36, 48, 62)
$FgMain = [System.Drawing.Color]::FromArgb(230, 230, 230)
$FgMuted = [System.Drawing.Color]::FromArgb(158, 170, 182)
$Accent = [System.Drawing.Color]::FromArgb(33, 150, 243)
$Border = [System.Drawing.Color]::FromArgb(56, 66, 78)

# =========================
# Hilfsfunktionen
# =========================
function Get-UiLanguage {
    $language = [string]$script:UiState.Language
    if ([string]::IsNullOrWhiteSpace($language) -or -not $Translations.ContainsKey($language)) {
        return "de"
    }
    return $language
}

function T {
    param(
        [Parameter(Mandatory = $true)][string]$Key,
        [object[]]$FormatArgs
    )

    $language = if ($null -ne $script:UiState) { Get-UiLanguage } else { "de" }
    $table = $Translations[$language]
    $text = if ($table.ContainsKey($Key)) { [string]$table[$Key] } else { $Key }

    if ($null -ne $FormatArgs -and $FormatArgs.Count -gt 0) {
        return [string]::Format($text, $FormatArgs)
    }

    return $text
}

function Get-CurrentSortMode {
    $sortMode = [string]$script:UiState.SortMode
    if ($sortMode -notin @("host", "tag")) {
        return "host"
    }
    return $sortMode
}

function Get-HostSortKey {
    param([Parameter(Mandatory = $true)][string]$HostName)

    $tags = @(Get-HostTags -HostName $HostName | Sort-Object)
    if ($tags.Count -eq 0) {
        return "~"
    }

    return (($tags -join "|").ToLowerInvariant())
}

function Get-SortedHosts {
    param([Parameter(Mandatory = $true)][array]$Hosts)

    if ((Get-CurrentSortMode) -eq "tag") {
        return @(
            $Hosts |
                Sort-Object -Property `
                    @{ Expression = { Get-HostSortKey -HostName $_.Host } }, `
                    @{ Expression = { $_.Host.ToLowerInvariant() } }
        )
    }

    return @($Hosts | Sort-Object -Property @{ Expression = { $_.Host.ToLowerInvariant() } })
}

function Resolve-SshConfigIncludes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    $allFiles = New-Object System.Collections.Generic.List[string]

    function Add-ConfigFileRecursive {
        param([string]$Path)

        if (-not (Test-Path $Path)) {
            return
        }

        $fullPath = [System.IO.Path]::GetFullPath($Path)
        if ($allFiles.Contains($fullPath)) {
            return
        }

        $allFiles.Add($fullPath) | Out-Null

        $baseDir = Split-Path -Path $fullPath -Parent
        $lines = Get-Content -Path $fullPath -ErrorAction SilentlyContinue

        foreach ($line in $lines) {
            $trimmed = $line.Trim()
            if ($trimmed -match '^(?i)Include\s+(.+)$') {
                $includeValue = $matches[1].Trim()

                if ($includeValue -match '^(.*?)(\s+#.*)?$') {
                    $includeValue = $matches[1].Trim()
                }

                $parts = $includeValue -split '\s+'
                foreach ($part in $parts) {
                    if ([string]::IsNullOrWhiteSpace($part)) {
                        continue
                    }

                    $expanded = $part.Replace('~', $HOME)

                    if (-not [System.IO.Path]::IsPathRooted($expanded)) {
                        $expanded = Join-Path $baseDir $expanded
                    }

                    $matchesFound = Get-ChildItem -Path $expanded -File -ErrorAction SilentlyContinue
                    foreach ($match in $matchesFound) {
                        Add-ConfigFileRecursive -Path $match.FullName
                    }
                }
            }
        }
    }

    Add-ConfigFileRecursive -Path $ConfigPath
    return $allFiles
}

function Get-SshHostsFromConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    $result = New-Object System.Collections.Generic.List[object]
    $seen = @{}

    $files = Resolve-SshConfigIncludes -ConfigPath $ConfigPath

    foreach ($file in $files) {
        $lines = Get-Content -Path $file -ErrorAction SilentlyContinue

        foreach ($line in $lines) {
            $trimmed = $line.Trim()

            if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
            if ($trimmed.StartsWith('#')) { continue }

            if ($trimmed -match '^(?i)Host\s+(.+)$') {
                $hostPart = $matches[1].Trim()

                if ($hostPart -match '^(.*?)(\s+#.*)?$') {
                    $hostPart = $matches[1].Trim()
                }

                $hosts = $hostPart -split '\s+'

                foreach ($hostName in $hosts) {
                    if ([string]::IsNullOrWhiteSpace($hostName)) { continue }
                    if ($hostName -match '[\*\?]' -or $hostName.StartsWith('!')) { continue }

                    $key = $hostName.ToLowerInvariant()
                    if (-not $seen.ContainsKey($key)) {
                        $seen[$key] = $true
                        $result.Add([PSCustomObject]@{
                            Host       = $hostName
                            SourceFile = $file
                        }) | Out-Null
                    }
                }
            }
        }
    }

    return $result | Sort-Object Host
}

function Get-DefaultHostMeta {
    return @{
        IncludeTags = @()
        ExcludeTags = @()
        HostTags = @{}
        KnownTags = @()
    }
}

function Import-HostMeta {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MetaPath
    )

    $meta = Get-DefaultHostMeta
    if (-not (Test-Path $MetaPath)) {
        $json = ([ordered]@{
            IncludeTags   = @()
            ExcludeTags   = @()
            HostTags      = [ordered]@{}
            KnownTags     = @()
        } | ConvertTo-Json -Depth 5)
        Set-Content -Path $MetaPath -Value $json -Encoding UTF8
        return $meta
    }

    try {
        $raw = Get-Content -Path $MetaPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($null -ne $raw.IncludeTags) {
            $meta.IncludeTags = @($raw.IncludeTags | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        }
        elseif ($null -ne $raw.LastTagFilter -and -not [string]::IsNullOrWhiteSpace([string]$raw.LastTagFilter) -and [string]$raw.LastTagFilter -ne "Alle") {
            $meta.IncludeTags = @([string]$raw.LastTagFilter)
        }
        if ($null -ne $raw.ExcludeTags) {
            $meta.ExcludeTags = @($raw.ExcludeTags | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        }
        if ($null -ne $raw.HostTags) {
            foreach ($p in $raw.HostTags.PSObject.Properties) {
                $tagList = @()
                if ($p.Value -is [System.Array]) {
                    $tagList = @($p.Value | ForEach-Object { [string]$_ })
                }
                elseif ($null -ne $p.Value -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
                    $tagList = @([string]$p.Value)
                }
                $meta.HostTags[[string]$p.Name] = @($tagList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
            }
        }
        if ($null -ne $raw.KnownTags) {
            $meta.KnownTags = @($raw.KnownTags | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        }
    }
    catch {
        # Fallback to defaults when metadata is invalid.
        $meta = Get-DefaultHostMeta
    }

    return $meta
}

function Export-HostMeta {
    param(
        [Parameter(Mandatory = $true)] [string]$MetaPath,
        [Parameter(Mandatory = $true)] [hashtable]$Meta
    )

    $hostTags = [ordered]@{}
    foreach ($key in ($Meta.HostTags.Keys | Sort-Object)) {
        $hostTags[$key] = $Meta.HostTags[$key]
    }

    $payload = [ordered]@{
        IncludeTags = @($Meta.IncludeTags | Sort-Object -Unique)
        ExcludeTags = @($Meta.ExcludeTags | Sort-Object -Unique)
        HostTags = $hostTags
        KnownTags = @($Meta.KnownTags | Sort-Object -Unique)
    }

    $json = $payload | ConvertTo-Json -Depth 5
    Set-Content -Path $MetaPath -Value $json -Encoding UTF8
}

function Get-HostTags {
    param([Parameter(Mandatory = $true)][string]$HostName)
    $key = $HostName.ToLowerInvariant()
    if ($script:HostMeta.HostTags.ContainsKey($key)) {
        return @($script:HostMeta.HostTags[$key])
    }
    return @()
}

function Set-HostTags {
    param(
        [Parameter(Mandatory = $true)][string]$HostName,
        [string[]]$TagNames
    )

    $key = $HostName.ToLowerInvariant()
    $normalized = @($TagNames | ForEach-Object { [string]$_ } | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

    if ($normalized.Count -eq 0) {
        $null = $script:HostMeta.HostTags.Remove($key)
    }
    else {
        $script:HostMeta.HostTags[$key] = $normalized
        $script:HostMeta.KnownTags = @($script:HostMeta.KnownTags + $normalized | Sort-Object -Unique)
    }

    Export-HostMeta -MetaPath $HostMetaPath -Meta $script:HostMeta
}

function Get-AllTags {
    $all = @()
    $all += @($script:HostMeta.KnownTags)
    foreach ($tagSet in $script:HostMeta.HostTags.Values) {
        $all += @($tagSet)
    }
    return @($all | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
}

function Start-SshHost {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostName
    )

    $sshPath = (Get-Command ssh.exe -ErrorAction SilentlyContinue).Source
    if (-not $sshPath) {
        [System.Windows.Forms.MessageBox]::Show(
            "ssh.exe wurde nicht gefunden. OpenSSH Client ist nicht installiert oder nicht im PATH.",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    try {
        Start-Process -FilePath "cmd.exe" -ArgumentList @("/k", "ssh $HostName")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "SSH-Verbindung zu '$HostName' konnte nicht gestartet werden.`n`n$($_.Exception.Message)",
            "Startfehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

function Repair-SshConfigPermissions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    try {
        if (-not (Test-Path $ConfigPath)) {
            New-Item -Path $ConfigPath -ItemType File -Force | Out-Null
        }

        $currentUser = [System.Security.Principal.NTAccount]::new("$env:USERDOMAIN\$env:USERNAME")
        $userSid = $currentUser.Translate([System.Security.Principal.SecurityIdentifier])
        $systemSid = [System.Security.Principal.SecurityIdentifier]::new("S-1-5-18")
        $adminsSid = [System.Security.Principal.SecurityIdentifier]::new("S-1-5-32-544")

        $acl = New-Object System.Security.AccessControl.FileSecurity
        $acl.SetOwner($currentUser)
        $acl.SetAccessRuleProtection($true, $false)

        $userRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $userSid,
            [System.Security.AccessControl.FileSystemRights]::FullControl,
            [System.Security.AccessControl.AccessControlType]::Allow
        )
        $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $systemSid,
            [System.Security.AccessControl.FileSystemRights]::FullControl,
            [System.Security.AccessControl.AccessControlType]::Allow
        )
        $adminsRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $adminsSid,
            [System.Security.AccessControl.FileSystemRights]::FullControl,
            [System.Security.AccessControl.AccessControlType]::Allow
        )

        [void]$acl.AddAccessRule($userRule)
        [void]$acl.AddAccessRule($systemRule)
        [void]$acl.AddAccessRule($adminsRule)

        Set-Acl -Path $ConfigPath -AclObject $acl
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Rechte fuer '$ConfigPath' konnten nicht gesetzt werden.`n`n$($_.Exception.Message)",
            "Berechtigungsfehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $false
    }
}

function New-HostEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = T "NewHostDialogTitle"
    $dialog.Size = New-Object System.Drawing.Size(460, 300)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.BackColor = $BgPanel
    $dialog.ForeColor = $FgMain

    $labelAlias = New-Object System.Windows.Forms.Label
    $labelAlias.Text = "{0}:" -f (T "AliasLabel")
    $labelAlias.Location = New-Object System.Drawing.Point(20, 20)
    $labelAlias.AutoSize = $true

    $txtAlias = New-Object System.Windows.Forms.TextBox
    $txtAlias.Location = New-Object System.Drawing.Point(140, 16)
    $txtAlias.Size = New-Object System.Drawing.Size(280, 24)
    $txtAlias.BackColor = $BgInput
    $txtAlias.ForeColor = $FgMain
    $txtAlias.BorderStyle = "FixedSingle"

    $labelHostName = New-Object System.Windows.Forms.Label
    $labelHostName.Text = "{0}:" -f (T "HostNameLabel")
    $labelHostName.Location = New-Object System.Drawing.Point(20, 60)
    $labelHostName.AutoSize = $true

    $txtHostName = New-Object System.Windows.Forms.TextBox
    $txtHostName.Location = New-Object System.Drawing.Point(140, 56)
    $txtHostName.Size = New-Object System.Drawing.Size(280, 24)
    $txtHostName.BackColor = $BgInput
    $txtHostName.ForeColor = $FgMain
    $txtHostName.BorderStyle = "FixedSingle"

    $labelUser = New-Object System.Windows.Forms.Label
    $labelUser.Text = "{0}:" -f (T "UserOptionalLabel")
    $labelUser.Location = New-Object System.Drawing.Point(20, 100)
    $labelUser.AutoSize = $true

    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(140, 96)
    $txtUser.Size = New-Object System.Drawing.Size(280, 24)
    $txtUser.BackColor = $BgInput
    $txtUser.ForeColor = $FgMain
    $txtUser.BorderStyle = "FixedSingle"

    $labelPort = New-Object System.Windows.Forms.Label
    $labelPort.Text = "{0}:" -f (T "PortOptionalLabel")
    $labelPort.Location = New-Object System.Drawing.Point(20, 140)
    $labelPort.AutoSize = $true

    $txtPort = New-Object System.Windows.Forms.TextBox
    $txtPort.Location = New-Object System.Drawing.Point(140, 136)
    $txtPort.Size = New-Object System.Drawing.Size(280, 24)
    $txtPort.BackColor = $BgInput
    $txtPort.ForeColor = $FgMain
    $txtPort.BorderStyle = "FixedSingle"

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = T "Save"
    $btnSave.Location = New-Object System.Drawing.Point(260, 190)
    $btnSave.Size = New-Object System.Drawing.Size(160, 34)
    $btnSave.BackColor = $Accent
    $btnSave.ForeColor = $FgMain
    $btnSave.FlatStyle = "Flat"
    $btnSave.FlatAppearance.BorderSize = 0

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = T "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(140, 190)
    $btnCancel.Size = New-Object System.Drawing.Size(110, 34)
    $btnCancel.BackColor = $BgInput
    $btnCancel.ForeColor = $FgMain
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.FlatAppearance.BorderColor = $FgMuted

    $dialog.Controls.AddRange(@(
            $labelAlias, $txtAlias,
            $labelHostName, $txtHostName,
            $labelUser, $txtUser,
            $labelPort, $txtPort,
            $btnSave, $btnCancel
        ))

    $script:NewHostAlias = $null

    $btnCancel.Add_Click({
            $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $dialog.Close()
        })

    $btnSave.Add_Click({
            $alias = $txtAlias.Text.Trim()
            $targetHost = $txtHostName.Text.Trim()
            $user = $txtUser.Text.Trim()
            $port = $txtPort.Text.Trim()

            if ([string]::IsNullOrWhiteSpace($alias) -or [string]::IsNullOrWhiteSpace($targetHost)) {
                [System.Windows.Forms.MessageBox]::Show(
                    (T "InputIncomplete"),
                    (T "InputIncompleteTitle"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }

            if ($alias -match '\s') {
                [System.Windows.Forms.MessageBox]::Show(
                    (T "InvalidAlias"),
                    (T "InvalidAliasTitle"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }

            if (-not [string]::IsNullOrWhiteSpace($port) -and $port -notmatch '^\d+$') {
                [System.Windows.Forms.MessageBox]::Show(
                    (T "InvalidPort"),
                    (T "InvalidPortTitle"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }

            try {
                if (-not (Test-Path $ConfigPath)) {
                    New-Item -Path $ConfigPath -ItemType File -Force | Out-Null
                }

                $entryLines = New-Object System.Collections.Generic.List[string]
                $entryLines.Add("") | Out-Null
                $entryLines.Add("Host $alias") | Out-Null
                $entryLines.Add("    HostName $targetHost") | Out-Null
                if (-not [string]::IsNullOrWhiteSpace($user)) {
                    $entryLines.Add("    User $user") | Out-Null
                }
                if (-not [string]::IsNullOrWhiteSpace($port)) {
                    $entryLines.Add("    Port $port") | Out-Null
                }

                Add-Content -Path $ConfigPath -Value ($entryLines -join [Environment]::NewLine)
                $script:NewHostAlias = $alias
                $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $dialog.Close()
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    (T "CreateHostSaveError" @($_.Exception.Message)),
                    (T "SaveErrorTitle"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
            }
        })

    $dialog.AcceptButton = $btnSave
    $dialog.CancelButton = $btnCancel

    [void]$dialog.ShowDialog()
    return $script:NewHostAlias
}

function Set-HostTagDialog {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Hosts
    )

    if ($Hosts.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            (T "NoHostsAvailable"),
            (T "Notice"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return $null
    }

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = T "TagDialogTitle"
    $dialog.Size = New-Object System.Drawing.Size(560, 340)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.BackColor = $BgPanel
    $dialog.ForeColor = $FgMain
    $dialog.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $dialogLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $dialogLayout.Dock = "Fill"
    $dialogLayout.Padding = New-Object System.Windows.Forms.Padding(16)
    $dialogLayout.BackColor = $BgPanel
    $dialogLayout.ColumnCount = 2
    $dialogLayout.RowCount = 5
    $dialogLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110))) | Out-Null
    $dialogLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $dialogLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
    $dialogLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 120))) | Out-Null
    $dialogLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42))) | Out-Null
    $dialogLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $dialogLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42))) | Out-Null

    $labelHost = New-Object System.Windows.Forms.Label
    $labelHost.Text = T "HostLabel"
    $labelHost.Dock = "Fill"
    $labelHost.TextAlign = "MiddleLeft"
    $labelHost.AutoSize = $true

    $comboHost = New-Object System.Windows.Forms.ComboBox
    $comboHost.Dock = "Fill"
    $comboHost.DropDownStyle = "DropDownList"
    $comboHost.BackColor = $BgInput
    $comboHost.ForeColor = $FgMain
    $comboHost.FlatStyle = "Popup"

    $labelTag = New-Object System.Windows.Forms.Label
    $labelTag.Text = T "ExistingTags"
    $labelTag.Dock = "Fill"
    $labelTag.TextAlign = "MiddleLeft"
    $labelTag.AutoSize = $true

    $tagList = New-Object System.Windows.Forms.CheckedListBox
    $tagList.Dock = "Fill"
    $tagList.CheckOnClick = $true
    $tagList.BackColor = $BgInput
    $tagList.ForeColor = $FgMain
    $tagList.BorderStyle = "FixedSingle"

    $labelNewTags = New-Object System.Windows.Forms.Label
    $labelNewTags.Text = T "NewTags"
    $labelNewTags.Dock = "Fill"
    $labelNewTags.TextAlign = "MiddleLeft"
    $labelNewTags.AutoSize = $true

    $newTagsBox = New-Object System.Windows.Forms.TextBox
    $newTagsBox.Dock = "Fill"
    $newTagsBox.BackColor = $BgInput
    $newTagsBox.ForeColor = $FgMain
    $newTagsBox.BorderStyle = "FixedSingle"

    $info = New-Object System.Windows.Forms.Label
    $info.Text = T "TagDialogInfo"
    $info.Dock = "Fill"
    $info.TextAlign = "TopLeft"
    $info.ForeColor = $FgMuted

    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock = "Fill"
    $buttonPanel.FlowDirection = "RightToLeft"
    $buttonPanel.WrapContents = $false

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = T "Save"
    $btnSave.Width = 110
    $btnSave.Height = 30
    $btnSave.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
    $btnSave.BackColor = $Accent
    $btnSave.ForeColor = $FgMain
    $btnSave.FlatStyle = "Flat"
    $btnSave.FlatAppearance.BorderSize = 0

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = T "Cancel"
    $btnCancel.Width = 110
    $btnCancel.Height = 30
    $btnCancel.BackColor = $BgInput
    $btnCancel.ForeColor = $FgMain
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.FlatAppearance.BorderColor = $Border

    [void]$buttonPanel.Controls.Add($btnSave)
    [void]$buttonPanel.Controls.Add($btnCancel)

    foreach ($h in ($Hosts | Sort-Object Host)) {
        [void]$comboHost.Items.Add($h.Host)
    }

    if ($comboHost.Items.Count -gt 0) {
        $comboHost.SelectedIndex = 0
        foreach ($tag in (Get-AllTags)) {
            [void]$tagList.Items.Add($tag)
        }
        $existingTags = @(Get-HostTags -HostName ([string]$comboHost.SelectedItem))
        for ($index = 0; $index -lt $tagList.Items.Count; $index++) {
            if ($existingTags -contains [string]$tagList.Items[$index]) {
                $tagList.SetItemChecked($index, $true)
            }
        }
    }

    $comboHost.Add_SelectedIndexChanged({
        $selected = [string]$comboHost.SelectedItem
        for ($index = 0; $index -lt $tagList.Items.Count; $index++) {
            $tagList.SetItemChecked($index, $false)
        }
        $newTagsBox.Text = ""
        $existingTags = @(Get-HostTags -HostName $selected)
        for ($index = 0; $index -lt $tagList.Items.Count; $index++) {
            if ($existingTags -contains [string]$tagList.Items[$index]) {
                $tagList.SetItemChecked($index, $true)
            }
        }
    })

    $script:TaggedHostResult = $null

    $btnCancel.Add_Click({
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Close()
    })

    $btnSave.Add_Click({
        $selected = [string]$comboHost.SelectedItem
        if ([string]::IsNullOrWhiteSpace($selected)) {
            return
        }
        $selectedTags = @($tagList.CheckedItems | ForEach-Object { [string]$_ })
        $typedTags = @($newTagsBox.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        Set-HostTags -HostName $selected -TagNames @($selectedTags + $typedTags)
        $script:TaggedHostResult = $selected
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    })

    [void]$dialogLayout.Controls.Add($labelHost, 0, 0)
    [void]$dialogLayout.Controls.Add($comboHost, 1, 0)
    [void]$dialogLayout.Controls.Add($labelTag, 0, 1)
    [void]$dialogLayout.Controls.Add($tagList, 1, 1)
    [void]$dialogLayout.Controls.Add($labelNewTags, 0, 2)
    [void]$dialogLayout.Controls.Add($newTagsBox, 1, 2)
    [void]$dialogLayout.Controls.Add($info, 1, 3)
    [void]$dialogLayout.Controls.Add($buttonPanel, 0, 4)
    $dialogLayout.SetColumnSpan($buttonPanel, 2)
    $dialog.Controls.Add($dialogLayout)
    $dialog.AcceptButton = $btnSave
    $dialog.CancelButton = $btnCancel

    [void]$dialog.ShowDialog()
    return $script:TaggedHostResult
}

function Edit-TagFilterDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = T "FilterDialogTitle"
    $dialog.Size = New-Object System.Drawing.Size(640, 360)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.BackColor = $BgPanel
    $dialog.ForeColor = $FgMain
    $dialog.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = "Fill"
    $layout.Padding = New-Object System.Windows.Forms.Padding(16)
    $layout.BackColor = $BgPanel
    $layout.ColumnCount = 2
    $layout.RowCount = 3
    $layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42))) | Out-Null

    $includeLabel = New-Object System.Windows.Forms.Label
    $includeLabel.Text = T "FilterIncludeLabel"
    $includeLabel.Dock = "Fill"
    $includeLabel.ForeColor = $FgMain

    $excludeLabel = New-Object System.Windows.Forms.Label
    $excludeLabel.Text = T "FilterExcludeLabel"
    $excludeLabel.Dock = "Fill"
    $excludeLabel.ForeColor = $FgMain

    $includeList = New-Object System.Windows.Forms.CheckedListBox
    $includeList.Dock = "Fill"
    $includeList.CheckOnClick = $true
    $includeList.BackColor = $BgInput
    $includeList.ForeColor = $FgMain
    $includeList.BorderStyle = "FixedSingle"

    $excludeList = New-Object System.Windows.Forms.CheckedListBox
    $excludeList.Dock = "Fill"
    $excludeList.CheckOnClick = $true
    $excludeList.BackColor = $BgInput
    $excludeList.ForeColor = $FgMain
    $excludeList.BorderStyle = "FixedSingle"

    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock = "Fill"
    $buttonPanel.FlowDirection = "RightToLeft"
    $buttonPanel.WrapContents = $false

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = T "Apply"
    $btnSave.Width = 110
    $btnSave.Height = 30
    $btnSave.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
    $btnSave.BackColor = $Accent
    $btnSave.ForeColor = $FgMain
    $btnSave.FlatStyle = "Flat"
    $btnSave.FlatAppearance.BorderSize = 0

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = T "Cancel"
    $btnCancel.Width = 110
    $btnCancel.Height = 30
    $btnCancel.BackColor = $BgInput
    $btnCancel.ForeColor = $FgMain
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.FlatAppearance.BorderColor = $Border

    [void]$buttonPanel.Controls.Add($btnSave)
    [void]$buttonPanel.Controls.Add($btnCancel)

    $allTags = @(Get-AllTags)
    foreach ($tag in $allTags) {
        [void]$includeList.Items.Add($tag)
        [void]$excludeList.Items.Add($tag)
    }

    for ($index = 0; $index -lt $allTags.Count; $index++) {
        if ($script:HostMeta.IncludeTags -contains $allTags[$index]) {
            $includeList.SetItemChecked($index, $true)
        }
        if ($script:HostMeta.ExcludeTags -contains $allTags[$index]) {
            $excludeList.SetItemChecked($index, $true)
        }
    }

    $script:FilterDialogAccepted = $false
    $btnCancel.Add_Click({
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Close()
    })

    $btnSave.Add_Click({
        $script:HostMeta.IncludeTags = @($includeList.CheckedItems | ForEach-Object { [string]$_ } | Sort-Object -Unique)
        $script:HostMeta.ExcludeTags = @($excludeList.CheckedItems | ForEach-Object { [string]$_ } | Sort-Object -Unique)
        Export-HostMeta -MetaPath $HostMetaPath -Meta $script:HostMeta
        $script:FilterDialogAccepted = $true
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    })

    [void]$layout.Controls.Add($includeLabel, 0, 0)
    [void]$layout.Controls.Add($excludeLabel, 1, 0)
    [void]$layout.Controls.Add($includeList, 0, 1)
    [void]$layout.Controls.Add($excludeList, 1, 1)
    [void]$layout.Controls.Add($buttonPanel, 0, 2)
    $layout.SetColumnSpan($buttonPanel, 2)
    $dialog.Controls.Add($layout)
    $dialog.AcceptButton = $btnSave
    $dialog.CancelButton = $btnCancel

    [void]$dialog.ShowDialog()
    return $script:FilterDialogAccepted
}

function New-UiButton {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [switch]$Primary
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Height = 30
    $button.Margin = New-Object System.Windows.Forms.Padding(3)
    $button.Dock = "Fill"
    $button.ForeColor = $FgMain
    $button.FlatStyle = "Flat"
    $button.FlatAppearance.BorderColor = $Border
    $button.FlatAppearance.MouseOverBackColor = $BgCardHover
    $button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand

    if ($Primary) {
        $button.BackColor = $Accent
        $button.FlatAppearance.BorderSize = 0
    }
    else {
        $button.BackColor = $BgInput
    }

    return $button
}

function Get-DefaultUiState {
    return @{
        Width = 620
        Height = 400
        Left = -1
        Top = -1
        Language = "de"
        SortMode = "host"
    }
}

function Import-UiState {
    param([Parameter(Mandatory = $true)][string]$UiPath)
    $state = Get-DefaultUiState
    if (-not (Test-Path $UiPath)) {
        ($state | ConvertTo-Json -Depth 3) | Set-Content -Path $UiPath -Encoding UTF8
        return $state
    }
    try {
        $raw = Get-Content -Path $UiPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($raw.Width -gt 400) { $state.Width = [int]$raw.Width }
        if ($raw.Height -gt 300) { $state.Height = [int]$raw.Height }
        if ($raw.Left -ge 0) { $state.Left = [int]$raw.Left }
        if ($raw.Top -ge 0) { $state.Top = [int]$raw.Top }
        if ([string]$raw.Language -in @("de", "en")) { $state.Language = [string]$raw.Language }
        if ([string]$raw.SortMode -in @("host", "tag")) { $state.SortMode = [string]$raw.SortMode }
    }
    catch {}

    # Keep the app compact by default even when older UI state files were saved very wide.
    return $state
}

function Export-UiState {
    param(
        [Parameter(Mandatory = $true)][string]$UiPath,
        [Parameter(Mandatory = $true)][hashtable]$State
    )
    ($State | ConvertTo-Json -Depth 3) | Set-Content -Path $UiPath -Encoding UTF8
}

function Get-LatestReleaseInfo {
    try {
        return Invoke-RestMethod -Uri "https://api.github.com/repos/$GitHubRepo/releases/latest" -Headers (Get-GitHubHeaders) -TimeoutSec 15
    }
    catch {
        return $null
    }
}

function Get-RemoteAppVersion {
    try {
        $rawUrl = "https://raw.githubusercontent.com/$GitHubRepo/$GitHubBranch/ssh-vault.ps1"
        $response = Invoke-WebRequest -Uri $rawUrl -Headers (Get-GitHubHeaders) -UseBasicParsing -TimeoutSec 20
        if ($response.Content -match '(?m)^\$AppVersion\s*=\s*"([^"]+)"\s*$') {
            return $matches[1]
        }
    }
    catch {
        return $null
    }

    return $null
}

function ConvertTo-NormalizedVersion {
    param([string]$VersionText)

    if ([string]::IsNullOrWhiteSpace($VersionText)) { return $null }
    $clean = $VersionText.Trim().TrimStart('v', 'V')
    try { return [version]$clean } catch { return $null }
}

function Format-UpdateStatusReason {
    param([Parameter(Mandatory = $true)]$UpdateStatus)

    if ($null -eq $UpdateStatus.RemoteVersion -or $null -eq $UpdateStatus.LocalVersion) {
        return (T "VersionOnly" @($AppVersion))
    }

    if ($UpdateStatus.IsAvailable) {
        return (T "VersionWithUpdate" @($AppVersion, $UpdateStatus.RemoteVersionText))
    }

    return (T "VersionOnly" @($AppVersion))
}

function Get-UpdateStatus {
    $localVersion = ConvertTo-NormalizedVersion -VersionText $AppVersion
    $remoteVersionText = Get-RemoteAppVersion
    $remoteVersion = ConvertTo-NormalizedVersion -VersionText $remoteVersionText

    if ($null -eq $remoteVersion) {
        $release = Get-LatestReleaseInfo
        if ($null -ne $release) {
            $remoteVersionText = [string]$release.tag_name
            $remoteVersion = ConvertTo-NormalizedVersion -VersionText $release.tag_name
        }
    }

    if ($null -eq $remoteVersion -or $null -eq $localVersion) {
        $status = [PSCustomObject]@{
            IsAvailable = $false
            Reason = $null
            Release = $null
            RemoteVersion = $remoteVersion
            LocalVersion = $localVersion
            RemoteVersionText = $remoteVersionText
        }
        $status.Reason = Format-UpdateStatusReason -UpdateStatus $status
        return $status
    }

    $status = [PSCustomObject]@{
        IsAvailable = ($remoteVersion -gt $localVersion)
        Reason = $null
        Release = $null
        RemoteVersion = $remoteVersion
        LocalVersion = $localVersion
        RemoteVersionText = $remoteVersionText
    }
    $status.Reason = Format-UpdateStatusReason -UpdateStatus $status
    return $status
}

function Get-CurrentScriptPath {
    if ($PSCommandPath) { return $PSCommandPath }
    if ($MyInvocation.MyCommand.Path) { return $MyInvocation.MyCommand.Path }
    return (Join-Path $AppRoot "ssh-vault.ps1")
}

function Get-ShortUiText {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [int]$MaxLength = 24
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    if ($Text.Length -le $MaxLength) {
        return $Text
    }

    if ($MaxLength -le 3) {
        return "..."
    }

    return $Text.Substring(0, ($MaxLength - 3)) + "..."
}

function Resolve-AppLogoPath {
    param([Parameter(Mandatory = $true)][string]$ImagePath)

    if (Test-Path $ImagePath) {
        return $ImagePath
    }

    try {
        $logoUrl = "https://raw.githubusercontent.com/$GitHubRepo/$GitHubBranch/logo.png"
        Invoke-WebRequest -Uri $logoUrl -Headers (Get-GitHubHeaders) -OutFile $ImagePath -UseBasicParsing -TimeoutSec 15
        if (Test-Path $ImagePath) {
            return $ImagePath
        }
    }
    catch {
        return $null
    }

    return $null
}

function Get-AppIcon {
    param([Parameter(Mandatory = $true)][string]$ImagePath)

    if (-not (Test-Path $ImagePath)) {
        return $null
    }

    try {
        $bitmap = New-Object System.Drawing.Bitmap($ImagePath)
        $hIcon = $bitmap.GetHicon()
        $icon = [System.Drawing.Icon]::FromHandle($hIcon)
        $clonedIcon = [System.Drawing.Icon]$icon.Clone()
        $icon.Dispose()
        $bitmap.Dispose()
        return $clonedIcon
    }
    catch {
        return $null
    }
}

function Get-GitHubHeaders {
    return @{
        "User-Agent" = "$AppName-Updater"
        "Accept" = "application/vnd.github+json"
    }
}

function Update-AppFromGitHub {
    if ($GitHubRepo -like "REPLACE_*") {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte zuerst `\$GitHubRepo` im Script auf dein Repo setzen.",
            "Repo fehlt",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    try {
        $updateStatus = Get-UpdateStatus
        if (-not $updateStatus.IsAvailable) {
            [System.Windows.Forms.MessageBox]::Show(
                $updateStatus.Reason,
                (T "NoUpdateInstalled"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            $versionStatusLabel.Text = $updateStatus.Reason
            return
        }

        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            (T "UpdateConfirm" @($AppVersion, $updateStatus.RemoteVersionText)),
            (T "UpdateConfirmTitle"),
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }

        $rawUrl = "https://raw.githubusercontent.com/$GitHubRepo/$GitHubBranch/ssh-vault.ps1"
        $newScript = Invoke-WebRequest -Uri $rawUrl -Headers (Get-GitHubHeaders) -UseBasicParsing -TimeoutSec 20
        if ([string]::IsNullOrWhiteSpace($newScript.Content)) {
            throw "Leerer Download vom Repo."
        }

        $scriptPath = Get-CurrentScriptPath
        $backupPath = "$scriptPath.bak"
        Copy-Item -Path $scriptPath -Destination $backupPath -Force
        Set-Content -Path $scriptPath -Value $newScript.Content -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show(
            (T "UpdateSuccess" @($backupPath)),
            (T "UpdateSuccessTitle"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            (T "UpdateFailed" @($_.Exception.Message)),
            (T "UpdateFailedTitle"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

# =========================
# GUI aufbauen
# =========================
$script:AllHosts = @()
$script:HostMeta = Import-HostMeta -MetaPath $HostMetaPath
$script:HostToolTip = New-Object System.Windows.Forms.ToolTip
$script:UiState = Import-UiState -UiPath $UiStatePath
$script:LastUpdateStatus = $null

$form = New-Object System.Windows.Forms.Form
$form.Text = $WindowTitle
$form.Size = New-Object System.Drawing.Size($script:UiState.Width, $script:UiState.Height)
$form.StartPosition = "CenterScreen"
if ($script:UiState.Left -ge 0 -and $script:UiState.Top -ge 0) {
    $form.StartPosition = "Manual"
    $form.Location = New-Object System.Drawing.Point($script:UiState.Left, $script:UiState.Top)
}
$resolvedAppLogoPath = Resolve-AppLogoPath -ImagePath $AppLogoPath
$appIcon = if ($null -ne $resolvedAppLogoPath) { Get-AppIcon -ImagePath $resolvedAppLogoPath } else { $null }
if ($null -ne $appIcon) {
    $form.Icon = $appIcon
}
$form.MinimumSize = New-Object System.Drawing.Size(580, 520)
$form.BackColor = $BgMain
$form.ForeColor = $FgMain
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = "Top"
$topPanel.AutoSize = $true
$topPanel.AutoSizeMode = "GrowAndShrink"
$topPanel.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 8)
$topPanel.BackColor = $BgPanel

$headerGrid = New-Object System.Windows.Forms.TableLayoutPanel
$headerGrid.Dock = "Top"
$headerGrid.AutoSize = $true
$headerGrid.AutoSizeMode = "GrowAndShrink"
$headerGrid.ColumnCount = 2
$headerGrid.RowCount = 2
$headerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$headerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 180))) | Out-Null
$headerGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34))) | Out-Null
$headerGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

$searchWrap = New-Object System.Windows.Forms.Panel
$searchWrap.Dock = "Fill"
$searchWrap.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$searchWrap.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)
$searchWrap.BackColor = $BgInput

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Dock = "Fill"
$searchBox.Margin = New-Object System.Windows.Forms.Padding(0)
$searchBox.BackColor = $BgInput
$searchBox.ForeColor = $FgMain
$searchBox.BorderStyle = "None"
$searchBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$searchWrap.Controls.Add($searchBox)

$filterResetButton = New-Object System.Windows.Forms.Button
$filterResetButton.Dock = "Fill"
$filterResetButton.Margin = New-Object System.Windows.Forms.Padding(3)
$filterResetButton.Text = "Filter: Alle"
$filterResetButton.BackColor = $BgInput
$filterResetButton.ForeColor = $FgMain
$filterResetButton.FlatStyle = "Flat"
$filterResetButton.FlatAppearance.BorderColor = $Border
$filterResetButton.FlatAppearance.MouseOverBackColor = $BgCardHover
$filterResetButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$filterResetButton.Cursor = [System.Windows.Forms.Cursors]::Hand

$headerActions = New-Object System.Windows.Forms.FlowLayoutPanel
$headerActions.Dock = "Top"
$headerActions.AutoSize = $true
$headerActions.AutoSizeMode = "GrowAndShrink"
$headerActions.WrapContents = $true
$headerActions.FlowDirection = "LeftToRight"
$headerActions.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)

$tagButton = New-UiButton -Text "Tags"
$tagButton.Width = 100
$tagButton.Dock = "None"
$filterButton = New-UiButton -Text "Filter"
$filterButton.Width = 90
$filterButton.Dock = "None"
$newHostButton = New-UiButton -Text "+ Host" -Primary
$newHostButton.Width = 100
$newHostButton.Dock = "None"
$refreshButton = New-UiButton -Text "Refresh"
$refreshButton.Width = 90
$refreshButton.Dock = "None"

[void]$headerGrid.Controls.Add($searchWrap, 0, 0)
[void]$headerGrid.Controls.Add($filterResetButton, 1, 0)
[void]$headerActions.Controls.Add($tagButton)
[void]$headerActions.Controls.Add($filterButton)
[void]$headerActions.Controls.Add($newHostButton)
[void]$headerActions.Controls.Add($refreshButton)
[void]$headerGrid.Controls.Add($headerActions, 0, 1)
$headerGrid.SetColumnSpan($headerActions, 2)
$topPanel.Controls.Add($headerGrid)

$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Dock = "Fill"
$mainPanel.BackColor = $BgMain

$tabNavPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$tabNavPanel.Dock = "Top"
$tabNavPanel.Height = 34
$tabNavPanel.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 0)
$tabNavPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$tabNavPanel.WrapContents = $false
$tabNavPanel.FlowDirection = "LeftToRight"
$tabNavPanel.BackColor = $BgMain

$hostsTabButton = New-Object System.Windows.Forms.Button
$hostsTabButton.Text = "Hosts"
$hostsTabButton.Width = 90
$hostsTabButton.Height = 28
$hostsTabButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
$hostsTabButton.FlatStyle = "Flat"
$hostsTabButton.FlatAppearance.BorderSize = 0
$hostsTabButton.BackColor = $BgCardHover
$hostsTabButton.ForeColor = $FgMain
$hostsTabButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$hostsTabButton.Cursor = [System.Windows.Forms.Cursors]::Hand

$infoTabButton = New-Object System.Windows.Forms.Button
$infoTabButton.Text = "Info"
$infoTabButton.Width = 90
$infoTabButton.Height = 28
$infoTabButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
$infoTabButton.FlatStyle = "Flat"
$infoTabButton.FlatAppearance.BorderSize = 0
$infoTabButton.BackColor = $BgPanel
$infoTabButton.ForeColor = $FgMuted
$infoTabButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$infoTabButton.Cursor = [System.Windows.Forms.Cursors]::Hand

$settingsTabButton = New-Object System.Windows.Forms.Button
$settingsTabButton.Text = "Settings"
$settingsTabButton.Width = 110
$settingsTabButton.Height = 28
$settingsTabButton.Margin = New-Object System.Windows.Forms.Padding(0)
$settingsTabButton.FlatStyle = "Flat"
$settingsTabButton.FlatAppearance.BorderSize = 0
$settingsTabButton.BackColor = $BgPanel
$settingsTabButton.ForeColor = $FgMuted
$settingsTabButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$settingsTabButton.Cursor = [System.Windows.Forms.Cursors]::Hand

[void]$tabNavPanel.Controls.Add($hostsTabButton)
[void]$tabNavPanel.Controls.Add($settingsTabButton)
[void]$tabNavPanel.Controls.Add($infoTabButton)

$hostPanel = New-Object System.Windows.Forms.Panel
$hostPanel.Dock = "Fill"
$hostPanel.AutoScroll = $true
$hostPanel.Padding = New-Object System.Windows.Forms.Padding(6)
$hostPanel.BackColor = $BgMain

$infoPanel = New-Object System.Windows.Forms.Panel
$infoPanel.Dock = "Fill"
$infoPanel.Padding = New-Object System.Windows.Forms.Padding(24)
$infoPanel.BackColor = $BgMain
$infoPanel.Visible = $false

$infoTitle = New-Object System.Windows.Forms.Label
$infoTitle.AutoSize = $true
$infoTitle.Location = New-Object System.Drawing.Point(0, 0)
$infoTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 16)
$infoTitle.ForeColor = $FgMain
$infoTitle.Text = $AppName

$infoVersion = New-Object System.Windows.Forms.Label
$infoVersion.AutoSize = $true
$infoVersion.Location = New-Object System.Drawing.Point(2, 46)
$infoVersion.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$infoVersion.ForeColor = $FgMuted
$infoVersion.Text = "Version: v$AppVersion"

$infoAuthor = New-Object System.Windows.Forms.Label
$infoAuthor.AutoSize = $true
$infoAuthor.Location = New-Object System.Drawing.Point(2, 74)
$infoAuthor.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$infoAuthor.ForeColor = $FgMuted
$infoAuthor.Text = "Author: $AppAuthor"

$infoRepoLabel = New-Object System.Windows.Forms.Label
$infoRepoLabel.AutoSize = $true
$infoRepoLabel.Location = New-Object System.Drawing.Point(2, 102)
$infoRepoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$infoRepoLabel.ForeColor = $FgMuted
$infoRepoLabel.Text = "Repo:"

$infoRepoLink = New-Object System.Windows.Forms.LinkLabel
$infoRepoLink.AutoSize = $false
$infoRepoLink.Location = New-Object System.Drawing.Point(2, 126)
$infoRepoLink.Size = New-Object System.Drawing.Size(520, 24)
$infoRepoLink.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$infoRepoLink.LinkColor = $Accent
$infoRepoLink.ActiveLinkColor = $FgMain
$infoRepoLink.VisitedLinkColor = $Accent
$infoRepoLink.Text = $GitHubRepoUrl
$infoRepoLink.Add_LinkClicked({ Start-Process $GitHubRepoUrl })

$settingsPanel = New-Object System.Windows.Forms.Panel
$settingsPanel.Dock = "Fill"
$settingsPanel.Padding = New-Object System.Windows.Forms.Padding(24)
$settingsPanel.BackColor = $BgMain
$settingsPanel.Visible = $false

$settingsTitle = New-Object System.Windows.Forms.Label
$settingsTitle.AutoSize = $true
$settingsTitle.Location = New-Object System.Drawing.Point(0, 0)
$settingsTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 16)
$settingsTitle.ForeColor = $FgMain

$languageLabel = New-Object System.Windows.Forms.Label
$languageLabel.AutoSize = $true
$languageLabel.Location = New-Object System.Drawing.Point(2, 48)
$languageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$languageLabel.ForeColor = $FgMuted

$languageCombo = New-Object System.Windows.Forms.ComboBox
$languageCombo.Location = New-Object System.Drawing.Point(180, 44)
$languageCombo.Size = New-Object System.Drawing.Size(180, 28)
$languageCombo.DropDownStyle = "DropDownList"
$languageCombo.BackColor = $BgInput
$languageCombo.ForeColor = $FgMain
$languageCombo.FlatStyle = "Popup"
$null = $languageCombo.Items.Add([PSCustomObject]@{ Text = ""; Value = "de" })
$null = $languageCombo.Items.Add([PSCustomObject]@{ Text = ""; Value = "en" })
$languageCombo.DisplayMember = "Text"
$languageCombo.ValueMember = "Value"

$sortLabel = New-Object System.Windows.Forms.Label
$sortLabel.AutoSize = $true
$sortLabel.Location = New-Object System.Drawing.Point(2, 84)
$sortLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$sortLabel.ForeColor = $FgMuted

$sortCombo = New-Object System.Windows.Forms.ComboBox
$sortCombo.Location = New-Object System.Drawing.Point(180, 80)
$sortCombo.Size = New-Object System.Drawing.Size(260, 28)
$sortCombo.DropDownStyle = "DropDownList"
$sortCombo.BackColor = $BgInput
$sortCombo.ForeColor = $FgMain
$sortCombo.FlatStyle = "Popup"
$null = $sortCombo.Items.Add([PSCustomObject]@{ Text = ""; Value = "host" })
$null = $sortCombo.Items.Add([PSCustomObject]@{ Text = ""; Value = "tag" })
$sortCombo.DisplayMember = "Text"
$sortCombo.ValueMember = "Value"

$languageRestartNotice = New-Object System.Windows.Forms.Label
$languageRestartNotice.AutoSize = $false
$languageRestartNotice.Location = New-Object System.Drawing.Point(180, 112)
$languageRestartNotice.Size = New-Object System.Drawing.Size(360, 38)
$languageRestartNotice.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$languageRestartNotice.ForeColor = $FgMuted

$settingsUpdateTitle = New-Object System.Windows.Forms.Label
$settingsUpdateTitle.AutoSize = $true
$settingsUpdateTitle.Location = New-Object System.Drawing.Point(2, 176)
$settingsUpdateTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
$settingsUpdateTitle.ForeColor = $FgMain

$infoUpdateStatus = New-Object System.Windows.Forms.Label
$infoUpdateStatus.AutoSize = $true
$infoUpdateStatus.Location = New-Object System.Drawing.Point(2, 208)
$infoUpdateStatus.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$infoUpdateStatus.ForeColor = $FgMuted
$infoUpdateStatus.Text = "Version"

$settingsActions = New-Object System.Windows.Forms.FlowLayoutPanel
$settingsActions.AutoSize = $true
$settingsActions.AutoSizeMode = "GrowAndShrink"
$settingsActions.WrapContents = $true
$settingsActions.FlowDirection = "LeftToRight"
$settingsActions.Location = New-Object System.Drawing.Point(0, 242)
$settingsActions.BackColor = $BgMain

$checkUpdateButton = New-UiButton -Text "Check"
$checkUpdateButton.Width = 130
$checkUpdateButton.Dock = "None"

$updateButton = New-UiButton -Text "Update"
$updateButton.Width = 110
$updateButton.Dock = "None"

[void]$settingsActions.Controls.Add($checkUpdateButton)
[void]$settingsActions.Controls.Add($updateButton)

$infoPanel.Controls.Add($infoTitle)
$infoPanel.Controls.Add($infoVersion)
$infoPanel.Controls.Add($infoAuthor)
$infoPanel.Controls.Add($infoRepoLabel)
$infoPanel.Controls.Add($infoRepoLink)
$settingsPanel.Controls.Add($settingsTitle)
$settingsPanel.Controls.Add($languageLabel)
$settingsPanel.Controls.Add($languageCombo)
$settingsPanel.Controls.Add($sortLabel)
$settingsPanel.Controls.Add($sortCombo)
$settingsPanel.Controls.Add($languageRestartNotice)
$settingsPanel.Controls.Add($settingsUpdateTitle)
$settingsPanel.Controls.Add($infoUpdateStatus)
$settingsPanel.Controls.Add($settingsActions)
$mainPanel.Controls.Add($hostPanel)
$mainPanel.Controls.Add($infoPanel)
$mainPanel.Controls.Add($settingsPanel)
$mainPanel.Controls.Add($tabNavPanel)

$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = $BgPanel
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = T "StatusLoadingHosts"
$statusLabel.ForeColor = $FgMuted
$versionStatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$versionStatusLabel.Alignment = "Right"
$versionStatusLabel.Spring = $true
$versionStatusLabel.TextAlign = "MiddleRight"
$versionStatusLabel.Text = T "VersionOnly" @($AppVersion)
$versionStatusLabel.ForeColor = $FgMuted
$statusBar.Items.Add($statusLabel) | Out-Null
$statusBar.Items.Add($versionStatusLabel) | Out-Null

$form.Controls.Add($mainPanel)
$form.Controls.Add($topPanel)
$form.Controls.Add($statusBar)

function Update-FilterSummary {
    $include = @($script:HostMeta.IncludeTags)
    $exclude = @($script:HostMeta.ExcludeTags)
    $search = $searchBox.Text.Trim()

    if ($include.Count -eq 0 -and $exclude.Count -eq 0 -and [string]::IsNullOrWhiteSpace($search)) {
        $filterResetButton.Text = T "FilterAll"
        return
    }

    $parts = @()
    if (-not [string]::IsNullOrWhiteSpace($search)) {
        $parts += "$(T "SearchPrefix"): $search"
    }
    if ($include.Count -gt 0) {
        $parts += "+" + ($include -join ", +")
    }
    if ($exclude.Count -gt 0) {
        $parts += "-" + ($exclude -join ", -")
    }
    $summary = $parts -join " | "
    $filterResetButton.Text = "$(T "FilterPrefix"): $(Get-ShortUiText -Text $summary -MaxLength 24)"
}

function Update-Language {
    $form.Text = "$AppName v$AppVersion"

    $tagButton.Text = T "Tags"
    $filterButton.Text = T "Filter"
    $newHostButton.Text = T "NewHost"
    $refreshButton.Text = T "Refresh"
    $hostsTabButton.Text = T "Hosts"
    $infoTabButton.Text = T "Info"
    $settingsTabButton.Text = T "Settings"

    $infoTitle.Text = $AppName
    $infoVersion.Text = "{0}: v{1}" -f (T "VersionLabel"), $AppVersion
    $infoAuthor.Text = "{0}: {1}" -f (T "AuthorLabel"), $AppAuthor
    $infoRepoLabel.Text = "{0}:" -f (T "RepoLabel")

    $settingsTitle.Text = T "SettingsTitle"
    $languageLabel.Text = T "LanguageLabel"
    $sortLabel.Text = T "SortLabel"
    $languageRestartNotice.Text = T "RestartRequiredNote"
    $settingsUpdateTitle.Text = T "UpdateSection"
    $checkUpdateButton.Text = T "CheckUpdate"
    $updateButton.Text = T "Update"

    foreach ($item in $languageCombo.Items) {
        if ($item.Value -eq "de") { $item.Text = T "LanguageGerman" }
        if ($item.Value -eq "en") { $item.Text = T "LanguageEnglish" }
    }
    foreach ($item in $sortCombo.Items) {
        if ($item.Value -eq "host") { $item.Text = T "SortByHost" }
        if ($item.Value -eq "tag") { $item.Text = T "SortByTag" }
    }

    $updateReason = if ($null -ne $script:LastUpdateStatus) { Format-UpdateStatusReason -UpdateStatus $script:LastUpdateStatus } else { T "VersionOnly" @($AppVersion) }
    $infoUpdateStatus.Text = $updateReason
    $versionStatusLabel.Text = $updateReason
    Update-FilterSummary
    if ($script:AllHosts.Count -gt 0) {
        Update-HostFilter
    }
}

function Show-HostButtons {
    param([array]$Hosts)

    $hostPanel.SuspendLayout()
    $hostPanel.Controls.Clear()
    $maxBottom = $hostPanel.Padding.Top
    $gap = 8
    $buttonHeight = 42
    $availableWidth = [math]::Max(200, $hostPanel.ClientSize.Width - $hostPanel.Padding.Left - $hostPanel.Padding.Right)

    if ((Get-CurrentSortMode) -eq "tag") {
        $groupMap = [ordered]@{}
        foreach ($entry in $Hosts) {
            $entryTags = @(Get-HostTags -HostName $entry.Host | Sort-Object)
            if ($entryTags.Count -eq 0) {
                $groupName = T "Untagged"
                if (-not $groupMap.Contains($groupName)) {
                    $groupMap[$groupName] = New-Object System.Collections.Generic.List[object]
                }
                $groupMap[$groupName].Add($entry) | Out-Null
                continue
            }

            foreach ($tagName in $entryTags) {
                if (-not $groupMap.Contains($tagName)) {
                    $groupMap[$tagName] = New-Object System.Collections.Generic.List[object]
                }
                $groupMap[$tagName].Add($entry) | Out-Null
            }
        }

        $y = $hostPanel.Padding.Top
        foreach ($groupName in ($groupMap.Keys | Sort-Object)) {
            $header = New-Object System.Windows.Forms.Label
            $header.Text = $groupName
            $header.Location = New-Object System.Drawing.Point($hostPanel.Padding.Left, $y)
            $header.Size = New-Object System.Drawing.Size($availableWidth, 24)
            $header.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
            $header.ForeColor = $FgMain
            $hostPanel.Controls.Add($header)
            $y += 24

            $separator = New-Object System.Windows.Forms.Panel
            $separator.Location = New-Object System.Drawing.Point($hostPanel.Padding.Left, $y)
            $separator.Size = New-Object System.Drawing.Size([math]::Min(220, $availableWidth), 1)
            $separator.BackColor = $Border
            $hostPanel.Controls.Add($separator)
            $y += 8

            foreach ($entry in ($groupMap[$groupName] | Sort-Object Host -Unique)) {
                $button = New-Object System.Windows.Forms.Button
                $button.Location = New-Object System.Drawing.Point($hostPanel.Padding.Left, $y)
                $button.Size = New-Object System.Drawing.Size($availableWidth, $buttonHeight)
                $button.FlatStyle = "Flat"
                $button.FlatAppearance.BorderColor = $Border
                $button.FlatAppearance.MouseOverBackColor = $BgCardHover
                $button.BackColor = $BgCard
                $button.ForeColor = $FgMain
                $button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
                $button.TextAlign = "MiddleLeft"
                $button.Padding = New-Object System.Windows.Forms.Padding(12, 0, 12, 0)
                $button.Tag = $entry.Host
                $button.Text = $entry.Host
                $button.Cursor = [System.Windows.Forms.Cursors]::Hand
                $tags = @(Get-HostTags -HostName $entry.Host)
                $tooltipSourceLabel = if ((Get-UiLanguage) -eq "en") { "Source" } else { "Quelle" }
                $tooltipTagsLabel = "Tags"
                $tooltipHostLabel = "Host"
                $script:HostToolTip.SetToolTip($button, "${tooltipHostLabel}: $($entry.Host)`r`n${tooltipTagsLabel}: $($tags -join ', ')`r`n${tooltipSourceLabel}: $($entry.SourceFile)")
                $button.Add_Click({
                    param($controlSender, $clickEvent)
                    Start-SshHost -HostName $controlSender.Tag
                })
                $hostPanel.Controls.Add($button)
                $y += $buttonHeight + 6
            }

            $y += 14
            $maxBottom = [math]::Max($maxBottom, $y)
        }
    }
    else {
        $sortedHosts = @(Get-SortedHosts -Hosts $Hosts)
        $minCardWidth = 130
        $columns = [math]::Max(1, [math]::Floor(($availableWidth + $gap) / ($minCardWidth + $gap)))
        $cardWidth = [math]::Floor(($availableWidth - (($columns - 1) * $gap)) / $columns)
        if ($cardWidth -lt $minCardWidth) { $cardWidth = $minCardWidth }

        $index = 0
        foreach ($entry in $sortedHosts) {
            $button = New-Object System.Windows.Forms.Button
            $button.Width = $cardWidth
            $button.Height = 46
            $button.FlatStyle = "Flat"
            $button.FlatAppearance.BorderColor = $Border
            $button.FlatAppearance.MouseOverBackColor = $BgCardHover
            $button.BackColor = $BgCard
            $button.ForeColor = $FgMain
            $button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
            $button.TextAlign = "MiddleCenter"
            $button.Tag = $entry.Host
            $button.Text = $entry.Host
            $button.Cursor = [System.Windows.Forms.Cursors]::Hand
            $tags = @(Get-HostTags -HostName $entry.Host)
            $tooltipSourceLabel = if ((Get-UiLanguage) -eq "en") { "Source" } else { "Quelle" }
            $tooltipTagsLabel = "Tags"
            $tooltipHostLabel = "Host"
            $script:HostToolTip.SetToolTip($button, "${tooltipHostLabel}: $($entry.Host)`r`n${tooltipTagsLabel}: $($tags -join ', ')`r`n${tooltipSourceLabel}: $($entry.SourceFile)")
            $button.Add_Click({
                param($controlSender, $clickEvent)
                Start-SshHost -HostName $controlSender.Tag
            })

            $columnIndex = $index % $columns
            $rowIndex = [math]::Floor($index / $columns)
            $x = $hostPanel.Padding.Left + ($columnIndex * ($cardWidth + $gap))
            $y = $hostPanel.Padding.Top + ($rowIndex * ($button.Height + $gap))
            $button.Location = New-Object System.Drawing.Point($x, $y)
            $index++
            $maxBottom = [math]::Max($maxBottom, $y + $button.Height)
            $hostPanel.Controls.Add($button)
        }
    }

    $hostPanel.AutoScrollMinSize = [System.Drawing.Size]::new(0, ($maxBottom + $hostPanel.Padding.Bottom))
    $hostPanel.ResumeLayout()
}

function Update-Hosts {
    if (-not (Test-Path $SshConfigPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            (T "MissingConfig" @($SshConfigPath)),
            (T "MissingConfigTitle"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        $script:AllHosts = @()
        Show-HostButtons -Hosts @()
        return
    }

    try {
        $script:AllHosts = @(Get-SshHostsFromConfig -ConfigPath $SshConfigPath)
        Update-FilterSummary
        Update-HostFilter
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            (T "ReadError" @($_.Exception.Message)),
            (T "ReadErrorTitle"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

function Update-HostFilter {
    $filter = $searchBox.Text.Trim()
    $includeTags = @($script:HostMeta.IncludeTags)
    $excludeTags = @($script:HostMeta.ExcludeTags)

    $filtered = $script:AllHosts | Where-Object {
        $matchSearch = [string]::IsNullOrWhiteSpace($filter) -or $_.Host -like "*$filter*"
        if (-not $matchSearch) { return $false }
        $hostTags = @(Get-HostTags -HostName $_.Host)
        $matchesInclude = ($includeTags.Count -eq 0) -or (@($hostTags | Where-Object { $includeTags -contains $_ }).Count -gt 0)
        $matchesExclude = (@($hostTags | Where-Object { $excludeTags -contains $_ }).Count -eq 0)
        return ($matchesInclude -and $matchesExclude)
    }

    Show-HostButtons -Hosts $filtered
    $statusLabel.Text = T "StatusHosts" @($filtered.Count, $script:AllHosts.Count)
}

function Test-AppVersion {
    $checkUpdateButton.Enabled = $false
    $infoUpdateStatus.Text = T "UpdateChecking"
    $versionStatusLabel.Text = T "UpdateChecking"
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $script:LastUpdateStatus = Get-UpdateStatus
        $updateReason = Format-UpdateStatusReason -UpdateStatus $script:LastUpdateStatus
        $infoUpdateStatus.Text = $updateReason
        $versionStatusLabel.Text = $updateReason
    }
    finally {
        $checkUpdateButton.Enabled = $true
    }
}

function Set-ActiveMainView {
    param([Parameter(Mandatory = $true)][string]$ViewName)

    $showHosts = $ViewName -eq "Hosts"
    $showInfo = $ViewName -eq "Info"
    $showSettings = $ViewName -eq "Settings"
    $hostPanel.Visible = $showHosts
    $infoPanel.Visible = $showInfo
    $settingsPanel.Visible = $showSettings

    $hostsTabButton.BackColor = if ($showHosts) { $BgCardHover } else { $BgPanel }
    $hostsTabButton.ForeColor = if ($showHosts) { $FgMain } else { $FgMuted }
    $infoTabButton.BackColor = if ($showInfo) { $BgCardHover } else { $BgPanel }
    $infoTabButton.ForeColor = if ($showInfo) { $FgMain } else { $FgMuted }
    $settingsTabButton.BackColor = if ($showSettings) { $BgCardHover } else { $BgPanel }
    $settingsTabButton.ForeColor = if ($showSettings) { $FgMain } else { $FgMuted }
}

$refreshButton.Add_Click({ Update-Hosts })
$searchBox.Add_TextChanged({ Update-HostFilter })
$hostsTabButton.Add_Click({ Set-ActiveMainView -ViewName "Hosts" })
$infoTabButton.Add_Click({ Set-ActiveMainView -ViewName "Info" })
$settingsTabButton.Add_Click({ Set-ActiveMainView -ViewName "Settings" })
$filterResetButton.Add_Click({
    $searchBox.Text = ""
    $script:HostMeta.IncludeTags = @()
    $script:HostMeta.ExcludeTags = @()
    Export-HostMeta -MetaPath $HostMetaPath -Meta $script:HostMeta
    Update-FilterSummary
    Update-HostFilter
})

$tagButton.Add_Click({
    $changedHost = Set-HostTagDialog -Hosts $script:AllHosts
    if (-not [string]::IsNullOrWhiteSpace($changedHost)) {
        Update-FilterSummary
        Update-HostFilter
    }
})

$filterButton.Add_Click({
    if (Edit-TagFilterDialog) {
        Update-FilterSummary
        Update-HostFilter
    }
})

$newHostButton.Add_Click({
    $createdAlias = New-HostEntry -ConfigPath $SshConfigPath
    if (-not [string]::IsNullOrWhiteSpace($createdAlias)) {
        $searchBox.Text = ""
        Update-Hosts
    }
})

$checkUpdateButton.Add_Click({ Test-AppVersion })
$updateButton.Add_Click({ Update-AppFromGitHub })
$languageCombo.Add_SelectedIndexChanged({
    if ($null -eq $languageCombo.SelectedItem) { return }
    $script:UiState.Language = [string]$languageCombo.SelectedItem.Value
    Update-Language
})
$sortCombo.Add_SelectedIndexChanged({
    if ($null -eq $sortCombo.SelectedItem) { return }
    $script:UiState.SortMode = [string]$sortCombo.SelectedItem.Value
    if ($script:AllHosts.Count -gt 0) {
        Update-HostFilter
    }
})
$hostPanel.Add_SizeChanged({
    if ($script:AllHosts.Count -gt 0) { Update-HostFilter }
})

$form.Add_FormClosing({
    $script:UiState.Width = $form.Width
    $script:UiState.Height = $form.Height
    $script:UiState.Left = $form.Left
    $script:UiState.Top = $form.Top
    Export-UiState -UiPath $UiStatePath -State $script:UiState
})

$form.Add_Shown({
    foreach ($item in $languageCombo.Items) {
        if ($item.Value -eq $script:UiState.Language) {
            $languageCombo.SelectedItem = $item
            break
        }
    }
    if ($null -eq $languageCombo.SelectedItem -and $languageCombo.Items.Count -gt 0) {
        $languageCombo.SelectedIndex = 0
    }
    foreach ($item in $sortCombo.Items) {
        if ($item.Value -eq $script:UiState.SortMode) {
            $sortCombo.SelectedItem = $item
            break
        }
    }
    if ($null -eq $sortCombo.SelectedItem -and $sortCombo.Items.Count -gt 0) {
        $sortCombo.SelectedIndex = 0
    }
    Update-Language
    Set-ActiveMainView -ViewName "Hosts"
    Update-FilterSummary
    Update-Hosts
    Test-AppVersion
})

[void]$form.ShowDialog()

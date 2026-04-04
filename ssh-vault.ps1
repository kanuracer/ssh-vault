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
$AppVersion = "0.8.4"
$GitHubRepo = "kanuracer/ssh-vault"
$GitHubRepoUrl = "https://github.com/$GitHubRepo"
$GitHubBranch = "main"
$WindowTitle = "$AppName v$AppVersion"

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
        LastTagFilter = "Alle"
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
            LastTagFilter = $meta.LastTagFilter
            HostTags      = [ordered]@{}
            KnownTags     = @()
        } | ConvertTo-Json -Depth 5)
        Set-Content -Path $MetaPath -Value $json -Encoding UTF8
        return $meta
    }

    try {
        $raw = Get-Content -Path $MetaPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($null -ne $raw.LastTagFilter -and -not [string]::IsNullOrWhiteSpace([string]$raw.LastTagFilter)) {
            $meta.LastTagFilter = [string]$raw.LastTagFilter
        }
        if ($null -ne $raw.HostTags) {
            foreach ($p in $raw.HostTags.PSObject.Properties) {
                $meta.HostTags[[string]$p.Name] = [string]$p.Value
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
        LastTagFilter = $Meta.LastTagFilter
        HostTags = $hostTags
        KnownTags = @($Meta.KnownTags | Sort-Object -Unique)
    }

    $json = $payload | ConvertTo-Json -Depth 5
    Set-Content -Path $MetaPath -Value $json -Encoding UTF8
}

function Get-HostTag {
    param([Parameter(Mandatory = $true)][string]$HostName)
    $key = $HostName.ToLowerInvariant()
    if ($script:HostMeta.HostTags.ContainsKey($key)) {
        return $script:HostMeta.HostTags[$key]
    }
    return ""
}

function Set-HostTag {
    param(
        [Parameter(Mandatory = $true)][string]$HostName,
        [string]$TagName
    )

    $key = $HostName.ToLowerInvariant()
    $trimmed = $TagName.Trim()

    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        $null = $script:HostMeta.HostTags.Remove($key)
    }
    else {
        $script:HostMeta.HostTags[$key] = $trimmed
        $script:HostMeta.KnownTags = @($script:HostMeta.KnownTags + $trimmed | Sort-Object -Unique)
    }

    Export-HostMeta -MetaPath $HostMetaPath -Meta $script:HostMeta
}

function Get-AllTags {
    $all = @()
    $all += @($script:HostMeta.KnownTags)
    $all += @($script:HostMeta.HostTags.Values)
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
    $dialog.Text = "Neuen SSH Host anlegen"
    $dialog.Size = New-Object System.Drawing.Size(460, 300)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.BackColor = $BgPanel
    $dialog.ForeColor = $FgMain

    $labelAlias = New-Object System.Windows.Forms.Label
    $labelAlias.Text = "Alias (Host):"
    $labelAlias.Location = New-Object System.Drawing.Point(20, 20)
    $labelAlias.AutoSize = $true

    $txtAlias = New-Object System.Windows.Forms.TextBox
    $txtAlias.Location = New-Object System.Drawing.Point(140, 16)
    $txtAlias.Size = New-Object System.Drawing.Size(280, 24)
    $txtAlias.BackColor = $BgInput
    $txtAlias.ForeColor = $FgMain
    $txtAlias.BorderStyle = "FixedSingle"

    $labelHostName = New-Object System.Windows.Forms.Label
    $labelHostName.Text = "HostName:"
    $labelHostName.Location = New-Object System.Drawing.Point(20, 60)
    $labelHostName.AutoSize = $true

    $txtHostName = New-Object System.Windows.Forms.TextBox
    $txtHostName.Location = New-Object System.Drawing.Point(140, 56)
    $txtHostName.Size = New-Object System.Drawing.Size(280, 24)
    $txtHostName.BackColor = $BgInput
    $txtHostName.ForeColor = $FgMain
    $txtHostName.BorderStyle = "FixedSingle"

    $labelUser = New-Object System.Windows.Forms.Label
    $labelUser.Text = "User (optional):"
    $labelUser.Location = New-Object System.Drawing.Point(20, 100)
    $labelUser.AutoSize = $true

    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(140, 96)
    $txtUser.Size = New-Object System.Drawing.Size(280, 24)
    $txtUser.BackColor = $BgInput
    $txtUser.ForeColor = $FgMain
    $txtUser.BorderStyle = "FixedSingle"

    $labelPort = New-Object System.Windows.Forms.Label
    $labelPort.Text = "Port (optional):"
    $labelPort.Location = New-Object System.Drawing.Point(20, 140)
    $labelPort.AutoSize = $true

    $txtPort = New-Object System.Windows.Forms.TextBox
    $txtPort.Location = New-Object System.Drawing.Point(140, 136)
    $txtPort.Size = New-Object System.Drawing.Size(280, 24)
    $txtPort.BackColor = $BgInput
    $txtPort.ForeColor = $FgMain
    $txtPort.BorderStyle = "FixedSingle"

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Speichern"
    $btnSave.Location = New-Object System.Drawing.Point(260, 190)
    $btnSave.Size = New-Object System.Drawing.Size(160, 34)
    $btnSave.BackColor = $Accent
    $btnSave.ForeColor = $FgMain
    $btnSave.FlatStyle = "Flat"
    $btnSave.FlatAppearance.BorderSize = 0

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
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
                    "Alias und HostName sind Pflichtfelder.",
                    "Eingabe unvollstaendig",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }

            if ($alias -match '\s') {
                [System.Windows.Forms.MessageBox]::Show(
                    "Der Alias darf keine Leerzeichen enthalten.",
                    "Ungueltiger Alias",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }

            if (-not [string]::IsNullOrWhiteSpace($port) -and $port -notmatch '^\d+$') {
                [System.Windows.Forms.MessageBox]::Show(
                    "Port muss numerisch sein.",
                    "Ungueltiger Port",
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
                    "Host konnte nicht gespeichert werden.`n`n$($_.Exception.Message)",
                    "Speicherfehler",
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
            "Keine Hosts verfuegbar.",
            "Hinweis",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return $null
    }

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Tag setzen"
    $dialog.Size = New-Object System.Drawing.Size(520, 240)
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
    $dialogLayout.RowCount = 4
    $dialogLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110))) | Out-Null
    $dialogLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $dialogLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
    $dialogLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
    $dialogLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $dialogLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42))) | Out-Null

    $labelHost = New-Object System.Windows.Forms.Label
    $labelHost.Text = "Host"
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
    $labelTag.Text = "Tag/Kategorie"
    $labelTag.Dock = "Fill"
    $labelTag.TextAlign = "MiddleLeft"
    $labelTag.AutoSize = $true

    $comboTag = New-Object System.Windows.Forms.ComboBox
    $comboTag.Dock = "Fill"
    $comboTag.DropDownStyle = "DropDown"
    $comboTag.BackColor = $BgInput
    $comboTag.ForeColor = $FgMain
    $comboTag.FlatStyle = "Popup"

    $info = New-Object System.Windows.Forms.Label
    $info.Text = "Bestehende Tags auswaehlen oder neuen Tag eintippen. Leer = Tag entfernen."
    $info.Dock = "Fill"
    $info.TextAlign = "TopLeft"
    $info.ForeColor = $FgMuted

    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock = "Fill"
    $buttonPanel.FlowDirection = "RightToLeft"
    $buttonPanel.WrapContents = $false

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Speichern"
    $btnSave.Width = 110
    $btnSave.Height = 30
    $btnSave.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
    $btnSave.BackColor = $Accent
    $btnSave.ForeColor = $FgMain
    $btnSave.FlatStyle = "Flat"
    $btnSave.FlatAppearance.BorderSize = 0

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
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
            [void]$comboTag.Items.Add($tag)
        }
        $comboTag.Text = Get-HostTag -HostName ([string]$comboHost.SelectedItem)
    }

    $comboHost.Add_SelectedIndexChanged({
        $selected = [string]$comboHost.SelectedItem
        $comboTag.Text = Get-HostTag -HostName $selected
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
        Set-HostTag -HostName $selected -TagName $comboTag.Text
        $script:TaggedHostResult = $selected
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    })

    [void]$dialogLayout.Controls.Add($labelHost, 0, 0)
    [void]$dialogLayout.Controls.Add($comboHost, 1, 0)
    [void]$dialogLayout.Controls.Add($labelTag, 0, 1)
    [void]$dialogLayout.Controls.Add($comboTag, 1, 1)
    [void]$dialogLayout.Controls.Add($info, 1, 2)
    [void]$dialogLayout.Controls.Add($buttonPanel, 0, 3)
    $dialogLayout.SetColumnSpan($buttonPanel, 2)
    $dialog.Controls.Add($dialogLayout)
    $dialog.AcceptButton = $btnSave
    $dialog.CancelButton = $btnCancel

    [void]$dialog.ShowDialog()
    return $script:TaggedHostResult
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
        Width = 820
        Height = 560
        Left = -1
        Top = -1
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
    }
    catch {}
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
        return [PSCustomObject]@{
            IsAvailable = $false
            Reason = "Version v$AppVersion"
            Release = $null
            RemoteVersion = $remoteVersion
            LocalVersion = $localVersion
            RemoteVersionText = $remoteVersionText
        }
    }

    return [PSCustomObject]@{
        IsAvailable = ($remoteVersion -gt $localVersion)
        Reason = if ($remoteVersion -gt $localVersion) { "Version v$AppVersion | Update verfuegbar: v$remoteVersionText" } else { "Version v$AppVersion" }
        Release = $null
        RemoteVersion = $remoteVersion
        LocalVersion = $localVersion
        RemoteVersionText = $remoteVersionText
    }
}

function Get-CurrentScriptPath {
    if ($PSCommandPath) { return $PSCommandPath }
    if ($MyInvocation.MyCommand.Path) { return $MyInvocation.MyCommand.Path }
    return (Join-Path $AppRoot "ssh-vault.ps1")
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
                "Kein Update installiert",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            $versionStatusLabel.Text = $updateStatus.Reason
            return
        }

        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Soll SSH Vault wirklich von v$AppVersion auf v$($updateStatus.RemoteVersionText) aktualisiert werden?",
            "Update bestaetigen",
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
            "Update erfolgreich installiert.`nBackup: $backupPath`nBitte App neu starten.",
            "Update erfolgreich",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Update fehlgeschlagen.`n`n$($_.Exception.Message)",
            "Updatefehler",
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

$form = New-Object System.Windows.Forms.Form
$form.Text = $WindowTitle
$form.Size = New-Object System.Drawing.Size($script:UiState.Width, $script:UiState.Height)
$form.StartPosition = "CenterScreen"
if ($script:UiState.Left -ge 0 -and $script:UiState.Top -ge 0) {
    $form.StartPosition = "Manual"
    $form.Location = New-Object System.Drawing.Point($script:UiState.Left, $script:UiState.Top)
}
$form.MinimumSize = New-Object System.Drawing.Size(620, 400)
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

$tagFilterCombo = New-Object System.Windows.Forms.ComboBox
$tagFilterCombo.Dock = "Fill"
$tagFilterCombo.Margin = New-Object System.Windows.Forms.Padding(3)
$tagFilterCombo.DropDownStyle = "DropDownList"
$tagFilterCombo.BackColor = $BgInput
$tagFilterCombo.ForeColor = $FgMain
$tagFilterCombo.FlatStyle = "Flat"

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
$newHostButton = New-UiButton -Text "+ Host" -Primary
$newHostButton.Width = 100
$newHostButton.Dock = "None"
$refreshButton = New-UiButton -Text "Refresh"
$refreshButton.Width = 90
$refreshButton.Dock = "None"
$checkUpdateButton = New-UiButton -Text "Check"
$checkUpdateButton.Width = 80
$checkUpdateButton.Dock = "None"
$updateButton = New-UiButton -Text "Update"
$updateButton.Width = 80
$updateButton.Dock = "None"

[void]$headerGrid.Controls.Add($searchWrap, 0, 0)
[void]$headerGrid.Controls.Add($tagFilterCombo, 1, 0)
[void]$headerActions.Controls.Add($tagButton)
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
$infoTabButton.Margin = New-Object System.Windows.Forms.Padding(0)
$infoTabButton.FlatStyle = "Flat"
$infoTabButton.FlatAppearance.BorderSize = 0
$infoTabButton.BackColor = $BgPanel
$infoTabButton.ForeColor = $FgMuted
$infoTabButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$infoTabButton.Cursor = [System.Windows.Forms.Cursors]::Hand

[void]$tabNavPanel.Controls.Add($hostsTabButton)
[void]$tabNavPanel.Controls.Add($infoTabButton)

$hostPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$hostPanel.Dock = "Fill"
$hostPanel.AutoScroll = $true
$hostPanel.WrapContents = $true
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
$infoRepoLink.AutoSize = $true
$infoRepoLink.Location = New-Object System.Drawing.Point(52, 102)
$infoRepoLink.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$infoRepoLink.LinkColor = $Accent
$infoRepoLink.ActiveLinkColor = $FgMain
$infoRepoLink.VisitedLinkColor = $Accent
$infoRepoLink.Text = $GitHubRepoUrl
$infoRepoLink.Add_LinkClicked({ Start-Process $GitHubRepoUrl })

$infoUpdateStatus = New-Object System.Windows.Forms.Label
$infoUpdateStatus.AutoSize = $true
$infoUpdateStatus.Location = New-Object System.Drawing.Point(2, 140)
$infoUpdateStatus.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$infoUpdateStatus.ForeColor = $FgMuted
$infoUpdateStatus.Text = "Version v$AppVersion"

$infoActions = New-Object System.Windows.Forms.FlowLayoutPanel
$infoActions.AutoSize = $true
$infoActions.AutoSizeMode = "GrowAndShrink"
$infoActions.WrapContents = $true
$infoActions.FlowDirection = "LeftToRight"
$infoActions.Location = New-Object System.Drawing.Point(0, 174)
$infoActions.BackColor = $BgMain

$checkUpdateButton.Dock = "None"
$checkUpdateButton.Width = 110
$checkUpdateButton.Text = "Check Update"

$updateButton.Dock = "None"
$updateButton.Width = 110

[void]$infoActions.Controls.Add($checkUpdateButton)
[void]$infoActions.Controls.Add($updateButton)

$infoPanel.Controls.Add($infoTitle)
$infoPanel.Controls.Add($infoVersion)
$infoPanel.Controls.Add($infoAuthor)
$infoPanel.Controls.Add($infoRepoLabel)
$infoPanel.Controls.Add($infoRepoLink)
$infoPanel.Controls.Add($infoUpdateStatus)
$infoPanel.Controls.Add($infoActions)
$mainPanel.Controls.Add($hostPanel)
$mainPanel.Controls.Add($infoPanel)
$mainPanel.Controls.Add($tabNavPanel)

$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = $BgPanel
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Lade Hosts..."
$statusLabel.ForeColor = $FgMuted
$versionStatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$versionStatusLabel.Alignment = "Right"
$versionStatusLabel.Spring = $true
$versionStatusLabel.TextAlign = "MiddleRight"
$versionStatusLabel.Text = "Version v$AppVersion"
$versionStatusLabel.ForeColor = $FgMuted
$statusBar.Items.Add($statusLabel) | Out-Null
$statusBar.Items.Add($versionStatusLabel) | Out-Null

$form.Controls.Add($mainPanel)
$form.Controls.Add($topPanel)
$form.Controls.Add($statusBar)

function Update-TagFilterOptions {
    $current = $script:HostMeta.LastTagFilter
    if ($tagFilterCombo.SelectedItem) {
        $current = [string]$tagFilterCombo.SelectedItem
    }

    $tagFilterCombo.BeginUpdate()
    $tagFilterCombo.Items.Clear()
    [void]$tagFilterCombo.Items.Add("Alle")
    foreach ($tag in (Get-AllTags)) {
        [void]$tagFilterCombo.Items.Add($tag)
    }

    if (-not [string]::IsNullOrWhiteSpace($current) -and $tagFilterCombo.Items.Contains($current)) {
        $tagFilterCombo.SelectedItem = $current
    }
    else {
        $tagFilterCombo.SelectedIndex = 0
    }
    $tagFilterCombo.EndUpdate()
}

function Show-HostButtons {
    param([array]$Hosts)

    $sortedHosts = @($Hosts | Sort-Object Host)
    $availableWidth = [math]::Max(200, $hostPanel.ClientSize.Width - $hostPanel.Padding.Left - $hostPanel.Padding.Right)
    $gap = 4
    $minCardWidth = 130
    $columns = [math]::Max(1, [math]::Floor(($availableWidth + $gap) / ($minCardWidth + $gap)))
    $cardWidth = [math]::Floor(($availableWidth - (($columns - 1) * $gap)) / $columns)
    if ($cardWidth -gt 220) { $cardWidth = 220 }
    if ($cardWidth -lt $minCardWidth) { $cardWidth = $minCardWidth }

    $hostPanel.SuspendLayout()
    $hostPanel.Controls.Clear()
    foreach ($entry in $sortedHosts) {
        $tag = Get-HostTag -HostName $entry.Host
        $button = New-Object System.Windows.Forms.Button
        $button.Width = $cardWidth
        $button.Height = 46
        $button.Margin = New-Object System.Windows.Forms.Padding($gap)
        $button.FlatStyle = "Flat"
        $button.FlatAppearance.BorderColor = $Border
        $button.FlatAppearance.MouseOverBackColor = $BgCardHover
        $button.BackColor = $BgCard
        $button.ForeColor = $FgMain
        $button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
        $button.TextAlign = "MiddleCenter"
        $button.Tag = $entry.Host
        $button.Text = $entry.Host
        if (-not [string]::IsNullOrWhiteSpace($tag)) {
            $button.Text = "$($entry.Host)`r`n[$tag]"
        }
        $button.Cursor = [System.Windows.Forms.Cursors]::Hand
        $script:HostToolTip.SetToolTip($button, "Host: $($entry.Host)`r`nTag: $tag`r`nQuelle: $($entry.SourceFile)")
        $button.Add_Click({
            param($controlSender, $clickEvent)
            Start-SshHost -HostName $controlSender.Tag
        })
        $hostPanel.Controls.Add($button)
    }
    $hostPanel.ResumeLayout()
}

function Update-Hosts {
    if (-not (Test-Path $SshConfigPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Die Datei '$SshConfigPath' wurde nicht gefunden.",
            "SSH-Config fehlt",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        $script:AllHosts = @()
        Show-HostButtons -Hosts @()
        return
    }

    try {
        $script:AllHosts = @(Get-SshHostsFromConfig -ConfigPath $SshConfigPath)
        Update-TagFilterOptions
        Update-HostFilter
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Fehler beim Einlesen der SSH-Config.`n`n$($_.Exception.Message)",
            "Lesefehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

function Update-HostFilter {
    $filter = $searchBox.Text.Trim()
    $selectedTag = "Alle"
    if ($tagFilterCombo.SelectedItem) {
        $selectedTag = [string]$tagFilterCombo.SelectedItem
    }

    $filtered = $script:AllHosts | Where-Object {
        $matchSearch = [string]::IsNullOrWhiteSpace($filter) -or $_.Host -like "*$filter*"
        if (-not $matchSearch) { return $false }
        if ($selectedTag -eq "Alle") { return $true }
        return (Get-HostTag -HostName $_.Host) -eq $selectedTag
    }

    Show-HostButtons -Hosts $filtered
    $statusLabel.Text = "$($filtered.Count) / $($script:AllHosts.Count) Host(s)"
}

function Test-AppVersion {
    $checkUpdateButton.Enabled = $false
    $infoUpdateStatus.Text = "Pruefe Update..."
    $versionStatusLabel.Text = "Pruefe Update..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $updateReason = (Get-UpdateStatus).Reason
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
    $hostPanel.Visible = $showHosts
    $infoPanel.Visible = -not $showHosts

    $hostsTabButton.BackColor = if ($showHosts) { $BgCardHover } else { $BgPanel }
    $hostsTabButton.ForeColor = if ($showHosts) { $FgMain } else { $FgMuted }
    $infoTabButton.BackColor = if ($showHosts) { $BgPanel } else { $BgCardHover }
    $infoTabButton.ForeColor = if ($showHosts) { $FgMuted } else { $FgMain }
}

$refreshButton.Add_Click({ Update-Hosts })
$searchBox.Add_TextChanged({ Update-HostFilter })
$hostsTabButton.Add_Click({ Set-ActiveMainView -ViewName "Hosts" })
$infoTabButton.Add_Click({ Set-ActiveMainView -ViewName "Info" })

$tagFilterCombo.Add_SelectedIndexChanged({
    if (-not $tagFilterCombo.SelectedItem) { return }
    $script:HostMeta.LastTagFilter = [string]$tagFilterCombo.SelectedItem
    Export-HostMeta -MetaPath $HostMetaPath -Meta $script:HostMeta
    Update-HostFilter
})

$tagButton.Add_Click({
    $changedHost = Set-HostTagDialog -Hosts $script:AllHosts
    if (-not [string]::IsNullOrWhiteSpace($changedHost)) {
        Update-TagFilterOptions
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
    Set-ActiveMainView -ViewName "Hosts"
    Update-TagFilterOptions
    Update-Hosts
    Test-AppVersion
})

[void]$form.ShowDialog()

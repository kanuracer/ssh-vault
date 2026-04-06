# SSH Vault

SSH Vault is a compact PowerShell WinForms app for managing and launching SSH hosts from your local OpenSSH config on Windows.

It was built as an SSH shortcut and bookmark tool for an OpenSSH-based workflow with Bitwarden-managed SSH keys.

## Features

- Reads hosts from `%USERPROFILE%\.ssh\config`
- Supports `Include` directives in SSH config files
- Starts SSH sessions via `cmd.exe`
- Search and filter hosts
- Multi-tag support per host
- Include and exclude tag filters
- Optional sorting by:
  - host name
  - tag, then host name
- Grouped host view when tag sorting is enabled
- Host management menu for creating and deleting hosts
- Reusable tags
- Persistent window size and position
- Persistent language and sorting settings
- German and English UI
- Built-in version check and self-update from GitHub

## Requirements

- Windows
- Windows PowerShell 5.1
- OpenSSH client (`ssh.exe`) available in `PATH`

## Usage

Run the app:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\ssh-vault.ps1
```

## SSH Config

SSH Vault reads hosts from:

```text
%USERPROFILE%\.ssh\config
```

Example:

```sshconfig
Host web-prod
    HostName 203.0.113.10
    User root
    Port 22
```

The app also resolves OpenSSH `Include` directives and shows hosts from included config files.

## Host Management

You can manage hosts directly in the UI:

- `Host` opens a menu with:
  - `Neu`
  - `Loeschen`
- `Neu` creates a new `Host` block in the SSH config
- `Loeschen` removes the selected host from the source config file

If a `Host` line contains multiple aliases, SSH Vault removes only the selected alias from that line.

## Tags and Filters

Each host can have multiple tags.

Available functions:

- assign and remove tags
- reuse previously created tags
- filter by included tags
- filter by excluded tags
- combine tag filters with text search

When sorting by tag, the host list is grouped by tag.

Hosts without tags are shown in a separate group.

## Settings

The `Settings` tab contains:

- language selection (`Deutsch`, `English`)
- sorting mode selection
- update check
- self-update from GitHub

Default language is `Deutsch`.

Note: after changing the language, restart the app.

## Info

The `Info` tab contains:

- app name
- author
- repository link
- current version

## Local State

SSH Vault creates these files next to `ssh-vault.ps1` if they do not exist:

- `ssh-host-meta.json`
- `ssh-host-ui.json`

They store:

- host tags
- known tags
- active include/exclude filters
- window size and position
- selected language
- selected sorting mode

## Updates

SSH Vault compares the local `$AppVersion` with the version in:

```text
https://raw.githubusercontent.com/kanuracer/ssh-vault/main/ssh-vault.ps1
```

An update is only offered when the GitHub version is newer than the local version.

Before updating, SSH Vault asks for confirmation and creates a backup:

```text
ssh-vault.ps1.bak
```

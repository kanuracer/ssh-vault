# SSH Vault

SSH Vault is a small PowerShell WinForms launcher for SSH hosts from your local OpenSSH config.
The project originated from the fact that I use Bitwarden as an SSH key storage, and I found it very annoying to type ssh user@ip into the command prompt every time.

<p align="center">
  <img src="https://github.com/kanuracer/ssh-vault/blob/main/media/1.png" width="45%" />
  <img src="https://github.com/kanuracer/ssh-vault/blob/main/media/new-host.png" width="45%" />
</p>

<p align="center">
  <img src="https://github.com/kanuracer/ssh-vault/blob/main/media/tag-managment.png" width="45%" />
</p>

## Features

- Reads hosts from `~/.ssh/config`
- Supports `Include` directives in SSH config files
- One-click SSH connect via `cmd.exe`
- Host search
- Host tags/categories with reusable tag values
- Persistent tag filter
- Persistent window size and position
- Info page with app metadata and repository link
- Built-in update check and self-update from GitHub

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

SSH Vault reads host entries from:

```text
%USERPROFILE%\.ssh\config
```

Example:

```sshconfig
Host my-server
    HostName 192.168.178.102
    User root
    Port 22
```

## Local State

On first start, SSH Vault automatically creates these files next to `ssh-vault.ps1`:

- `ssh-host-meta.json`
- `ssh-host-ui.json`

They store host tags, the selected tag filter, and window size/position.

## Updates

SSH Vault checks the remote `$AppVersion` from `main/ssh-vault.ps1` and compares it with the local `$AppVersion`.

An update is only offered if the GitHub version is newer than the local version.

Before updating, the app asks for confirmation and creates a backup:

```text
ssh-vault.ps1.bak
```


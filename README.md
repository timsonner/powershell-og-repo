# powershell-og-repo

Short Windows PowerShell scripts for Active Directory account work and a few local Windows lookups. Most scripts call one cmdlet, print the result, and exit. There is no module manifest, installer, build, or test suite.

`ad-powershell/ad-commands.md` is a cheat sheet of `Invoke-Command`, `Enter-PSSession`, and Active Directory cmdlets. `latest/powershell-history.txt` is notes for `Get-History`, `Export-Clixml`, and clearing the PSReadLine history file. Neither file is a script.

## Prerequisites

- PowerShell. The Active Directory and local scripts target Windows PowerShell cmdlets (`Get-WinEvent`, `Get-CimInstance`, `Get-LocalGroupMember`, `Get-Acl`, `WindowsIdentity`).
- The `ActiveDirectory` module for every script that starts with `Import-Module ActiveDirectory`. On Windows that module is installed with RSAT. These scripts do not install it.
- A domain account that can run the cmdlet the script calls. Remote scripts prompt with `Get-Credential` and use `New-PSSession` or `Enter-PSSession`.
- Desktop Outlook for `latest/output-user-inbox.ps1` (`Outlook.Application` COM).
- For `ad-powershell/unlock-users-from-list.ps1`, a `users.txt` file in `ad-powershell/` with one username per line. That file is not in the repo.
- For `latest/get-deps-locations.ps1`, a dumpbin-style text file named `test.txt` in the current directory. The repo sample is `latest/dllttest.txt`.

## Install, build, run, test

There is nothing to install or build from this repo. There is no test script.

Run a script from the repository root:

```powershell
pwsh -NoProfile -File ./ad-powershell/get-all-users.ps1
```

On Windows PowerShell the same file is:

```powershell
powershell -NoProfile -File .\ad-powershell\get-all-users.ps1
```

Scripts do not set `$ErrorActionPreference`. A missing cmdlet is printed and the process can still exit 0.

Checked with PowerShell 7.5.4: every `.ps1` file parses. `get-all-users.ps1`, `get-all-objects.ps1`, `add-org-unit.ps1`, `move-object.ps1`, and `unlock-account.ps1` stop because `ActiveDirectory` is not installed here. The local scripts stop on Windows-only APIs. `unlock-users-from-list.ps1` prints that `ad-powershell/users.txt` is missing. `get-deps-locations.ps1` reads `.\test.txt` from the current directory.

## Usage

### List and query

These take no arguments. They need the `ActiveDirectory` module.

```powershell
pwsh -NoProfile -File ./ad-powershell/get-all-users.ps1
pwsh -NoProfile -File ./ad-powershell/get-disabled-accounts.ps1
pwsh -NoProfile -File ./ad-powershell/get-locked-accounts.ps1
pwsh -NoProfile -File ./ad-powershell/get-users-never-logged-in.ps1
pwsh -NoProfile -File ./ad-powershell/get-all-groups.ps1
pwsh -NoProfile -File ./ad-powershell/get-all-computers.ps1
pwsh -NoProfile -File ./ad-powershell/get-all-org-units.ps1
pwsh -NoProfile -File ./ad-powershell/get-all-objects.ps1
```

What each one queries:

| Script | What it runs |
| --- | --- |
| `get-all-users.ps1` | `Get-ADUser` where `Enabled` is true. Prints `SamAccountName` and `SID`. |
| `get-disabled-accounts.ps1` | `Get-ADUser` where `Enabled` is false. |
| `get-locked-accounts.ps1` | `Search-ADAccount -LockedOut`. |
| `get-users-never-logged-in.ps1` | Enabled users whose `LastLogonTimestamp` is null. |
| `get-all-groups.ps1` | `Get-ADGroup -Filter *`. Prints `Name` and `SamAccountName`. |
| `get-all-computers.ps1` | `Get-ADComputer -Filter *`. Prints `Name`. |
| `get-all-org-units.ps1` | `Get-ADOrganizationalUnit -Filter *`. Prints `Name` and `DistinguishedName`. |
| `get-all-objects.ps1` | `Get-ADObject -Filter *`. |

`get-group-info.ps1` takes the group name as the first argument. With no argument it prints its usage line and exits. The usage line inside the file says `Get-GroupMembers.ps1`. The file name is `get-group-info.ps1`.

```powershell
pwsh -NoProfile -File ./ad-powershell/get-group-info.ps1 "GroupName"
```

`get-user-account-control-matches.ps1` takes one `userAccountControl` integer. With no argument it prints the values it knows (`512`, `514`, `66048`, `66050`, `66082`, `262656`, `262658`, `16`, `65536`, `524288`) and exits. With a known value it prints the description, then runs `Get-ADUser -LDAPFilter` for that bit.

```powershell
pwsh -NoProfile -File ./ad-powershell/get-user-account-control-matches.ps1
pwsh -NoProfile -File ./ad-powershell/get-user-account-control-matches.ps1 512
```

These prompt in the console:

```powershell
pwsh -NoProfile -File ./ad-powershell/get-user-info.ps1
pwsh -NoProfile -File ./ad-powershell/get-user-by-sid.ps1
pwsh -NoProfile -File ./ad-powershell/get-objects-from-ou.ps1
```

- `get-user-info.ps1` asks for a username and runs `Get-ADUser -Identity` with `-Properties *`.
- `get-user-by-sid.ps1` asks for a SID and filters `Get-ADUser` on `ObjectSID`.
- `get-objects-from-ou.ps1` asks for an OU name and a domain name, then searches `OU=<name>,DC=<domain>,DC=com`.

### Change accounts

Edit the hardcoded values in the script before running these. They are not parameters.

```powershell
pwsh -NoProfile -File ./ad-powershell/add-org-unit.ps1
pwsh -NoProfile -File ./ad-powershell/add-user.ps1
pwsh -NoProfile -File ./ad-powershell/add-user-multiple-groups.ps1
pwsh -NoProfile -File ./ad-powershell/remove-user-multiple-groups.ps1
pwsh -NoProfile -File ./ad-powershell/move-object.ps1
```

- `add-org-unit.ps1` creates an OU named `NewOU` at `DC=domain,DC=com`.
- `add-user.ps1` prompts for a username and a password (`Read-Host -AsSecureString`). Name, UPN, given name, surname, OU path, and email are fixed strings in the `New-ADUser` call.
- `add-user-multiple-groups.ps1` prompts for a username and calls `Add-ADGroupMember` for the names in `$GroupNames` (`1`, `2`, `etc`).
- `remove-user-multiple-groups.ps1` prompts for a username and calls `Remove-ADGroupMember` for those same placeholder group names.
- `move-object.ps1` is one `Move-ADObject` call with a fixed identity, target path, and target server. Change those three values in the file before running it.

These prompt for a username (and a password, where noted):

```powershell
pwsh -NoProfile -File ./ad-powershell/disable-user-account.ps1
pwsh -NoProfile -File ./ad-powershell/reset-user-password.ps1
pwsh -NoProfile -File ./ad-powershell/remove-user-all-groups.ps1
pwsh -NoProfile -File ./ad-powershell/unlock-users-from-list.ps1
pwsh -NoProfile -File ./latest/unlock-account.ps1
```

- `disable-user-account.ps1` calls `Disable-ADAccount`.
- `reset-user-password.ps1` prompts for a secure password and calls `Set-ADAccountPassword -Reset`.
- `remove-user-all-groups.ps1` removes the user from every group returned by `Get-ADPrincipalGroupMembership` (`-Confirm:$false`).
- `unlock-users-from-list.ps1` reads `users.txt` next to the script and calls `Unlock-ADAccount` for each line. If the file is missing it prints the path it expected.
- `unlock-account.ps1` prompts for credentials and calls `Unlock-ADAccount` for the identity text `<username>` written in the script. Replace that text before running it.

`get-account-lockout-events.ps1` reads Security event `4740` with `Get-WinEvent` from the computer stored in `$remoteComputer` (`x.x.x.x` in the file). Change that value before running it.

```powershell
pwsh -NoProfile -File ./ad-powershell/get-account-lockout-events.ps1
```

`get-all-hosts-running-winrm.ps1` calls `Test-WSMan` for host numbers `$startHostNumber` through `$endHostNumber` (1 through 254) on `$baseIP` (`x.x.x.` in the file). Change the base before running it. It prints only hosts where `Test-WSMan` succeeds.

```powershell
pwsh -NoProfile -File ./ad-powershell/get-all-hosts-running-winrm.ps1
```

### Remote menu and sessions

These prompt for a domain admin credential and a domain controller name:

```powershell
pwsh -NoProfile -File ./latest/ps-admin-tool.ps1
pwsh -NoProfile -File ./latest/ps-pssession.ps1
pwsh -NoProfile -File ./latest/ps-run-command-dc.ps1
```

- `ps-admin-tool.ps1` opens a `PSSession` and shows a menu on that session: reset password, add or remove a group member, unlock, create user, disable, enable, move to an OU, get user info, list group members, list locked accounts, list groups, require a password change at next logon, list OUs, exit.
- `ps-pssession.ps1` calls `Enter-PSSession`. Leave with `Exit-PSSession`.
- `ps-run-command-dc.ps1` runs `Get-PSSession` inside the remote session, then removes the session.

### Local Windows

```powershell
pwsh -NoProfile -File ./ad-powershell/local/get-current-privs.ps1
pwsh -NoProfile -File ./ad-powershell/local/get-local-admins.ps1
pwsh -NoProfile -File ./ad-powershell/local/get-local-folder-acl.ps1
pwsh -NoProfile -File ./ad-powershell/local/get-local-users.ps1
pwsh -NoProfile -File ./ad-powershell/local/get-logon-events.ps1
pwsh -NoProfile -File ./latest/check-priv.ps1
```

- `get-current-privs.ps1` prints group names for the current Windows identity.
- `get-local-admins.ps1` lists `Administrators` members whose `PrincipalSource` is `Local`.
- `get-local-folder-acl.ps1` prints ACL entries for `$FolderPath` (`C:\windows\system32` in the file).
- `get-local-users.ps1` lists local `Win32_UserAccount` objects.
- `get-logon-events.ps1` prints the newest 10 Security log events with id `4624`.
- `check-priv.ps1` checks `WindowsIdentity.Token.IsElevated`. The printed text says `SeTcbPrivilege`. When the token is elevated it lists local `Win32_UserAccount` names and tries a WMI filter `Domain='True'`.

### File search, inbox, and DLL names

```powershell
pwsh -NoProfile -File ./latest/search-drive-filename.ps1
pwsh -NoProfile -File ./latest/search-drive-file-contents.ps1
pwsh -NoProfile -File ./latest/output-user-inbox.ps1
```

- Both search scripts prompt for credentials, a directory path, and a search string. The credential object is not passed into `Get-ChildItem`. Filename search matches `$_.Name`. Content search uses `Select-String`.
- `output-user-inbox.ps1` prints each Inbox item subject and body from the default Outlook profile, then releases the COM objects.

`get-deps-locations.ps1` reads `.\test.txt` from the current directory, takes DLL names indented four spaces, and runs `Get-Command` for each name. Place the included sample next to the working directory first:

```powershell
Copy-Item ./latest/dllttest.txt ./test.txt
pwsh -NoProfile -File ./latest/get-deps-locations.ps1
```

On the PowerShell 7.5.4 check, that sample produced an empty success list and printed each DLL name under `Errors`, because those Windows DLLs are not on `PATH` here. With no `test.txt` in the current directory, `Get-Content` fails.

## Layout

```
ad-powershell/           Active Directory scripts and ad-commands.md
ad-powershell/local/     Local user, privilege, ACL, and logon scripts
latest/                  Remote session scripts, search scripts, inbox script,
                         check-priv.ps1, get-deps-locations.ps1, dllttest.txt,
                         powershell-history.txt
```

## License

This repository has no `LICENSE` file.

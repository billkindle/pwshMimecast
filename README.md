# pwshMimecast

A PowerShell module collection for Mimecast tenant management. Currently includes delegate mailbox
management via both API 2.0 (OAuth 2.0 / Bearer token) and API 1.0 (HMAC-SHA1).

## Requirements

- PowerShell 7.0 or later (recommended)
- PowerShell 5.1 is also supported
- A Mimecast tenant with API credentials (see [Authentication](#authentication))

## Installation

### Option 1 — Clone the repository

```powershell
git clone https://github.com/bkindle/pwshMimecast.git
```

Then import the module by path:

```powershell
Import-Module ".\pwshMimecast\pwshMimecast.psd1"
```

### Option 2 — Copy the module folder

Copy the `pwshMimecast` folder to a directory on your `$env:PSModulePath`, then import by name:

```powershell
# List available module paths
$env:PSModulePath -split ';'

# Copy the folder to your preferred path, e.g.:
# C:\Users\<you>\Documents\PowerShell\Modules\pwshMimecast\

Import-Module pwshMimecast
```

### Verify the import

```powershell
Get-Module pwshMimecast
Get-Command -Module pwshMimecast
```

## Authentication

This module supports two authentication modes depending on your Mimecast tenant type.

### API 2.0 — OAuth 2.0 (Cloud Gateway tenants)

Used by `Add-MimecastDelegateUsers`, `Get-MimecastDelegateUsers`, `Remove-MimecastDelegateUser`.

Obtain a **Client ID** and **Client Secret** from your Mimecast API 2.0 app registration.

```powershell
$creds = @{
    ClientId     = $env:MIMECAST_CLIENT_ID
    ClientSecret = $env:MIMECAST_CLIENT_SECRET
    BaseUri      = 'https://api.services.mimecast.com'   # adjust for your region
}
```

> **Note:** If your login URL is `login-usb.mimecast.com` (US-B / Cloud Integrated), the API 2.0
> User Management endpoints return 404. Use API 1.0 instead.

### API 1.0 — HMAC-SHA1 (Cloud Integrated tenants)

Used by `Add-MimecastDelegateV1`, `Get-MimecastDelegateV1`, `Remove-MimecastDelegateV1`.

Obtain credentials from **Mimecast Admin → Administration → Services → API 1.0 Applications**.

```powershell
$creds = @{
    AppId       = $env:MIMECAST_APP_ID
    AppKey      = $env:MIMECAST_APP_KEY
    AccessKey   = $env:MIMECAST_ACCESS_KEY
    SecretKey   = $env:MIMECAST_SECRET_KEY
    BaseUri     = 'https://us-api.mimecast.com'   # adjust for your region
}
```

> Store credentials in environment variables or a secrets manager. Never hardcode them in scripts.

## Available Commands

| Command | API | Description |
| --- | --- | --- |
| `Add-MimecastDelegateUsers` | 2.0 | Grant delegate access to a mailbox |
| `Get-MimecastDelegateUsers` | 2.0 | List current delegates for a mailbox |
| `Remove-MimecastDelegateUser` | 2.0 | Revoke delegate access |
| `Add-MimecastDelegateV1` | 1.0 | Grant delegate access (HMAC-SHA1) |
| `Get-MimecastDelegateV1` | 1.0 | List delegates (HMAC-SHA1) |
| `Remove-MimecastDelegateV1` | 1.0 | Revoke delegate access (HMAC-SHA1) |
| `Test-MimecastApiPath` | Both | Verify API endpoint connectivity |

For full parameter reference and examples, see [Mimecast-Delegates-README.md](Mimecast-Delegates-README.md).

## Quick Start

```powershell
Import-Module .\pwshMimecast\pwshMimecast.psd1

# --- API 2.0 example ---
$creds = @{
    ClientId     = $env:MIMECAST_CLIENT_ID
    ClientSecret = $env:MIMECAST_CLIENT_SECRET
    BaseUri      = 'https://api.services.mimecast.com'
}

# List delegates for a mailbox
Get-MimecastDelegateUsers -PrimaryAddress 'shared@contoso.com' @creds

# Add a delegate
Add-MimecastDelegateUsers -PrimaryAddress 'shared@contoso.com' `
    -DelegateAddresses @('user@contoso.com') @creds

# Add delegates to a shared/functional mailbox (InvertedMode)
Add-MimecastDelegateUsers -PrimaryAddress 'helpdesk@contoso.com' `
    -DelegateAddresses @('agent1@contoso.com', 'agent2@contoso.com') `
    -InvertedMode @creds

# Remove a delegate
Remove-MimecastDelegateUser -PrimaryAddress 'shared@contoso.com' `
    -DelegateAddress 'user@contoso.com' @creds
```

## License

MIT License — see [LICENSE](LICENSE) for details.

## Author

Bill Kindle

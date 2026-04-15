# Mimecast Delegate Mailbox Management — PowerShell Module

> **Module:** `Mimecast-Delegates.psm1`
> **API:** Mimecast API 2.0 (OAuth 2.0 Bearer token)
> **Requires:** PowerShell 5.1+
> **Reference:** [Mimecast API 2.0 — User and Group Management](https://developer.services.mimecast.com/docs/userandgroupmanagement/1/overview)

---

## Table of Contents

- [Mimecast Delegate Mailbox Management — PowerShell Module](#mimecast-delegate-mailbox-management--powershell-module)
  - [Table of Contents](#table-of-contents)
  - [Overview](#overview)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
  - [Authentication](#authentication)
    - [Setting credentials via environment variables](#setting-credentials-via-environment-variables)
    - [API Gateway URLs](#api-gateway-urls)
    - [Credential splat pattern](#credential-splat-pattern)
  - [Functions](#functions)
    - [Add-MimecastDelegateUsers](#add-mimecastdelegateusers)
      - [Syntax](#syntax)
      - [Parameters](#parameters)
      - [Output](#output)
      - [Shared mailboxes and -InvertedMode](#shared-mailboxes-and--invertedmode)
    - [Get-MimecastDelegateUsers](#get-mimecastdelegateusers)
      - [Syntax](#syntax-1)
      - [Parameters](#parameters-1)
      - [Output](#output-1)
    - [Remove-MimecastDelegateUser](#remove-mimecastdelegateuser)
      - [Syntax](#syntax-2)
      - [Parameters](#parameters-2)
      - [Output](#output-2)
  - [Usage Examples](#usage-examples)
    - [Add a fixed list of delegates](#add-a-fixed-list-of-delegates)
    - [Bulk-add from a CSV file](#bulk-add-from-a-csv-file)
    - [Bulk-add to a shared mailbox (InvertedMode)](#bulk-add-to-a-shared-mailbox-invertedmode)
    - [Bulk-add from a text file](#bulk-add-from-a-text-file)
    - [List current delegates](#list-current-delegates)
    - [Remove a delegate](#remove-a-delegate)
    - [Audit multiple shared mailboxes](#audit-multiple-shared-mailboxes)
    - [Preview changes with -WhatIf](#preview-changes-with--whatif)
  - [Token Caching](#token-caching)
  - [Credential Management](#credential-management)
  - [Error Handling](#error-handling)
  - [Terminology](#terminology)
  - [API Endpoints Used](#api-endpoints-used)

---

## Overview

This module provides three PowerShell functions for managing Mimecast **delegate mailbox permissions**.
Delegate access allows one user (the *primary address*) to read and search another user's archived
mailbox (the *delegate address*) within Mimecast's end-user applications.

All API calls use **Mimecast API 2.0** with **OAuth 2.0 client credentials** authentication.
The module automatically obtains and caches a Bearer token (valid 30 minutes) and refreshes it
transparently. You only need a `ClientId` and `ClientSecret` — no HMAC signing or App keys required.

---

## Prerequisites

| Requirement | Detail |
| --- | --- |
| PowerShell version | 5.1 or later |
| Mimecast admin role | Full Administrator or Super Administrator |
| Required API permission | User and Group Management — Delegate Access |
| Mimecast API credentials | Client ID and Client Secret (from API 2.0 application) |

To create API credentials, navigate to **Administration → Services → API and Platform
Integrations → Available Integrations**, locate the **Mimecast API 2.0** tile, and select
**Generate Keys**. Assign the application a custom role with least-privilege permissions scoped
to User and Group Management.

---

## Installation

No installation is required. Import the module directly from the file:

```powershell
Import-Module .\Mimecast-Delegates.psm1
```

To make it available in all sessions, copy it to a directory on your `$env:PSModulePath` and import by name:

```powershell
# Example: place in user module directory
Copy-Item .\Mimecast-Delegates.psm1 `
    "$HOME\Documents\PowerShell\Modules\Mimecast-Delegates\Mimecast-Delegates.psm1"

Import-Module Mimecast-Delegates
```

---

## Authentication

Mimecast API 2.0 uses **OAuth 2.0 client credentials**. The module calls `POST /oauth/token`
with your `ClientId` and `ClientSecret` to obtain a Bearer token, which is then sent as the
`Authorization: Bearer <token>` header on every API request.

The token is valid for **30 minutes**. The module caches it in a module-scoped variable and
auto-refreshes it 60 seconds before expiry — you do not need to manage token lifecycle yourself.

The `ClientSecret` parameter is typed as `[securestring]` throughout the module. Never pass it as a plain string.

### Setting credentials via environment variables

Store the Client ID as an environment variable. Retrieve the secret securely at runtime:

```powershell
$env:MC_CLIENT_ID = "your-client-id"

# Retrieve secret securely — never store plain text
$secret = Read-Host 'Mimecast Client Secret' -AsSecureString
```

### API Gateway URLs

API 2.0 uses a new gateway domain (`services.mimecast.com`), not the old `mimecast.com` endpoints.

| Option | Base URL | Notes |
| --- | --- | --- |
| Global (recommended) | `https://us-api.services.mimecast.com` | Auto-failover; default |
| US instance | `https://us-api.services.mimecast.com` | US data residency |
| UK instance | `https://uk-api.services.mimecast.com` | UK data residency |

The `BaseUrl` parameter defaults to `https://us-api.services.mimecast.com` in all functions.

### Credential splat pattern

```powershell
$secret    = Read-Host 'Mimecast Client Secret' -AsSecureString
$credSplat = @{
    ClientId     = $env:MC_CLIENT_ID
    ClientSecret = $secret
}
```

All functions accept and pass through `BaseUrl`, `ClientId`, and `ClientSecret`, so the splat
works the same way it did with API 1.0.

---

## Functions

---

### Add-MimecastDelegateUsers

Grants one or more users delegate access to a target mailbox.

**Standard mode:** `POST /api/user-management/v1/users/{delegateAddress}/delegates`

**Inverted mode** (`-InvertedMode`)**:** `POST /api/user-management/v1/users/{primaryAddress}/delegates`

Use `-InvertedMode` when `DelegateAddress` is a shared or functional mailbox that is not a full
user account in Mimecast. Without it the API returns HTTP 404 for those mailboxes.

#### Syntax

```powershell
Add-MimecastDelegateUsers
    -DelegateAddress  <string>
    -PrimaryAddresses <string[]>
    -ClientId         <string>
    -ClientSecret     <securestring>
    [-InvertedMode]
    [-BaseUrl         <string>]
    [-DelayMs         <int>]
    [-Verbose]
    [-WhatIf]
```

#### Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `DelegateAddress` | `string` | Yes | — | The mailbox being shared (the one others will gain access to). Must be a valid email address (`user@domain.com`) |
| `PrimaryAddresses` | `string[]` | Yes | — | One or more email addresses that will receive delegate access. Each value must be a valid email address (`user@domain.com`); plain usernames without `@` are rejected at bind time. Blank or whitespace-only entries are silently skipped. Accepts pipeline input |
| `InvertedMode` | `switch` | No | off | Inverts the API call so each primary address is in the URL path and `DelegateAddress` is in the body. Required for shared/functional mailboxes that are not full user accounts in Mimecast |
| `ClientId` | `string` | Yes | — | Mimecast API 2.0 Client ID |
| `ClientSecret` | `securestring` | Yes | — | Mimecast API 2.0 Client Secret |
| `BaseUrl` | `string` | No | `https://us-api.services.mimecast.com` | API 2.0 gateway base URL |
| `DelayMs` | `int` | No | `300` | Milliseconds to pause between API calls (rate-limit protection) |

#### Output

Returns a `PSCustomObject[]` with one entry per address:

| Property | Type | Description |
| --- | --- | --- |
| `PrimaryAddress` | `string` | The address that was processed |
| `DelegateAddress` | `string` | The target mailbox |
| `Success` | `bool` | `$true` if the delegate was created successfully |
| `DelegateId` | `string` | Mimecast secure ID for the delegate right |
| `FailReason` | `string` | HTTP status and API error detail if the call failed; otherwise `$null` |

#### Shared mailboxes and -InvertedMode

Mimecast's user-management API resolves the target mailbox by looking up a full user account.
Shared and functional mailboxes (e.g. `materialreceipt@`, `info@`, `helpdesk@`) are often not
full user accounts in Mimecast, so `POST /users/{sharedMailbox}/delegates` returns **HTTP 404**.

The `-InvertedMode` switch works around this by reversing which address goes in the URL:

| Mode | URL path | Body |
| --- | --- | --- |
| Standard | `{delegateAddress}` | `{ "delegateEmailAddress": "{primaryAddress}" }` |
| Inverted | `{primaryAddress}` | `{ "delegateEmailAddress": "{delegateAddress}" }` |

Both modes produce the same result — the primary user gains access to the target mailbox — but
inverted mode succeeds when the target is a shared mailbox.

> **Important — querying and removing delegates for shared mailboxes:**
> Because Mimecast registers the delegate relationship against each *user's* account (not against
> the shared mailbox), `Get-MimecastDelegateUsers` called with the shared mailbox address will
> return an empty list even when delegates exist. To list or remove a delegate relationship that
> was added with `-InvertedMode`, query using the **user's address** as `-PrimaryAddress`, then
> filter by `EmailAddress` to find the shared mailbox entry:
>
> ```powershell
> # List shared mailboxes delegated TO a user
> Get-MimecastDelegateUsers -PrimaryAddress 'user@contoso.com' @creds
>
> # Remove a specific shared mailbox from a user's delegate list
> $delegate = Get-MimecastDelegateUsers -PrimaryAddress 'user@contoso.com' @creds |
>                 Where-Object EmailAddress -EQ 'shared@contoso.com'
> Remove-MimecastDelegateUser -PrimaryAddress 'user@contoso.com' `
>     -DelegateId $delegate.DelegateId @creds
> ```

---

### Get-MimecastDelegateUsers

Returns all users who currently have delegate access to a mailbox.

**API endpoint:** `POST /api/user/find-delegate-users`

#### Syntax

```powershell
Get-MimecastDelegateUsers
    -PrimaryAddress <string>
    -ClientId       <string>
    -ClientSecret   <securestring>
    [-BaseUrl       <string>]
```

#### Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `PrimaryAddress` | `string` | Yes | — | The email address of the mailbox whose delegates you want to list |
| `ClientId` | `string` | Yes | — | Mimecast API 2.0 Client ID |
| `ClientSecret` | `securestring` | Yes | — | Mimecast API 2.0 Client Secret |
| `BaseUrl` | `string` | No | `https://us-api.services.mimecast.com` | API 2.0 gateway base URL |

#### Output

Returns a `PSCustomObject[]` with one entry per delegate:

| Property | Type | Description |
| --- | --- | --- |
| `EmailAddress` | `string` | Email address of the delegate user |
| `DisplayName` | `string` | Display name of the delegate user |
| `DelegateId` | `string` | Mimecast secure ID of the delegate right (required for `Remove-MimecastDelegateUser`) |
| `Source` | `string` | Provisioning source returned by Mimecast (e.g. `ADCON`) |
| `PrimaryAddress` | `string` | The mailbox that was queried |

---

### Remove-MimecastDelegateUser

Revokes a user's delegate access to a mailbox.

**API endpoint:** `DELETE /api/user-management/v1/users/{emailAddress}/delegates/{delegateId}`

> **Note:** The API 2.0 DELETE endpoint identifies the record by its **ID**, not by address pair.
> Call `Get-MimecastDelegateUsers` first to obtain the `DelegateId` for the delegate you want to remove.

#### Syntax

```powershell
Remove-MimecastDelegateUser
    -PrimaryAddress <string>
    -DelegateId     <string>
    -ClientId       <string>
    -ClientSecret   <securestring>
    [-BaseUrl       <string>]
    [-Confirm]
    [-WhatIf]
```

#### Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `PrimaryAddress` | `string` | Yes | — | The email address of the mailbox to remove the delegate from |
| `DelegateId` | `string` | Yes | — | The delegate record ID (from `Get-MimecastDelegateUsers`) |
| `ClientId` | `string` | Yes | — | Mimecast API 2.0 Client ID |
| `ClientSecret` | `securestring` | Yes | — | Mimecast API 2.0 Client Secret |
| `BaseUrl` | `string` | No | `https://us-api.services.mimecast.com` | API 2.0 gateway base URL |

> **Note:** This function has a `ConfirmImpact` of `High`. PowerShell will prompt for confirmation
> unless `-Confirm:$false` is specified.

#### Output

Returns a single `PSCustomObject`:

| Property | Type | Description |
| --- | --- | --- |
| `PrimaryAddress` | `string` | The mailbox that was modified |
| `DelegateId` | `string` | The delegate record ID that was removed |
| `Success` | `bool` | `$true` if the delegate was removed successfully |

---

## Usage Examples

### Add a fixed list of delegates

```powershell
Import-Module .\Mimecast-Delegates.psm1

$secret    = Read-Host 'Mimecast Client Secret' -AsSecureString
$credSplat = @{
    ClientId     = $env:MC_CLIENT_ID
    ClientSecret = $secret
}

$results = Add-MimecastDelegateUsers `
    -DelegateAddress  'shared-inbox@contoso.com' `
    -PrimaryAddresses 'alice@contoso.com', 'bob@contoso.com', 'carol@contoso.com' `
    -Verbose `
    @credSplat

$results | Format-Table PrimaryAddress, Success, DelegateId, FailReason -AutoSize
```

---

### Bulk-add from a CSV file

The CSV must have a `PrimaryAddress` column header (case-sensitive). There are three ways to create it:

**Option 1 — PowerShell from a hard-coded list (recommended):**

```powershell
@(
    'alice@contoso.com'
    'bob@contoso.com'
    'carol@contoso.com'
) | ForEach-Object { [PSCustomObject]@{ PrimaryAddress = $_ } } |
    Export-Csv -Path '.\delegates-to-add.csv' -NoTypeInformation
```

**Option 2 — PowerShell from an Exchange Online distribution group:**

```powershell
Get-DistributionGroupMember -Identity 'somegroup@contoso.com' |
    Select-Object @{ Name = 'PrimaryAddress'; Expression = { $_.PrimarySmtpAddress } } |
    Export-Csv -Path '.\delegates-to-add.csv' -NoTypeInformation
```

**Option 3 — Manually in Excel or Notepad:**

Create a file with exactly this content and save it as `.csv` (UTF-8 encoding):

```csv
PrimaryAddress
alice@contoso.com
bob@contoso.com
carol@contoso.com
```

Once the CSV is ready, import and run:

```powershell
$users = Import-Csv .\delegates-to-add.csv | Select-Object -ExpandProperty PrimaryAddress

$results = Add-MimecastDelegateUsers `
    -DelegateAddress  'shared-inbox@contoso.com' `
    -PrimaryAddresses $users `
    @credSplat

# Export results for auditing
$results | Export-Csv -Path .\delegate-results.csv -NoTypeInformation
```

---

### Bulk-add to a shared mailbox (InvertedMode)

When `DelegateAddress` is a shared or functional mailbox, add `-InvertedMode`:

```csv
PrimaryAddress
cpowers@contoso.com
jsmith@contoso.com
jdoe@contoso.com
```

```powershell
$users = Import-Csv .\delegates-to-add.csv | Select-Object -ExpandProperty PrimaryAddress

$results = Add-MimecastDelegateUsers `
    -DelegateAddress  'materialreceipt@contoso.com' `
    -PrimaryAddresses $users `
    -InvertedMode `
    @credSplat

$results | Format-Table PrimaryAddress, Success, DelegateId, FailReason -AutoSize

# Export results
$results | Export-Csv -Path ".\delegate-results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation
```

---

### Bulk-add from a text file

The text file should contain one email address per line:

```text
alice@contoso.com
bob@contoso.com
carol@contoso.com
```

```powershell
Get-Content .\users.txt |
    Add-MimecastDelegateUsers -DelegateAddress 'shared-inbox@contoso.com' @credSplat |
    Format-Table -AutoSize
```

---

### List current delegates

```powershell
$delegates = Get-MimecastDelegateUsers `
    -PrimaryAddress 'shared-inbox@contoso.com' `
    @credSplat

$delegates | Format-Table EmailAddress, DisplayName, DelegateId, Source -AutoSize
```

---

### Remove a delegate

Because the API 2.0 DELETE endpoint requires the delegate record ID, use `Get-MimecastDelegateUsers` to look it up first:

```powershell
# Step 1: find the record ID
$delegate = Get-MimecastDelegateUsers -PrimaryAddress 'shared-inbox@contoso.com' @credSplat |
                Where-Object EmailAddress -eq 'alice@contoso.com'

# Step 2: remove it (will prompt for confirmation)
Remove-MimecastDelegateUser `
    -PrimaryAddress 'shared-inbox@contoso.com' `
    -DelegateId     $delegate.DelegateId `
    @credSplat

# Skip confirmation prompt
Remove-MimecastDelegateUser `
    -PrimaryAddress 'shared-inbox@contoso.com' `
    -DelegateId     $delegate.DelegateId `
    -Confirm:$false `
    @credSplat
```

---

### Audit multiple shared mailboxes

Export all delegate relationships across several shared mailboxes to a CSV:

```powershell
$sharedMailboxes = @(
    'helpdesk@contoso.com'
    'finance@contoso.com'
    'legal@contoso.com'
)

$audit = foreach ($mailbox in $sharedMailboxes) {
    $delegates = Get-MimecastDelegateUsers -PrimaryAddress $mailbox @credSplat
    foreach ($d in $delegates) {
        [PSCustomObject]@{
            SharedMailbox  = $mailbox
            DelegateEmail  = $d.EmailAddress
            DisplayName    = $d.DisplayName
            DelegateId     = $d.DelegateId
            Source         = $d.Source
        }
    }
}

$audit | Export-Csv -Path ".\delegate-audit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation
Write-Host "[OK] Audit saved ($($audit.Count) records)"
```

---

### Preview changes with -WhatIf

Both `Add-MimecastDelegateUsers` and `Remove-MimecastDelegateUser` support `-WhatIf`, letting
you see what would happen without making any API calls:

```powershell
Add-MimecastDelegateUsers `
    -DelegateAddress  'shared-inbox@contoso.com' `
    -PrimaryAddresses 'alice@contoso.com', 'bob@contoso.com' `
    -WhatIf `
    @credSplat
```

---

## Token Caching

The module maintains a module-scoped token cache (`$script:_TokenCache`). The first call to any
function triggers authentication against `/oauth/token`. Subsequent calls within the same session
reuse the cached token until it is within 60 seconds of its 30-minute expiry, at which point a
new token is fetched automatically.

The cache is **per PowerShell session**. Re-importing the module or starting a new session clears it.

To force an immediate token refresh (e.g. after rotating credentials):

```powershell
Get-MimecastAccessToken -BaseUrl 'https://us-api.services.mimecast.com' `
    -ClientId $env:MC_CLIENT_ID -ClientSecret $secret -ForceRefresh
```

---

## Credential Management

Avoid storing the Client Secret in plain text. Recommended approaches:

**Interactive prompt (ad-hoc scripts):**

```powershell
$secret = Read-Host 'Mimecast Client Secret' -AsSecureString
```

**PowerShell SecretManagement (recommended for interactive use):**

```powershell
Install-Module Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore
Register-SecretVault -Name 'LocalVault' -ModuleName Microsoft.PowerShell.SecretStore
Set-Secret -Name 'MC_CLIENT_SECRET' -Secret 'your-client-secret'

# Retrieve as SecureString at runtime
$secret = Get-Secret -Name 'MC_CLIENT_SECRET'
```

**Azure Key Vault (recommended for production/automation):**

```powershell
$plainSecret = Get-AzKeyVaultSecret -VaultName 'MyVault' -Name 'MC-CLIENT-SECRET' -AsPlainText
$secret      = ConvertTo-SecureString $plainSecret -AsPlainText -Force
```

---

## Error Handling

`Add-MimecastDelegateUsers` writes failures to the output stream as `PSCustomObject` entries
with `Success = $false` and a `FailReason` (HTTP status code + message), rather than throwing
terminating errors. This ensures a bulk operation processes all addresses even if some fail.

```powershell
$results = Add-MimecastDelegateUsers -DelegateAddress 'shared@contoso.com' `
               -PrimaryAddresses $users @credSplat

$succeeded = $results | Where-Object Success
$failed    = $results | Where-Object { -not $_.Success }

Write-Host "Succeeded: $($succeeded.Count)  Failed: $($failed.Count)"
$failed | Select-Object PrimaryAddress, FailReason | Format-Table -AutoSize
```

`Get-MimecastDelegateUsers` and `Remove-MimecastDelegateUser` use `Write-Error` for failures since
they operate on a single target. A `429 Too Many Requests` response includes the `X-RateLimit-Reset`
header (milliseconds until quota resets) — add a `Start-Sleep` based on that value before retrying.

---

## Terminology

| Term | Definition |
| --- | --- |
| **Primary Address** | The user who *gains* access to another mailbox |
| **Delegate Address** | The mailbox being shared (the one being accessed) |
| **Delegate Right** | The permission record linking a primary address to a delegate address |

---

## API Endpoints Used

| Function | Mode | Method | Endpoint |
| --- | --- | --- | --- |
| `Add-MimecastDelegateUsers` | Standard | POST | `/api/user-management/v1/users/{delegateAddress}/delegates` |
| `Add-MimecastDelegateUsers` | `-InvertedMode` | POST | `/api/user-management/v1/users/{primaryAddress}/delegates` |
| `Get-MimecastDelegateUsers` | — | GET | `/api/user-management/v1/users/{emailAddress}/delegates` |
| `Remove-MimecastDelegateUser` | — | DELETE | `/api/user-management/v1/users/{emailAddress}/delegates/{delegateId}` |
| *(internal)* `Get-MimecastAccessToken` | — | POST | `/oauth/token` |

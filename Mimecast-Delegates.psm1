<#
.SYNOPSIS
    Mimecast Delegate Mailbox Management Module

.DESCRIPTION
    PowerShell module for managing Mimecast delegate mailbox permissions.
    Supports two authentication modes:

    API 2.0 (OAuth 2.0) - functions: Add/Get/Remove-MimecastDelegateUsers
      Use when your tenant is Email Security Cloud Gateway.
      Credentials: ClientId + ClientSecret (from API 2.0 app registration).
      NOTE: User Management API endpoints are marked 'Cloud Gateway only' by
      Mimecast. If your login URL is login-usb.mimecast.com (US-B / Cloud
      Integrated), these endpoints return 404 and you must use API 1.0 instead.

    API 1.0 (HMAC-SHA1) - functions: Add/Get/Remove-MimecastDelegateV1
      Use when your tenant is Email Security Cloud Integrated, or when API 2.0
      User Management endpoints return 404.
      Credentials: AppId + AppKey + AccessKey + SecretKey
      Obtained from: Mimecast Admin -> Administration -> Services ->
        API 1.0 Applications -> Add API Application, then
        Account -> Roles -> [your role] -> API Application Authentication

    Shared/Functional Mailboxes (API 2.0 only):
      Use the -InvertedMode switch on Add-MimecastDelegateUsers when the target
      mailbox is a shared/functional mailbox not registered as a full user.

.NOTES
    Author            : Bill Kindle (with AI assistance)
    Version           : 1.0
    Created           : 2026-04-02
    API 2.0 Reference : https://developer.services.mimecast.com/docs/userandgroupmanagement/1/overview
    API 1.0 Reference : https://integrations.mimecast.com/documentation/endpoint-reference/
    Auth Model 2.0    : OAuth 2.0 Client Credentials (Bearer token, 30 min TTL)
    Auth Model 1.0    : HMAC-SHA1 per-request signed headers

    Terminology (Mimecast API 2.0 field names):
      primaryAddress  - the mailbox being shared (the owner whose mailbox others access)
      delegateAddress - the user who GAINS access to the shared mailbox
#>

#region Module-scope token cache

$script:_TokenCache = [PSCustomObject]@{
    AccessToken = [string]::Empty
    ExpiresAt   = [datetime]::MinValue
    BaseUrl     = [string]::Empty
}

#endregion Module-scope token cache

#region Helper Functions - Authentication

function Get-MimecastAccessToken {
    <#
    .SYNOPSIS
        Obtains (or returns a cached) Mimecast API 2.0 Bearer access token.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)] [string]       $BaseUrl,
        [Parameter(Mandatory)] [string]       $ClientId,
        [Parameter(Mandatory)] [securestring] $ClientSecret,
        [switch]                               $ForceRefresh
    )

    # Return cached token when still valid (60-second safety buffer)
    if (-not $ForceRefresh -and
        $script:_TokenCache.AccessToken -and
        $script:_TokenCache.BaseUrl -eq $BaseUrl.TrimEnd('/') -and
        [datetime]::UtcNow -lt $script:_TokenCache.ExpiresAt.AddSeconds(-60)) {

        Write-Verbose '[INFO] Using cached Mimecast access token'
        return $script:_TokenCache.AccessToken
    }

    Write-Verbose '[INFO] Requesting new Mimecast access token...'

    $tokenUrl = $BaseUrl.TrimEnd('/') + '/oauth/token'

    # Decrypt SecureString only long enough to build the form body
    $bstr        = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
    $plainSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

    $formBody = "grant_type=client_credentials" +
                "&client_id=$([uri]::EscapeDataString($ClientId))" +
                "&client_secret=$([uri]::EscapeDataString($plainSecret))"

    # Overwrite plain secret in memory before any potential exception
    $plainSecret = [string]::Empty

    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenUrl `
                        -ContentType 'application/x-www-form-urlencoded' `
                        -Body $formBody -ErrorAction Stop

        $ttl = if ($response.expires_in) { [int]$response.expires_in } else { 1800 }

        $script:_TokenCache.AccessToken = $response.access_token
        $script:_TokenCache.ExpiresAt   = [datetime]::UtcNow.AddSeconds($ttl)
        $script:_TokenCache.BaseUrl     = $BaseUrl.TrimEnd('/')

        Write-Verbose "[OK] Token acquired. Expires at $($script:_TokenCache.ExpiresAt.ToString('u'))"
        return $script:_TokenCache.AccessToken
    }
    catch {
        throw "Failed to obtain Mimecast access token: $($_.Exception.Message)"
    }
    finally {
        $formBody = [string]::Empty
    }
}

#endregion Helper Functions - Authentication

#region Helper Functions - HTTP

function Invoke-MimecastApiCall {
    <#
    .SYNOPSIS
        Sends an authenticated request to the Mimecast API 2.0 gateway.
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory)] [string] $BaseUrl,
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Token,

        [ValidateSet('GET', 'POST', 'DELETE')]
        [string] $Method = 'GET',

        [hashtable] $Body
    )

    $url     = $BaseUrl.TrimEnd('/') + $Path
    $headers = @{
        'Authorization' = "Bearer $Token"
        'Content-Type'  = 'application/json'
        'Accept'        = 'application/json'
    }

    $params = @{
        Method      = $Method
        Uri         = $url
        Headers     = $headers
        ErrorAction = 'Stop'
    }

    if ($Body) {
        $params['Body'] = ($Body | ConvertTo-Json -Depth 10)
    }

    try {
        return Invoke-RestMethod @params
    }
    catch {
        # Extract a meaningful error message from the response body.
        # PS 7 (.NET): response body is in $_.ErrorDetails.Message
        # PS 5.1 (.NET Framework): must read GetResponseStream()
        $statusCode = $null
        $rawBody    = $null

        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }

        # PS 7 path
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $rawBody = $_.ErrorDetails.Message
        }

        # PS 5.1 fallback
        if (-not $rawBody -and $_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                if ($stream) {
                    $reader  = [System.IO.StreamReader]::new($stream)
                    $rawBody = $reader.ReadToEnd()
                    $reader.Dispose()
                }
            }
            catch { <# stream read failed #> }
        }

        $apiDetail = $null
        if ($rawBody) {
            try {
                $errObj = $rawBody | ConvertFrom-Json -ErrorAction Stop
                # Mimecast API 2.0: { "fail": [{ "errors": [{ "message": "..." }] }] }
                if ($errObj.fail) {
                    $apiDetail = ($errObj.fail[0].errors | ForEach-Object { $_.message }) -join '; '
                }
                # RFC 7807: { "title": "...", "detail": "..." }
                elseif ($errObj.detail) {
                    $apiDetail = "$($errObj.title): $($errObj.detail)"
                }
                elseif ($errObj.message) {
                    $apiDetail = $errObj.message
                }
                else {
                    $apiDetail = $rawBody
                }
            }
            catch {
                $apiDetail = $rawBody
            }
        }

        $msg = "HTTP $statusCode - $($_.Exception.Message)"
        if ($apiDetail) { $msg += " | API detail: $apiDetail" }
        throw $msg
    }
}

#endregion Helper Functions - HTTP

#region Public Functions

# ---------------------------------------------------------------------------
# Public: Add one or more delegate users to a mailbox
# ---------------------------------------------------------------------------
function Add-MimecastDelegateUsers {
    <#
    .SYNOPSIS
        Grants one or more users delegate access to a target mailbox.

    .DESCRIPTION
        Calls POST /api/user/add-delegate-user for each primary address supplied.
        Both addresses are sent in the request body -- no address in the URL path.

        Request body: { "data": [{ "primaryAddress": "<mailbox-owner>", "delegateAddress": "<user-getting-access>" }] }

        Default mode:
            primaryAddress = DelegateAddress (the mailbox being shared)
            delegateAddress = each PrimaryAddress (the user receiving access)

        Inverted mode (-InvertedMode switch, for shared/functional mailboxes):
            primaryAddress = each PrimaryAddress
            delegateAddress = DelegateAddress
            Use if the default mode returns errors indicating an address is not
            found as a primary mailbox holder.

    .PARAMETER DelegateAddress
        The mailbox to be shared (the mailbox others will access).
        Must be a valid email address formatted as 'user@domain.com'.

    .PARAMETER PrimaryAddresses
        One or more email addresses that will receive delegate access.
        Each entry must be a valid email address formatted as 'user@domain.com'.
        Plain usernames without '@' are rejected at bind time.
        Blank and whitespace-only entries are silently skipped.
        Accepts pipeline input and array values.

    .PARAMETER InvertedMode
        Use this switch when DelegateAddress is a shared/functional mailbox that
        is not a full user account in Mimecast. Inverts the request body so each
        PrimaryAddress becomes the body's primaryAddress field and DelegateAddress
        becomes the delegateAddress field. Use if the default mode returns errors
        indicating an address is not found as a primary mailbox holder.

    .PARAMETER BaseUrl
        Mimecast API 2.0 gateway base URL.
        Default: https://api.services.mimecast.com

    .PARAMETER ClientId
        Your Mimecast API 2.0 Client ID.

    .PARAMETER ClientSecret
        Your Mimecast API 2.0 Client Secret as a SecureString.

    .PARAMETER DelayMs
        Milliseconds to pause between API calls. Default: 300.

    .EXAMPLE
        # Standard mode - full user mailbox
        $secret = Read-Host 'Client Secret' -AsSecureString
        Add-MimecastDelegateUsers `
            -DelegateAddress  'shared@contoso.com' `
            -PrimaryAddresses 'alice@contoso.com','bob@contoso.com' `
            -ClientId         $env:MC_CLIENT_ID `
            -ClientSecret     $secret

    .EXAMPLE
        # Inverted mode - shared/functional mailbox as target
        $secret = Read-Host 'Client Secret' -AsSecureString
        $users  = Import-Csv .\delegates.csv | Select-Object -ExpandProperty PrimaryAddress
        Add-MimecastDelegateUsers `
            -DelegateAddress  'materialreceipt@contoso.com' `
            -PrimaryAddresses $users `
            -InvertedMode `
            -ClientId         $env:MC_CLIENT_ID `
            -ClientSecret     $secret

    .EXAMPLE
        # WhatIf preview - no changes made
        Add-MimecastDelegateUsers -DelegateAddress 'shared@contoso.com' `
            -PrimaryAddresses 'alice@contoso.com' @creds -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'DelayMs',
        Justification = 'Used inside the process block foreach loop')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'InvertedMode',
        Justification = 'Used inside the process block via $InvertedMode variable')]
    param(
        [Parameter(Mandatory)]
        [ValidateScript({
            if ($_ -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return $true }
            throw "'$_' is not a valid email address. Expected format: user@domain.com"
        })]
        [string] $DelegateAddress,

        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowEmptyString()]
        [ValidateScript({
            if ([string]::IsNullOrWhiteSpace($_)) { return $true }
            if ($_ -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return $true }
            throw "'$_' is not a valid email address. Expected format: user@domain.com"
        })]
        [string[]] $PrimaryAddresses,

        [switch] $InvertedMode,

        [string]       $BaseUrl      = 'https://api.services.mimecast.com',
        [Parameter(Mandatory)] [string]       $ClientId,
        [Parameter(Mandatory)] [securestring] $ClientSecret,

        [int] $DelayMs = 300
    )

    begin {
        $ErrorActionPreference = 'Stop'
        $results = [System.Collections.Generic.List[PSCustomObject]]::new()

        $token = Get-MimecastAccessToken -BaseUrl $BaseUrl `
                     -ClientId $ClientId -ClientSecret $ClientSecret
    }

    process {
        foreach ($primary in $PrimaryAddresses) {
            $primary = $primary.Trim()
            if ([string]::IsNullOrWhiteSpace($primary)) { continue }

            if ($PSCmdlet.ShouldProcess($primary, "Add delegate access to [$DelegateAddress]")) {

                # Refresh token if it expired during a long batch
                $token = Get-MimecastAccessToken -BaseUrl $BaseUrl `
                             -ClientId $ClientId -ClientSecret $ClientSecret

                # POST /api/user/add-delegate-user -- both addresses in the request body.
                # InvertedMode swaps which address is treated as the mailbox owner vs delegate.
                $path = '/api/user/add-delegate-user'
                if ($InvertedMode) {
                    $body = @{ data = @(@{ primaryAddress = $primary; delegateAddress = $DelegateAddress }) }
                } else {
                    $body = @{ data = @(@{ primaryAddress = $DelegateAddress; delegateAddress = $primary }) }
                }

                try {
                    $response = Invoke-MimecastApiCall -BaseUrl $BaseUrl -Path $path `
                                    -Token $token -Method POST -Body $body

                    $delegateId = if ($response.data -and $response.data.Count -gt 0) {
                        $response.data[0].id
                    } else { $null }

                    $resultObj = [PSCustomObject]@{
                        PrimaryAddress  = $primary
                        DelegateAddress = $DelegateAddress
                        Success         = $true
                        DelegateId      = $delegateId
                        FailReason      = $null
                    }

                    $results.Add($resultObj)
                    Write-Verbose "  [OK] $primary -> $DelegateAddress  (id: $delegateId)"
                }
                catch {
                    $resultObj = [PSCustomObject]@{
                        PrimaryAddress  = $primary
                        DelegateAddress = $DelegateAddress
                        Success         = $false
                        DelegateId      = $null
                        FailReason      = $_.Exception.Message
                    }

                    $results.Add($resultObj)
                    Write-Warning "  [ERROR] $primary -> $DelegateAddress  $($resultObj.FailReason)"
                }

                if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
            }
        }
    }

    end {
        return $results
    }
}

# ---------------------------------------------------------------------------
# Public: List existing delegate users for a mailbox
# ---------------------------------------------------------------------------
function Get-MimecastDelegateUsers {
    <#
    .SYNOPSIS
        Returns the list of users who have delegate access to a mailbox.

    .DESCRIPTION
        Calls POST /api/user/find-delegate-users with the primary address in the
        request body and returns structured delegate objects.

        The API response shape is: { "data": [ { "delegateUsers": [...] } ], "fail": [] }
        The delegateUsers array is nested inside data[0]. Each entry is
        mapped to a PSCustomObject with the following properties:
            EmailAddress   - delegate's email address
            DisplayName    - delegate's display name (from the 'name' API field)
            DelegateId     - Mimecast secure ID required by Remove-MimecastDelegateUser
            Source         - provisioning source (e.g. 'ADCON')
            PrimaryAddress - the mailbox that was queried

        SHARED MAILBOXES: Delegate relationships added with -InvertedMode are stored
        against each user's account in Mimecast, not the shared mailbox. Querying the
        shared mailbox address directly returns an empty list. To list or remove those
        relationships, supply the user's address as -PrimaryAddress and filter results
        by EmailAddress.

    .PARAMETER PrimaryAddress
        The email address of the mailbox whose delegates you want to list.

    .PARAMETER BaseUrl
        Mimecast API 2.0 gateway base URL.
        Default: https://api.services.mimecast.com

    .PARAMETER ClientId
        Your Mimecast API 2.0 Client ID.

    .PARAMETER ClientSecret
        Your Mimecast API 2.0 Client Secret as a SecureString.

    .EXAMPLE
        # List all delegates and display key properties
        $secret = Read-Host 'Client Secret' -AsSecureString
        Get-MimecastDelegateUsers -PrimaryAddress 'shared@contoso.com' `
            -ClientId $env:MC_CLIENT_ID -ClientSecret $secret |`
            Format-Table EmailAddress, DisplayName, DelegateId, Source -AutoSize

    .EXAMPLE
        # Export delegate list to CSV
        $secret = Read-Host 'Client Secret' -AsSecureString
        Get-MimecastDelegateUsers -PrimaryAddress 'shared@contoso.com' `
            -ClientId $env:MC_CLIENT_ID -ClientSecret $secret |
            Export-Csv -Path "Delegates_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [string] $PrimaryAddress,

        [string]       $BaseUrl      = 'https://api.services.mimecast.com',
        [Parameter(Mandatory)] [string]       $ClientId,
        [Parameter(Mandatory)] [securestring] $ClientSecret
    )

    $ErrorActionPreference = 'Stop'

    $token = Get-MimecastAccessToken -BaseUrl $BaseUrl `
                 -ClientId $ClientId -ClientSecret $ClientSecret

    $path = '/api/user/find-delegate-users'
    $body = @{ data = @(@{ primaryAddress = $PrimaryAddress }) }

    try {
        $response = Invoke-MimecastApiCall -BaseUrl $BaseUrl -Path $path `
                        -Token $token -Method POST -Body $body

        # Real API response shape: { data: [ { delegateUsers: [...] } ], fail: [] }
        # delegateUsers is nested inside data[0]. Fall back through flatter shapes
        # used by some tenant variants.
        $delegates = if ($response.data -and $response.data[0].delegateUsers) {
                         $response.data[0].delegateUsers
                     } elseif ($response.delegateUsers) {
                         $response.delegateUsers
                     } elseif ($response.data) {
                         $response.data
                     } else {
                         @()
                     }

        if (-not $delegates -or $delegates.Count -eq 0) {
            Write-Verbose "[INFO] No delegates found for $PrimaryAddress"
            return @()
        }

        return $delegates | ForEach-Object {
            [PSCustomObject]@{
                EmailAddress   = if ($_.emailAddress) { $_.emailAddress } elseif ($_.delegateAddress) { $_.delegateAddress } else { $_.delegateEmailAddress }
                DisplayName    = $_.name
                DelegateId     = if ($_.id) { $_.id } elseif ($_.delegateId) { $_.delegateId } else { $_.secureId }
                Source         = $_.source
                PrimaryAddress = $PrimaryAddress
            }
        }
    }
    catch {
        Write-Error "Failed to retrieve delegates for ${PrimaryAddress}: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Public: Remove a delegate user from a mailbox
# ---------------------------------------------------------------------------
function Remove-MimecastDelegateUser {
    <#
    .SYNOPSIS
        Revokes a delegate user's access to a mailbox.

    .DESCRIPTION
        Calls POST /api/user/remove-delegate-user with the primary address and
        delegate record ID in the request body.
        Use Get-MimecastDelegateUsers first to obtain the DelegateId.

    .PARAMETER PrimaryAddress
        The email address of the mailbox to remove the delegate from.

    .PARAMETER DelegateId
        The delegate record ID returned by Add-MimecastDelegateUsers or
        Get-MimecastDelegateUsers.

    .PARAMETER BaseUrl
        Mimecast API 2.0 gateway base URL.
        Default: https://api.services.mimecast.com

    .PARAMETER ClientId
        Your Mimecast API 2.0 Client ID.

    .PARAMETER ClientSecret
        Your Mimecast API 2.0 Client Secret as a SecureString.

    .EXAMPLE
        $secret = Read-Host 'Client Secret' -AsSecureString
        $creds  = @{ ClientId = $env:MC_CLIENT_ID; ClientSecret = $secret }

        # Step 1: find the delegate record ID
        $delegate = Get-MimecastDelegateUsers -PrimaryAddress 'shared@contoso.com' @creds |
                        Where-Object EmailAddress -eq 'alice@contoso.com'

        # Step 2: remove it
        Remove-MimecastDelegateUser -PrimaryAddress 'shared@contoso.com' `
            -DelegateId $delegate.DelegateId @creds

    .EXAMPLE
        # WhatIf preview - no changes made
        Remove-MimecastDelegateUser -PrimaryAddress 'shared@contoso.com' `
            -DelegateId 'abc-123' @creds -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)] [string]       $PrimaryAddress,
        [Parameter(Mandatory)] [string]       $DelegateId,

        [string]       $BaseUrl      = 'https://api.services.mimecast.com',
        [Parameter(Mandatory)] [string]       $ClientId,
        [Parameter(Mandatory)] [securestring] $ClientSecret
    )

    $ErrorActionPreference = 'Stop'

    if ($PSCmdlet.ShouldProcess("$PrimaryAddress (delegate: $DelegateId)", 'Remove delegate access')) {

        $token = Get-MimecastAccessToken -BaseUrl $BaseUrl `
                     -ClientId $ClientId -ClientSecret $ClientSecret

        $path = '/api/user/remove-delegate-user'
        $body = @{ data = @(@{ primaryAddress = $PrimaryAddress; id = $DelegateId }) }

        try {
            Invoke-MimecastApiCall -BaseUrl $BaseUrl -Path $path `
                -Token $token -Method POST -Body $body | Out-Null

            Write-Verbose "[OK] Removed delegate $DelegateId from $PrimaryAddress"

            return [PSCustomObject]@{
                PrimaryAddress = $PrimaryAddress
                DelegateId     = $DelegateId
                Success        = $true
            }
        }
        catch {
            Write-Error "Failed to remove delegate $DelegateId from ${PrimaryAddress}: $($_.Exception.Message)"
        }
    }
}

# ---------------------------------------------------------------------------
# Diagnostic: Probe API path variants to find the correct base path
# ---------------------------------------------------------------------------
function Test-MimecastApiPath {
    <#
    .SYNOPSIS
        Probes known API path variants to identify the correct delegate endpoint.
    .DESCRIPTION
        Tests several path prefix combinations (with and without /api/) using a
        known valid user address to determine which variant the tenant responds to.
        Run this once when debugging 404 errors to confirm the correct path.
    .EXAMPLE
        $secret = Read-Host 'Client Secret' -AsSecureString
        Test-MimecastApiPath -KnownUserEmail 'jsmith@company.com' `
            -ClientId $env:MC_CLIENT_ID -ClientSecret $secret
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]       $KnownUserEmail,
        [string]               $BaseUrl  = 'https://api.services.mimecast.com',
        [Parameter(Mandatory)] [string]       $ClientId,
        [Parameter(Mandatory)] [securestring] $ClientSecret
    )

    $ErrorActionPreference = 'Stop'

    # Try the supplied BaseUrl first for the token, then test all base URL variants.
    # The login DC (e.g. login-usb.mimecast.com = US-B) determines the correct API base URL.
    $token = Get-MimecastAccessToken -BaseUrl $BaseUrl -ClientId $ClientId -ClientSecret $ClientSecret

    $encoded = [uri]::EscapeDataString($KnownUserEmail)

    # All known API 2.0 base URLs including DC-specific variants
    $baseUrlsToTry = @(
        'https://api.services.mimecast.com',       # Global (auto-routes)
        'https://us-api.services.mimecast.com',    # US instance
        'https://usb-api.mimecast.com',            # US-B legacy gateway
        'https://us-api.mimecast.com'              # US legacy gateway
    )

    # Correct API 2.0 delegate endpoint paths (all use POST)
    $probeBody = @{ data = @(@{ primaryAddress = $KnownUserEmail }) }
    $pathsToTry = @(
        '/api/user/find-delegate-users',
        '/api/user/add-delegate-user',
        '/api/user/remove-delegate-user'
    )

    Write-Host "`n[INFO] Your login DC appears to be US-B (login-usb.mimecast.com)." -ForegroundColor Cyan
    Write-Host "[INFO] Step 1: Testing a basic API 2.0 endpoint to confirm data plane is reachable..." -ForegroundColor Cyan

    $dataPlaneWorks = $false
    $workingBase    = $null
    foreach ($base in $baseUrlsToTry) {
        try {
            $headers = @{
                'Authorization' = "Bearer $token"
                'Content-Type'  = 'application/json'
                'Accept'        = 'application/json'
            }
            # /ttp/url/get-all-managed-urls is available on all API 2.0 Cloud Gateway tenants
            $null = Invoke-RestMethod -Method GET -Uri ($base.TrimEnd('/') + '/ttp/url/get-all-managed-urls') `
                        -Headers $headers -ErrorAction Stop
            Write-Host "[OK] API 2.0 data plane reachable at: $base" -ForegroundColor Green
            $dataPlaneWorks = $true
            $workingBase    = $base
            break
        }
        catch {
            $sc = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { '?' }
            Write-Host "  [$sc] $base/ttp/url/get-all-managed-urls" -ForegroundColor DarkYellow
        }
    }

    if (-not $dataPlaneWorks) {
        Write-Host "`n[WARNING] No API 2.0 data endpoint responded successfully on any base URL." -ForegroundColor Yellow
        Write-Host "" -ForegroundColor Yellow
        Write-Host "  ROOT CAUSE: Your Mimecast tenant is almost certainly on 'Email Security Cloud" -ForegroundColor Yellow
        Write-Host "  Integrated' NOT 'Email Security Cloud Gateway'. The Mimecast API 2.0 User" -ForegroundColor Yellow
        Write-Host "  Management endpoints (including delegates) are explicitly marked in the" -ForegroundColor Yellow
        Write-Host "  documentation as 'Email Security Cloud Gateway only' and are not yet" -ForegroundColor Yellow
        Write-Host "  available for Cloud Integrated tenants." -ForegroundColor Yellow
        Write-Host "" -ForegroundColor Yellow
        Write-Host "  CONFIRM TENANT TYPE:" -ForegroundColor Cyan
        Write-Host "    Mimecast Admin Console -> Administration -> Account -> Account Settings" -ForegroundColor Cyan
        Write-Host "    Look for 'Email Security Cloud Gateway' vs 'Cloud Integrated'" -ForegroundColor Cyan
        Write-Host "" -ForegroundColor Yellow
        Write-Host "  YOUR OPTIONS:" -ForegroundColor Cyan
        Write-Host "    Option A: Use Mimecast API 1.0 (HMAC auth) -- works on ALL tenant types" -ForegroundColor Cyan
        Write-Host "      Requires different credentials from Mimecast Admin:" -ForegroundColor Cyan
        Write-Host "        Administration -> Services -> API 1.0 Applications" -ForegroundColor Cyan
        Write-Host "      Credentials needed: App ID, App Key, Access Key, Secret Key" -ForegroundColor Cyan
        Write-Host "      Delegate endpoint: POST /api/delegate/add-delegate" -ForegroundColor Cyan
        Write-Host "" -ForegroundColor Yellow
        Write-Host "    Option B: Configure delegates manually via the Mimecast Admin Console:" -ForegroundColor Cyan
        Write-Host "      Administration -> Directories -> Internal Directories" -ForegroundColor Cyan
        Write-Host "      -> select user -> Identities -> Delegate Access -> Add" -ForegroundColor Cyan
        return
    }

    Write-Host "`n[INFO] Step 2: Testing User Management delegate path variants..." -ForegroundColor Cyan

    foreach ($base in $baseUrlsToTry) {
        Write-Host "--- BaseUrl: $base ---" -ForegroundColor DarkCyan
        foreach ($path in $pathsToTry) {
            try {
                $resp = Invoke-MimecastApiCall -BaseUrl $base -Path $path -Token $token -Method POST -Body $probeBody
                Write-Host "[OK] SUCCESS!" -ForegroundColor Green
                Write-Host "     BaseUrl : $base" -ForegroundColor Green
                Write-Host "     Path    : $path" -ForegroundColor Green
                Write-Host "`n[ACTION] Re-run your commands with -BaseUrl '$base'" -ForegroundColor Yellow
                return
            }
            catch {
                $sc     = if ($_ -match 'HTTP (\d+)') { $Matches[1] } else { '?' }
                $detail = if ($_ -match 'API detail: (.+)') { $Matches[1] } else { $_.Exception.Message }
                Write-Host "  [$sc] $path | $detail" -ForegroundColor Red
            }
        }
    }

    Write-Host "`n[WARNING] Data plane works but User Management endpoints all failed." -ForegroundColor Yellow
    Write-Host "  The API application Role may not include 'User and Group Management' permission." -ForegroundColor Yellow
    Write-Host "  Check: Mimecast Admin -> Services -> API and Platform Integrations" -ForegroundColor Yellow
    Write-Host "  -> Your application -> Role -> ensure 'User and Group Management' is assigned." -ForegroundColor Yellow
}

#endregion Public Functions

#region API 1.0 Helper Functions - HMAC Auth

function Invoke-MimecastV1ApiCall {
    <#
    .SYNOPSIS
        Sends a signed Mimecast API 1.0 request using HMAC-SHA1 authentication.
    .DESCRIPTION
        Builds the per-request Authorization header (MC scheme) and calls the
        Mimecast API 1.0 gateway. Use this for tenants where API 2.0 User
        Management endpoints are not available (Cloud Integrated / US-B DC).
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory)] [string]    $BaseUrl,
        [Parameter(Mandatory)] [string]    $Uri,
        [Parameter(Mandatory)] [string]    $AppId,
        [Parameter(Mandatory)] [string]    $AppKey,
        [Parameter(Mandatory)] [string]    $AccessKey,
        [Parameter(Mandatory)] [securestring] $SecretKey,
        [hashtable] $Data = @{}
    )

    # Decode SecureString secret key
    $bstr      = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecretKey)
    $plainSKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

    try {
        $reqId   = [System.Guid]::NewGuid().ToString()
        $date    = (Get-Date).ToUniversalTime().ToString('ddd, dd MMM yyyy HH:mm:ss UTC')

        # Signature data: date:requestId:uri:appKey
        $sigData  = "$date`:$reqId`:$Uri`:$AppKey"
        $keyBytes = [System.Convert]::FromBase64String($plainSKey)
        $hmac     = [System.Security.Cryptography.HMACSHA1]::new($keyBytes)
        $sigBytes = $hmac.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($sigData))
        $hmac.Dispose()
        $signature = [System.Convert]::ToBase64String($sigBytes)

        $body = @{
            meta = @{ requestId = $reqId }
            data = @($Data)
        } | ConvertTo-Json -Depth 10

        $headers = @{
            'Authorization' = "MC $AccessKey`:$signature"
            'x-mc-app-id'   = $AppId
            'x-mc-date'     = $date
            'x-mc-req-id'   = $reqId
            'Content-Type'  = 'application/json'
            'Accept'        = 'application/json'
        }

        $response = Invoke-RestMethod -Method Post `
                        -Uri ($BaseUrl.TrimEnd('/') + $Uri) `
                        -Headers $headers `
                        -Body $body `
                        -ErrorAction Stop

        return $response
    }
    catch {
        $statusCode = $null
        $rawBody    = $null
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $rawBody = $_.ErrorDetails.Message
        }
        if (-not $rawBody -and $_.Exception.Response) {
            try {
                $stream  = $_.Exception.Response.GetResponseStream()
                $reader  = [System.IO.StreamReader]::new($stream)
                $rawBody = $reader.ReadToEnd()
                $reader.Dispose()
            } catch {}
        }
        $msg = "HTTP $statusCode - $($_.Exception.Message)"
        if ($rawBody) { $msg += " | Response: $rawBody" }
        throw $msg
    }
    finally {
        $plainSKey = [string]::Empty
    }
}

#endregion API 1.0 Helper Functions - HMAC Auth

#region API 1.0 Public Functions

# ---------------------------------------------------------------------------
# Public (API 1.0): Add delegate access to a mailbox
# ---------------------------------------------------------------------------
function Add-MimecastDelegateV1 {
    <#
    .SYNOPSIS
        Grants one or more users delegate access to a mailbox using Mimecast API 1.0.

    .DESCRIPTION
        Calls POST /api/delegate/add-delegate for each primary address supplied.
        Use this function when your Mimecast tenant is on Email Security Cloud
        Integrated or the US-B data center (login-usb.mimecast.com), where the
        API 2.0 User Management endpoints return 404.

        Credentials are obtained from Mimecast Admin Console:
          Administration -> Services -> API 1.0 Applications (App ID + App Key)
          Account -> Roles -> [role] -> API Application Authentication
            (Access Key + Secret Key - per user/application binding)

    .PARAMETER DelegateAddress
        The mailbox to be shared (the one others will gain access to).

    .PARAMETER PrimaryAddresses
        One or more email addresses that will receive delegate access.

    .PARAMETER BaseUrl
        Mimecast API 1.0 base URL. Use the URL matching your login DC:
          US-A : https://us-api.mimecast.com
          US-B : https://usb-api.mimecast.com
          EU   : https://eu-api.mimecast.com
          DE   : https://de-api.mimecast.com
          CA   : https://ca-api.mimecast.com

    .PARAMETER AppId
        Application ID from the API 1.0 application registration.

    .PARAMETER AppKey
        Application Key from the API 1.0 application registration.

    .PARAMETER AccessKey
        Access Key from the API application authentication binding.

    .PARAMETER SecretKey
        Secret Key from the API application authentication binding (SecureString).

    .PARAMETER DelayMs
        Milliseconds to pause between API calls. Default: 300.

    .EXAMPLE
        $secretKey = Read-Host 'Secret Key' -AsSecureString
        Add-MimecastDelegateV1 `
            -DelegateAddress  'shared@company.com' `
            -PrimaryAddresses 'jsmith@company.com' `
            -BaseUrl    'https://usb-api.mimecast.com' `
            -AppId      'your-app-id' `
            -AppKey     'your-app-key' `
            -AccessKey  'your-access-key' `
            -SecretKey  $secretKey

    .EXAMPLE
        # Bulk add from CSV
        $secretKey = Read-Host 'Secret Key' -AsSecureString
        $v1Creds = @{
            BaseUrl   = 'https://usb-api.mimecast.com'
            AppId     = 'your-app-id'
            AppKey    = 'your-app-key'
            AccessKey = 'your-access-key'
            SecretKey = $secretKey
        }
        $users = Import-Csv .\delegates.csv | Select-Object -ExpandProperty PrimaryAddress
        Add-MimecastDelegateV1 -DelegateAddress 'shared@company.com' `
            -PrimaryAddresses $users @v1Creds |
            Format-Table PrimaryAddress, Success, FailReason -AutoSize
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'DelayMs',
        Justification = 'Used inside the process block foreach loop')]
    param(
        [Parameter(Mandatory)]
        [string] $DelegateAddress,

        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowEmptyString()]
        [string[]] $PrimaryAddresses,

        [string] $BaseUrl   = 'https://usb-api.mimecast.com',
        [Parameter(Mandatory)] [string]        $AppId,
        [Parameter(Mandatory)] [string]        $AppKey,
        [Parameter(Mandatory)] [string]        $AccessKey,
        [Parameter(Mandatory)] [securestring]  $SecretKey,

        [int] $DelayMs = 300
    )

    begin {
        $ErrorActionPreference = 'Stop'
        $results = [System.Collections.Generic.List[PSCustomObject]]::new()
        $v1Params = @{
            BaseUrl   = $BaseUrl
            AppId     = $AppId
            AppKey    = $AppKey
            AccessKey = $AccessKey
            SecretKey = $SecretKey
        }
    }

    process {
        foreach ($primary in $PrimaryAddresses) {
            $primary = $primary.Trim()
            if ([string]::IsNullOrWhiteSpace($primary)) { continue }

            if ($PSCmdlet.ShouldProcess($primary, "Add delegate access to [$DelegateAddress]")) {
                try {
                    $response = Invoke-MimecastV1ApiCall @v1Params `
                        -Uri  '/api/delegate/add-delegate' `
                        -Data @{
                            delegate = @{ emailAddress = $primary }
                            mailbox  = @{ emailAddress = $DelegateAddress }
                        }

                    # API 1.0 returns fail array on error even with HTTP 200
                    if ($response.fail -and $response.fail.Count -gt 0) {
                        $errMsg = ($response.fail[0].errors | ForEach-Object { $_.message }) -join '; '
                        throw $errMsg
                    }

                    $delegateId = if ($response.data -and $response.data[0].id) {
                        $response.data[0].id
                    } else { $null }

                    $results.Add([PSCustomObject]@{
                        PrimaryAddress  = $primary
                        DelegateAddress = $DelegateAddress
                        Success         = $true
                        DelegateId      = $delegateId
                        FailReason      = $null
                    })
                    Write-Verbose "  [OK] $primary -> $DelegateAddress"
                }
                catch {
                    $results.Add([PSCustomObject]@{
                        PrimaryAddress  = $primary
                        DelegateAddress = $DelegateAddress
                        Success         = $false
                        DelegateId      = $null
                        FailReason      = $_.Exception.Message
                    })
                    Write-Warning "  [ERROR] $primary -> $DelegateAddress  $($_.Exception.Message)"
                }

                if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
            }
        }
    }

    end { return $results }
}

# ---------------------------------------------------------------------------
# Public (API 1.0): List delegates for a mailbox
# ---------------------------------------------------------------------------
function Get-MimecastDelegateV1 {
    <#
    .SYNOPSIS
        Returns users who have delegate access to a mailbox using Mimecast API 1.0.

    .PARAMETER DelegateAddress
        The shared mailbox whose delegate list you want to retrieve.

    .PARAMETER BaseUrl
        Mimecast API 1.0 base URL (default: https://usb-api.mimecast.com).

    .PARAMETER AppId
        Application ID from the API 1.0 application registration.

    .PARAMETER AppKey
        Application Key from the API 1.0 application registration.

    .PARAMETER AccessKey
        Access Key from the API application authentication binding.

    .PARAMETER SecretKey
        Secret Key from the API application authentication binding (SecureString).

    .EXAMPLE
        $secretKey = Read-Host 'Secret Key' -AsSecureString
        Get-MimecastDelegateV1 -DelegateAddress 'shared@company.com' `
            -BaseUrl 'https://usb-api.mimecast.com' `
            -AppId 'your-app-id' -AppKey 'your-app-key' `
            -AccessKey 'your-access-key' -SecretKey $secretKey
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)] [string]        $DelegateAddress,
        [string]                               $BaseUrl    = 'https://usb-api.mimecast.com',
        [Parameter(Mandatory)] [string]        $AppId,
        [Parameter(Mandatory)] [string]        $AppKey,
        [Parameter(Mandatory)] [string]        $AccessKey,
        [Parameter(Mandatory)] [securestring]  $SecretKey
    )

    $ErrorActionPreference = 'Stop'

    try {
        $response = Invoke-MimecastV1ApiCall `
            -BaseUrl   $BaseUrl `
            -Uri       '/api/delegate/get-delegates' `
            -AppId     $AppId `
            -AppKey    $AppKey `
            -AccessKey $AccessKey `
            -SecretKey $SecretKey `
            -Data      @{ mailbox = @{ emailAddress = $DelegateAddress } }

        if ($response.fail -and $response.fail.Count -gt 0) {
            $errMsg = ($response.fail[0].errors | ForEach-Object { $_.message }) -join '; '
            Write-Error "API error: $errMsg"
            return
        }

        if (-not $response.data -or $response.data.Count -eq 0) {
            Write-Verbose "[INFO] No delegates found for $DelegateAddress"
            return @()
        }

        $response.data | ForEach-Object {
            [PSCustomObject]@{
                PrimaryAddress  = $_.delegate.emailAddress
                DelegateAddress = $DelegateAddress
                DelegateId      = $_.id
            }
        }
    }
    catch {
        Write-Error "Failed to retrieve delegates for ${DelegateAddress}: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Public (API 1.0): Remove a delegate from a mailbox
# ---------------------------------------------------------------------------
function Remove-MimecastDelegateV1 {
    <#
    .SYNOPSIS
        Revokes a user's delegate access to a mailbox using Mimecast API 1.0.

    .PARAMETER DelegateAddress
        The shared mailbox to remove the delegate from.

    .PARAMETER PrimaryAddress
        The email address of the user whose delegate access will be removed.

    .PARAMETER BaseUrl
        Mimecast API 1.0 base URL (default: https://usb-api.mimecast.com).

    .PARAMETER AppId, AppKey, AccessKey, SecretKey
        API 1.0 credentials (see Add-MimecastDelegateV1 for details).

    .EXAMPLE
        $secretKey = Read-Host 'Secret Key' -AsSecureString
        Remove-MimecastDelegateV1 `
            -DelegateAddress 'shared@company.com' `
            -PrimaryAddress  'jsmith@company.com' `
            -BaseUrl 'https://usb-api.mimecast.com' `
            -AppId 'your-app-id' -AppKey 'your-app-key' `
            -AccessKey 'your-access-key' -SecretKey $secretKey
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)] [string]        $DelegateAddress,
        [Parameter(Mandatory)] [string]        $PrimaryAddress,
        [string]                               $BaseUrl    = 'https://usb-api.mimecast.com',
        [Parameter(Mandatory)] [string]        $AppId,
        [Parameter(Mandatory)] [string]        $AppKey,
        [Parameter(Mandatory)] [string]        $AccessKey,
        [Parameter(Mandatory)] [securestring]  $SecretKey
    )

    $ErrorActionPreference = 'Stop'

    if ($PSCmdlet.ShouldProcess("$PrimaryAddress from [$DelegateAddress]", 'Remove delegate access')) {
        try {
            $response = Invoke-MimecastV1ApiCall `
                -BaseUrl   $BaseUrl `
                -Uri       '/api/delegate/remove-delegate' `
                -AppId     $AppId `
                -AppKey    $AppKey `
                -AccessKey $AccessKey `
                -SecretKey $SecretKey `
                -Data      @{
                    delegate = @{ emailAddress = $PrimaryAddress }
                    mailbox  = @{ emailAddress = $DelegateAddress }
                }

            if ($response.fail -and $response.fail.Count -gt 0) {
                $errMsg = ($response.fail[0].errors | ForEach-Object { $_.message }) -join '; '
                throw $errMsg
            }

            Write-Verbose "[OK] Removed $PrimaryAddress from $DelegateAddress"
            return [PSCustomObject]@{
                PrimaryAddress  = $PrimaryAddress
                DelegateAddress = $DelegateAddress
                Success         = $true
            }
        }
        catch {
            Write-Error "Failed to remove delegate: $($_.Exception.Message)"
        }
    }
}

#endregion API 1.0 Public Functions

# Function exports are controlled by the root module manifest (pwshMimecast.psd1).
# Do not call Export-ModuleMember here; doing so when dot-sourced by pwshMimecast.psm1
# overrides the manifest's FunctionsToExport and causes Get-Command to return nothing.

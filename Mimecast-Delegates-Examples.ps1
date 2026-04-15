<#
.SYNOPSIS
    Example usage of the Mimecast-Delegates module (API 2.0).

.DESCRIPTION
    Demonstrates how to add multiple delegate users to a mailbox,
    list existing delegates, and remove a delegate using Mimecast API 2.0
    OAuth client credentials authentication.

    Includes examples for both normal user mailboxes and shared/functional
    mailboxes (which require the -InvertedMode switch).

    Run this after importing the module:
        Import-Module .\Mimecast-Delegates.psm1
#>

# ---------------------------------------------------------------------------
# 1. Store credentials
#    - ClientId can live in an environment variable (not sensitive)
#    - ClientSecret must be a SecureString -- never plain text
# ---------------------------------------------------------------------------
$secret    = Read-Host 'Mimecast Client Secret' -AsSecureString
$credSplat = @{
    ClientId     = $env:MC_CLIENT_ID   # e.g. set in your profile or pipeline
    ClientSecret = $secret
    # BaseUrl defaults to 'https://us-api.services.mimecast.com' (US region)
    # Override for other regions:
    #   EU: 'https://eu-api.services.mimecast.com'
    #   UK: 'https://uk-api.services.mimecast.com'
    #   DE: 'https://de-api.services.mimecast.com'
}

# ---------------------------------------------------------------------------
# 2a. Add a fixed list of delegates to a shared mailbox
# ---------------------------------------------------------------------------
$delegateMailbox = 'shared-inbox@contoso.com'

$usersToAdd = @(
    'alice@contoso.com'
    'bob@contoso.com'
    'carol@contoso.com'
)

Write-Host "`n[INFO] Adding delegates to $delegateMailbox" -ForegroundColor Cyan

$results = Add-MimecastDelegateUsers `
    -DelegateAddress  $delegateMailbox `
    -PrimaryAddresses $usersToAdd `
    -Verbose `
    @credSplat

# Show a summary table
$results | Format-Table PrimaryAddress, Success, DelegateId, FailReason -AutoSize


# ---------------------------------------------------------------------------
# 2b. Bulk-load delegates from a CSV file (normal user mailbox)
#     CSV format:  PrimaryAddress
#                  user1@contoso.com
#                  user2@contoso.com
# ---------------------------------------------------------------------------
<#
$csvPath  = '.\delegates-to-add.csv'
$csvUsers = Import-Csv $csvPath | Select-Object -ExpandProperty PrimaryAddress

$results = Add-MimecastDelegateUsers `
    -DelegateAddress  $delegateMailbox `
    -PrimaryAddresses $csvUsers `
    @credSplat

$outputPath = "delegate-results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation
Write-Host "[OK] Results exported to $outputPath"
#>


# ---------------------------------------------------------------------------
# 2c. Bulk-load delegates for a SHARED / FUNCTIONAL mailbox (-InvertedMode)
#
#     Use -InvertedMode when the target mailbox (DelegateAddress) is a shared
#     or functional mailbox that is not a full user account in Mimecast.
#     Without this switch the API returns HTTP 404.
#
#     The switch inverts the API call:
#       Standard : POST /users/{delegateAddress}/delegates  body: { primary }
#       Inverted : POST /users/{primaryAddress}/delegates   body: { shared }
#
#     CSV format:  PrimaryAddress
#                  user1@contoso.com
#                  user2@contoso.com
# ---------------------------------------------------------------------------
<#
$sharedMailbox = 'materialreceipt@contoso.com'
$csvPath       = '.\delegates-to-add.csv'
$csvUsers      = Import-Csv $csvPath | Select-Object -ExpandProperty PrimaryAddress

Write-Host "`n[INFO] Adding delegates to shared mailbox $sharedMailbox" -ForegroundColor Cyan

$results = Add-MimecastDelegateUsers `
    -DelegateAddress  $sharedMailbox `
    -PrimaryAddresses $csvUsers `
    -InvertedMode `
    -Verbose `
    @credSplat

$results | Format-Table PrimaryAddress, Success, DelegateId, FailReason -AutoSize

$outputPath = "delegate-results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation
Write-Host "[OK] Results exported to $outputPath"
#>


# ---------------------------------------------------------------------------
# 2d. Pipe addresses in from a plain text file (one per line)
# ---------------------------------------------------------------------------
<#
Get-Content '.\users.txt' |
    Add-MimecastDelegateUsers -DelegateAddress $delegateMailbox @credSplat |
    Format-Table -AutoSize
#>


# ---------------------------------------------------------------------------
# 3. List current delegates for a mailbox
# ---------------------------------------------------------------------------
Write-Host "`n[INFO] Current delegates for $delegateMailbox" -ForegroundColor Cyan

$existing = Get-MimecastDelegateUsers `
    -PrimaryAddress $delegateMailbox `
    @credSplat

if ($existing) {
    $existing | Format-Table EmailAddress, Id -AutoSize
} else {
    Write-Host '[INFO] No delegates found.' -ForegroundColor Yellow
}


# ---------------------------------------------------------------------------
# 4. Remove a delegate (prompts for confirmation by default)
#    API 2.0 requires the delegate record ID -- look it up first
# ---------------------------------------------------------------------------
<#
Write-Host "`n[INFO] Removing alice from $delegateMailbox" -ForegroundColor Yellow

$delegateRecord = Get-MimecastDelegateUsers -PrimaryAddress $delegateMailbox @credSplat |
                      Where-Object EmailAddress -eq 'alice@contoso.com'

if ($delegateRecord) {
    $removeResult = Remove-MimecastDelegateUser `
        -PrimaryAddress $delegateMailbox `
        -DelegateId     $delegateRecord.Id `
        @credSplat

    $removeResult | Format-List
} else {
    Write-Host '[WARNING] No delegate record found for alice@contoso.com' -ForegroundColor Yellow
}
#>


# ---------------------------------------------------------------------------
# 5. Audit: export all delegates for a list of shared mailboxes
# ---------------------------------------------------------------------------
<#
$sharedMailboxes = @(
    'helpdesk@contoso.com'
    'finance@contoso.com'
    'legal@contoso.com'
)

$audit = foreach ($mailbox in $sharedMailboxes) {
    $delegates = Get-MimecastDelegateUsers -PrimaryAddress $mailbox @credSplat
    foreach ($d in $delegates) {
        [PSCustomObject]@{
            SharedMailbox = $mailbox
            DelegateEmail = $d.EmailAddress
            DelegateId    = $d.Id
        }
    }
}

$outputPath = "delegate-audit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$audit | Export-Csv -Path $outputPath -NoTypeInformation
Write-Host "[OK] Audit exported to $outputPath ($($audit.Count) records)"
#>

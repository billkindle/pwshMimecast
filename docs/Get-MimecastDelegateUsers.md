---
external help file: Mimecast-Delegates-help.xml
Module Name: Mimecast-Delegates
online version:
schema: 2.0.0
---

# Get-MimecastDelegateUsers

## SYNOPSIS

Returns the list of users who have delegate access to a mailbox.

## SYNTAX

```text
Get-MimecastDelegateUsers [-PrimaryAddress] <String> [[-BaseUrl] <String>] [-ClientId] <String>
 [-ClientSecret] <SecureString> [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION

Calls POST /api/user/find-delegate-users with the primary address in the
request body and returns structured delegate objects.

The API response shape is `{ "data": [ { "delegateUsers": [...] } ], "fail": [] }`. The
`delegateUsers` array is nested inside `data[0]`. Each entry is mapped to
a PSCustomObject with the following properties:

- **EmailAddress** — delegate's email address
- **DisplayName** — delegate's display name
- **DelegateId** — Mimecast secure ID required by Remove-MimecastDelegateUser
- **Source** — provisioning source (e.g. `ADCON`)
- **PrimaryAddress** — the mailbox that was queried

**Shared mailboxes:** Delegate relationships added with `-InvertedMode` are stored against
the *user's* account in Mimecast, not the shared mailbox. Querying the shared mailbox
address directly will return an empty list. To list or remove those relationships, supply
the **user's address** as `-PrimaryAddress` and filter the results by `EmailAddress`.

## EXAMPLES

### EXAMPLE 1

```powershell
$secret = Read-Host 'Client Secret' -AsSecureString
Get-MimecastDelegateUsers -PrimaryAddress 'shared@contoso.com' `
    -ClientId $env:MC_CLIENT_ID -ClientSecret $secret |`
    Format-Table EmailAddress, DisplayName, DelegateId, Source -AutoSize
```powershell

### EXAMPLE 2

```powershell
# Export delegate list to CSV
$secret = Read-Host 'Client Secret' -AsSecureString
Get-MimecastDelegateUsers -PrimaryAddress 'shared@contoso.com' `
    -ClientId $env:MC_CLIENT_ID -ClientSecret $secret |
    Export-Csv -Path "Delegates_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation
```powershell

## PARAMETERS

### -PrimaryAddress

The email address of the mailbox whose delegates you want to list.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### -BaseUrl

Mimecast API 2.0 gateway base URL.
Default: <https://api.services.mimecast.com>

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: Https://api.services.mimecast.com
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### -ClientId

Your Mimecast API 2.0 Client ID.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### -ClientSecret

Your Mimecast API 2.0 Client Secret as a SecureString.

```yaml
Type: SecureString
Parameter Sets: (All)
Aliases:

Required: True
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### -ProgressAction

{{ Fill ProgressAction Description }}

```yaml
Type: ActionPreference
Parameter Sets: (All)
Aliases: proga

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable,
-Verbose, -WarningAction, and -WarningVariable.
For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## OUTPUTS

### System.Management.Automation.PSObject

Returns a PSCustomObject per delegate with properties: EmailAddress, DisplayName, DelegateId, Source, PrimaryAddress.

## NOTES

## RELATED LINKS

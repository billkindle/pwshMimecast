---
external help file: Mimecast-Delegates-help.xml
Module Name: Mimecast-Delegates
online version:
schema: 2.0.0
---

# Add-MimecastDelegateUsers

## SYNOPSIS

Grants one or more users delegate access to a target mailbox.

## SYNTAX

```text
Add-MimecastDelegateUsers [-DelegateAddress] <String> [-PrimaryAddresses] <String[]> [-InvertedMode]
 [[-BaseUrl] <String>] [-ClientId] <String> [-ClientSecret] <SecureString> [[-DelayMs] <Int32>]
 [-ProgressAction <ActionPreference>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION

Calls POST /api/user/add-delegate-user for each primary address supplied.
Both addresses are sent in the request body -- no address in the URL path.

Request body: { "data": \[{ "primaryAddress": "\<mailbox-owner\>",
"delegateAddress": "\<user-getting-access\>" }\] }

Default mode:
    primaryAddress = DelegateAddress (the mailbox being shared)
    delegateAddress = each PrimaryAddress (the user receiving access)

Inverted mode (-InvertedMode switch, for shared/functional mailboxes):
    primaryAddress = each PrimaryAddress
    delegateAddress = DelegateAddress
    Use if the default mode returns errors indicating an address is not
    found as a primary mailbox holder.

## EXAMPLES

### EXAMPLE 1

```powershell
# Standard mode - full user mailbox
$secret = Read-Host 'Client Secret' -AsSecureString
Add-MimecastDelegateUsers `
    -DelegateAddress  'shared@contoso.com' `
    -PrimaryAddresses 'alice@contoso.com','bob@contoso.com' `
    -ClientId         $env:MC_CLIENT_ID `
    -ClientSecret     $secret
```powershell

### EXAMPLE 2

```powershell
# Inverted mode - shared/functional mailbox as target
$secret = Read-Host 'Client Secret' -AsSecureString
$users  = Import-Csv .\delegates.csv | Select-Object -ExpandProperty PrimaryAddress
Add-MimecastDelegateUsers `
    -DelegateAddress  'materialreceipt@contoso.com' `
    -PrimaryAddresses $users `
    -InvertedMode `
    -ClientId         $env:MC_CLIENT_ID `
    -ClientSecret     $secret
```powershell

### EXAMPLE 3

```powershell
# WhatIf preview - no changes made
Add-MimecastDelegateUsers -DelegateAddress 'shared@contoso.com' `
    -PrimaryAddresses 'alice@contoso.com' @creds -WhatIf
```powershell

## PARAMETERS

### -DelegateAddress

The mailbox to be shared (the mailbox others will access).

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

### -PrimaryAddresses

One or more email addresses that will receive delegate access.
Accepts pipeline input and array values.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```powershell

### -InvertedMode

Use this switch when DelegateAddress is a shared/functional mailbox that
is not a full user account in Mimecast.
Inverts the request body so each
PrimaryAddress becomes the body's primaryAddress field and DelegateAddress
becomes the delegateAddress field.
Use if the default mode returns errors
indicating an address is not found as a primary mailbox holder.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
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
Position: 3
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
Position: 4
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
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### -DelayMs

Milliseconds to pause between API calls.
Default: 300.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: 300
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### -WhatIf

Shows what would happen if the cmdlet runs.
The cmdlet is not run.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### -Confirm

Prompts you for confirmation before running the cmdlet.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

Required: False
Position: Named
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

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

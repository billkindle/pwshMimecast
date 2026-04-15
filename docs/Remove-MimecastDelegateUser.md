---
external help file: Mimecast-Delegates-help.xml
Module Name: Mimecast-Delegates
online version:
schema: 2.0.0
---

# Remove-MimecastDelegateUser

## SYNOPSIS

Revokes a delegate user's access to a mailbox.

## SYNTAX

```text
Remove-MimecastDelegateUser [-PrimaryAddress] <String> [-DelegateId] <String> [[-BaseUrl] <String>]
 [-ClientId] <String> [-ClientSecret] <SecureString> [-ProgressAction <ActionPreference>]
 [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION

Calls POST /api/user/remove-delegate-user with the primary address and
delegate record ID in the request body.
Use Get-MimecastDelegateUsers first to obtain the DelegateId.

## EXAMPLES

### EXAMPLE 1

```powershell
$secret = Read-Host 'Client Secret' -AsSecureString
$creds  = @{ ClientId = $env:MC_CLIENT_ID; ClientSecret = $secret }

# Step 1: find the delegate record ID
$delegate = Get-MimecastDelegateUsers -PrimaryAddress 'shared@contoso.com' @creds |
                Where-Object EmailAddress -eq 'alice@contoso.com'

# Step 2: remove it
Remove-MimecastDelegateUser -PrimaryAddress 'shared@contoso.com' `
    -DelegateId $delegate.DelegateId @creds
```powershell

### EXAMPLE 2

```powershell
# WhatIf preview - no changes made
Remove-MimecastDelegateUser -PrimaryAddress 'shared@contoso.com' `
    -DelegateId 'abc-123' @creds -WhatIf
```powershell

## PARAMETERS

### -PrimaryAddress

The email address of the mailbox to remove the delegate from.

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

### -DelegateId

The delegate record ID returned by Add-MimecastDelegateUsers or
Get-MimecastDelegateUsers.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
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

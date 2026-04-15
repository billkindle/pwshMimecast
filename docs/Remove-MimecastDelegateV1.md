---
external help file: Mimecast-Delegates-help.xml
Module Name: Mimecast-Delegates
online version:
schema: 2.0.0
---

# Remove-MimecastDelegateV1

## SYNOPSIS

Revokes a user's delegate access to a mailbox using Mimecast API 1.0.

## SYNTAX

```text
Remove-MimecastDelegateV1 [-DelegateAddress] <String> [-PrimaryAddress] <String>
 [[-BaseUrl] <String>] [-AppId] <String> [-AppKey] <String> [-AccessKey] <String>
 [-SecretKey] <SecureString>
 [-ProgressAction <ActionPreference>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION

{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1

```powershell
$secretKey = Read-Host 'Secret Key' -AsSecureString
Remove-MimecastDelegateV1 `
    -DelegateAddress 'shared@company.com' `
    -PrimaryAddress  'jsmith@company.com' `
    -BaseUrl 'https://usb-api.mimecast.com' `
    -AppId 'your-app-id' -AppKey 'your-app-key' `
    -AccessKey 'your-access-key' -SecretKey $secretKey
```powershell

## PARAMETERS

### -DelegateAddress

The shared mailbox to remove the delegate from.

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

### -PrimaryAddress

The email address of the user whose delegate access will be removed.

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

Mimecast API 1.0 base URL (default: <https://usb-api.mimecast.com>).

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: Https://usb-api.mimecast.com
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### -AppId

{{ Fill AppId Description }}

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

### -AppKey

{{ Fill AppKey Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### -AccessKey

{{ Fill AccessKey Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### -SecretKey

{{ Fill SecretKey Description }}

```yaml
Type: SecureString
Parameter Sets: (All)
Aliases:

Required: True
Position: 7
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

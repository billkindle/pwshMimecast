---
external help file: Mimecast-Delegates-help.xml
Module Name: Mimecast-Delegates
online version:
schema: 2.0.0
---

# Get-MimecastDelegateV1

## SYNOPSIS

Returns users who have delegate access to a mailbox using Mimecast API 1.0.

## SYNTAX

```text
Get-MimecastDelegateV1 [-DelegateAddress] <String> [[-BaseUrl] <String>] [-AppId] <String>
 [-AppKey] <String> [-AccessKey] <String> [-SecretKey] <SecureString>
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION

{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1

```powershell
$secretKey = Read-Host 'Secret Key' -AsSecureString
Get-MimecastDelegateV1 -DelegateAddress 'shared@company.com' `
    -BaseUrl 'https://usb-api.mimecast.com' `
    -AppId 'your-app-id' -AppKey 'your-app-key' `
    -AccessKey 'your-access-key' -SecretKey $secretKey
```powershell

## PARAMETERS

### -DelegateAddress

The shared mailbox whose delegate list you want to retrieve.

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

Mimecast API 1.0 base URL (default: <https://usb-api.mimecast.com>).

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: Https://usb-api.mimecast.com
Accept pipeline input: False
Accept wildcard characters: False
```powershell

### -AppId

Application ID from the API 1.0 application registration.

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

### -AppKey

Application Key from the API 1.0 application registration.

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

### -AccessKey

Access Key from the API application authentication binding.

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

### -SecretKey

Secret Key from the API application authentication binding (SecureString).

```yaml
Type: SecureString
Parameter Sets: (All)
Aliases:

Required: True
Position: 6
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

### System.Management.Automation.PSObject

## NOTES

## RELATED LINKS

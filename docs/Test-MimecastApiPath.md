---
external help file: Mimecast-Delegates-help.xml
Module Name: Mimecast-Delegates
online version:
schema: 2.0.0
---

# Test-MimecastApiPath

## SYNOPSIS

Probes known API path variants to identify the correct delegate endpoint.

## SYNTAX

```text
Test-MimecastApiPath [-KnownUserEmail] <String> [[-BaseUrl] <String>] [-ClientId] <String>
 [-ClientSecret] <SecureString> [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION

Tests several path prefix combinations (with and without /api/) using a
known valid user address to determine which variant the tenant responds to.
Run this once when debugging 404 errors to confirm the correct path.

## EXAMPLES

### EXAMPLE 1

```powershell
$secret = Read-Host 'Client Secret' -AsSecureString
Test-MimecastApiPath -KnownUserEmail 'jsmith@company.com' `
    -ClientId $env:MC_CLIENT_ID -ClientSecret $secret
```powershell

## PARAMETERS

### -KnownUserEmail

{{ Fill KnownUserEmail Description }}

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

{{ Fill BaseUrl Description }}

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

{{ Fill ClientId Description }}

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

{{ Fill ClientSecret Description }}

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

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

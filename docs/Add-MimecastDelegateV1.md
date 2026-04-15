---
external help file: Mimecast-Delegates-help.xml
Module Name: Mimecast-Delegates
online version:
schema: 2.0.0
---

# Add-MimecastDelegateV1

## SYNOPSIS

Grants one or more users delegate access to a mailbox using Mimecast API 1.0.

## SYNTAX

```text
Add-MimecastDelegateV1 [-DelegateAddress] <String> [-PrimaryAddresses] <String[]>
 [[-BaseUrl] <String>] [-AppId] <String> [-AppKey] <String> [-AccessKey] <String>
 [-SecretKey] <SecureString> [[-DelayMs] <Int32>]
 [-ProgressAction <ActionPreference>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION

Calls POST /api/delegate/add-delegate for each primary address supplied.
Use this function when your Mimecast tenant is on Email Security Cloud
Integrated or the US-B data center (login-usb.mimecast.com), where the
API 2.0 User Management endpoints return 404.

Credentials are obtained from Mimecast Admin Console:
  Administration -\> Services -\> API 1.0 Applications (App ID + App Key)
  Account -\> Roles -\> \[role\] -\> API Application Authentication
    (Access Key + Secret Key - per user/application binding)

## EXAMPLES

### EXAMPLE 1

```powershell
$secretKey = Read-Host 'Secret Key' -AsSecureString
Add-MimecastDelegateV1 `
    -DelegateAddress  'shared@company.com' `
    -PrimaryAddresses 'jsmith@company.com' `
    -BaseUrl    'https://usb-api.mimecast.com' `
    -AppId      'your-app-id' `
    -AppKey     'your-app-key' `
    -AccessKey  'your-access-key' `
    -SecretKey  $secretKey
```powershell

### EXAMPLE 2

```powershell
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
```powershell

## PARAMETERS

### -DelegateAddress

The mailbox to be shared (the one others will gain access to).

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

### -BaseUrl

Mimecast API 1.0 base URL.
Use the URL matching your login DC:
  US-A : <https://us-api.mimecast.com>
  US-B : <https://usb-api.mimecast.com>
  EU   : <https://eu-api.mimecast.com>
  DE   : <https://de-api.mimecast.com>
  CA   : <https://ca-api.mimecast.com>

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

Application ID from the API 1.0 application registration.

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

Application Key from the API 1.0 application registration.

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

Access Key from the API application authentication binding.

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

Secret Key from the API application authentication binding (SecureString).

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

### -DelayMs

Milliseconds to pause between API calls.
Default: 300.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
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

#Requires -Module Pester
#Requires -Version 5.1

<#
.SYNOPSIS
    Pester 5 tests for the Mimecast-Delegates sub-module (pwshMimecast).

.NOTES
    Author  : Bill Kindle (with AI assistance)
    Version : 1.0
    Created : 2026-04-02
    Requires: Pester 5.x, pwshMimecast module

    Run:
        Invoke-Pester .\tests\Mimecast-Delegates.Tests.ps1 -Output Detailed
#>

BeforeAll {
    # Import via the root manifest so NestedModules are loaded
    $modulePsd1 = Join-Path $PSScriptRoot '..' 'pwshMimecast.psd1'
    Import-Module (Resolve-Path $modulePsd1).Path -Force
}

AfterAll {
    Remove-Module 'Mimecast-Delegates' -Force -ErrorAction SilentlyContinue
    Remove-Module 'pwshMimecast'       -Force -ErrorAction SilentlyContinue
}

#region Module structure

Describe 'Module: pwshMimecast' {

    It 'loads without error' {
        Get-Module pwshMimecast | Should -Not -BeNullOrEmpty
    }

    It 'reports version 1.0.0' {
        (Get-Module pwshMimecast).Version.ToString() | Should -Be '1.0.0'
    }

    It 'exports exactly 7 public functions' {
        (Get-Command -Module pwshMimecast).Count | Should -Be 7
    }

    It 'exports the expected function <_>' -ForEach @(
        'Add-MimecastDelegateUsers'
        'Get-MimecastDelegateUsers'
        'Remove-MimecastDelegateUser'
        'Add-MimecastDelegateV1'
        'Get-MimecastDelegateV1'
        'Remove-MimecastDelegateV1'
        'Test-MimecastApiPath'
    ) {
        Get-Command -Module pwshMimecast -Name $_ | Should -Not -BeNullOrEmpty
    }

    It 'does not export private helper <_>' -ForEach @(
        'Get-MimecastAccessToken'
        'Invoke-MimecastApiCall'
        'Invoke-MimecastV1ApiCall'
    ) {
        (Get-Command -Module pwshMimecast).Name | Should -Not -Contain $_
    }

    Context 'Comment-based help' {
        It '<_> has a non-empty Synopsis' -ForEach @(
            'Add-MimecastDelegateUsers'
            'Get-MimecastDelegateUsers'
            'Remove-MimecastDelegateUser'
            'Add-MimecastDelegateV1'
            'Get-MimecastDelegateV1'
            'Remove-MimecastDelegateV1'
            'Test-MimecastApiPath'
        ) {
            (Get-Help $_).Synopsis | Should -Not -BeNullOrEmpty
        }
    }
}

#endregion Module structure

#region Private: Get-MimecastAccessToken

Describe 'Get-MimecastAccessToken' {

    BeforeEach {
        # Reset the token cache before each test
        InModuleScope 'Mimecast-Delegates' {
            $script:_TokenCache.AccessToken = [string]::Empty
            $script:_TokenCache.ExpiresAt   = [datetime]::MinValue
            $script:_TokenCache.BaseUrl     = [string]::Empty
        }
    }

    It 'requests a new token when the cache is empty' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' {
            @{ access_token = 'new-token'; expires_in = 1800 }
        }

        $result = InModuleScope 'Mimecast-Delegates' {
            Get-MimecastAccessToken `
                -BaseUrl      'https://api.services.mimecast.com' `
                -ClientId     'test-id' `
                -ClientSecret (ConvertTo-SecureString 'secret' -AsPlainText -Force)
        }

        $result | Should -Be 'new-token'
        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1
    }

    It 'returns the cached token when it has not expired' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' {
            @{ access_token = 'should-not-be-called'; expires_in = 1800 }
        }

        InModuleScope 'Mimecast-Delegates' {
            $script:_TokenCache.AccessToken = 'cached-token'
            $script:_TokenCache.ExpiresAt   = [datetime]::UtcNow.AddMinutes(10)
            $script:_TokenCache.BaseUrl     = 'https://api.services.mimecast.com'

            $result = Get-MimecastAccessToken `
                -BaseUrl      'https://api.services.mimecast.com' `
                -ClientId     'any-id' `
                -ClientSecret (ConvertTo-SecureString 'any' -AsPlainText -Force)

            $result | Should -Be 'cached-token'
        }

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 0
    }

    It 'bypasses the cache when -ForceRefresh is used' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' {
            @{ access_token = 'refreshed-token'; expires_in = 1800 }
        }

        InModuleScope 'Mimecast-Delegates' {
            $script:_TokenCache.AccessToken = 'stale-cached-token'
            $script:_TokenCache.ExpiresAt   = [datetime]::UtcNow.AddMinutes(10)
            $script:_TokenCache.BaseUrl     = 'https://api.services.mimecast.com'
        }

        $result = InModuleScope 'Mimecast-Delegates' {
            Get-MimecastAccessToken `
                -BaseUrl      'https://api.services.mimecast.com' `
                -ClientId     'test-id' `
                -ClientSecret (ConvertTo-SecureString 'secret' -AsPlainText -Force) `
                -ForceRefresh
        }

        $result | Should -Be 'refreshed-token'
        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1
    }

    It 'stores the acquired token and expiry in the cache' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' {
            @{ access_token = 'stored-token'; expires_in = 1800 }
        }

        InModuleScope 'Mimecast-Delegates' {
            Get-MimecastAccessToken `
                -BaseUrl      'https://api.services.mimecast.com' `
                -ClientId     'test-id' `
                -ClientSecret (ConvertTo-SecureString 'secret' -AsPlainText -Force) | Out-Null

            $script:_TokenCache.AccessToken | Should -Be 'stored-token'
            $script:_TokenCache.BaseUrl     | Should -Be 'https://api.services.mimecast.com'
            $script:_TokenCache.ExpiresAt   | Should -BeGreaterThan ([datetime]::UtcNow.AddSeconds(1700))
        }
    }

    It 'defaults to 1800 second TTL when expires_in is absent from the response' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' {
            @{ access_token = 'no-expiry-token' }
        }

        InModuleScope 'Mimecast-Delegates' {
            Get-MimecastAccessToken `
                -BaseUrl      'https://api.services.mimecast.com' `
                -ClientId     'test-id' `
                -ClientSecret (ConvertTo-SecureString 'secret' -AsPlainText -Force) | Out-Null

            $script:_TokenCache.ExpiresAt | Should -BeGreaterThan ([datetime]::UtcNow.AddSeconds(1700))
        }
    }

    It 'throws a descriptive error when the token endpoint fails' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' {
            throw [System.Net.WebException]::new('Unauthorized')
        }

        {
            InModuleScope 'Mimecast-Delegates' {
                Get-MimecastAccessToken `
                    -BaseUrl      'https://api.services.mimecast.com' `
                    -ClientId     'bad-id' `
                    -ClientSecret (ConvertTo-SecureString 'bad' -AsPlainText -Force)
            }
        } | Should -Throw -ExpectedMessage '*Failed to obtain Mimecast access token*'
    }

    It 'does not re-request a token when BaseUrl differs from cached BaseUrl' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' {
            @{ access_token = 'different-base-token'; expires_in = 1800 }
        }

        InModuleScope 'Mimecast-Delegates' {
            $script:_TokenCache.AccessToken = 'cached-for-other-url'
            $script:_TokenCache.ExpiresAt   = [datetime]::UtcNow.AddMinutes(10)
            $script:_TokenCache.BaseUrl     = 'https://us-api.services.mimecast.com'
        }

        # Different BaseUrl -- should trigger a new request
        InModuleScope 'Mimecast-Delegates' {
            Get-MimecastAccessToken `
                -BaseUrl      'https://api.services.mimecast.com' `
                -ClientId     'test-id' `
                -ClientSecret (ConvertTo-SecureString 'secret' -AsPlainText -Force) | Out-Null
        }

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1
    }
}

#endregion Private: Get-MimecastAccessToken

#region Private: Invoke-MimecastApiCall

Describe 'Invoke-MimecastApiCall' {

    It 'builds the correct URL by joining BaseUrl and Path' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -eq 'https://api.services.mimecast.com/api/user/find-delegate-users'
        } { @{ data = @() } }

        InModuleScope 'Mimecast-Delegates' {
            Invoke-MimecastApiCall `
                -BaseUrl 'https://api.services.mimecast.com/' `
                -Path    '/api/user/find-delegate-users' `
                -Token   'test-token'
        }

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1
    }

    It 'sets the Bearer authorization header' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Headers['Authorization'] -eq 'Bearer my-token'
        } { @{ data = @() } }

        InModuleScope 'Mimecast-Delegates' {
            Invoke-MimecastApiCall `
                -BaseUrl 'https://api.services.mimecast.com' `
                -Path    '/api/user/find-delegate-users' `
                -Token   'my-token'
        }

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1
    }

    It 'serializes the Body hashtable as JSON for POST requests' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Method -eq 'POST' -and $Body -ne $null
        } { @{ data = @() } }

        InModuleScope 'Mimecast-Delegates' {
            Invoke-MimecastApiCall `
                -BaseUrl 'https://api.services.mimecast.com' `
                -Path    '/api/user/add-delegate-user' `
                -Token   'tok' `
                -Method  POST `
                -Body    @{ data = @(@{ primaryAddress = 'a@b.com' }) }
        }

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1
    }

    It 'uses GET method by default' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Method -eq 'GET'
        } { @{} }

        InModuleScope 'Mimecast-Delegates' {
            Invoke-MimecastApiCall `
                -BaseUrl 'https://api.services.mimecast.com' `
                -Path    '/ttp/url/get-all-managed-urls' `
                -Token   'tok'
        }

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1
    }
}

#endregion Private: Invoke-MimecastApiCall

#region Public: Add-MimecastDelegateUsers

Describe 'Add-MimecastDelegateUsers' {

    BeforeAll {
        $script:Secret = ConvertTo-SecureString 'test-secret' -AsPlainText -Force
        $script:Creds = @{
            BaseUrl      = 'https://api.services.mimecast.com'
            ClientId     = 'test-id'
            ClientSecret = $script:Secret
            DelayMs      = 0
        }
    }

    BeforeEach {
        # Pre-warm token cache so tests don't hit the token endpoint
        InModuleScope 'Mimecast-Delegates' {
            $script:_TokenCache.AccessToken = 'pre-cached-token'
            $script:_TokenCache.ExpiresAt   = [datetime]::UtcNow.AddHours(1)
            $script:_TokenCache.BaseUrl     = 'https://api.services.mimecast.com'
        }
    }

    It 'standard mode: sends primaryAddress=DelegateAddress, delegateAddress=PrimaryAddress' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/add-delegate-user' -and
            ($Body | ConvertFrom-Json).data[0].primaryAddress  -eq 'shared@test.com' -and
            ($Body | ConvertFrom-Json).data[0].delegateAddress -eq 'alice@test.com'
        } { @{ data = @(@{ id = 'del-001' }) } }

        Add-MimecastDelegateUsers `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'alice@test.com' `
            @script:Creds

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1 `
            -ParameterFilter { $Uri -like '*/add-delegate-user' }
    }

    It 'inverted mode: swaps primaryAddress and delegateAddress in the request body' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/add-delegate-user' -and
            ($Body | ConvertFrom-Json).data[0].primaryAddress  -eq 'alice@test.com' -and
            ($Body | ConvertFrom-Json).data[0].delegateAddress -eq 'shared@test.com'
        } { @{ data = @(@{ id = 'del-002' }) } }

        Add-MimecastDelegateUsers `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'alice@test.com' `
            -InvertedMode `
            @script:Creds

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1 `
            -ParameterFilter { $Uri -like '*/add-delegate-user' }
    }

    It 'returns a success result object with all expected properties' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/add-delegate-user'
        } { @{ data = @(@{ id = 'id-abc' }) } }

        $result = Add-MimecastDelegateUsers `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'alice@test.com' `
            @script:Creds

        $result.PrimaryAddress  | Should -Be 'alice@test.com'
        $result.DelegateAddress | Should -Be 'shared@test.com'
        $result.Success         | Should -BeTrue
        $result.DelegateId      | Should -Be 'id-abc'
        $result.FailReason      | Should -BeNullOrEmpty
    }

    It 'returns a failure result per address when the API throws, without terminating the batch' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/add-delegate-user'
        } { throw 'HTTP 404 - Not Found' }

        $result = Add-MimecastDelegateUsers `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'alice@test.com' `
            @script:Creds `
            -WarningAction SilentlyContinue

        $result.Success    | Should -BeFalse
        $result.FailReason | Should -Not -BeNullOrEmpty
    }

    It 'processes multiple addresses and returns one result per address' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/add-delegate-user'
        } { @{ data = @(@{ id = 'id-x' }) } }

        $results = Add-MimecastDelegateUsers `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'a@t.com', 'b@t.com', 'c@t.com' `
            @script:Creds

        $results.Count | Should -Be 3
        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 3 `
            -ParameterFilter { $Uri -like '*/add-delegate-user' }
    }

    It 'skips blank and whitespace-only entries in PrimaryAddresses' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/add-delegate-user'
        } { @{ data = @(@{ id = 'id-x' }) } }

        $results = Add-MimecastDelegateUsers `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'alice@test.com', '   ', '', 'bob@test.com' `
            @script:Creds

        $results.Count | Should -Be 2
        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 2 `
            -ParameterFilter { $Uri -like '*/add-delegate-user' }
    }

    It 'rejects a plain username (no @ sign) in -PrimaryAddresses' {
        {
            Add-MimecastDelegateUsers `
                -DelegateAddress  'shared@test.com' `
                -PrimaryAddresses 'notanemail' `
                @script:Creds
        } | Should -Throw
    }

    It 'rejects a plain username (no @ sign) as -DelegateAddress' {
        {
            Add-MimecastDelegateUsers `
                -DelegateAddress  'notanemail' `
                -PrimaryAddresses 'alice@test.com' `
                @script:Creds
        } | Should -Throw
    }

    It 'makes no API call when -WhatIf is specified' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' { @{ data = @() } }

        Add-MimecastDelegateUsers `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'alice@test.com' `
            @script:Creds `
            -WhatIf

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 0 `
            -ParameterFilter { $Uri -like '*/add-delegate-user' }
    }

    It 'accepts PrimaryAddresses from the pipeline' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/add-delegate-user'
        } { @{ data = @(@{ id = 'id-pipe' }) } }

        $results = 'alice@test.com', 'bob@test.com' |
            Add-MimecastDelegateUsers -DelegateAddress 'shared@test.com' @script:Creds

        $results.Count | Should -Be 2
    }
}

#endregion Public: Add-MimecastDelegateUsers

#region Public: Get-MimecastDelegateUsers

Describe 'Get-MimecastDelegateUsers' {

    BeforeAll {
        $script:Secret = ConvertTo-SecureString 'test-secret' -AsPlainText -Force
        $script:Creds = @{
            BaseUrl      = 'https://api.services.mimecast.com'
            ClientId     = 'test-id'
            ClientSecret = $script:Secret
        }
    }

    BeforeEach {
        InModuleScope 'Mimecast-Delegates' {
            $script:_TokenCache.AccessToken = 'pre-cached-token'
            $script:_TokenCache.ExpiresAt   = [datetime]::UtcNow.AddHours(1)
            $script:_TokenCache.BaseUrl     = 'https://api.services.mimecast.com'
        }
    }

    It 'maps the delegateUsers response shape (real API) to all output properties' {
        # Regression: real Mimecast API shape is { data: [ { delegateUsers: [...] } ], fail: [] }
        # delegateUsers is nested inside data[0] -- NOT at the top level.
        # Previously the wrong nesting caused EmailAddress and DelegateId to always be empty.
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/find-delegate-users'
        } {
            @{
                data = @(
                    @{ delegateUsers = @(
                        @{ id = 'del-001'; emailAddress = 'alice@test.com'; name = 'Alice Smith'; source = 'ADCON' }
                    )}
                )
                fail = @()
            }
        }

        $result = Get-MimecastDelegateUsers -PrimaryAddress 'shared@test.com' @script:Creds

        $result.EmailAddress   | Should -Be 'alice@test.com'
        $result.DisplayName    | Should -Be 'Alice Smith'
        $result.DelegateId     | Should -Be 'del-001'
        $result.Source         | Should -Be 'ADCON'
        $result.PrimaryAddress | Should -Be 'shared@test.com'
    }

    It 'falls back to data key and delegateAddress field when delegateUsers is absent' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/find-delegate-users'
        } {
            @{ data = @(@{ id = 'del-001'; delegateAddress = 'alice@test.com' }) }
        }

        $result = Get-MimecastDelegateUsers -PrimaryAddress 'shared@test.com' @script:Creds

        $result.EmailAddress   | Should -Be 'alice@test.com'
        $result.DelegateId     | Should -Be 'del-001'
        $result.PrimaryAddress | Should -Be 'shared@test.com'
    }

    It 'falls back to data key when delegateUsers is absent' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/find-delegate-users'
        } {
            @{ data = @(@{ id = 'del-001'; emailAddress = 'alice@test.com' }) }
        }

        $result = Get-MimecastDelegateUsers -PrimaryAddress 'shared@test.com' @script:Creds

        $result.EmailAddress | Should -Be 'alice@test.com'
        $result.DelegateId   | Should -Be 'del-001'
    }

    It 'falls back to delegateAddress when emailAddress is absent' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/find-delegate-users'
        } {
            @{ data = @(@{ delegateUsers = @(@{ id = 'del-002'; delegateAddress = 'bob@test.com' }) }) }
        }

        $result = Get-MimecastDelegateUsers -PrimaryAddress 'shared@test.com' @script:Creds

        $result.EmailAddress | Should -Be 'bob@test.com'
    }

    It 'falls back to delegateEmailAddress when neither emailAddress nor delegateAddress is present' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/find-delegate-users'
        } {
            @{ data = @(@{ delegateUsers = @(@{ id = 'del-003'; delegateEmailAddress = 'carol@test.com' }) }) }
        }

        $result = Get-MimecastDelegateUsers -PrimaryAddress 'shared@test.com' @script:Creds

        $result.EmailAddress | Should -Be 'carol@test.com'
    }

    It 'returns an empty result when no delegates are found' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/find-delegate-users'
        } { @{ data = @() } }

        $result = Get-MimecastDelegateUsers -PrimaryAddress 'shared@test.com' @script:Creds `
            -WarningAction SilentlyContinue

        @($result).Count | Should -Be 0
    }

    It 'returns multiple results when multiple delegates are present' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/find-delegate-users'
        } {
            @{
                data = @(@{
                    delegateUsers = @(
                        @{ id = 'del-1'; emailAddress = 'a@test.com'; name = 'User A'; source = 'ADCON' }
                        @{ id = 'del-2'; emailAddress = 'b@test.com'; name = 'User B'; source = 'ADCON' }
                        @{ id = 'del-3'; emailAddress = 'c@test.com'; name = 'User C'; source = 'ADCON' }
                    )
                })
            }
        }

        $results = Get-MimecastDelegateUsers -PrimaryAddress 'shared@test.com' @script:Creds

        $results | Should -HaveCount 3
    }

    It 'calls POST /api/user/find-delegate-users' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri  -like '*/find-delegate-users' -and $Method -eq 'POST'
        } { @{ data = @() } }

        Get-MimecastDelegateUsers -PrimaryAddress 'shared@test.com' @script:Creds `
            -WarningAction SilentlyContinue

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1 `
            -ParameterFilter { $Uri -like '*/find-delegate-users' -and $Method -eq 'POST' }
    }
}

#endregion Public: Get-MimecastDelegateUsers

#region Public: Remove-MimecastDelegateUser

Describe 'Remove-MimecastDelegateUser' {

    BeforeAll {
        $script:Secret = ConvertTo-SecureString 'test-secret' -AsPlainText -Force
        $script:Creds = @{
            BaseUrl      = 'https://api.services.mimecast.com'
            ClientId     = 'test-id'
            ClientSecret = $script:Secret
        }
    }

    BeforeEach {
        InModuleScope 'Mimecast-Delegates' {
            $script:_TokenCache.AccessToken = 'pre-cached-token'
            $script:_TokenCache.ExpiresAt   = [datetime]::UtcNow.AddHours(1)
            $script:_TokenCache.BaseUrl     = 'https://api.services.mimecast.com'
        }
    }

    It 'calls POST /api/user/remove-delegate-user' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/remove-delegate-user'
        } { @{ data = @() } }

        Remove-MimecastDelegateUser `
            -PrimaryAddress 'shared@test.com' `
            -DelegateId     'del-abc' `
            @script:Creds `
            -Confirm:$false

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1 `
            -ParameterFilter { $Uri -like '*/remove-delegate-user' }
    }

    It 'sends PrimaryAddress and DelegateId in the request body' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/remove-delegate-user' -and
            ($Body | ConvertFrom-Json).data[0].primaryAddress -eq 'shared@test.com' -and
            ($Body | ConvertFrom-Json).data[0].id             -eq 'del-abc'
        } { @{ data = @() } }

        Remove-MimecastDelegateUser `
            -PrimaryAddress 'shared@test.com' `
            -DelegateId     'del-abc' `
            @script:Creds `
            -Confirm:$false

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1 `
            -ParameterFilter { $Uri -like '*/remove-delegate-user' }
    }

    It 'returns a success PSCustomObject with the correct properties' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/remove-delegate-user'
        } { @{ data = @() } }

        $result = Remove-MimecastDelegateUser `
            -PrimaryAddress 'shared@test.com' `
            -DelegateId     'del-abc' `
            @script:Creds `
            -Confirm:$false

        $result.PrimaryAddress | Should -Be 'shared@test.com'
        $result.DelegateId     | Should -Be 'del-abc'
        $result.Success        | Should -BeTrue
    }

    It 'makes no API call when -WhatIf is specified' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' { @{ data = @() } }

        Remove-MimecastDelegateUser `
            -PrimaryAddress 'shared@test.com' `
            -DelegateId     'del-abc' `
            @script:Creds `
            -WhatIf

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 0 `
            -ParameterFilter { $Uri -like '*/remove-delegate-user' }
    }

    It 'has ConfirmImpact set to High' {
        $attrs = (Get-Command Remove-MimecastDelegateUser).ScriptBlock.Attributes
        $binding = $attrs | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $binding.ConfirmImpact | Should -Be 'High'
    }

    It 'regression: DelegateId from Get-MimecastDelegateUsers pipes correctly into Remove-MimecastDelegateUser' {
        # Regression: Get-MimecastDelegateUsers previously returned an empty DelegateId due to
        # incorrect response field mapping (data[0].delegateUsers nesting was not handled).
        # This caused Remove-MimecastDelegateUser to throw "Cannot bind argument to parameter
        # 'DelegateId' because it is an empty string."
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/find-delegate-users'
        } {
            @{
                data = @(@{
                    delegateUsers = @(
                        @{ id = 'del-xyz'; emailAddress = 'alice@test.com'; name = 'Alice'; source = 'ADCON' }
                    )
                })
                fail = @()
            }
        }

        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/remove-delegate-user'
        } { @{ data = @() } }

        $delegate = Get-MimecastDelegateUsers -PrimaryAddress 'shared@test.com' @script:Creds |
                        Where-Object EmailAddress -EQ 'alice@test.com'

        # DelegateId must be populated -- this was the failing condition
        $delegate.DelegateId | Should -Be 'del-xyz'

        # Must not throw when passed to Remove
        { Remove-MimecastDelegateUser -PrimaryAddress 'shared@test.com' `
              -DelegateId $delegate.DelegateId @script:Creds -Confirm:$false } | Should -Not -Throw

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1 `
            -ParameterFilter { $Uri -like '*/remove-delegate-user' }
    }
}

#endregion Public: Remove-MimecastDelegateUser

#region Private: Invoke-MimecastV1ApiCall

Describe 'Invoke-MimecastV1ApiCall' {

    BeforeAll {
        # Must be valid base64 -- the function decodes it with FromBase64String
        $script:V1SecretKey = ConvertTo-SecureString 'dGVzdA==' -AsPlainText -Force
        $script:V1CommonParams = @{
            BaseUrl   = 'https://usb-api.mimecast.com'
            Uri       = '/api/delegate/get-delegates'
            AppId     = 'app-001'
            AppKey    = 'app-key'
            AccessKey = 'access-key'
            SecretKey = $script:V1SecretKey
        }
    }

    It 'sends a POST request with an MC Authorization header' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Headers['Authorization'] -like 'MC *'
        } { @{ data = @() } }

        InModuleScope 'Mimecast-Delegates' {
            Invoke-MimecastV1ApiCall `
                -BaseUrl   'https://usb-api.mimecast.com' `
                -Uri       '/api/delegate/get-delegates' `
                -AppId     'app-001' `
                -AppKey    'app-key' `
                -AccessKey 'access-key' `
                -SecretKey (ConvertTo-SecureString 'dGVzdA==' -AsPlainText -Force)
        }

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1
    }

    It 'includes all required x-mc-* headers' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Headers['x-mc-app-id'] -eq 'my-app' -and
            $Headers['x-mc-date'] -ne $null -and
            $Headers['x-mc-req-id'] -ne $null
        } { @{ data = @() } }

        InModuleScope 'Mimecast-Delegates' {
            Invoke-MimecastV1ApiCall `
                -BaseUrl   'https://usb-api.mimecast.com' `
                -Uri       '/api/delegate/get-delegates' `
                -AppId     'my-app' `
                -AppKey    'app-key' `
                -AccessKey 'access-key' `
                -SecretKey (ConvertTo-SecureString 'dGVzdA==' -AsPlainText -Force)
        }

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1
    }

    It 'throws on HTTP error' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' {
            throw [System.Net.WebException]::new('Internal Server Error')
        }

        {
            InModuleScope 'Mimecast-Delegates' {
                Invoke-MimecastV1ApiCall `
                    -BaseUrl   'https://usb-api.mimecast.com' `
                    -Uri       '/api/delegate/get-delegates' `
                    -AppId     'app-001' `
                    -AppKey    'app-key' `
                    -AccessKey 'access-key' `
                    -SecretKey (ConvertTo-SecureString 'dGVzdA==' -AsPlainText -Force)
            }
        } | Should -Throw
    }
}

#endregion Private: Invoke-MimecastV1ApiCall

#region Public: Add-MimecastDelegateV1

Describe 'Add-MimecastDelegateV1' {

    BeforeAll {
        $script:V1Key   = ConvertTo-SecureString 'dGVzdA==' -AsPlainText -Force
        $script:V1Creds = @{
            BaseUrl   = 'https://usb-api.mimecast.com'
            AppId     = 'app-001'
            AppKey    = 'app-key'
            AccessKey = 'access-key'
            SecretKey = $script:V1Key
            DelayMs   = 0
        }
    }

    It 'calls POST /api/delegate/add-delegate' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/add-delegate'
        } { @{ data = @(@{ id = 'v1-id-001' }) } }

        Add-MimecastDelegateV1 `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'alice@test.com' `
            @script:V1Creds

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1 `
            -ParameterFilter { $Uri -like '*/add-delegate' }
    }

    It 'returns a success result with the delegate Id' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/add-delegate'
        } { @{ data = @(@{ id = 'v1-abc' }) } }

        $result = Add-MimecastDelegateV1 `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'alice@test.com' `
            @script:V1Creds

        $result.Success         | Should -BeTrue
        $result.DelegateId      | Should -Be 'v1-abc'
        $result.PrimaryAddress  | Should -Be 'alice@test.com'
        $result.DelegateAddress | Should -Be 'shared@test.com'
        $result.FailReason      | Should -BeNullOrEmpty
    }

    It 'returns a failure result when the API returns an HTTP-200 fail array' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/add-delegate'
        } {
            @{
                fail = @(@{
                    errors = @(@{ message = 'Delegate already exists' })
                })
            }
        }

        $result = Add-MimecastDelegateV1 `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'alice@test.com' `
            @script:V1Creds `
            -WarningAction SilentlyContinue

        $result.Success    | Should -BeFalse
        $result.FailReason | Should -Match 'Delegate already exists'
    }

    It 'processes multiple addresses and returns one result per address' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/add-delegate'
        } { @{ data = @(@{ id = 'v1-id' }) } }

        $results = Add-MimecastDelegateV1 `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'a@t.com', 'b@t.com', 'c@t.com' `
            @script:V1Creds

        $results.Count | Should -Be 3
    }

    It 'makes no API call when -WhatIf is specified' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' { @{ data = @() } }

        Add-MimecastDelegateV1 `
            -DelegateAddress  'shared@test.com' `
            -PrimaryAddresses 'alice@test.com' `
            @script:V1Creds `
            -WhatIf

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 0 `
            -ParameterFilter { $Uri -like '*/add-delegate' }
    }
}

#endregion Public: Add-MimecastDelegateV1

#region Public: Get-MimecastDelegateV1

Describe 'Get-MimecastDelegateV1' {

    BeforeAll {
        $script:V1Key   = ConvertTo-SecureString 'dGVzdA==' -AsPlainText -Force
        $script:V1Creds = @{
            BaseUrl   = 'https://usb-api.mimecast.com'
            AppId     = 'app-001'
            AppKey    = 'app-key'
            AccessKey = 'access-key'
            SecretKey = $script:V1Key
        }
    }

    It 'maps response data to PrimaryAddress, DelegateAddress and DelegateId' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/get-delegates'
        } {
            @{
                data = @(@{
                    id       = 'v1-del-1'
                    delegate = @{ emailAddress = 'alice@test.com' }
                })
            }
        }

        $result = Get-MimecastDelegateV1 -DelegateAddress 'shared@test.com' @script:V1Creds

        $result.PrimaryAddress  | Should -Be 'alice@test.com'
        $result.DelegateAddress | Should -Be 'shared@test.com'
        $result.DelegateId      | Should -Be 'v1-del-1'
    }

    It 'returns an empty result when the data array is empty' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/get-delegates'
        } { @{ data = @() } }

        $result = Get-MimecastDelegateV1 -DelegateAddress 'shared@test.com' @script:V1Creds `
            -WarningAction SilentlyContinue

        @($result).Count | Should -Be 0
    }

    It 'returns multiple results when the API returns multiple delegates' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/get-delegates'
        } {
            @{
                data = @(
                    @{ id = 'v1-del-1'; delegate = @{ emailAddress = 'a@test.com' } }
                    @{ id = 'v1-del-2'; delegate = @{ emailAddress = 'b@test.com' } }
                )
            }
        }

        $results = Get-MimecastDelegateV1 -DelegateAddress 'shared@test.com' @script:V1Creds

        $results | Should -HaveCount 2
    }

    It 'writes a terminating error when the API returns a fail array' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/get-delegates'
        } {
            @{
                fail = @(@{
                    errors = @(@{ message = 'Mailbox not found' })
                })
            }
        }

        {
            Get-MimecastDelegateV1 -DelegateAddress 'shared@test.com' @script:V1Creds `
                -ErrorAction Stop
        } | Should -Throw -ExpectedMessage '*Mailbox not found*'
    }
}

#endregion Public: Get-MimecastDelegateV1

#region Public: Remove-MimecastDelegateV1

Describe 'Remove-MimecastDelegateV1' {

    BeforeAll {
        $script:V1Key   = ConvertTo-SecureString 'dGVzdA==' -AsPlainText -Force
        $script:V1Creds = @{
            BaseUrl   = 'https://usb-api.mimecast.com'
            AppId     = 'app-001'
            AppKey    = 'app-key'
            AccessKey = 'access-key'
            SecretKey = $script:V1Key
        }
    }

    It 'calls POST /api/delegate/remove-delegate' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/remove-delegate'
        } { @{ data = @() } }

        Remove-MimecastDelegateV1 `
            -DelegateAddress 'shared@test.com' `
            -PrimaryAddress  'alice@test.com' `
            @script:V1Creds `
            -Confirm:$false

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 1 `
            -ParameterFilter { $Uri -like '*/remove-delegate' }
    }

    It 'returns a success PSCustomObject' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/remove-delegate'
        } { @{ data = @() } }

        $result = Remove-MimecastDelegateV1 `
            -DelegateAddress 'shared@test.com' `
            -PrimaryAddress  'alice@test.com' `
            @script:V1Creds `
            -Confirm:$false

        $result.PrimaryAddress  | Should -Be 'alice@test.com'
        $result.DelegateAddress | Should -Be 'shared@test.com'
        $result.Success         | Should -BeTrue
    }

    It 'makes no API call when -WhatIf is specified' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' { @{ data = @() } }

        Remove-MimecastDelegateV1 `
            -DelegateAddress 'shared@test.com' `
            -PrimaryAddress  'alice@test.com' `
            @script:V1Creds `
            -WhatIf

        Should -Invoke Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -Exactly 0 `
            -ParameterFilter { $Uri -like '*/remove-delegate' }
    }

    It 'has ConfirmImpact set to High' {
        $attrs   = (Get-Command Remove-MimecastDelegateV1).ScriptBlock.Attributes
        $binding = $attrs | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $binding.ConfirmImpact | Should -Be 'High'
    }

    It 'writes a terminating error when the API call throws' {
        Mock Invoke-RestMethod -ModuleName 'Mimecast-Delegates' -ParameterFilter {
            $Uri -like '*/remove-delegate'
        } { throw 'HTTP 500 - Server Error' }

        {
            Remove-MimecastDelegateV1 `
                -DelegateAddress 'shared@test.com' `
                -PrimaryAddress  'alice@test.com' `
                @script:V1Creds `
                -Confirm:$false `
                -ErrorAction Stop
        } | Should -Throw
    }
}

#endregion Public: Remove-MimecastDelegateV1

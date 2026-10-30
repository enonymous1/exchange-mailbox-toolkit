BeforeAll {
    $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $repoRoot 'ExchangeMailboxToolkit' 'ExchangeMailboxToolkit.psd1') -Force
}

Describe 'Module manifest' {
    It 'Has a valid manifest for ExchangeMailboxToolkit 2.1.0 on PowerShell 7.2' {
        $manifest = Test-ModuleManifest -Path (Join-Path $repoRoot 'ExchangeMailboxToolkit' 'ExchangeMailboxToolkit.psd1')
        $manifest.Name | Should -Be 'ExchangeMailboxToolkit'
        $manifest.Version.ToString() | Should -Be '2.1.0'
        $manifest.PowerShellVersion.ToString() | Should -Be '7.2'
    }

    It 'Exports the documented functions' {
        $expected = @(
            'Connect-ExchangeEnvironment', 'Write-LogMessage', 'Set-LogFile', 'Clear-LogFile', 'Write-ProgressUpdate',
            'Get-SafeFilename', 'Get-EncodedGraphId', 'Invoke-GraphWithRetry',
            'Move-GraphMessage', 'Get-MailFoldersWithPaths', 'Invoke-EnsureFolderPath',
            'Get-GraphMailboxes', 'Get-Mailboxes'
        )
        $actual = (Get-Command -Module ExchangeMailboxToolkit).Name
        foreach ($name in $expected) { $actual | Should -Contain $name }
    }

    It 'Does not require Graph/EXO at import time, so helpers work without them' {
        # If RequiredModules forced Graph/EXO, this BeforeAll import would
        # already have failed on a machine without those modules installed.
        (Get-Module ExchangeMailboxToolkit) | Should -Not -BeNullOrEmpty
    }
}

Describe 'Pure helpers (no Graph session needed)' {
    It 'Get-SafeFilename replaces unsafe characters' {
        Get-SafeFilename -Name 'user@contoso.com:folder/path' | Should -Be 'user@contoso.com_folder_path'
        Get-SafeFilename -Name '' | Should -Be 'unnamed'
    }

    It 'Get-EncodedGraphId URL-encodes Graph ids' {
        Get-EncodedGraphId -Value 'AAMk-Test+Value/With=Special' | Should -Be 'AAMk-Test%2BValue%2FWith%3DSpecial'
    }

    It 'Write-LogMessage writes to the configured log file' {
        $log = Join-Path $TestDrive 'test.log'
        Set-LogFile -Path $log
        Write-LogMessage -Message 'hello from test' -Level 'INFO' -NoConsole
        Get-Content $log -Raw | Should -Match 'hello from test'
        Clear-LogFile
    }
}

Describe 'Every script parses' {
    It 'All .ps1 files in src, ExchangeMailboxToolkit and examples parse with no errors' {
        $files = Get-ChildItem -Recurse -Filter *.ps1 -Path $repoRoot |
            Where-Object { $_.FullName -notmatch 'tests' -and $_.FullName -notmatch '\.git' }
        $files.Count | Should -BeGreaterThan 10
        foreach ($file in $files) {
            $tokens = $null; $errors = $null
            [void][System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            $errors.Count | Should -Be 0 -Because "$($file.FullName) must parse"
        }
    }
}

Describe 'Loader fixes (regressions from v1)' {
    It 'Connect-ExchangeEnvironment.ps1 exists and loads the function' {
        $shim = Join-Path $repoRoot 'src' 'Utils' 'Connect-ExchangeEnvironment.ps1'
        Test-Path $shim | Should -BeTrue
        Get-Command -Name Connect-ExchangeEnvironment -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }

    It 'Initialize-Environment.ps1 no longer auto-connects on import' {
        $content = Get-Content (Join-Path $repoRoot 'src' 'Utils' 'Initialize-Environment.ps1') -Raw
        $content | Should -Not -Match "Connect-ExchangeEnvironment `@PSBoundParameters"
    }

    It 'No script references the never-existent Helpers.ps1 as a real file' {
        $hits = Get-ChildItem -Recurse -Filter *.ps1 -Path (Join-Path $repoRoot 'src') |
            Select-String -Pattern "Helpers\.ps1" -SimpleMatch
        $hits | Should -BeNullOrEmpty
    }
}

Describe 'Cross-platform paths (regression from v1)' {
    It 'No Join-Path child argument contains a Windows backslash' {
        $hits = Get-ChildItem -Recurse -Filter *.ps1 -Path (Join-Path $repoRoot 'src') |
            Select-String -Pattern 'Join-Path.*\\' |
            Where-Object { $_.Line -notmatch 'FolderPath|ParentFolderPath|currentPath|parentPath' }
        $hits | Should -BeNullOrEmpty
    }
}

Describe 'Destructive operations honour -WhatIf / -Confirm' {
    It 'Ensure, Remove-MailboxFolders and Import-MoveMessages declare SupportsShouldProcess' {
        foreach ($script in @('Ensure_Directory_Structure.ps1', 'Remove-MailboxFolders.ps1', 'Import-MoveMessages.ps1')) {
            $content = Get-Content (Join-Path $repoRoot 'src' 'Exchange' $script) -Raw
            $content | Should -Match 'SupportsShouldProcess'
        }
    }
}

Describe 'Check_For_Messages (non-interactive in v2)' {
    It 'Compares comma- and newline-separated ids and saves found/missing files' {
        $csv = Join-Path $TestDrive 'messages.csv'
        $ids = Join-Path $TestDrive 'ids.txt'
        $out = Join-Path $TestDrive 'out'
        @('MessageId,Subject', 'AAMk1,Hello', 'AAMk2,World') | Set-Content $csv
        @('AAMk1, AAMk999', 'AAMk2') | Set-Content $ids
        & (Join-Path $repoRoot 'src' 'Exchange' 'Check_For_Messages.ps1') -CsvPath $csv -IdFilePath $ids -Save -OutputDir $out | Out-Null
        (Get-ChildItem $out '*_found_*.csv').Count | Should -Be 1
        (Get-Content (Get-ChildItem $out '*_missing_*.txt') -Raw).Trim() | Should -Be 'AAMk999'
    }
}

Describe 'Examples use only fictional addresses' {
    It 'Every email address in src, examples, ExchangeMailboxToolkit, README and docs uses an example domain' {
        $paths = @((Join-Path $repoRoot 'src'), (Join-Path $repoRoot 'examples'), (Join-Path $repoRoot 'ExchangeMailboxToolkit'), (Join-Path $repoRoot 'README.md'), (Join-Path $repoRoot 'docs'))
        $addresses = Get-ChildItem -Recurse -File -Path $paths |
            Select-String -Pattern '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}' -AllMatches |
            ForEach-Object { $_.Matches.Value } | Select-Object -Unique
        $addresses.Count | Should -BeGreaterThan 0
        foreach ($address in $addresses) {
            $address | Should -Match '@(contoso\.com|domain\.com|example\.com)$' -Because "example address $address must use a fictional domain"
        }
    }
}

Describe 'mp-powershell fold-in (2.1.0)' {
    It 'Export-DistributionGroupMembers uses unlimited paging, not a hard-coded path or a += loop' {
        $path = Join-Path $repoRoot 'src' 'Exchange' 'Export-DistributionGroupMembers.ps1'
        Test-Path $path | Should -BeTrue
        $content = Get-Content $path -Raw
        $content | Should -Match 'ResultSize Unlimited'
        $content | Should -Not -Match 'C:\\distribution'
        $content | Should -Not -Match '\$results \+= \$result'
        $content | Should -Match 'List\[PSCustomObject\]'
        $content | Should -Match 'Connect-ExchangeEnvironment.*-LoadExchange'
    }

    It 'Test-GraphUserAccess corrects the impersonation claim and drops the unused Exchange connection' {
        $path = Join-Path $repoRoot 'src' 'Exchange' 'Test-GraphUserAccess.ps1'
        Test-Path $path | Should -BeTrue
        $content = Get-Content $path -Raw
        $content | Should -Match 'not an\s+impersonation'
        $content | Should -Not -Match 'Impersonation successful'
        # The help text mentions the original's calls by name; check the
        # code after the help block, not the prose, for what it actually does.
        $code = $content.Substring($content.IndexOf('#>') + 2)
        $code | Should -Not -Match 'Connect-ExchangeOnline'
        $code | Should -Match 'User\.Read\.All'
        $code | Should -Not -Match 'Group\.Read\.All'
        $code | Should -Not -Match 'millenniumpartners'
    }

    It 'Export-DistributionGroupMembers exports members and one blank row for an empty group (mocked EXO)' {
        . (Join-Path $repoRoot 'src' 'Exchange' 'Export-DistributionGroupMembers.ps1')
        function Connect-ExchangeEnvironment {
            [CmdletBinding()]
            param([switch]$LoadExchange)
            return [PSCustomObject]@{
                Graph    = [PSCustomObject]@{ Connected = $false; Message = $null }
                Exchange = [PSCustomObject]@{ Connected = $true; Message = $null }
            }
        }
        function Get-DistributionGroup {
            [CmdletBinding()]
            param([string]$Identity, [string]$ResultSize)
            return @(
                [PSCustomObject]@{ Identity = 'group-a'; DisplayName = 'All Staff'; PrimarySmtpAddress = 'all-staff@contoso.com' },
                [PSCustomObject]@{ Identity = 'group-b'; DisplayName = 'Empty Group'; PrimarySmtpAddress = 'empty@contoso.com' }
            )
        }
        function Get-DistributionGroupMember {
            [CmdletBinding()]
            param([string]$Identity, [string]$ResultSize)
            if ($Identity -eq 'group-a') {
                return @(
                    [PSCustomObject]@{ DisplayName = 'Ada'; PrimarySmtpAddress = 'ada@contoso.com' },
                    [PSCustomObject]@{ DisplayName = 'Ben'; PrimarySmtpAddress = 'ben@contoso.com' }
                )
            }
            return @()
        }
        $out = Join-Path $TestDrive 'groups.csv'
        Export-DistributionGroupMembers -OutputPath $out | Out-Null
        $rows = Import-Csv $out
        $rows.Count | Should -Be 3
        ($rows | Where-Object { $_.MemberEmail -eq 'ada@contoso.com' }).GroupName | Should -Be 'All Staff'
        $emptyRow = $rows | Where-Object { $_.GroupName -eq 'Empty Group' }
        $emptyRow.Count | Should -Be 1
        $emptyRow.MemberEmail | Should -BeNullOrEmpty

        $outFiltered = Join-Path $TestDrive 'groups-filtered.csv'
        Export-DistributionGroupMembers -OutputPath $outFiltered -ExcludeEmptyGroups | Out-Null
        (Import-Csv $outFiltered).Count | Should -Be 2
    }

    It 'Test-GraphUserAccess returns CanRead for a readable user and for a failure (mocked Graph)' {
        . (Join-Path $repoRoot 'src' 'Exchange' 'Test-GraphUserAccess.ps1')
        function Connect-ExchangeEnvironment {
            [CmdletBinding()]
            param([string[]]$GraphScopes)
            return [PSCustomObject]@{
                Graph    = [PSCustomObject]@{ Connected = $true; Message = $null }
                Exchange = [PSCustomObject]@{ Connected = $false; Message = $null }
            }
        }
        function Get-MgUser {
            [CmdletBinding()]
            param([string]$UserId)
            return [PSCustomObject]@{ DisplayName = 'Ada Lovelace'; UserPrincipalName = 'ada@contoso.com' }
        }
        $ok = Test-GraphUserAccess -UserPrincipalName 'ada@contoso.com'
        $ok.CanRead | Should -BeTrue
        $ok.DisplayName | Should -Be 'Ada Lovelace'

        function Get-MgUser {
            [CmdletBinding()]
            param([string]$UserId)
            throw 'Not found'
        }
        $bad = Test-GraphUserAccess -UserPrincipalName 'nobody@contoso.com' -WarningAction SilentlyContinue
        $bad.CanRead | Should -BeFalse
        $bad.Error | Should -Not -BeNullOrEmpty
    }
}

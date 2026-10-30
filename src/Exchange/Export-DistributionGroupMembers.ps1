<#
.SYNOPSIS
    Export Exchange Online distribution-group membership to CSV.

.DESCRIPTION
    v2.1.0: folded in from mp-powershell ("Get Distribution Group + Members.ps1").
    The original was a 34-line script with its Connect/Disconnect lines
    commented out, a hard-coded C:\ output path, no -ResultSize Unlimited
    (so large tenants could be silently truncated), and a $results +=
    loop. This version keeps the original output columns and adds what
    the original lacked:

    - Connects through Connect-ExchangeEnvironment (-LoadExchange) and
      fails with a clear message if Exchange Online is not connected.
      The session is left open afterwards, matching the rest of this
      toolkit: disconnect with Disconnect-ExchangeOnline when done.
    - -ResultSize Unlimited on both the group and member enumerations.
    - A List[PSCustomObject] instead of $results += in a loop.
    - Groups with no members are not silently dropped: they produce one
      row with blank MemberName/MemberEmail, unless -ExcludeEmptyGroups
      is given.
    - A per-group try/catch, so one unreadable group warns and the
      export continues instead of aborting the whole tenant.
    - -GroupIdentity to scope the export to one group for testing, and
      -OutputPath instead of a hard-coded drive path.

    This is read-only: it changes nothing in the tenant.

.PARAMETER GroupIdentity
    Optional. Export only this distribution group instead of the whole
    tenant. Run scoped to one group first.

.PARAMETER OutputPath
    Path for the CSV. Defaults to data/output/distribution_group_members.csv
    under the current directory. The parent directory is created if needed.

.PARAMETER ExcludeEmptyGroups
    Omit groups with no members from the CSV. By default an empty group
    produces one row with blank member fields, so no group disappears
    silently from the export.

.EXAMPLE
    pwsh -File src/Exchange/Export-DistributionGroupMembers.ps1 -GroupIdentity 'all-staff@contoso.com'

.EXAMPLE
    pwsh -File src/Exchange/Export-DistributionGroupMembers.ps1 -OutputPath 'data/output/groups.csv'
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$GroupIdentity,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = 'data/output/distribution_group_members.csv',

    [Parameter(Mandatory = $false)]
    [switch]$ExcludeEmptyGroups
)

. (Join-Path $PSScriptRoot '..' 'Utils' 'Initialize-Environment.ps1')

function Export-DistributionGroupMembers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$GroupIdentity,

        [Parameter(Mandatory = $false)]
        [string]$OutputPath = 'data/output/distribution_group_members.csv',

        [Parameter(Mandatory = $false)]
        [switch]$ExcludeEmptyGroups
    )

    $envStatus = Connect-ExchangeEnvironment -LoadExchange -ErrorAction Stop
    if (-not $envStatus.Exchange.Connected) {
        throw "Exchange Online is not connected: $($envStatus.Exchange.Message) Install the ExchangeOnlineManagement module and sign in, then retry."
    }

    if (-not (Get-Command -Name Get-DistributionGroup -ErrorAction SilentlyContinue)) {
        throw 'Get-DistributionGroup is not available. Install the ExchangeOnlineManagement module and reconnect with -LoadExchange.'
    }
    if (-not (Get-Command -Name Get-DistributionGroupMember -ErrorAction SilentlyContinue)) {
        throw 'Get-DistributionGroupMember is not available. Install the ExchangeOnlineManagement module and reconnect with -LoadExchange.'
    }

    if ($GroupIdentity) {
        $groups = @(Get-DistributionGroup -Identity $GroupIdentity -ErrorAction Stop)
    } else {
        $groups = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)
    }

    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($group in $groups) {
        try {
            $members = @(Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop)
        } catch {
            Write-Warning "Could not read members of $($group.DisplayName): $($_.Exception.Message)"
            continue
        }

        if ($members.Count -eq 0) {
            if (-not $ExcludeEmptyGroups) {
                $rows.Add([PSCustomObject]@{
                    GroupName   = $group.DisplayName
                    GroupEmail  = $group.PrimarySmtpAddress
                    MemberName  = $null
                    MemberEmail = $null
                })
            }
            continue
        }

        foreach ($member in $members) {
            $rows.Add([PSCustomObject]@{
                GroupName   = $group.DisplayName
                GroupEmail  = $group.PrimarySmtpAddress
                MemberName  = $member.DisplayName
                MemberEmail = $member.PrimarySmtpAddress
            })
        }
    }

    $outDir = Split-Path -Path $OutputPath -Parent
    if ($outDir -and -not (Test-Path -Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Output "Exported $($rows.Count) distribution-group member records for $($groups.Count) groups to: $OutputPath"
}

if ($MyInvocation.InvocationName -ne '.') {
    Export-DistributionGroupMembers -GroupIdentity $GroupIdentity -OutputPath $OutputPath -ExcludeEmptyGroups:$ExcludeEmptyGroups
}

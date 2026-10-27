<#

.SYNOPSIS
    Retrieve mailbox accounts from Microsoft Graph with filters (Shared/User/All, internal/external).

.DESCRIPTION
    Graph-based mailbox discovery. Uses server-side OData filtering by default to
    reduce memory and network usage. The function will fall back to a client-side fetch if a server-side
    query is not supported by the service. For testing or when server-side filtering is undesirable, use
    `-ForceClientSide` to force local filtering.

.PARAMETER Type
    One of All|Shared|User. Defaults to Shared.

.PARAMETER InternalOnly
    Return only mailboxes whose SMTP domain matches those in -Domains (or inferred from signed-in account).

.PARAMETER ExternalOnly
    Return only mailboxes whose SMTP domain does NOT match -Domains.

.PARAMETER Domains
    Array of domains used for InternalOnly/ExternalOnly filtering. If omitted and domain filtering
    is requested, the signed-in account's domain is used (if available). If no domain can be determined
    and domain filtering is requested, the function will throw.

.PARAMETER Max
    Limit results to the first N entries (client-side).

.PARAMETER ForceClientSide
    When set, force client-side fetching and filtering (do not attempt server-side OData filter).

.EXAMPLE
    Get-GraphMailboxes -Type Shared -InternalOnly -Domains contoso.com

.EXAMPLE
    Get-GraphMailboxes -Type User -Max 100 -UseServerFilter
#>
#requires -Version 7.2

function Get-GraphMailboxes {
    [CmdletBinding()]
    param(
        [ValidateSet('All','Shared','User')]
        [string]$Type = 'Shared',

        [Parameter()]
        [switch]$InternalOnly,

        [Parameter()]
        [switch]$ExternalOnly,

        [Parameter()]
        [string[]]$Domains,

        [Parameter()]
        [int]$Max = 0,

        [Parameter()]
        [switch]$ForceClientSide
    )

    # Ensure Graph is connected; prefer Connect-ExchangeEnvironment if available
    if (Get-Command -Name Connect-ExchangeEnvironment -ErrorAction SilentlyContinue) {
        $status = Connect-ExchangeEnvironment -GraphScopes @('User.Read.All','Directory.Read.All') -Verbose -LogFile $null
    }

    # Ensure we have an authenticated Microsoft Graph context for queries
    $mgCtxCmd = Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue
    if ($mgCtxCmd) { $ctx = Get-MgContext -ErrorAction SilentlyContinue }
    if (-not $ctx -or -not $ctx.Account) {
        if (Get-Command -Name Connect-ExchangeEnvironment -ErrorAction SilentlyContinue) {
            Write-Verbose "No Graph context found; attempting Connect-ExchangeEnvironment for mailbox discovery..."
            $status = Connect-ExchangeEnvironment -GraphScopes @('User.Read.All','Directory.Read.All') -LogFile $null -Verbose
            if (-not $status -or -not $status.Graph.Connected) {
                throw "Authentication required: Connect-ExchangeEnvironment did not establish a Graph session."
            }
        }
        else {
            throw "Authentication required: no Graph session found and Connect-ExchangeEnvironment is unavailable."
        }
    }

    # Helper to build server-side filter
    function Build-FilterString {
        param(
            [string]$type,
            [string[]]$domains,
            [switch]$internalOnly,
            [switch]$externalOnly
        )
        $parts = @()

        # Prefer explicit equals; avoid 'ne' because some Graph endpoints reject NotEqualsMatch
        if ($type -eq 'Shared') { $parts += "userType eq 'Shared'" }
        elseif ($type -eq 'User') { $parts += "not (userType eq 'Shared')" }

        # require mail present - avoid 'ne' operator which some Graph endpoints reject
        $parts += "not (mail eq null)"

        if ($domains -and $domains.Count -gt 0) {
            $domParts = @()
            foreach ($d in $domains) { $domParts += "endsWith(mail,'@{0}')" -f $d }
            $domExpr = ($domParts -join ' or ')
            if ($internalOnly) { $parts += "($domExpr)" }
            elseif ($externalOnly) { $parts += "not ($domExpr)" }
        }

        return ($parts -join ' and ')
    }

    # If domain-based filtering requested and no Domains provided, infer from signed-in account
    if (($InternalOnly -or $ExternalOnly) -and (-not $Domains -or $Domains.Count -eq 0)) {
        if ($status -and $status.Graph -and $status.Graph.Account) {
            $acctDomain = ($status.Graph.Account -split '@')[-1].ToLower()
            $Domains = @($acctDomain)
        }
        else {
            throw "Domain filtering requested (InternalOnly/ExternalOnly) but no -Domains provided and no signed-in account available to infer a domain."
        }
    }

    $selectFields = 'id,displayName,mail,userType'
    $users = @()

    # Decide mode: server-side by default, unless user requests client-side
    $useServer = -not $ForceClientSide
    if ($useServer) {
        $filter = Build-FilterString -type $Type -domains $Domains -internalOnly:$InternalOnly -externalOnly:$ExternalOnly
        try {
            Write-Verbose "Querying Microsoft Graph with server-side filter: $filter"
            $users = Get-MgUser -Filter $filter -Select $selectFields -All -ErrorAction Stop | Select-Object Id,DisplayName,Mail,UserType
            # Fallback: if asking for Shared and none returned, attempt relaxed client-side fetch
            if ($Type -eq 'Shared' -and (-not $users -or $users.Count -eq 0)) {
                Write-Verbose "Server-side filter returned no Shared users; falling back to client-side retrieval."
                $users = Get-MgUser -All -Select $selectFields -ErrorAction Stop | Where-Object { $_.Mail } | Select-Object Id,DisplayName,Mail,UserType
            }
        }
        catch {
            Write-Warning "Server-side query failed: $($_.Exception.Message); falling back to client-side retrieval."
            $users = Get-MgUser -All -Select $selectFields -ErrorAction Stop | Where-Object { $_.Mail } | Select-Object Id,DisplayName,Mail,UserType
        }
    }
    else {
        # Client-side: fetch once and apply filters locally (may be memory heavy for large tenants)
        Write-Verbose "Fetching all users client-side (may be slow for large tenants)."
        try {
            $all = Get-MgUser -All -Select $selectFields -ErrorAction Stop | Where-Object { $_.Mail }
        }
        catch {
            throw "Failed to query users from Microsoft Graph: $($_.Exception.Message)"
        }

        switch ($Type) {
            'Shared' { $users = $all | Where-Object { $_.UserType -ieq 'Shared' } }
            'User'   { $users = $all | Where-Object { $_.UserType -ne 'Shared' } }
            default  { $users = $all }
        }

        if ($Domains -and $Domains.Count -gt 0) {
            $doms = $Domains | ForEach-Object { $_.ToLower() }
            if ($InternalOnly) {
                $users = $users | Where-Object { ($_.Mail -split '@')[-1].ToLower() -in $doms }
            }
            elseif ($ExternalOnly) {
                $users = $users | Where-Object { -not ( ($_.Mail -split '@')[-1].ToLower() -in $doms ) }
            }
        }
    }

    if ($Max -gt 0) { $users = $users | Select-Object -First $Max }

    # Ensure consistent PSCustomObject output shape
    $out = $users | ForEach-Object {
        [PSCustomObject]@{
            Id = $_.Id
            DisplayName = $_.DisplayName
            Mail = $_.Mail
            UserType = $_.UserType
        }
    }

    Write-Verbose ("Get-GraphMailboxes returning {0} entries" -f ($out.Count))
    return $out
}

try {
    Export-ModuleMember -Function Get-GraphMailboxes -ErrorAction SilentlyContinue
}
catch {
    # Not running as a module - ignore
}

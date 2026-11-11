<#
.SYNOPSIS
    Creates a Teams team, imports channel threads with original authors/timestamps, and adds members
    using Microsoft Graph. Optionally automates the Entra app registration and admin consent via Azure CLI.

.PREREQUISITES
    - Azure CLI, signed in with: az login --tenant <tenant> --allow-no-subscriptions
      The signed-in account must be able to grant admin consent (Global Admin or Privileged Role Admin).
    - All UPNs in the config must belong to the signed-in tenant and have Teams licenses.
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'sample-conversation.json'),
    [string]$AppName = 'Teams Conversation Import',
    [int]$SecretLifetimeDays = 7,
    # Team owner; its domain is appended to every user name in the config. Prompted if omitted.
    [string]$AdminUpn,
    # Defaults to the domain of the admin UPN.
    [string]$TenantId,
    [switch]$UseDeviceCode
)

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -ne 'Core') {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
}

$GraphAppId = '00000003-0000-0000-c000-000000000000'
$RequiredRoles = @(
    'Teamwork.Migrate.All',
    'Team.Create',
    'Team.ReadBasic.All',
    'Channel.ReadBasic.All',
    'TeamMember.ReadWrite.All',
    'User.Read.All'
)

#region Helpers

function Invoke-Az {
    param([string[]]$Arguments, [switch]$AllowFailure)
    # Windows PowerShell 5.1 turns native stderr into terminating errors when EAP is Stop.
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $output = & az @Arguments 2>&1 } finally { $ErrorActionPreference = $previous }

    $stdout = @($output | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] })
    $stderr = @($output | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
    if ($LASTEXITCODE -ne 0) {
        if ($AllowFailure) { return $null }
        throw "az $($Arguments[0..2] -join ' ') failed: $($stderr -join [Environment]::NewLine)"
    }
    return ($stdout -join [Environment]::NewLine).Trim()
}

function Invoke-AzRestPost {
    param([string]$Url, $Body)
    $file = New-TemporaryFile
    try {
        [IO.File]::WriteAllText($file.FullName, ($Body | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
        Invoke-Az @('rest', '--method', 'post', '--url', $Url, '--headers', 'Content-Type=application/json', '--body', "@$($file.FullName)")
    }
    finally { Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue }
}

function ConvertTo-GraphDate {
    param($Value)
    if ($Value -is [datetime]) { return [datetimeoffset]$Value.ToUniversalTime() }
    return [datetimeoffset]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture)
}

function Format-GraphDate {
    param([datetimeoffset]$Value)
    return $Value.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", [Globalization.CultureInfo]::InvariantCulture)
}

function Get-TokenRoles {
    param([string]$Token)
    $payload = $Token.Split('.')[1].Replace('-', '+').Replace('_', '/')
    $payload = $payload.PadRight($payload.Length + (4 - $payload.Length % 4) % 4, '=')
    $claims = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
    return @($claims.roles)
}

function Get-GraphToken {
    param([string]$TenantId, [string]$ClientId, [string]$ClientSecret)
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }
    # New secrets and role assignments take time to replicate.
    for ($attempt = 1; $attempt -le 30; $attempt++) {
        try {
            $token = (Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body -ContentType 'application/x-www-form-urlencoded').access_token
            $missing = @($RequiredRoles | Where-Object { (Get-TokenRoles $token) -notcontains $_ })
            if ($missing.Count -eq 0) { return $token }
            Write-Host "  Waiting for permissions to propagate (missing: $($missing -join ', '))..."
        }
        catch {
            Write-Host "  Waiting for app credentials to propagate..."
        }
        Start-Sleep -Seconds 10
    }
    throw 'Timed out waiting for an app token with the required Graph roles.'
}

function Invoke-Graph {
    param(
        [string]$Method = 'GET',
        [string]$Path,
        $Body,
        [switch]$RetryNotFound,
        [switch]$IncludeHeaders
    )
    $uri = if ($Path -match '^https://') { $Path } else { "https://graph.microsoft.com/v1.0$Path" }
    $params = @{
        Method          = $Method
        Uri             = $uri
        Headers         = @{ Authorization = "Bearer $script:GraphToken" }
        UseBasicParsing = $true
    }
    if ($null -ne $Body) {
        $params.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 20))
        $params.ContentType = 'application/json; charset=utf-8'
    }

    for ($attempt = 1; ; $attempt++) {
        try {
            $response = Invoke-WebRequest @params
            $content = if ($response.Content) { $response.Content | ConvertFrom-Json } else { $null }
            if ($IncludeHeaders) { return [pscustomobject]@{ Body = $content; Headers = $response.Headers } }
            return $content
        }
        catch {
            $status = 0
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            $retryable = $status -eq 429 -or $status -ge 500 -or ($RetryNotFound -and $status -eq 404)
            if (-not $retryable -or $attempt -ge 8) {
                $detail = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
                throw "Graph $Method $uri failed ($status): $detail"
            }
            Start-Sleep -Seconds ([Math]::Min(60, 5 * $attempt))
        }
    }
}

function New-ImportMessageBody {
    param($Message, $UserMap)
    $author = $UserMap[$Message.author]
    return @{
        createdDateTime = Format-GraphDate (ConvertTo-GraphDate $Message.createdDateTime)
        from            = @{
            user = @{
                id               = $author.id
                displayName      = $author.displayName
                userIdentityType = 'aadUser'
            }
        }
        body            = @{
            contentType = 'html'
            content     = [Net.WebUtility]::HtmlEncode([string]$Message.content)
        }
    }
}

#endregion

#region Load config

$jsonArgs = @{}
if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { $jsonArgs.DateKind = 'String' }
$config = Get-Content -Raw -Encoding UTF8 -Path $ConfigPath | ConvertFrom-Json @jsonArgs
$threads = @($config.channel.threads | Sort-Object { (ConvertTo-GraphDate $_.message.createdDateTime).UtcDateTime })

while ($AdminUpn -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
    if ($AdminUpn) { Write-Warning "'$AdminUpn' is not a valid UPN." }
    $AdminUpn = (Read-Host 'Input the username (eg: PattiF@domain.onmicrosoft.com)').Trim()
}
$domain = $AdminUpn.Split('@')[1]
Write-Host "Using domain: $domain"
foreach ($u in @($config.users)) {
    $u | Add-Member -NotePropertyName userPrincipalName -NotePropertyValue "$($u.userName)@$domain" -Force
}

#endregion

#region App registration (Azure CLI)

Write-Host '[1/6] Configuring app registration...'
if (-not $TenantId) { $TenantId = $domain }

$accountJson = Invoke-Az @('account', 'show', '-o', 'json') -AllowFailure
$account = if ($accountJson) { $accountJson | ConvertFrom-Json } else { $null }
$tenantInfo = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/v2.0/.well-known/openid-configuration"
$expectedTenantId = ($tenantInfo.issuer -split '/')[3]

if (-not $account -or $account.tenantId -ne $expectedTenantId) {
    Write-Host "  Signing in to Azure CLI (tenant $TenantId)..."
    $loginArgs = @('login', '--tenant', $TenantId, '--allow-no-subscriptions', '-o', 'none')
    if ($UseDeviceCode) { $loginArgs += '--use-device-code' }
    # Not captured so the browser/device-code prompt stays visible.
    & az @loginArgs
    if ($LASTEXITCODE -ne 0) { throw 'az login failed.' }
    $account = Invoke-Az @('account', 'show', '-o', 'json') | ConvertFrom-Json
}
$tenantId = $account.tenantId

$app = @(Invoke-Az @('ad', 'app', 'list', '--display-name', $AppName, '-o', 'json') | ConvertFrom-Json) | Select-Object -First 1
if (-not $app) {
    $app = Invoke-Az @('ad', 'app', 'create', '--display-name', $AppName, '--sign-in-audience', 'AzureADMyOrg', '-o', 'json') | ConvertFrom-Json
    Write-Host "  Created app '$AppName' ($($app.appId))"
}
else {
    Write-Host "  Reusing app '$AppName' ($($app.appId))"
}
$appId = $app.appId

$graphSp = Invoke-Az @('ad', 'sp', 'show', '--id', $GraphAppId, '-o', 'json') | ConvertFrom-Json
$roleIds = @{}
foreach ($roleName in $RequiredRoles) {
    $role = $graphSp.appRoles | Where-Object { $_.value -eq $roleName -and $_.allowedMemberTypes -contains 'Application' }
    if (-not $role) { throw "Graph application role '$roleName' not found." }
    $roleIds[$roleName] = $role.id
}

# Declare permissions on the app so they show in the portal.
$rraFile = New-TemporaryFile
try {
    $rra = @(@{
            resourceAppId  = $GraphAppId
            resourceAccess = @($roleIds.Values | ForEach-Object { @{ id = $_; type = 'Role' } })
        })
    [IO.File]::WriteAllText($rraFile.FullName, (ConvertTo-Json -InputObject $rra -Depth 10), [Text.UTF8Encoding]::new($false))
    Invoke-Az @('ad', 'app', 'update', '--id', $appId, '--required-resource-accesses', "@$($rraFile.FullName)") | Out-Null
}
finally { Remove-Item $rraFile.FullName -Force -ErrorAction SilentlyContinue }

$spJson = Invoke-Az @('ad', 'sp', 'show', '--id', $appId, '-o', 'json') -AllowFailure
$sp = if ($spJson) { $spJson | ConvertFrom-Json } else { Invoke-Az @('ad', 'sp', 'create', '--id', $appId, '-o', 'json') | ConvertFrom-Json }

Write-Host '[2/6] Granting admin consent...'
$existingRoleIds = @(Invoke-Az @('rest', '--method', 'get', '--url', "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/appRoleAssignments", '-o', 'json') | ConvertFrom-Json).value.appRoleId
foreach ($roleName in $RequiredRoles) {
    if ($existingRoleIds -contains $roleIds[$roleName]) { continue }
    $assignment = @{ principalId = $sp.id; resourceId = $graphSp.id; appRoleId = $roleIds[$roleName] }
    for ($attempt = 1; ; $attempt++) {
        try {
            Invoke-AzRestPost -Url "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/appRoleAssignments" -Body $assignment | Out-Null
            Write-Host "  Granted $roleName"
            break
        }
        catch {
            # A new service principal may not have replicated yet.
            if ($attempt -ge 6) { throw }
            Start-Sleep -Seconds 10
        }
    }
}

$endDate = (Get-Date).ToUniversalTime().AddDays($SecretLifetimeDays).ToString('yyyy-MM-dd')
$clientSecret = Invoke-Az @('ad', 'app', 'credential', 'reset', '--id', $appId, '--append', '--end-date', $endDate, '--query', 'password', '-o', 'tsv')

Write-Host '[3/6] Acquiring Graph app token...'
$script:GraphToken = Get-GraphToken -TenantId $tenantId -ClientId $appId -ClientSecret $clientSecret

#endregion

#region Resolve users

Write-Host '[4/6] Resolving users...'
function Resolve-User([string]$Upn) {
    Invoke-Graph -Path "/users/$([uri]::EscapeDataString($Upn))?`$select=id,displayName,userPrincipalName"
}

$userMap = @{}
foreach ($u in @($config.users)) { $userMap[$u.key] = Resolve-User $u.userPrincipalName }
$owner = Resolve-User $AdminUpn

$allMessages = @($threads | ForEach-Object { $_.message; @($_.replies) })
$unknownAuthors = @($allMessages.author | Sort-Object -Unique | Where-Object { -not $userMap.ContainsKey($_) })
if ($unknownAuthors.Count -gt 0) { throw "Unknown author keys: $($unknownAuthors -join ', ')" }

#endregion

#region Create team and channel in migration mode, import threads

Write-Host '[5/6] Creating team and importing threads...'
$earliest = ($allMessages | ForEach-Object { ConvertTo-GraphDate $_.createdDateTime } | Sort-Object | Select-Object -First 1)

$teamResponse = Invoke-Graph -Method POST -Path '/teams' -IncludeHeaders -Body @{
    '@microsoft.graph.teamCreationMode' = 'migration'
    'template@odata.bind'               = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
    displayName                         = $config.team.displayName
    description                         = $config.team.description
    createdDateTime                     = Format-GraphDate $earliest.AddMinutes(-2)
}
$location = [string](@($teamResponse.Headers['Location'])[0])
if ($location -notmatch "teams\('([^']+)'\)") { throw "Unexpected team creation Location header: $location" }
$teamId = $Matches[1]

for ($attempt = 1; ; $attempt++) {
    $operation = Invoke-Graph -Path $location -RetryNotFound
    if ($operation.status -eq 'succeeded') { break }
    if ($operation.status -eq 'failed') { throw "Team creation failed: $($operation.error | ConvertTo-Json -Compress)" }
    if ($attempt -ge 60) { throw 'Timed out waiting for team creation.' }
    Start-Sleep -Seconds 5
}
Write-Host "  Team created: $teamId"

$channel = Invoke-Graph -Path "/teams/$teamId/primaryChannel" -RetryNotFound
if ($channel.migrationMode -and $channel.migrationMode -ne 'inProgress') {
    Invoke-Graph -Method POST -Path "/teams/$teamId/channels/$($channel.id)/startMigration" -Body @{
        conversationCreationDateTime = Format-GraphDate $earliest.AddMinutes(-1)
    } | Out-Null
}
Write-Host "  Using channel: $($channel.displayName)"

$messagesPath = "/teams/$teamId/channels/$($channel.id)/messages"
foreach ($thread in $threads) {
    $root = Invoke-Graph -Method POST -Path $messagesPath -RetryNotFound -Body (New-ImportMessageBody $thread.message $userMap)
    $replies = @($thread.replies | Sort-Object { (ConvertTo-GraphDate $_.createdDateTime).UtcDateTime })
    foreach ($reply in $replies) {
        Invoke-Graph -Method POST -Path "$messagesPath/$($root.id)/replies" -Body (New-ImportMessageBody $reply $userMap) | Out-Null
    }
    Write-Host "  Imported thread with $($replies.Count) replies"
}

foreach ($ch in @((Invoke-Graph -Path "/teams/$teamId/channels").value)) {
    try { Invoke-Graph -Method POST -Path "/teams/$teamId/channels/$($ch.id)/completeMigration" | Out-Null }
    catch { Write-Warning "completeMigration for channel '$($ch.displayName)': $_" }
}
Invoke-Graph -Method POST -Path "/teams/$teamId/completeMigration" | Out-Null
Write-Host '  Migration completed'

#endregion

#region Add members

Write-Host '[6/6] Adding owner and members...'
$members = @([pscustomobject]@{ User = $owner; Roles = @('owner') })
$members += @($userMap.Values | Where-Object { $_.id -ne $owner.id } | Sort-Object id -Unique | ForEach-Object { [pscustomobject]@{ User = $_; Roles = @() } })

foreach ($m in $members) {
    Invoke-Graph -Method POST -Path "/teams/$teamId/members" -RetryNotFound -Body @{
        '@odata.type'     = '#microsoft.graph.aadUserConversationMember'
        roles             = $m.Roles
        'user@odata.bind' = "https://graph.microsoft.com/v1.0/users('$($m.User.id)')"
    } | Out-Null
    Write-Host "  Added $($m.User.userPrincipalName) $(if ($m.Roles) { '(owner)' })"
}

$team = Invoke-Graph -Path "/teams/$teamId"
Write-Host '#############################################################' -ForegroundColor Green
Write-Host "Successfully Created Team: $($team.displayName)" -ForegroundColor Green
Write-Host "URL:  $($team.webUrl)" -ForegroundColor Cyan

#endregion

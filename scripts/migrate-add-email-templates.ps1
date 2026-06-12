<#
.SYNOPSIS
  One-time migration for the email-template module:
    1. Adds the PersonalEmail (single-line text) column to the Employees list
    2. Creates the ELC_EmailTemplates list

  Idempotent — re-running is a no-op once both are present.
#>
[CmdletBinding()]
param(
    [string]$SiteId   = "vanrockre.sharepoint.com,a02c1cd8-9f1f-4827-8286-7b6b7ce74232,01202419-6625-4499-b0d5-8ceb1cffdba3",
    [string]$TenantId = "33575d04-ca7b-4396-8011-9eaea4030b46",
    [string]$ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
)
$ErrorActionPreference = 'Stop'

function Write-Step($m) { Write-Host "==> $m" -ForegroundColor Cyan }
function Write-OK   ($m) { Write-Host "    [OK]    $m" -ForegroundColor Green }
function Write-Skip ($m) { Write-Host "    [SKIP]  $m" -ForegroundColor DarkYellow }
function Write-New  ($m) { Write-Host "    [APPLY] $m" -ForegroundColor Yellow }

Write-Step "Requesting device code"
$scope  = "https://graph.microsoft.com/Sites.Manage.All https://graph.microsoft.com/Sites.ReadWrite.All offline_access"
$dcResp = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
    -Body @{ client_id = $ClientId; scope = $scope } -ContentType 'application/x-www-form-urlencoded'

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
Write-Host "  ║  ACTION REQUIRED — sign in to authorise the change           ║" -ForegroundColor Yellow
Write-Host "  ║                                                              ║" -ForegroundColor Yellow
Write-Host "  ║  1. Open:   https://login.microsoft.com/device               ║" -ForegroundColor Yellow
Write-Host "  ║  2. Enter:  $($dcResp.user_code.PadRight(50))║" -ForegroundColor Yellow
Write-Host "  ║  3. Sign in as bturner@newshirepm.com                        ║" -ForegroundColor Yellow
Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
Write-Host ""

$expiresAt    = (Get-Date).AddSeconds([int]$dcResp.expires_in - 5)
$pollInterval = [int]$dcResp.interval; if ($pollInterval -lt 5) { $pollInterval = 5 }
$token = $null
while ((Get-Date) -lt $expiresAt) {
    Start-Sleep -Seconds $pollInterval
    try {
        $tokenResp = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Body @{ grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; client_id = $ClientId; device_code = $dcResp.device_code } `
            -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        $token = $tokenResp.access_token; break
    } catch {
        $err = $null; try { $err = ($_.ErrorDetails.Message | ConvertFrom-Json) } catch {}
        if ($err -and $err.error -eq 'authorization_pending') { Write-Host '.' -NoNewline -ForegroundColor DarkGray; continue }
        if ($err -and $err.error -eq 'slow_down')              { $pollInterval += 5; continue }
        Write-Host ""; Write-Error ($err.error_description ?? $_.Exception.Message); exit 1
    }
}
Write-Host ""
if (-not $token) { Write-Error "Authentication did not complete within the device-code lifetime."; exit 1 }
Write-OK "Authenticated"
$headers = @{ Authorization = "Bearer $token" }

function Invoke-Graph {
    param([string]$Method,[string]$Path,[object]$Body)
    $uri = "https://graph.microsoft.com/v1.0$Path"
    if ($Body) { return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -Body ($Body | ConvertTo-Json -Depth 20 -Compress) -ContentType 'application/json' }
    return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers
}

# ── 1. Add PersonalEmail column to Employees ──
Write-Step "Locating Employees list"
$lists = Invoke-Graph GET "/sites/$SiteId/lists?`$select=id,name,displayName"
$empList = $lists.value | Where-Object { $_.name -eq 'Employees' -or $_.displayName -eq 'Employees' } | Select-Object -First 1
if (-not $empList) { Write-Error "Employees list not found."; exit 1 }

$empCols = Invoke-Graph GET "/sites/$SiteId/lists/$($empList.id)/columns?`$select=id,name,displayName"
$existsCol = $empCols.value | Where-Object { $_.name -eq 'PersonalEmail' -or $_.displayName -eq 'PersonalEmail' } | Select-Object -First 1
if ($existsCol) {
    Write-Skip "Employees.PersonalEmail already exists"
} else {
    Write-New "Adding Employees.PersonalEmail"
    Invoke-Graph POST "/sites/$SiteId/lists/$($empList.id)/columns" @{ name='PersonalEmail'; text=@{} } | Out-Null
    Write-OK "Added Employees.PersonalEmail"
}

# ── 2. Create ELC_EmailTemplates list ──
Write-Step "Checking for ELC_EmailTemplates"
$existsList = $lists.value | Where-Object { $_.name -eq 'ELC_EmailTemplates' -or $_.displayName -eq 'ELC_EmailTemplates' } | Select-Object -First 1
if ($existsList) {
    Write-Skip "ELC_EmailTemplates already exists (id=$($existsList.id))"
} else {
    Write-New "Creating ELC_EmailTemplates"
    $body = @{
        displayName = 'ELC_EmailTemplates'
        description = 'Reusable email templates with {{var}} substitution'
        list        = @{ template = 'genericList' }
        columns     = @(
            @{ name='Category';  choice=@{ allowTextEntry=$false; displayAs='dropDownMenu'; choices=@('Welcome','Onboarding','Reminder','Performance','Offboarding','Celebration','Other') }; defaultValue=@{ value='Other' } },
            @{ name='Subject';   text=@{} },
            @{ name='Body';      text=@{ allowMultipleLines=$true; appendChangesToExistingText=$false; linesForEditing=10 } },
            @{ name='DefaultTo'; text=@{} },
            @{ name='DefaultCc'; text=@{} },
            @{ name='Active';    boolean=@{}; defaultValue=@{ value='true' } },
            @{ name='Notes';     text=@{ allowMultipleLines=$true; appendChangesToExistingText=$false; linesForEditing=6 } }
        )
    }
    Invoke-Graph POST "/sites/$SiteId/lists" $body | Out-Null
    Write-OK "Created ELC_EmailTemplates"
}

Write-Host ""
Write-Step "Done."
Write-Host "  Hard-refresh the app, then Setup -> Email Templates -> 'Seed defaults' to load the starter set." -ForegroundColor Gray

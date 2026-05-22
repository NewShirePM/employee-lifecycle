<#
.SYNOPSIS
  One-time migration: adds "VA" as a valid choice on Employees.ELCRole
  and updates the ELC_Apps "elc" registry row to include VA in its Roles.
  Idempotent — re-running is a no-op once VA is present.

.DESCRIPTION
  Same hand-rolled device-code auth as provision-lists.ps1.
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

# ── DEVICE CODE AUTH ──
Write-Step "Requesting device code"
$scope  = "https://graph.microsoft.com/Sites.Manage.All https://graph.microsoft.com/Sites.ReadWrite.All offline_access"
$dcResp = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
    -Body @{ client_id = $ClientId; scope = $scope } -ContentType 'application/x-www-form-urlencoded'

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
Write-Host "  ║  ACTION REQUIRED — sign in to authorise list creation        ║" -ForegroundColor Yellow
Write-Host "  ║                                                              ║" -ForegroundColor Yellow
Write-Host "  ║  1. Open:   https://login.microsoft.com/device               ║" -ForegroundColor Yellow
Write-Host "  ║  2. Enter:  $($dcResp.user_code.PadRight(50))║" -ForegroundColor Yellow
Write-Host "  ║  3. Sign in as bturner@newshirepm.com                        ║" -ForegroundColor Yellow
Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
Write-Host ""

$expiresAt   = (Get-Date).AddSeconds([int]$dcResp.expires_in - 5)
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

# ── 1. PATCH Employees.ELCRole column to include "VA" ──
Write-Step "Locating Employees.ELCRole column"
$lists = Invoke-Graph GET "/sites/$SiteId/lists?`$select=id,name,displayName"
$empList = $lists.value | Where-Object { $_.name -eq 'Employees' -or $_.displayName -eq 'Employees' } | Select-Object -First 1
if (-not $empList) { Write-Error "Employees list not found on this site."; exit 1 }

$empCols = Invoke-Graph GET "/sites/$SiteId/lists/$($empList.id)/columns?`$select=id,name,displayName,choice"
$elcCol  = $empCols.value | Where-Object { $_.name -eq 'ELCRole' -or $_.displayName -eq 'ELCRole' } | Select-Object -First 1
if (-not $elcCol) {
    Write-New "ELCRole column missing — creating with full choice set"
    $newCol = @{ name='ELCRole'; choice = @{ allowTextEntry=$false; displayAs='dropDownMenu'; choices=@('None','Employee','Manager','HR','IT','Admin','VA') }; defaultValue=@{ value='None' } }
    Invoke-Graph POST "/sites/$SiteId/lists/$($empList.id)/columns" $newCol | Out-Null
    Write-OK "Created Employees.ELCRole"
} else {
    $current = @($elcCol.choice.choices)
    if ($current -contains 'VA') {
        Write-Skip "Employees.ELCRole already includes VA"
    } else {
        $updated = $current + 'VA'
        Write-New ("Adding VA to Employees.ELCRole (was: " + ($current -join ', ') + ")")
        Invoke-Graph PATCH "/sites/$SiteId/lists/$($empList.id)/columns/$($elcCol.id)" @{ choice = @{ allowTextEntry=$false; displayAs='dropDownMenu'; choices=$updated } } | Out-Null
        Write-OK "Patched Employees.ELCRole"
    }
}

# ── 2. UPDATE the ELC_Apps "elc" row to include VA in Roles JSON ──
Write-Step "Locating ELC_Apps registry row for 'elc'"
$appsList = $lists.value | Where-Object { $_.name -eq 'ELC_Apps' -or $_.displayName -eq 'ELC_Apps' } | Select-Object -First 1
if (-not $appsList) {
    Write-Skip "ELC_Apps list not found — skipping registry update (the in-app DEFAULT_APPS fallback will pick up the change)."
} else {
    $items = Invoke-Graph GET "/sites/$SiteId/lists/$($appsList.id)/items?expand=fields(`$select=Title,AppKey,Roles)&`$top=200"
    $elcRow = $items.value | Where-Object { $_.fields.AppKey -eq 'elc' } | Select-Object -First 1
    if (-not $elcRow) {
        Write-Skip "ELC_Apps row with AppKey='elc' not found — skip."
    } else {
        $rolesArr = @()
        try { $rolesArr = @($elcRow.fields.Roles | ConvertFrom-Json) } catch { $rolesArr = @() }
        if ($rolesArr -contains 'VA') {
            Write-Skip "ELC_Apps elc row already lists VA"
        } else {
            $updatedRoles = $rolesArr + 'VA'
            Write-New ("Adding VA to ELC_Apps elc.Roles (was: " + ($rolesArr -join ', ') + ")")
            Invoke-Graph PATCH "/sites/$SiteId/lists/$($appsList.id)/items/$($elcRow.id)/fields" @{ Roles = ($updatedRoles | ConvertTo-Json -Compress) } | Out-Null
            Write-OK "Patched ELC_Apps elc.Roles"
        }
    }
}

Write-Host ""
Write-Step "Done."
Write-Host "  Hard-refresh the app — the ELC Role dropdown should now include 'VA'." -ForegroundColor Gray

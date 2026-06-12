<#
.SYNOPSIS
  One-time migration: adds eight columns to the Employees list for
  primary + secondary emergency-contact info. Idempotent.
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

Write-Step "Locating Employees list"
$lists = Invoke-Graph GET "/sites/$SiteId/lists?`$select=id,name,displayName"
$empList = $lists.value | Where-Object { $_.name -eq 'Employees' -or $_.displayName -eq 'Employees' } | Select-Object -First 1
if (-not $empList) { Write-Error "Employees list not found."; exit 1 }

$empCols = Invoke-Graph GET "/sites/$SiteId/lists/$($empList.id)/columns?`$select=id,name,displayName"
$have = @{}; foreach ($c in $empCols.value) { $have[$c.name] = $true; if ($c.displayName) { $have[$c.displayName] = $true } }

$columnsToAdd = @(
    'EC1Name','EC1Relationship','EC1Phone','EC1Email',
    'EC2Name','EC2Relationship','EC2Phone','EC2Email'
)

foreach ($colName in $columnsToAdd) {
    if ($have[$colName]) {
        Write-Skip "Employees.$colName already exists"
    } else {
        Write-New "Adding Employees.$colName"
        Invoke-Graph POST "/sites/$SiteId/lists/$($empList.id)/columns" @{ name=$colName; text=@{} } | Out-Null
        Write-OK "Added Employees.$colName"
    }
}

Write-Host ""
Write-Step "Done."
Write-Host "  Hard-refresh the app, then open any employee -> Profile sub-tab to enter contacts." -ForegroundColor Gray

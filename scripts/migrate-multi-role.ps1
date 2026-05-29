<#
.SYNOPSIS
  One-time migration: lets Employees.ELCRole hold MULTIPLE roles for one person
  (e.g. "Admin; HR; IT") by enabling fill-in on the choice column. Also makes sure
  the choice set includes the four combinable roles. Idempotent — re-running is a no-op.

.DESCRIPTION
  Graph cannot convert a single-choice column to a true multi-choice column in place,
  so instead we set choice.allowTextEntry = $true. The app writes a delimited list of
  roles as the fill-in value (e.g. "Admin; HR; IT") and parses it back into a set. A
  plain single value like "Admin" still works unchanged.

  Same hand-rolled device-code auth as provision-lists.ps1 / migrate-add-va-role.ps1.
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

# -- DEVICE CODE AUTH --
Write-Step "Requesting device code"
$scope  = "https://graph.microsoft.com/Sites.Manage.All https://graph.microsoft.com/Sites.ReadWrite.All offline_access"
$dcResp = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
    -Body @{ client_id = $ClientId; scope = $scope } -ContentType 'application/x-www-form-urlencoded'

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Yellow
Write-Host "  ACTION REQUIRED - sign in to authorise the column change" -ForegroundColor Yellow
Write-Host "    1. Open:   https://login.microsoft.com/device" -ForegroundColor Yellow
Write-Host "    2. Enter:  $($dcResp.user_code)" -ForegroundColor Yellow
Write-Host "    3. Sign in as bturner@newshirepm.com" -ForegroundColor Yellow
Write-Host "  ============================================================" -ForegroundColor Yellow
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

# -- Enable fill-in on Employees.ELCRole so it can hold a delimited multi-role value --
Write-Step "Locating Employees.ELCRole column"
$lists = Invoke-Graph GET "/sites/$SiteId/lists?`$select=id,name,displayName"
$empList = $lists.value | Where-Object { $_.name -eq 'Employees' -or $_.displayName -eq 'Employees' } | Select-Object -First 1
if (-not $empList) { Write-Error "Employees list not found on this site."; exit 1 }

$empCols = Invoke-Graph GET "/sites/$SiteId/lists/$($empList.id)/columns?`$select=id,name,displayName,choice"
$elcCol  = $empCols.value | Where-Object { $_.name -eq 'ELCRole' -or $_.displayName -eq 'ELCRole' } | Select-Object -First 1
$wantChoices = @('None','Employee','Manager','HR','IT','Admin','VA')
if (-not $elcCol) {
    Write-New "ELCRole column missing - creating with fill-in enabled"
    $newCol = @{ name='ELCRole'; choice = @{ allowTextEntry=$true; displayAs='dropDownMenu'; choices=$wantChoices }; defaultValue=@{ value='None' } }
    Invoke-Graph POST "/sites/$SiteId/lists/$($empList.id)/columns" $newCol | Out-Null
    Write-OK "Created Employees.ELCRole (fill-in enabled)"
} else {
    $current   = @($elcCol.choice.choices)
    $merged    = @($current + ($wantChoices | Where-Object { $current -notcontains $_ }))
    $alreadyOk = ($elcCol.choice.allowTextEntry -eq $true) -and (@($wantChoices | Where-Object { $current -notcontains $_ }).Count -eq 0)
    if ($alreadyOk) {
        Write-Skip "Employees.ELCRole already allows fill-in and has all roles"
    } else {
        Write-New ("Enabling fill-in on Employees.ELCRole (allowTextEntry was: " + $elcCol.choice.allowTextEntry + ")")
        Invoke-Graph PATCH "/sites/$SiteId/lists/$($empList.id)/columns/$($elcCol.id)" @{ choice = @{ allowTextEntry=$true; displayAs='dropDownMenu'; choices=$merged } } | Out-Null
        Write-OK "Patched Employees.ELCRole"
    }
}

Write-Host ""
Write-Step "Done."
Write-Host "  ELCRole now accepts a delimited list, e.g. 'Admin; HR; IT'." -ForegroundColor Gray
Write-Host "  Assign multiple roles from the app: Setup -> Access to this app." -ForegroundColor Gray

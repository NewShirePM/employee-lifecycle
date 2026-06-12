<#
.SYNOPSIS
  Provisions all SharePoint lists + document library required by the
  Employee Lifecycle app on the NewShire (vanrockre) SharePoint site.

.DESCRIPTION
  Hand-rolled OAuth2 device-code flow against the Microsoft identity
  platform, then raw Invoke-RestMethod calls against Microsoft Graph.
  No Connect-MgGraph (its 120s inactivity timeout was unworkable).

  Idempotent: re-running is safe; existing lists/columns are left alone.

.NOTES
  Required Graph scopes: Sites.Manage.All, Sites.ReadWrite.All
  Tested on PowerShell 7.x.
#>
[CmdletBinding()]
param(
    [string]$SiteId   = "vanrockre.sharepoint.com,a02c1cd8-9f1f-4827-8286-7b6b7ce74232,01202419-6625-4499-b0d5-8ceb1cffdba3",
    [string]$TenantId = "33575d04-ca7b-4396-8011-9eaea4030b46",
    # Public "Microsoft Graph Command Line Tools" client id (device-flow enabled, no app registration needed)
    [string]$ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e",
    [switch]$SkipSeed
)

$ErrorActionPreference = 'Stop'
$InformationPreference  = 'Continue'

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-OK   ($msg) { Write-Host "    [OK]    $msg" -ForegroundColor Green }
function Write-Skip ($msg) { Write-Host "    [SKIP]  $msg" -ForegroundColor DarkYellow }
function Write-New  ($msg) { Write-Host "    [CREATE]$msg" -ForegroundColor Yellow }
function Write-Warn ($msg) { Write-Host "    [WARN]  $msg" -ForegroundColor Magenta }

# ─────────────────────────────────────────────────────────────
# DEVICE CODE AUTH  (hand-rolled, 15-minute polling window)
# ─────────────────────────────────────────────────────────────
Write-Step "Requesting device code"
$scope = "https://graph.microsoft.com/Sites.Manage.All https://graph.microsoft.com/Sites.ReadWrite.All offline_access"
$dcResp = Invoke-RestMethod -Method POST `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
    -Body @{ client_id = $ClientId; scope = $scope } `
    -ContentType 'application/x-www-form-urlencoded'

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
Write-Host "  ║  ACTION REQUIRED — sign in to authorise list creation        ║" -ForegroundColor Yellow
Write-Host "  ║                                                              ║" -ForegroundColor Yellow
Write-Host "  ║  1. Open:   https://login.microsoft.com/device               ║" -ForegroundColor Yellow
Write-Host "  ║  2. Enter:  $($dcResp.user_code.PadRight(50))║" -ForegroundColor Yellow
Write-Host "  ║  3. Sign in as bturner@newshirepm.com                        ║" -ForegroundColor Yellow
Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
Write-Host ""
Write-Host "Polling for token (device code expires in $([int]($dcResp.expires_in/60)) minutes)..." -ForegroundColor DarkGray

$token = $null
$expiresAt = (Get-Date).AddSeconds([int]$dcResp.expires_in - 5)
$pollInterval = [int]$dcResp.interval
if ($pollInterval -lt 5) { $pollInterval = 5 }

while ((Get-Date) -lt $expiresAt) {
    Start-Sleep -Seconds $pollInterval
    try {
        $tokenResp = Invoke-RestMethod -Method POST `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Body @{
                grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id   = $ClientId
                device_code = $dcResp.device_code
            } `
            -ContentType 'application/x-www-form-urlencoded' `
            -ErrorAction Stop
        $token = $tokenResp.access_token
        break
    } catch {
        $err = $null
        try { $err = ($_.ErrorDetails.Message | ConvertFrom-Json) } catch { }
        if ($err -and $err.error -eq 'authorization_pending') {
            Write-Host "." -NoNewline -ForegroundColor DarkGray
            continue
        }
        if ($err -and $err.error -eq 'slow_down') {
            $pollInterval += 5; continue
        }
        if ($err -and $err.error -eq 'expired_token') {
            Write-Host ""; Write-Error "Device code expired before sign-in completed. Re-run the script."; exit 1
        }
        if ($err -and $err.error -eq 'authorization_declined') {
            Write-Host ""; Write-Error "Authorization declined by user."; exit 1
        }
        Write-Host ""
        Write-Error "Token poll failed: $($err.error_description ?? $_.Exception.Message)"
        exit 1
    }
}
Write-Host ""
if (-not $token) { Write-Error "Authentication did not complete within the device-code lifetime."; exit 1 }
Write-OK "Authenticated"

$headers = @{ Authorization = "Bearer $token" }

# ─────────────────────────────────────────────────────────────
# GRAPH HELPERS
# ─────────────────────────────────────────────────────────────
function Invoke-Graph {
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST','PATCH','DELETE')][string]$Method,
        [Parameter(Mandatory)][string]$Path,
        [object]$Body
    )
    $uri = "https://graph.microsoft.com/v1.0$Path"
    if ($Body) {
        $json = $Body | ConvertTo-Json -Depth 20 -Compress
        return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -Body $json -ContentType 'application/json'
    }
    return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers
}

# ─────────────────────────────────────────────────────────────
# COLUMN BUILDERS
# ─────────────────────────────────────────────────────────────
function ColText      ($n)         { @{ name = $n; text = @{ } } }
function ColMultiText ($n)         { @{ name = $n; text = @{ allowMultipleLines = $true; appendChangesToExistingText = $false; linesForEditing = 6 } } }
function ColNumber    ($n)         { @{ name = $n; number = @{ } } }
function ColDate      ($n)         { @{ name = $n; dateTime = @{ format = 'dateOnly' } } }
function ColDateTime  ($n)         { @{ name = $n; dateTime = @{ format = 'dateTime' } } }
function ColYesNo     ($n,$def=$false) { @{ name = $n; boolean = @{ } ; defaultValue = @{ value = ([string]$def).ToLower() } } }
function ColChoice    ($n,[string[]]$choices,[string]$default = '') {
    $c = @{ allowTextEntry = $false; choices = $choices; displayAs = 'dropDownMenu' }
    $h = @{ name = $n; choice = $c }
    if ($default) { $h.defaultValue = @{ value = $default } }
    $h
}

# ─────────────────────────────────────────────────────────────
# DISCOVERY
# ─────────────────────────────────────────────────────────────
Write-Step "Reading existing lists on site"
$listsResp = Invoke-Graph -Method GET -Path "/sites/$SiteId/lists?`$select=id,displayName,name"
$existing = @{}
foreach ($l in $listsResp.value) {
    $existing[$l.displayName.ToLower()] = $l
    if ($l.name) { $existing[$l.name.ToLower()] = $l }
}
Write-OK "Found $($listsResp.value.Count) existing list(s)/library/libraries"

function Ensure-List {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][array]$Columns,
        [string]$Template = 'genericList',
        [string]$Description = ''
    )
    $found = $existing[$Name.ToLower()]
    if ($found) {
        Write-Skip "$Name already exists  (id=$($found.id))"
        try {
            $colResp = Invoke-Graph -Method GET -Path "/sites/$SiteId/lists/$($found.id)/columns?`$select=name,displayName"
            $haveCols = @{}
            foreach ($c in $colResp.value) {
                $haveCols[$c.name.ToLower()] = $true
                if ($c.displayName) { $haveCols[$c.displayName.ToLower()] = $true }
            }
            foreach ($col in $Columns) {
                if (-not $haveCols[$col.name.ToLower()]) {
                    Write-New "  + column $($col.name)"
                    Invoke-Graph -Method POST -Path "/sites/$SiteId/lists/$($found.id)/columns" -Body $col | Out-Null
                }
            }
        } catch {
            Write-Warn "Column reconcile failed for $Name : $($_.Exception.Message)"
        }
        return $found
    }
    Write-New "$Name"
    $body = @{
        displayName = $Name
        description = $Description
        columns     = $Columns
        list        = @{ template = $Template }
    }
    try {
        $created = Invoke-Graph -Method POST -Path "/sites/$SiteId/lists" -Body $body
        $existing[$Name.ToLower()] = $created
        Write-OK "Created $Name  (id=$($created.id))"
        return $created
    } catch {
        Write-Warn "Create failed for $Name : $($_.Exception.Message)"
        throw
    }
}

# ─────────────────────────────────────────────────────────────
# LISTS
# ─────────────────────────────────────────────────────────────
Write-Step "Provisioning ELC_Journeys"
$ELC_Journeys = Ensure-List -Name 'ELC_Journeys' -Description 'One row per onboarding/offboarding instance' -Columns @(
    (ColChoice    'JourneyType' @('Onboarding','Offboarding') 'Onboarding')
    (ColText      'EmployeeEmail')
    (ColText      'EmployeeName')
    (ColText      'JobTitle')
    (ColText      'ManagerEmail')
    (ColText      'Department')
    (ColDate      'StartDate')
    (ColDate      'EndDate')
    (ColChoice    'Status' @('In Progress','On Hold','Complete','Cancelled') 'In Progress')
    (ColMultiText 'Notes')
    (ColText      'CreatedBy')
    (ColText      'OffboardReason')
    (ColText      'TemplateGroup')
    (ColYesNo     'SharedWithEmployee' $false)
)

Write-Step "Provisioning ELC_TemplateTasks"
$ELC_TemplateTasks = Ensure-List -Name 'ELC_TemplateTasks' -Description 'Reusable task templates for onboarding/offboarding flows' -Columns @(
    (ColChoice    'JourneyType' @('Onboarding','Offboarding') 'Onboarding')
    (ColText      'TemplateGroup')
    (ColText      'Phase')
    (ColChoice    'AssigneeRole' @('HR','IT','Manager','Employee','Admin','Accounting') 'HR')
    (ColNumber    'OffsetDays')
    (ColYesNo     'Required' $true)
    (ColYesNo     'Active' $true)
    (ColMultiText 'Notes')
    (ColText      'RoleScope')
    (ColNumber    'OrderIdx')
)

Write-Step "Provisioning ELC_JourneyTasks"
$ELC_JourneyTasks = Ensure-List -Name 'ELC_JourneyTasks' -Description 'Per-journey tasks materialised from templates' -Columns @(
    (ColText      'JourneyId')
    (ColText      'EmployeeEmail')
    (ColText      'Phase')
    (ColChoice    'AssigneeRole' @('HR','IT','Manager','Employee','Admin','Accounting') 'HR')
    (ColText      'AssigneeEmail')
    (ColDate      'DueDate')
    (ColNumber    'OffsetDays')
    (ColYesNo     'Required' $true)
    (ColChoice    'Status' @('Pending','In Progress','Done','Blocked','N/A') 'Pending')
    (ColMultiText 'Notes')
    (ColDate      'CompletedDate')
    (ColText      'CompletedBy')
    (ColNumber    'OrderIdx')
    (ColText      'TemplateId')
)

Write-Step "Provisioning ELC_Config"
$ELC_Config = Ensure-List -Name 'ELC_Config' -Description 'Single-row JSON blob for app-wide settings' -Columns @(
    (ColMultiText 'ConfigJSON')
)

Write-Step "Provisioning ELC_Apps"
$ELC_Apps = Ensure-List -Name 'ELC_Apps' -Description 'Registry of NewShire internal apps and the role columns they read from Employees' -Columns @(
    (ColText      'AppKey')
    (ColText      'ColumnName')
    (ColMultiText 'Roles')
    (ColText      'IconLetter')
    (ColText      'Color')
    (ColYesNo     'Active' $true)
    (ColMultiText 'Description')
    (ColNumber    'OrderIdx')
    (ColText      'AppUrl')
    (ColChoice    'OnboardingDefault' @('None','Lowest','Standard','Manager','Admin') 'None')
)

Write-Step "Provisioning ELC_EmployeeNotes"
$ELC_EmployeeNotes = Ensure-List -Name 'ELC_EmployeeNotes' -Description 'Coaching notes, disciplinary records, performance check-ins' -Columns @(
    (ColText      'EmployeeEmail')
    (ColChoice    'NoteType'  @('Coaching','Discipline','Praise','PIP','1:1','Termination','Policy Acknowledgement','General') 'Coaching')
    (ColDate      'NoteDate')
    (ColText      'AuthorEmail')
    (ColMultiText 'Body')
    (ColYesNo     'Confidential' $true)
    (ColYesNo     'VisibleToManager' $true)
    (ColYesNo     'VisibleToEmployee' $false)
    (ColDate      'FollowUpDate')
    (ColChoice    'Status' @('Open','In Progress','Resolved','Escalated') 'Open')
    (ColChoice    'Severity' @('Info','Low','Medium','High','Critical') 'Info')
    (ColMultiText 'AttachmentLinks')
    (ColText      'RelatedJourneyId')
    (ColText      'AcknowledgedBy')
    (ColDate      'AcknowledgedDate')
)

Write-Step "Provisioning ELC_PermissionAudit"
$ELC_PermissionAudit = Ensure-List -Name 'ELC_PermissionAudit' -Description 'Audit trail for changes to per-app employee permissions' -Columns @(
    (ColText      'EmployeeEmail')
    (ColText      'AppKey')
    (ColText      'OldRole')
    (ColText      'NewRole')
    (ColText      'ChangedBy')
    (ColDateTime  'ChangedAt')
    (ColMultiText 'Reason')
    (ColText      'JourneyId')
)

Write-Step "Provisioning ELC_EmployeeFiles (document library)"
$ELC_EmployeeFiles = Ensure-List -Name 'ELC_EmployeeFiles' -Template 'documentLibrary' -Description 'Per-employee documents (offer letters, signed policies, disciplinary docs, exit paperwork)' -Columns @(
    (ColText   'EmployeeEmail')
    (ColChoice 'DocCategory' @('Onboarding','Offboarding','Policy','Disciplinary','Coaching','Benefits','I-9 / W-4','Performance Review','Other') 'Other')
    (ColYesNo  'Confidential' $true)
    (ColDate   'DocDate')
)

Write-Step "Provisioning ELC_QuarterlyReviews"
$ELC_QuarterlyReviews = Ensure-List -Name 'ELC_QuarterlyReviews' -Description 'Quarterly performance reviews per employee' -Columns @(
    (ColText      'EmployeeEmail')
    (ColText      'ReviewPeriod')
    (ColDate      'DueDate')
    (ColDate      'ConductedDate')
    (ColText      'ReviewerEmail')
    (ColChoice    'Rating' @('Outstanding','Exceeds Expectations','Meets Expectations','Needs Improvement','Below Expectations') 'Meets Expectations')
    (ColChoice    'Status' @('Scheduled','In Progress','Conducted','Acknowledged','Cancelled') 'Scheduled')
    (ColMultiText 'Strengths')
    (ColMultiText 'GrowthAreas')
    (ColMultiText 'GoalsNextQuarter')
    (ColMultiText 'EmployeeComments')
    (ColMultiText 'AttachmentLinks')
    (ColDate      'AcknowledgedDate')
    (ColYesNo     'Confidential' $true)
)

Write-Step "Provisioning ELC_PayChanges"
$ELC_PayChanges = Ensure-List -Name 'ELC_PayChanges' -Description 'Compensation history per employee — every pay change is one row' -Columns @(
    (ColText      'EmployeeEmail')
    (ColDate      'EffectiveDate')
    (ColNumber    'PreviousPay')
    (ColNumber    'NewPay')
    (ColChoice    'PayType' @('Hourly','Salary','Salary + Commission','Commission Only','1099 Contract','Other') 'Hourly')
    (ColChoice    'ChangeType' @('Initial Hire','Merit Increase','Promotion','Cost of Living','Market Adjustment','Schedule Change','Demotion','Other') 'Merit Increase')
    (ColText      'ApprovedBy')
    (ColMultiText 'Reason')
    (ColText      'RelatedReviewId')
    (ColYesNo     'Confidential' $true)
)

Write-Step "Provisioning ELC_EmailTemplates"
$ELC_EmailTemplates = Ensure-List -Name 'ELC_EmailTemplates' -Description 'Reusable email templates with {{var}} substitution' -Columns @(
    (ColChoice    'Category' @('Welcome','Onboarding','Reminder','Performance','Offboarding','Celebration','Other') 'Other')
    (ColText      'Subject')
    (ColMultiText 'Body')
    (ColText      'DefaultTo')
    (ColText      'DefaultCc')
    (ColYesNo     'Active' $true)
    (ColMultiText 'Notes')
)

# ─────────────────────────────────────────────────────────────
# SEED — Apps registry
# ─────────────────────────────────────────────────────────────
if (-not $SkipSeed) {
    Write-Step "Seeding ELC_Apps registry"
    $appsExist = Invoke-Graph -Method GET -Path "/sites/$SiteId/lists/$($ELC_Apps.id)/items?expand=fields(`$select=Title,AppKey)&`$top=100"
    $appsHave = @{}
    foreach ($it in $appsExist.value) { if ($it.fields.AppKey) { $appsHave[$it.fields.AppKey.ToLower()] = $true } }

    $appsSeed = @(
        @{ Title='VA Tracker';          AppKey='vatracker';   ColumnName='VATrackerRole';   IconLetter='V'; Color='#3A6577'; Roles='["VA","Manager","Regional","Admin"]';        OrderIdx=10; Description='Productivity tracker for virtual assistants';      OnboardingDefault='None' }
        @{ Title='PM Hub';              AppKey='pmhub';       ColumnName='PMHubRole';        IconLetter='P'; Color='#1C3740'; Roles='["PM","Regional","Owner","Admin"]';         OrderIdx=20; Description='Property manager activity & help queue';            OnboardingDefault='None' }
        @{ Title='NewShire University'; AppKey='university';  ColumnName='UniversityRole';   IconLetter='U'; Color='#CDA04B'; Roles='["Employee","Manager","Admin"]';            OrderIdx=30; Description='Training, courses, compliance learning paths';      OnboardingDefault='Lowest' }
        @{ Title='Expense Manager';     AppKey='expense';     ColumnName='ExpenseRole';      IconLetter='E'; Color='#2D8A5A'; Roles='["Employee","Accounting","Admin"]';         OrderIdx=40; Description='Mileage & expense reimbursement';                  OnboardingDefault='Lowest' }
        @{ Title='CAHP Compliance Hub'; AppKey='cahp';        ColumnName='CAHPRole';         IconLetter='C'; Color='#4A78B0'; Roles='["Viewer","Editor","Admin"]';               OrderIdx=50; Description='Affordable housing compliance tracker';            OnboardingDefault='None' }
        @{ Title='AppFolio Dashboard';  AppKey='appfolio';    ColumnName='AppFolioRole';     IconLetter='A'; Color='#B8922E'; Roles='["Viewer","Editor","Admin"]';               OrderIdx=60; Description='NewShire AppFolio analytics dashboard';            OnboardingDefault='None' }
        @{ Title='ShowMojo Sync';       AppKey='showmojo';    ColumnName='ShowMojoRole';     IconLetter='S'; Color='#5B3FA8'; Roles='["Viewer","Editor","Admin"]';               OrderIdx=70; Description='ShowMojo listing/showings sync';                   OnboardingDefault='None' }
        @{ Title='Renewal Manager';     AppKey='renewal';     ColumnName='RenewalRole';      IconLetter='R'; Color='#C44B3B'; Roles='["Viewer","Editor","Admin"]';               OrderIdx=80; Description='Lease renewal workflow & tracking';                OnboardingDefault='None' }
        @{ Title='Employee Lifecycle';  AppKey='elc';         ColumnName='ELCRole';          IconLetter='L'; Color='#CDA04B'; Roles='["Employee","Manager","HR","IT","Admin"]';  OrderIdx=90; Description='Onboarding / offboarding / HR file system';        OnboardingDefault='Lowest' }
    )

    foreach ($a in $appsSeed) {
        if ($appsHave[$a.AppKey.ToLower()]) { Write-Skip "$($a.Title)"; continue }
        Write-New $a.Title
        $body = @{ fields = $a }
        Invoke-Graph -Method POST -Path "/sites/$SiteId/lists/$($ELC_Apps.id)/items" -Body $body | Out-Null
    }
}

# ─────────────────────────────────────────────────────────────
# Per-app role columns on Employees list (additive)
# ─────────────────────────────────────────────────────────────
Write-Step "Reconciling per-app role columns on Employees list"
$empList = $existing['employees']
if (-not $empList) {
    Write-Warn "Employees list not found on this site — skipping per-app columns."
} else {
    $empColsResp = Invoke-Graph -Method GET -Path "/sites/$SiteId/lists/$($empList.id)/columns?`$select=name,displayName"
    $empCols = @{}
    foreach ($c in $empColsResp.value) {
        $empCols[$c.name.ToLower()] = $true
        if ($c.displayName) { $empCols[$c.displayName.ToLower()] = $true }
    }
    $appsToCheck = Invoke-Graph -Method GET -Path "/sites/$SiteId/lists/$($ELC_Apps.id)/items?expand=fields(`$select=Title,AppKey,ColumnName,Roles)&`$top=100"
    foreach ($it in $appsToCheck.value) {
        $f = $it.fields
        if (-not $f.ColumnName) { continue }
        if ($empCols[$f.ColumnName.ToLower()]) { Write-Skip "Employees.$($f.ColumnName)"; continue }
        $roles = @()
        try { $roles = ($f.Roles | ConvertFrom-Json) } catch { $roles = @() }
        $choices = @('None') + $roles
        $col = ColChoice $f.ColumnName $choices 'None'
        Write-New "Employees.$($f.ColumnName)  (choices: $($choices -join ', '))"
        try {
            Invoke-Graph -Method POST -Path "/sites/$SiteId/lists/$($empList.id)/columns" -Body $col | Out-Null
        } catch {
            Write-Warn "Could not add column $($f.ColumnName) to Employees: $($_.Exception.Message)"
        }
    }
}

Write-Host ""
Write-Step "Done."
Write-Host "  Lists provisioned (or already present):" -ForegroundColor Gray
@(
    'ELC_Journeys', 'ELC_TemplateTasks', 'ELC_JourneyTasks', 'ELC_Config',
    'ELC_Apps', 'ELC_EmployeeNotes', 'ELC_PermissionAudit', 'ELC_EmployeeFiles',
    'ELC_QuarterlyReviews', 'ELC_PayChanges', 'ELC_EmailTemplates'
) | ForEach-Object { Write-Host "    - $_" -ForegroundColor Gray }
Write-Host ""

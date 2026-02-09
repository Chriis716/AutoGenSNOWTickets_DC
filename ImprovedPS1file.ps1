Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ----------------------------
# Startup / Assemblies
# ----------------------------
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Load DLLs relative to script root (assembly\*.dll)
$DllRoot = Join-Path $ScriptRoot "assembly"
$Dlls = @(
    "MahApps.Metro.dll",
    "System.Windows.Interactivity.dll",
    "LoadingIndicators.WPF.dll"
)

foreach ($dll in $Dlls) {
    $p = Join-Path $DllRoot $dll
    if (-not (Test-Path $p)) { throw "Missing dependency: $p" }
    [void][System.Reflection.Assembly]::LoadFrom($p)
}

function LoadXml {
    param([Parameter(Mandatory)][string]$Path)

    $doc = New-Object System.Xml.XmlDocument
    $doc.Load($Path)
    return $doc
}

# ----------------------------
# ServiceNow config
# ----------------------------
$Config = [ordered]@{
    BrowserExe   = "msedge.exe"
    HomeUrl      = "https://yourit.va.gov/home.do"
    TermsUrl     = "https://yourIT.VA.GOV/va?id=va_termsandconditions"
    InstanceBase = "https://yourit.va.gov"

    AssignmentGroupSysId = "4e965a49db5dd3c0b857fd721f9619d5"  # Team3
    AffectedCI           = "bbf1bbdf1b88ec5006020f6fe54bcb68"
    DueTime24            = "14:30:00"
}

# ----------------------------
# Helpers
# ----------------------------
function Invoke-SNowPreflightTerms {
    # Opens Terms page and pauses so user can click Accept (reliable w/o Selenium/Playwright)
    Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-window", $Config.TermsUrl) | Out-Null

    [void][System.Windows.MessageBox]::Show(
        "ServiceNow Terms & Conditions page was opened in Edge.`n`nClick ACCEPT, then press OK to continue.",
        "ServiceNow Preflight",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
}

function Get-WorkWeekDates {
    param([datetime]$ReferenceDate = (Get-Date))

    # Sunday=0 ... Saturday=6
    $dow = [int]$ReferenceDate.DayOfWeek

    if ($dow -eq 0) {        # Sunday -> next Monday
        $monday = $ReferenceDate.Date.AddDays(1)
    }
    elseif ($dow -eq 6) {    # Saturday -> next Monday
        $monday = $ReferenceDate.Date.AddDays(2)
    }
    else {
        # Monday(1) .. Friday(5)
        $monday = $ReferenceDate.Date.AddDays(-($dow - 1))
    }

    [pscustomobject]@{
        Monday    = $monday
        Tuesday   = $monday.AddDays(1)
        Wednesday = $monday.AddDays(2)
        Thursday  = $monday.AddDays(3)
        Friday    = $monday.AddDays(4)
    }
}

function To-SNowDate([datetime]$d) { $d.ToString("MM/dd/yyyy") }

function New-SNowClassicUrl {
    <#
      Builds Classic form URL:
      https://instance/now/nav/ui/classic/params/target/<table>.do%3Fsys_id%3D-1&sysparm_query=...
    #>
    param(
        [Parameter(Mandatory)][string]$Table,
        [Parameter(Mandatory)][hashtable]$Fields
    )

    $pairs = foreach ($k in $Fields.Keys) {
        $v = $Fields[$k]
        if ($null -eq $v) { continue }
        "$k=$v"
    }

    $raw = ($pairs -join "^")
    $encoded = [System.Uri]::EscapeDataString($raw)

    "$($Config.InstanceBase)/now/nav/ui/classic/params/target/$Table.do%3Fsys_id%3D-1&sysparm_query=$encoded"
}

function Get-SNowSysIdFromText {
    param([Parameter(Mandatory)][string]$Text)

    # Supports sys_id=xxxx or sys_id%3Dxxxx or any string containing it
    if ($Text -match '(?:sys_id=|sys_id%3D)([0-9a-fA-F]{32})') { return $matches[1] }
    return $null
}

# ----------------------------
# Ticket / Task builders
# ----------------------------
function New-WorkTicketUrl {
    param([Parameter(Mandatory)][string]$Site)

    $week = Get-WorkWeekDates
    $m = To-SNowDate $week.Monday
    $f = To-SNowDate $week.Friday
    $due = "$f $($Config.DueTime24)"

    $short = "$Site - Daily Checks Work Ticket $m to $f"
    $desc  = "$Site Daily Checks for Workweek: $m to $f"

    New-SNowClassicUrl -Table "u_work_ticket" -Fields @{
        type              = "normal"
        impact            = 4
        assignment_group  = $Config.AssignmentGroupSysId
        assigned_to       = "javascript:gs.user_id()"
        short_description = $short
        description       = $desc
        due_date          = $due
        cmdb_ci           = $Config.AffectedCI
    }
}

function New-WorkTaskUrls {
    param(
        [Parameter(Mandatory)][string]$Site,
        [Parameter(Mandatory)][string]$WorkTicketSysId
    )

    $week = Get-WorkWeekDates

    $taskDesc = @"
$Site Perform Daily Checks on the VistA Imaging System In Accordance With the SOP and Training documents

Training Modules:
https://dvagov.sharepoint.com/sites/oitspmhsphismclinternal/Training/Forms/AllItems.aspx?id=%2Fsites%2Foitspmhsphismclinternal%2FTraining%2FDaily%20Checks&viewid=6a448197%2Ddc45%2D4f14%2Da02d%2Db306278fe303

SPM-HISM-CL Daily Checks SOP:
https://dvagov.sharepoint.com/sites/oitspmhsphismclinternal/SitePages/Clinical-Imaging-SOP-Guidance.aspx
"@

    $days = @(
        @{ Name="Monday";    Date=$week.Monday    }
        @{ Name="Tuesday";   Date=$week.Tuesday   }
        @{ Name="Wednesday"; Date=$week.Wednesday }
        @{ Name="Thursday";  Date=$week.Thursday  }
        @{ Name="Friday";    Date=$week.Friday    }
    )

    $urls = New-Object System.Collections.Generic.List[string]

    foreach ($d in $days) {
        $dateString = To-SNowDate $d.Date
        $short = "$Site - $($d.Name) Daily Checks, Date: $dateString"

        $urls.Add( (New-SNowClassicUrl -Table "u_work_task" -Fields @{
            type              = "normal"
            impact            = 4
            urgency           = 4
            assignment_group  = $Config.AssignmentGroupSysId
            assigned_to       = "javascript:gs.user_id()"
            short_description = $short
            description       = $taskDesc
            cmdb_ci           = $Config.AffectedCI
            u_requestor       = "javascript:gs.user_id()"
            u_affected_user   = "javascript:gs.user_id()"
            u_work_ticket     = $WorkTicketSysId
        }) )
    }

    return $urls
}

function New-WeeklyWorkTaskUrl {
    param(
        [Parameter(Mandatory)][string]$Site,
        [Parameter(Mandatory)][string]$WorkTicketSysId
    )

    $week = Get-WorkWeekDates
    $m = To-SNowDate $week.Monday
    $f = To-SNowDate $week.Friday

    $short = "$Site - Weekly Task for Workweek $m to $f"

    $weeklyDesc = @"
Weekly Work Task Ticket for cleaning up Async
Check the Async Storage Request Errors under the Hybrid DICOM Gateway:

Login to VistA
At the prompt type
Hybrid DICOM Gateway Menu
At the Prompt type
- Find Async storage request errors: none found stop here
- if some found: requeue Async Storage request errors

Any Issues? Enter an Incident ticket with CLIN3

Training Modules:
https://dvagov.sharepoint.com/sites/oitspmhsphismclinternal/Training/Forms/AllItems.aspx?id=%2Fsites%2Foitspmhsphismclinternal%2FTraining%2FDaily%20Checks&viewid=6a448197%2Ddc45%2D4f14%2Da02d%2Db306278fe303

SPM-HISM-CL Daily Checks SOP:
https://dvagov.sharepoint.com/sites/oitspmhsphismclinternal/SitePages/Clinical-Imaging-SOP-Guidance.aspx
"@

    New-SNowClassicUrl -Table "u_work_task" -Fields @{
        type              = "normal"
        impact            = 4
        urgency           = 4
        assignment_group  = $Config.AssignmentGroupSysId
        assigned_to       = "javascript:gs.user_id()"
        short_description = $short
        description       = $weeklyDesc
        cmdb_ci           = $Config.AffectedCI
        u_requestor       = "javascript:gs.user_id()"
        u_affected_user   = "javascript:gs.user_id()"
        u_work_ticket     = $WorkTicketSysId
    }
}

# ----------------------------
# Load Main Window
# ----------------------------
$XamlPath = Join-Path $ScriptRoot "SNOW_DailyChecks.xaml"
$XamlMainWindow = LoadXml -Path $XamlPath
$Reader = New-Object System.Xml.XmlNodeReader $XamlMainWindow
$Form = [Windows.Markup.XamlReader]::Load($Reader)

$XamlMainWindow.SelectNodes("//*[@Name]") | ForEach-Object {
    Set-Variable -Name ("WPF_{0}" -f $_.Name) -Value $Form.FindName($_.Name)
}

# script-scoped state (cleaner than $Global:)
$script:CurrentSite = $null

# ----------------------------
# UI Events
# ----------------------------
$WPF_XMLbtnSelectInstallFile.Add_Click({
    try {
        $site = ($WPF_XMLAddInstallFilePath.Text ?? "").Trim()
        if ([string]::IsNullOrWhiteSpace($site)) {
            [void][System.Windows.MessageBox]::Show("Enter Site/Site Code (ex: TVH/626).")
            return
        }

        $script:CurrentSite = $site

        Invoke-SNowPreflightTerms

        Write-Host "Opening ServiceNow Home..." -ForegroundColor Cyan
        Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-window", $Config.HomeUrl) | Out-Null
        Start-Sleep -Seconds 4

        $ticketUrl = New-WorkTicketUrl -Site $site
        Write-Host "Creating Work Ticket for $site..." -ForegroundColor Cyan
        Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-window", $ticketUrl) | Out-Null

        Write-Host "Work Ticket window opened." -ForegroundColor Green
    }
    catch {
        [void][System.Windows.MessageBox]::Show($_.Exception.Message)
    }
})

$WPF_StartButton.Add_Click({
    try {
        $site = ($script:CurrentSite ?? "").Trim()
        if ([string]::IsNullOrWhiteSpace($site)) {
            # fallback if they skipped first button
            $site = ($WPF_XMLAddInstallFilePath.Text ?? "").Trim()
            if ([string]::IsNullOrWhiteSpace($site)) {
                [void][System.Windows.MessageBox]::Show("Enter Site/Site Code first (ex: TVH/626).")
                return
            }
            $script:CurrentSite = $site
        }

        $pasted = ($WPF_XMLEnterFQDNInfo.Text ?? "").Trim()
        if ([string]::IsNullOrWhiteSpace($pasted)) {
            [void][System.Windows.MessageBox]::Show("Paste the Work Ticket URL (or any text containing sys_id=...).")
            return
        }

        $id = Get-SNowSysIdFromText -Text $pasted
        if (-not $id) {
            [void][System.Windows.MessageBox]::Show("Could not find a 32-character sys_id in what you pasted.")
            return
        }

        Invoke-SNowPreflightTerms

        Write-Host "Creating Daily Work Tasks..." -ForegroundColor Cyan
        $taskUrls = New-WorkTaskUrls -Site $site -WorkTicketSysId $id
        foreach ($u in $taskUrls) {
            Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-tab", $u) | Out-Null
            Start-Sleep -Milliseconds 400
        }

        Write-Host "Creating Weekly Work Task..." -ForegroundColor Cyan
        $weeklyUrl = New-WeeklyWorkTaskUrl -Site $site -WorkTicketSysId $id
        Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-tab", $weeklyUrl) | Out-Null

        Write-Host "All tasks opened." -ForegroundColor Green
    }
    catch {
        [void][System.Windows.MessageBox]::Show($_.Exception.Message)
    }
})

$WPF_Close.Add_Click({ $Form.Close() })

[System.GC]::Collect()
$Form.ShowDialog() | Out-Null


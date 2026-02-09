Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# =========================
# Paths / Assemblies
# =========================
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Load local DLLs (assembly\*.dll next to script)
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

# =========================
# Config (edit these)
# =========================
$Config = [ordered]@{
    BrowserExe   = "msedge.exe"
    HomeUrl      = "https://yourit.va.gov/home.do"
    InstanceBase = "https://yourit.va.gov"

    AssignmentGroupSysId = "4e965a49db5dd3c0b857fd721f9619d5"   # Team3 sys_id
    AffectedCI           = "bbf1bbdf1b88ec5006020f6fe54bcb68"   # CI sys_id
    DueTime24            = "14:30:00"
}

# =========================
# Helpers
# =========================
function Get-TextSafe {
    param($Value)
    if ($null -eq $Value) { return "" }
    return [string]$Value
}

function Load-XmlXaml {
    param([Parameter(Mandatory)][string]$Path)

    $doc = New-Object System.Xml.XmlDocument
    $doc.Load($Path)
    return $doc
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

function ConvertTo-SNowQueryValue {
    <#
      Encodes a value for sysparm_query WITHOUT double-encoding existing %HH sequences.
      - Preserves newlines as %0D%0A
      - Preserves existing URL encodes like %2F %3A (so SharePoint URLs survive)
      - Encodes everything else as UTF-8 percent encoding
    #>
    param([Parameter(Mandatory)][string]$Text)

    # Normalize line breaks -> encode new lines for ServiceNow
    $t = $Text -replace "`r`n", "`n"
    $t = $t -replace "`r", "`n"
    $t = $t -replace "`n", "%0D%0A"

    $sb = New-Object System.Text.StringBuilder

    for ($i = 0; $i -lt $t.Length; $i++) {
        $ch = $t[$i]

        # Preserve existing %HH sequences
        if ($ch -eq '%' -and ($i + 2) -lt $t.Length) {
            $h1 = $t[$i + 1]
            $h2 = $t[$i + 2]
            if ($h1 -match '[0-9A-Fa-f]' -and $h2 -match '[0-9A-Fa-f]') {
                [void]$sb.Append('%').Append($h1).Append($h2)
                $i += 2
                continue
            }
        }

        # Allow unreserved characters through (RFC3986)
        if ($ch -match '[A-Za-z0-9\-\._~]') {
            [void]$sb.Append($ch)
            continue
        }

        # Encode everything else (UTF-8 bytes -> %HH)
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$ch)
        foreach ($b in $bytes) {
            [void]$sb.AppendFormat("%{0:X2}", $b)
        }
    }

    $sb.ToString()
}

function New-SNowClassicTargetUrl {
    <#
      Builds the exact classic "target" URL format that typically prefills correctly:
      https://instance/now/nav/ui/classic/params/target/<table>.do%3Fsys_id%3D-1%26sysparm_query%3D...
    #>
    param(
        [Parameter(Mandatory)][string]$InstanceBase,
        [Parameter(Mandatory)][string]$Table,
        [Parameter(Mandatory)][hashtable]$Fields
    )

    $pairs = New-Object System.Collections.Generic.List[string]

    foreach ($key in $Fields.Keys) {
        $val = $Fields[$key]
        if ($null -eq $val) { continue }

        $s = [string]$val
        if ($s.Trim() -eq "") { continue }

        $encodedVal = ConvertTo-SNowQueryValue -Text $s
        [void]$pairs.Add("$key=$encodedVal")
    }

    $sysparmQuery = ($pairs -join "^")
    $target = "$Table.do%3Fsys_id%3D-1%26sysparm_query%3D$sysparmQuery"
    "$InstanceBase/now/nav/ui/classic/params/target/$target"
}

function Get-SNowSysIdFromText {
    param([Parameter(Mandatory)][string]$Text)

    # Works with sys_id=... or sys_id%3D...
    if ($Text -match '(?:sys_id=|sys_id%3D)([0-9a-fA-F]{32})') { return $matches[1] }
    return $null
}

# =========================
# Ticket / Task URL builders
# =========================
function New-DailyChecksWorkTicketUrl {
    param([Parameter(Mandatory)][string]$SiteInfo)

    $week = Get-WorkWeekDates
    $m = To-SNowDate $week.Monday
    $f = To-SNowDate $week.Friday
    $due = "$f $($Config.DueTime24)"

    $short = "$SiteInfo - Daily Checks Work Ticket $m to $f"
    $desc  = "$SiteInfo Daily Checks for Workweek: $m to $f"

    $fields = @{
        type              = "normal"
        impact            = 4
        assignment_group  = $Config.AssignmentGroupSysId
        assigned_to       = "javascript:gs.user_id()"
        short_description = $short
        description       = $desc
        due_date          = $due
        cmdb_ci           = $Config.AffectedCI
    }

    New-SNowClassicTargetUrl -InstanceBase $Config.InstanceBase -Table "u_work_ticket" -Fields $fields
}

function New-DailyChecksWorkTaskUrls {
    param(
        [Parameter(Mandatory)][string]$SiteInfo,
        [Parameter(Mandatory)][string]$WorkTicketSysId
    )

    $week = Get-WorkWeekDates

    # NOTE: Leave these URLs exactly as you had them. This encoding function preserves %2F etc.
    $taskDesc = @"
$SiteInfo Perform Daily Checks on the VistA Imaging System In Accordance With the SOP and Training documents

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
        $short = "$SiteInfo - $($d.Name) Daily Checks, Date: $dateString"

        # If your instance uses different internal field names, swap them here.
        $fields = @{
            type              = "normal"
            impact            = 4
            urgency           = 4
            assignment_group  = $Config.AssignmentGroupSysId
            assigned_to       = "javascript:gs.user_id()"

            u_requestor       = "javascript:gs.user_id()"
            u_affected_user   = "javascript:gs.user_id()"
            cmdb_ci           = $Config.AffectedCI

            short_description = $short
            description       = $taskDesc

            u_work_ticket     = $WorkTicketSysId
        }

        $urls.Add((New-SNowClassicTargetUrl -InstanceBase $Config.InstanceBase -Table "u_work_task" -Fields $fields)) | Out-Null
    }

    return $urls
}

function New-WeeklyWorkTaskUrl {
    param(
        [Parameter(Mandatory)][string]$SiteInfo,
        [Parameter(Mandatory)][string]$WorkTicketSysId
    )

    $week = Get-WorkWeekDates
    $m = To-SNowDate $week.Monday
    $f = To-SNowDate $week.Friday

    $short = "$SiteInfo - Weekly Daily Checks Task ($m to $f)"

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

    $fields = @{
        type              = "normal"
        impact            = 4
        urgency           = 4
        assignment_group  = $Config.AssignmentGroupSysId
        assigned_to       = "javascript:gs.user_id()"

        u_requestor       = "javascript:gs.user_id()"
        u_affected_user   = "javascript:gs.user_id()"
        cmdb_ci           = $Config.AffectedCI

        short_description = $short
        description       = $weeklyDesc

        u_work_ticket     = $WorkTicketSysId
    }

    New-SNowClassicTargetUrl -InstanceBase $Config.InstanceBase -Table "u_work_task" -Fields $fields
}

# =========================
# Load XAML / Wire Controls
# =========================
$XamlPath = Join-Path $ScriptRoot "SNOW_DailyChecks.xaml"
if (-not (Test-Path $XamlPath)) { throw "XAML not found: $XamlPath" }

$XamlMainWindow = Load-XmlXaml -Path $XamlPath
$Reader = New-Object System.Xml.XmlNodeReader $XamlMainWindow
$Form = [Windows.Markup.XamlReader]::Load($Reader)

# Auto-bind controls by Name to $WPF_<Name>
$XamlMainWindow.SelectNodes("//*[@Name]") | ForEach-Object {
    Set-Variable -Name ("WPF_{0}" -f $_.Name) -Value $Form.FindName($_.Name)
}

# Script-scoped state (no globals)
$script:SiteInfo = ""

# =========================
# UI Events
# =========================

# Button: Create Work Ticket (your XAML button name)
$WPF_XMLbtnSelectInstallFile.Add_Click({
    try {
        $site = (Get-TextSafe $WPF_XMLAddInstallFilePath.Text).Trim()
        if ([string]::IsNullOrWhiteSpace($site)) {
            [void][System.Windows.MessageBox]::Show("Enter Site/Site Code (ex: TVH/626).")
            return
        }

        $script:SiteInfo = $site

        # Open ServiceNow; if Terms page appears, you click Accept.
        Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-window", $Config.HomeUrl) | Out-Null
        Start-Sleep -Seconds 4

        $ticketUrl = New-DailyChecksWorkTicketUrl -SiteInfo $site
        Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-window", $ticketUrl) | Out-Null
    }
    catch {
        [void][System.Windows.MessageBox]::Show($_.Exception.Message)
    }
})

# Button: Create Work Tasks (paste ticket URL containing sys_id)
$WPF_StartButton.Add_Click({
    try {
        $site = (Get-TextSafe $script:SiteInfo).Trim()
        if ([string]::IsNullOrWhiteSpace($site)) {
            $site = (Get-TextSafe $WPF_XMLAddInstallFilePath.Text).Trim()
            if ([string]::IsNullOrWhiteSpace($site)) {
                [void][System.Windows.MessageBox]::Show("Enter Site/Site Code first (ex: TVH/626).")
                return
            }
            $script:SiteInfo = $site
        }

        $pasted = (Get-TextSafe $WPF_XMLEnterFQDNInfo.Text).Trim()
        if ([string]::IsNullOrWhiteSpace($pasted)) {
            [void][System.Windows.MessageBox]::Show("Paste the Work Ticket URL (must contain sys_id=...).")
            return
        }

        $ticketSysId = Get-SNowSysIdFromText -Text $pasted
        if (-not $ticketSysId) {
            [void][System.Windows.MessageBox]::Show("Could not find sys_id (32 chars) in what you pasted.")
            return
        }

        # Open ServiceNow (you handle Terms if it appears)
        Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-window", $Config.HomeUrl) | Out-Null
        Start-Sleep -Seconds 2

        # Open Mon-Fri tasks
        $taskUrls = New-DailyChecksWorkTaskUrls -SiteInfo $site -WorkTicketSysId $ticketSysId
        foreach ($u in $taskUrls) {
            Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-tab", $u) | Out-Null
            Start-Sleep -Milliseconds 350
        }

        # Open weekly task
        $weeklyUrl = New-WeeklyWorkTaskUrl -SiteInfo $site -WorkTicketSysId $ticketSysId
        Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-tab", $weeklyUrl) | Out-Null
    }
    catch {
        [void][System.Windows.MessageBox]::Show($_.Exception.Message)
    }
})

$WPF_Close.Add_Click({ $Form.Close() })

# =========================
# Run UI
# =========================
[void]$Form.ShowDialog()

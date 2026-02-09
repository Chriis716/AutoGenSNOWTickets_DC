Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# =========================
# Paths / Assemblies
# =========================
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Load local DLLs the way you already do (assembly folder next to script)
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

    # Terms gate + ServiceNow base
    TermsUrl     = "https://yourIT.VA.GOV/va?id=va_termsandconditions"
    HomeUrl      = "https://yourit.va.gov/home.do"
    InstanceBase = "https://yourit.va.gov"

    # Defaults you said must prefill
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

function Invoke-SNowPreflightTerms {
    # Open Terms page and pause so you can click ACCEPT once.
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

function ConvertTo-SNowQueryValue {
    <#
      This is the critical piece to match your "old URL behavior":
      - Preserve new lines for description: %0D%0A
      - Encode characters that break sysparm_query parsing
      NOTE: We do NOT fully EscapeDataString the entire sysparm_query;
            we encode values similarly to your existing working URLs.
    #>
    param([Parameter(Mandatory)][string]$Text)

    # Normalize + preserve line breaks
    $t = $Text -replace "`r`n", "`n"
    $t = $t -replace "`r", "`n"
    $t = $t -replace "`n", "%0D%0A"

    # Encode common breakers (order matters; protect % first)
    $t = $t.Replace("%", "%25")
    $t = $t.Replace("^", "%5E")
    $t = $t.Replace("&", "%26")
    $t = $t.Replace("=", "%3D")
    $t = $t.Replace("#", "%23")
    $t = $t.Replace("+", "%2B")
    $t = $t.Replace(" ", "%20")

    return $t
}

function New-SNowClassicTargetUrl {
    <#
      Builds the SAME style link your old script uses:
      https://instance/now/nav/ui/classic/params/target/<table>.do%3Fsys_id%3D-1%26sysparm_query%3D...
      This is what makes ServiceNow prefill reliably in your environment.
    #>
    param(
        [Parameter(Mandatory)][string]$Table,    # u_work_ticket or u_work_task
        [Parameter(Mandatory)][hashtable]$Fields
    )

    # Build key=value^key=value with values encoded "SNOW-style"
    $pairs = foreach ($k in $Fields.Keys) {
        $v = $Fields[$k]
        if ($null -eq $v) { continue }

        $s = [string]$v
        if ($s.Trim() -eq "") { continue }

        $vv = ConvertTo-SNowQueryValue -Text $s
        "$k=$vv"
    }

    $rawQuery = ($pairs -join "^")

    # Embed sysparm_query in the encoded target (THIS is the key)
    $targetEncoded = "$Table.do%3Fsys_id%3D-1%26sysparm_query%3D$rawQuery"
    return "$($Config.InstanceBase)/now/nav/ui/classic/params/target/$targetEncoded"
}

function Get-SNowSysIdFromText {
    param([Parameter(Mandatory)][string]$Text)

    # Works with sys_id=... or sys_id%3D...
    if ($Text -match '(?:sys_id=|sys_id%3D)([0-9a-fA-F]{32})') { return $matches[1] }
    return $null
}

# =========================
# Builders: Work Ticket + Work Tasks
# =========================

function New-DailyChecksWorkTicketUrl {
    param([Parameter(Mandatory)][string]$SiteInfo)

    $week = Get-WorkWeekDates
    $m = To-SNowDate $week.Monday
    $f = To-SNowDate $week.Friday
    $due = "$f $($Config.DueTime24)"

    $short = "$SiteInfo - Daily Checks Work Ticket $m to $f"
    $desc  = "$SiteInfo Daily Checks for Workweek: $m to $f"

    # Only fields you know are correct for ticket table
    New-SNowClassicTargetUrl -Table "u_work_ticket" -Fields @{
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

function New-DailyChecksWorkTaskUrls {
    param(
        [Parameter(Mandatory)][string]$SiteInfo,
        [Parameter(Mandatory)][string]$WorkTicketSysId
    )

    $week = Get-WorkWeekDates

    # This preserves new lines exactly like your old URL did (via %0D%0A)
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

        # IMPORTANT: these keys must match your instance table field names
        $urls.Add(
            (New-SNowClassicTargetUrl -Table "u_work_task" -Fields @{
                type              = "normal"
                impact            = 4
                urgency           = 4
                assignment_group  = $Config.AssignmentGroupSysId
                assigned_to       = "javascript:gs.user_id()"

                # These were blank in my previous version due to URL shape; fixed now.
                u_requestor       = "javascript:gs.user_id()"
                u_affected_user   = "javascript:gs.user_id()"
                cmdb_ci           = $Config.AffectedCI

                short_description = $short
                description       = $taskDesc

                # Link the task to the ticket
                u_work_ticket     = $WorkTicketSysId
            })
        )
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

    New-SNowClassicTargetUrl -Table "u_work_task" -Fields @{
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
}

# =========================
# Load XAML / Wire Controls
# =========================
$XamlPath = Join-Path $ScriptRoot "SNOW_DailyChecks.xaml"
if (-not (Test-Path $XamlPath)) { throw "XAML not found: $XamlPath" }

$XamlMainWindow = Load-XmlXaml -Path $XamlPath
$Reader = New-Object System.Xml.XmlNodeReader $XamlMainWindow
$Form = [Windows.Markup.XamlReader]::Load($Reader)

# Auto-bind controls: $WPF_<Name>
$XamlMainWindow.SelectNodes("//*[@Name]") | ForEach-Object {
    Set-Variable -Name ("WPF_{0}" -f $_.Name) -Value $Form.FindName($_.Name)
}

# Keep site in script scope (cleaner than Global)
$script:SiteInfo = ""

# =========================
# UI Events
# =========================

# Button 1: Create Work Ticket
$WPF_XMLbtnSelectInstallFile.Add_Click({
    try {
        $site = (Get-TextSafe $WPF_XMLAddInstallFilePath.Text).Trim()
        if ([string]::IsNullOrWhiteSpace($site)) {
            [void][System.Windows.MessageBox]::Show("Enter Site/Site Code (ex: TVH/626).")
            return
        }

        $script:SiteInfo = $site

        # Preflight terms gate (SSO)
        Invoke-SNowPreflightTerms

        # Open home then the new ticket prefilled
        Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-window", $Config.HomeUrl) | Out-Null
        Start-Sleep -Seconds 4

        $ticketUrl = New-DailyChecksWorkTicketUrl -SiteInfo $site
        Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-window", $ticketUrl) | Out-Null
    }
    catch {
        [void][System.Windows.MessageBox]::Show($_.Exception.Message)
    }
})

# Button 2: Create Work Tasks (paste ticket URL containing sys_id)
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

        Invoke-SNowPreflightTerms

        $taskUrls = New-DailyChecksWorkTaskUrls -SiteInfo $site -WorkTicketSysId $ticketSysId
        foreach ($u in $taskUrls) {
            Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-tab", $u) | Out-Null
            Start-Sleep -Milliseconds 350
        }

        $weeklyUrl = New-WeeklyWorkTaskUrl -SiteInfo $site -WorkTicketSysId $ticketSysId
        Start-Process -FilePath $Config.BrowserExe -ArgumentList @("--new-tab", $weeklyUrl) | Out-Null
    }
    catch {
        [void][System.Windows.MessageBox]::Show($_.Exception.Message)
    }
})

$WPF_Close.Add_Click({ $Form.Close() })

# Show UI
[void]$Form.ShowDialog()

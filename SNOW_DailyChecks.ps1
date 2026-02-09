##Initialize######
[System.Void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')                                                     
[System.Void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')                                                     
[System.Void][System.Reflection.Assembly]::LoadFrom('assembly\MahApps.Metro.dll')                                                             
[System.Void][System.Reflection.Assembly]::LoadFrom('assembly\System.Windows.Interactivity.dll')
[System.Void][System.Reflection.Assembly]::LoadFrom('assembly\LoadingIndicators.WPF.dll') 

function LoadXml ($global:filename)
{
    $XamlLoader=(New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}

# Load MainWindow
$XamlMainWindow=LoadXml("SNOW_DailyChecks.xaml")
$Reader=(New-Object System.Xml.XmlNodeReader $XamlMainWindow)
$Form=[Windows.Markup.XamlReader]::Load($Reader)


$XamlMainWindow.SelectNodes("//*[@Name]") | %{
    try {Set-Variable -Name "$("WPF_"+$_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
    }

Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable *WPF*
}
#Get-FormVariables

# Get today's date
$today = Get-Date
$dayOfWeek = $today.DayOfWeek

# Calculate the date of Monday this week
$daysSinceMonday = $today.DayOfWeek.value__ - [System.DayOfWeek]::Monday
$mondayThisWeek = $today.AddDays(-$daysSinceMonday).ToString('MM/dd/yyyy')

# Calculate the date of Tuesday this week
$daysUntilTuesday = [System.DayOfWeek]::Tuesday - $today.DayOfWeek.value__
$tuesdayThisWeek = $today.AddDays($daysUntilTuesday).ToString('MM/dd/yyyy')

# Calculate the date of Wednesday this week
$daysUntilWednesday = [System.DayOfWeek]::Wednesday - $today.DayOfWeek.value__
$wednesdayThisWeek = $today.AddDays($daysUntilWednesday).ToString('MM/dd/yyyy')

# Calculate the date of Thursday this week
$daysUntilThursday = [System.DayOfWeek]::Thursday - $today.DayOfWeek.value__
$thursdayThisWeek = $today.AddDays($daysUntilThursday).ToString('MM/dd/yyyy')

# Calculate the date of Friday this week
$daysUntilFriday = [System.DayOfWeek]::Friday - $today.DayOfWeek.value__
$fridayThisWeek = $today.AddDays($daysUntilFriday).ToString('MM/dd/yyyy')

# Check if today is past Friday and adjust for next week
if ($today.DayOfWeek.value__ -gt [System.DayOfWeek]::Friday) {
    $mondayThisWeek = $today.AddDays(-$daysSinceMonday + 7).ToString('MM/dd/yyyy')
    $fridayThisWeek = $today.AddDays($daysUntilFriday + 7).ToString('MM/dd/yyyy')
}

$value="4e965a49db5dd3c0b857fd721f9619d5" # SPM.HEALTH.HISM.CLINICAL IMAGING-Team3
$assignmentGroup = $value
$AffectedCI = "bbf1bbdf1b88ec5006020f6fe54bcb68"
$notAvailable = "N/A"
$duedate = "$fridayThisWeek 14:30:00"
$encodedAssignmentGroup = [System.Uri]::EscapeDataString($assignmentGroup)
$encodedCategoryGroup = [System.Uri]::EscapeDataString($category)
$encodedContactGroup = [System.Uri]::EscapeDataString($ContactType)
$encodedAffectedSystemGroup = [System.Uri]::EscapeDataString($SystemName)
$encodedduedate = [System.Uri]::EscapeDataString($duedate)

Function Create-WorkTickets {
    param (
        [Parameter(Mandatory)]
        [string]
        $Site
    )

$Global:SiteInfo = "$Site" #Site/Site Code example TVH/626
$shortDescription = "$SiteInfo - Daily Checks Work Ticket $mondayThisWeek to $fridayThisWeek"
$detailedDescription = "$SiteInfo Daily Checks for Workweek: $mondayThisWeek to $fridayThisWeek"

# URL-encode the variable values 
$encodedShortDescription = [System.Uri]::EscapeDataString($shortDescription) 
$encodedDetailedDescription = [System.Uri]::EscapeDataString($detailedDescription) 

#Workweek decode
$encodedWorkTaskShortDescription = [System.Uri]::EscapeDataString($myOutputDailyChecks) 


# Create Incident for IaaMS Migration
$OpenSNOWsite = "https://yourit.va.gov/home.do"
$url = "https://yourit.va.gov/now/nav/ui/classic/params/target/u_work_ticket.do%3Fsys_id%3D-1&sysparm_query=type=normal^impact=4^assignment_group=$encodedAssignmentGroup^assigned_to=javascript:gs.user_id()^short_description=$encodedShortDescription^description=$encodedDetailedDescription^due_date=$encodedduedate^cmdb_ci=$AffectedCI"

Write-Host "Opening ServiceNow Webpage, please wait..."
Start-Process -FilePath "msedge.exe" -ArgumentList $OpenSNOWsite
Start-Sleep -Second 10
Write-Host "Creating Work Ticket, please wait..."
Start-Process -FilePath "msedge.exe" -ArgumentList "--new-window", $url

}#end Function InstallApplication


Function Create-WorkTasks {
    param (
        [Parameter(Mandatory)]
        [string]
        $WorkTicketID
    )

$workTicketInfo = $WorkTicketID

$MondayOutputDailyChecks = "$Global:SiteInfo - Monday Daily Checks, Date: $mondayThisWeek" 
$TuesdayOutputDailyChecks = "$Global:SiteInfo - Tuesday Daily Checks, Date: $tuesdayThisWeek" 
$WednesdayOutputDailyChecks = "$Global:SiteInfo - Wednesday Daily Checks, Date: $wednesdayThisWeek" 
$ThursdayOutputDailyChecks = "$Global:SiteInfo - Thursday Daily Checks, Date: $thursdayThisWeek" 
$FridayOutputDailyChecks = "$Global:SiteInfo - Friday Daily Checks, Date: $fridayThisWeek"

$workTaskdetailedDescription = "$Global:SiteInfo Perform Daily Checks on the VistA Imaging System In Accordance With the SOP and Training documents

Perform Daily Checks on the VistA Imaging System In Accordance With the SOP and Training documents

Training Modules:
https://dvagov.sharepoint.com/sites/oitspmhsphismclinternal/Training/Forms/AllItems.aspx?id=%2Fsites%2Foitspmhsphismclinternal%2FTraining%2FDaily%20Checks&viewid=6a448197%2Ddc45%2D4f14%2Da02d%2Db306278fe303

SPM-HISM-CL Daily Checks SOP:
https://dvagov.sharepoint.com/sites/oitspmhsphismclinternal/SitePages/Clinical-Imaging-SOP-Guidance.aspx                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
"

$encodedWorkTaskDetailedDescription = [System.Uri]::EscapeDataString($workTaskdetailedDescription)

Write-Host "Creating Work Task, please wait..."
$encodedWorkTicketInfo = [System.Uri]::EscapeDataString($workTicketInfo)

# Monday Work Task
$encodedMonWorkTaskShortDescription = [System.Uri]::EscapeDataString($MondayOutputDailyChecks)
$workTaskUrl1 = "https://yourit.va.gov/u_work_task.do?sys_id=-1&sysparm_query=type=normal^impact=4^urgency=4^assignment_group=$encodedAssignmentGroup^assigned_to=javascript:gs.user_id()^short_description=$encodedMonWorkTaskShortDescription^description=$encodedWorkTaskDetailedDescription^cmdb_ci=$AffectedCI^u_requestor=javascript:gs.user_id()^u_affected_user=javascript:gs.user_id()&sys_is_list=true&sys_is_related_list=true&sys_target=u_work_task&sysparm_checked_items=&sysparm_collection=u_work_ticket&sysparm_collectionID=$encodedWorkTicketInfo&sysparm_collection_key=u_work_ticket&sysparm_collection_label=Work+Tasks&sysparm_collection_related_field=&sysparm_collection_related_file=&sysparm_collection_related_relationship=u_work_task.u_work_ticket&sysparm_collection_relationship=&sysparm_fixed_query=&sysparm_group_sort=&sysparm_list_css=&sysparm_query=&sysparm_referring_url=u_work_ticket.do%3fsys_id%3d$encodedWorkTicketInfo%4099%40sysparm_record_rows%3d1329473%4099%40sysparm_record_target%3dtask%4099%40sysparm_record_list%3dactive%253Dtrue%255EAgetMyAssignments%2528%2529%255EORDERBYDESCsys_created_on%4099%40sysparm_record_row%3d6&sysparm_target=&sysparm_view=#"

Write-Host "Creating Work Task, $MondayOutputDailyChecks, please wait..."
Start-Process -FilePath "msedge.exe" -ArgumentList "--new-tab", $workTaskUrl1
$workTaskUrl1 = ""
Start-Sleep -Second 1

# Tuesday Work Task
$encodedTueWorkTaskShortDescription = [System.Uri]::EscapeDataString($TuesdayOutputDailyChecks)
$workTaskUrl2 = "https://yourit.va.gov/u_work_task.do?sys_id=-1&sysparm_query=type=normal^impact=4^urgency=4^assignment_group=$encodedAssignmentGroup^assigned_to=javascript:gs.user_id()^short_description=$encodedTueWorkTaskShortDescription^description=$encodedWorkTaskDetailedDescription^cmdb_ci=$AffectedCI^u_requestor=javascript:gs.user_id()^u_affected_user=javascript:gs.user_id()&sys_is_list=true&sys_is_related_list=true&sys_target=u_work_task&sysparm_checked_items=&sysparm_collection=u_work_ticket&sysparm_collectionID=$encodedWorkTicketInfo&sysparm_collection_key=u_work_ticket&sysparm_collection_label=Work+Tasks&sysparm_collection_related_field=&sysparm_collection_related_file=&sysparm_collection_related_relationship=u_work_task.u_work_ticket&sysparm_collection_relationship=&sysparm_fixed_query=&sysparm_group_sort=&sysparm_list_css=&sysparm_query=&sysparm_referring_url=u_work_ticket.do%3fsys_id%3d$encodedWorkTicketInfo%4099%40sysparm_record_rows%3d1329473%4099%40sysparm_record_target%3dtask%4099%40sysparm_record_list%3dactive%253Dtrue%255EAgetMyAssignments%2528%2529%255EORDERBYDESCsys_created_on%4099%40sysparm_record_row%3d6&sysparm_target=&sysparm_view=#"
Write-Host "Creating Work Task, $TuesdayOutputDailyChecks, please wait..."
Start-Process -FilePath "msedge.exe" -ArgumentList "--new-tab", $workTaskUrl2
Start-Sleep -Second 1

# Wednesday Work Task
$encodedWedWorkTaskShortDescription = [System.Uri]::EscapeDataString($WednesdayOutputDailyChecks)
$workTaskUrl3 = "https://yourit.va.gov/u_work_task.do?sys_id=-1&sysparm_query=type=normal^impact=4^urgency=4^assignment_group=$encodedAssignmentGroup^assigned_to=javascript:gs.user_id()^short_description=$encodedWedWorkTaskShortDescription^description=$encodedWorkTaskDetailedDescription^cmdb_ci=$AffectedCI^u_requestor=javascript:gs.user_id()^u_affected_user=javascript:gs.user_id()&sys_is_list=true&sys_is_related_list=true&sys_target=u_work_task&sysparm_checked_items=&sysparm_collection=u_work_ticket&sysparm_collectionID=$encodedWorkTicketInfo&sysparm_collection_key=u_work_ticket&sysparm_collection_label=Work+Tasks&sysparm_collection_related_field=&sysparm_collection_related_file=&sysparm_collection_related_relationship=u_work_task.u_work_ticket&sysparm_collection_relationship=&sysparm_fixed_query=&sysparm_group_sort=&sysparm_list_css=&sysparm_query=&sysparm_referring_url=u_work_ticket.do%3fsys_id%3d$encodedWorkTicketInfo%4099%40sysparm_record_rows%3d1329473%4099%40sysparm_record_target%3dtask%4099%40sysparm_record_list%3dactive%253Dtrue%255EAgetMyAssignments%2528%2529%255EORDERBYDESCsys_created_on%4099%40sysparm_record_row%3d6&sysparm_target=&sysparm_view=#"
Write-Host "Creating Work Task, $WednesdayOutputDailyChecks, please wait..."
Start-Process -FilePath "msedge.exe" -ArgumentList "--new-tab", $workTaskUrl3
Start-Sleep -Second 1

# Thursday Work Task
$encodedThurWorkTaskShortDescription = [System.Uri]::EscapeDataString($ThursdayOutputDailyChecks)
$workTaskUrl4 = "https://yourit.va.gov/u_work_task.do?sys_id=-1&sysparm_query=type=normal^impact=4^urgency=4^assignment_group=$encodedAssignmentGroup^assigned_to=javascript:gs.user_id()^short_description=$encodedThurWorkTaskShortDescription^description=$encodedWorkTaskDetailedDescription^cmdb_ci=$AffectedCI^u_requestor=javascript:gs.user_id()^u_affected_user=javascript:gs.user_id()&sys_is_list=true&sys_is_related_list=true&sys_target=u_work_task&sysparm_checked_items=&sysparm_collection=u_work_ticket&sysparm_collectionID=$encodedWorkTicketInfo&sysparm_collection_key=u_work_ticket&sysparm_collection_label=Work+Tasks&sysparm_collection_related_field=&sysparm_collection_related_file=&sysparm_collection_related_relationship=u_work_task.u_work_ticket&sysparm_collection_relationship=&sysparm_fixed_query=&sysparm_group_sort=&sysparm_list_css=&sysparm_query=&sysparm_referring_url=u_work_ticket.do%3fsys_id%3d$encodedWorkTicketInfo%4099%40sysparm_record_rows%3d1329473%4099%40sysparm_record_target%3dtask%4099%40sysparm_record_list%3dactive%253Dtrue%255EAgetMyAssignments%2528%2529%255EORDERBYDESCsys_created_on%4099%40sysparm_record_row%3d6&sysparm_target=&sysparm_view=#"
Write-Host "Creating Work Task, $ThursdayOutputDailyChecks, please wait..."
Start-Process -FilePath "msedge.exe" -ArgumentList "--new-tab", $workTaskUrl4
Start-Sleep -Second 1

# Friday Work Task
$encodedFriWorkTaskShortDescription = [System.Uri]::EscapeDataString($FridayOutputDailyChecks)
$workTaskUrl5 = "https://yourit.va.gov/u_work_task.do?sys_id=-1&sysparm_query=type=normal^impact=4^urgency=4^assignment_group=$encodedAssignmentGroup^assigned_to=javascript:gs.user_id()^short_description=$encodedFriWorkTaskShortDescription^description=$encodedWorkTaskDetailedDescription^cmdb_ci=$AffectedCI^u_requestor=javascript:gs.user_id()^u_affected_user=javascript:gs.user_id()&sys_is_list=true&sys_is_related_list=true&sys_target=u_work_task&sysparm_checked_items=&sysparm_collection=u_work_ticket&sysparm_collectionID=$encodedWorkTicketInfo&sysparm_collection_key=u_work_ticket&sysparm_collection_label=Work+Tasks&sysparm_collection_related_field=&sysparm_collection_related_file=&sysparm_collection_related_relationship=u_work_task.u_work_ticket&sysparm_collection_relationship=&sysparm_fixed_query=&sysparm_group_sort=&sysparm_list_css=&sysparm_query=&sysparm_referring_url=u_work_ticket.do%3fsys_id%3d$encodedWorkTicketInfo%4099%40sysparm_record_rows%3d1329473%4099%40sysparm_record_target%3dtask%4099%40sysparm_record_list%3dactive%253Dtrue%255EAgetMyAssignments%2528%2529%255EORDERBYDESCsys_created_on%4099%40sysparm_record_row%3d6&sysparm_target=&sysparm_view=#"
Write-Host "Creating Work Task, $FridayOutputDailyChecks, please wait..."
Start-Process -FilePath "msedge.exe" -ArgumentList "--new-tab", $workTaskUrl5
Start-Sleep -Second 1


}#end Function InstallApplication

Function Create-WeeklyWorkTasks {
    param (
        [Parameter(Mandatory)]
        [string]
        $WeeklyWorkTicketID
    )

	$workTicketInfo = $WeeklyWorkTicketID
	$encodedWorkTicketInfo = [System.Uri]::EscapeDataString($workTicketInfo)
	$WeeklyshortDescription = "$Global:SiteInfo - Weekly Task for Workweek $mondayThisWeek to $fridayThisWeek"

	$WeeklyDetailedDescription = "Weekly Work Task Ticket for  cleaning up Async
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

	"
	$encodedWeeklyShortDescription = [System.Uri]::EscapeDataString($WeeklyshortDescription) 
	$encodedWeeklyDetailedDescription = [System.Uri]::EscapeDataString($WeeklyDetailedDescription)

	$WeeklyWorkTaskUrl = "https://yourit.va.gov/u_work_task.do?sys_id=-1&sysparm_query=type=normal^impact=4^urgency=4^assignment_group=$encodedAssignmentGroup^assigned_to=javascript:gs.user_id()^short_description=$encodedWeeklyShortDescription^description=$encodedWeeklyDetailedDescription^cmdb_ci=$AffectedCI^u_requestor=javascript:gs.user_id()^u_affected_user=javascript:gs.user_id()&sys_is_list=true&sys_is_related_list=true&sys_target=u_work_task&sysparm_checked_items=&sysparm_collection=u_work_ticket&sysparm_collectionID=$encodedWorkTicketInfo&sysparm_collection_key=u_work_ticket&sysparm_collection_label=Work+Tasks&sysparm_collection_related_field=&sysparm_collection_related_file=&sysparm_collection_related_relationship=u_work_task.u_work_ticket&sysparm_collection_relationship=&sysparm_fixed_query=&sysparm_group_sort=&sysparm_list_css=&sysparm_query=&sysparm_referring_url=u_work_ticket.do%3fsys_id%3d$encodedWorkTicketInfo"
	Write-Host "Creating Weekly Task, please wait..."
	Start-Process -FilePath "msedge.exe" -ArgumentList "--new-tab", $WeeklyWorkTaskUrl



}#end Function InstallApplication

$WPF_XMLbtnSelectInstallFile.Add_Click({
                
                $SiteInfoName = $WPF_XMLAddInstallFilePath.text                          
                Create-WorkTickets -Site $SiteInfoName
                write-host "Script Completed"
})

$WPF_StartButton.Add_Click({
                
                $url = $WPF_XMLEnterFQDNInfo.text
                
                $pattern = 'u_work_ticket\.do%3Fsys_id%3D([a-zA-Z0-9]{32})%26sysparm_record_target'
                Write-Host "Creating Work Task, please wait..."

                if ($url -match $pattern) {
    $id = $matches[1]
    Write-Output "ID: $id"
                } else {
                                Write-Output "No match found."
                }
                write-host $id
                
                #Call Function
                Create-WorkTasks -WorkTicketID $id
                Create-WeeklyWorkTasks -WeeklyWorkTicketID $id
    
                write-host "Script Completed"
    #$Form.Close()
                
})
$WPF_Close.add_Click({

   $Form.Close()

})

# Force garbage collection just to start slightly lower RAM usage.
[System.GC]::Collect()

$Form.ShowDialog() | Out-Null

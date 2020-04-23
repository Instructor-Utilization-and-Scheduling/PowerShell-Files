$OutlookFormClose = {
    $FormOutlook.Close()
}
$GenerateAppts = {
    $ButtonOutlookCreate.Text = "Working..."
    $ButtonOutlookCreate.Enabled = $false
    $LabelOutlookNumEvents.Visible = $true
    $search = @($Script:FilteredEvents | 
        Select-Object -Property "course", "class" -Unique | 
            ForEach-Object {"$($_.course) / $($_.class)"})
    $calendar = $Script:AllCalendars | 
        Where-Object {$_.fullfolderpath -eq $ComboBoxOutlookCal.SelectedItem.fullfolderpath}
    $existingappts = @(Get-OutlookAppointments -OutlookCalendar $calendar)
    $existingappts = $existingappts | 
        Where-Object {(($_.subject -split " / ")[0..1] -join " / ") -in $search}
    if ($existingappts.count -ge 1) {
        $msg = "There are $($existingappts.count) appointments for this class already. Do you want to delete the old appointments before adding the new ones?"
        $caption = "Delete Existing Appts?"
        $result = [System.Windows.Forms.MessageBox]::Show($msg, $caption, 4, 32)
        $delevents = 0
        $curpass = -1
        if ($result -eq 6) {
            $LabelOutlookNumEvents.text = "Deleting existing events."
            while ($curpass -ne 0){
                $curpass = $search | 
                    ForEach-Object {"{0}*" -f $_} |
                        Remove-OutlookAppointment -OutlookCalendar $calendar
                $delevents += $curpass
            } # while
            $msg = "Deleted {0} Outlook appointments in {1}" -f $delevents, $calendar.name
            $caption = "Events Deleted"
            $result = [System.Windows.Forms.MessageBox]::Show($msg, $caption, 0, 64)
        } # if delete
    } # existing appts
    $eventnum = 1
    $CalAppts = $Script:FilteredEvents | 
        Group-Object -Property "Course", "Class", "Lesson", "room", "start", "end", "AsOf" |
        ForEach-Object {
            $LabelOutlookNumEvents.Text = "Generating event: $eventnum"
            $course, $class, $Lesson, $room, [datetime]$start, [datetime]$end, [datetime]$asof = $_.name -split ", "
            $Primary = @($_.group | Where-Object {$_.Role -eq "Primary"} | Select-Object -ExpandProperty "Instructor") -join ", "
            $Secondary = @($_.group | Where-Object {$_.Role -eq "Secondary"} | Select-Object -ExpandProperty "Instructor") -join ", "
            $SecondarySupport = @($_.group | Where-Object {$_.Role -eq "Secondary/Support"} | Select-Object -ExpandProperty "Instructor") -join ", "
            $Support = @($_.group | Where-Object {$_.Role -eq "Support"} | Select-Object -ExpandProperty "Instructor") -join ", "
            [pscustomobject]@{
                CalendarFolder = $calendar.fullfolderpath
                Start          = $start
                End            = $end
                Subject        = "{0} / {1} / {2} / [{3}]" -f $course, $class, $lesson, $AsOf.ToString("d")
                Location       = $room
                Category       = "Instruction"
                Body           = "Primary: {0}`nSecondary: {1}`nSecondary/Support: {2}`nSupport: {3}" -f $Primary, $Secondary, $SecondarySupport, $Support
            }
            $eventnum ++
        } # foreach-object
        $LabelOutlookNumEvents.Text = "Publishing to Outlook..."
        $eventspub = $CalAppts | 
            New-OutlookEvent -CalendarFolder $calendar.fullfolderpath 
        $ButtonOutlookCreate.Text = "Create Outlook Appointments" 
        $msg = "Published {0} Outlook appointments to {1}" -f $eventspub, $calendar.name
        $result = [System.Windows.Forms.MessageBox]::Show($msg, $caption, 0, 64)
        $ButtonOutlookCreate.Enabled = $true

    $FormOutlook.close()
}
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'OutlookForm.designer.ps1')

#Load data
$script:AllCalendars = Get-OutlookCalendars
$AllCalendars |
    Select-Object -Property fullfolderpath, name |
        ForEach-Object {
            $ComboBoxOutlookCal.items.add($_)
        }
$TotalEvents = $script:FilteredEvents | 
    Group-Object -Property "Course", "Class", "Lesson", "room", "start", "end", "AsOf" |
        Measure-Object | Select-Object -ExpandProperty count
$lbl = "Total Number of Appointments: {0}" -f $TotalEvents
$LabelOutlookNumEvents.text = $lbl
$FormOutlook.ShowDialog()
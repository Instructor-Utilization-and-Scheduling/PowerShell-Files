$OutlookFormClose = {
    $FormOutlook.Close()
}
$GenerateAppts = {
    $ButtonOutlookCreate.Text = "Working..."
    $ButtonOutlookCreate.Enabled = $false
    $LabelOutlookNumEvents.Text = "Analyzing {0}" -f $ComboBoxOutlookCal.SelectedItem.name
    $search = @($Script:FilteredEvents | 
        Select-Object -Property "course", "class" -Unique | 
            ForEach-Object {"$($_.course) / $($_.class)"})
    
    $calendar = $Script:AllCalendars | 
        Where-Object {$_.fullfolderpath -eq $ComboBoxOutlookCal.SelectedItem.fullfolderpath}
    $existingappts = @(Get-OutlookAppointments -OutlookCalendar $calendar)
    $deleteappts = @($existingappts | 
        Where-Object {(($_.subject -split " / ")[0..1] -join " / ") -in $search})
    if ($deleteappts.count -ge 1) {
        $msg = "There are $($deleteappts.count) appointments for this class already. Do you want to delete the old appointments before adding the new ones?"
        $caption = "Delete Existing Appts?"
        $result = [System.Windows.Forms.MessageBox]::Show($msg, $caption, 4, 32)
        $delevents = 0
        $curpass = -1
        if ($result -eq 6) {
            $deleteappts |
                ForEach-Object {
                    $_.delete()
                    $LabelOutlookNumEvents.text = "Deleting $delevents"
                    $delevents++
                } #foreach-object                
        } # if delete
    } # existing appts
    if ($script:SingleInstr) {
        
    }
    $eventnum = 1
    $CalAppts = $Events |
        ForEach-Object {
            $LabelOutlookNumEvents.Text = "Generating event: $eventnum"
            $course, $class, $Lesson, $room, [datetime]$start, [datetime]$end, [datetime]$asof = $_.name -split ", "
            $Primary = @($_.group | Where-Object {$_.Role -eq "Primary"} | Select-Object -ExpandProperty "Instructor") -join ", "
            $Secondary = @($_.group | Where-Object {$_.Role -eq "Secondary"} | Select-Object -ExpandProperty "Instructor") -join ", "
            $SecondarySupport = @($_.group | Where-Object {$_.Role -eq "Secondary/Support"} | Select-Object -ExpandProperty "Instructor") -join ", "
            $Support = @($_.group | Where-Object {$_.Role -eq "Support"} | Select-Object -ExpandProperty "Instructor") -join ", "
            if ($script:SingleInst.count -eq 1) {
                switch ("*$($script:SingleInst.Instructor)*") {
                    {$Primary -like $_}          {$Cat = "Primary";break }
                    {$Secondary -like $_}        {$Cat = "Secondary";break }
                    {$SecondarySupport -like $_} {$Cat = "Secondary/Support";break }
                    {$Support -like $_}          {$Cat = "Support";break }
                }
            }
            else {$Cat = "Instruction"}
            [pscustomobject]@{
                CalendarFolder = $calendar.fullfolderpath
                Start          = $start
                End            = $end
                Subject        = "{0} / {1} / {2} / [{3}]" -f $course, $class, $lesson, $AsOf.ToString("d")
                Location       = $room
                Category       = $Cat
                Body           = "Primary: {0}`nSecondary: {1}`nSecondary/Support: {2}`nSupport: {3}" -f $Primary, $Secondary, $SecondarySupport, $Support
            }
            $eventnum ++
        } # foreach-object
        $LabelOutlookNumEvents.Text = "Publishing to Outlook..."
        $eventspub = $CalAppts | 
            New-OutlookEvent -CalendarFolder $calendar.fullfolderpath 
        $ButtonOutlookCreate.Text = "Create Outlook Appointments" 
        $msg = "Published {0} Outlook appointments to {1}" -f $eventspub, $calendar.name
        $LabelOutlookNumEvents.Text = $msg
        $result = [System.Windows.Forms.MessageBox]::Show($msg, $caption, 0, 64)
        $ButtonOutlookCreate.Enabled = $true

    $FormOutlook.close()
}
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'OutlookForm.designer.ps1')

#Load data
$script:AllCalendars = Get-OutlookCalendars
$script:AllCalendars |
    Select-Object -Property fullfolderpath, name |
        ForEach-Object {
            $ComboBoxOutlookCal.items.add($_)
        }

$group = @($script:FilteredEvents | 
    Group-Object "Course", "Class", "Lesson", "room", "start", "end", "AsOf")
$script:Events = @($script:AllEvents | 
Group-Object -Property "Course", "Class", "Lesson", "room", "start", "end", "AsOf" |
    Where-Object {$_.Name -in $group.name})
$script:SingleInst = @($script:FilteredEvents | Sort-Object -Property Instructor -Unique)

$lbl = "Total Number of Appointments: {0}" -f $Events.count
$LabelOutlookNumEvents.text = $lbl
$FormOutlook.ShowDialog()
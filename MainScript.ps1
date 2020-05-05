# Including what will be our module....
. (Join-Path $PSScriptRoot 'InstructorUtilizationModule.ps1')

Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'MainScript.designer.ps1')

$ViewEventGrid = {
    $UpdateFiltered.Invoke()    
    $script:FilteredEvents | Out-GridView
}
$DeleteSched = {
    $ButtonRemoveClassSched.Text = "Working..."
    $ButtonRemoveClassSched.Enabled = $false
    $course = $DataGridViewClassesLoaded.SelectedCells[0].value
    $class  = $DataGridViewClassesLoaded.SelectedCells[1].Value
    $asof   = $DataGridViewClassesLoaded.SelectedCells[2].Value
    $msg      = "Are you sure you want to delete {0}" -f (@($course, $class, $asof) -join ", ")
    $caption  = "Are you sure"
    $buttons  = [System.Windows.Forms.MessageBoxButtons]::YesNoCancel
    $icon     = [System.Windows.Forms.MessageBoxIcon]::Question
    $selection   = [System.Windows.Forms.MessageBox]::Show($msg,$caption,$buttons,$icon)
    if ($selection -eq [System.Windows.Forms.DialogResult]::Yes) {
        $Script:AllEvents = $Script:AllEvents |
                                Where-Object {!($_.course -eq $course -and $_.class -eq $class)}
        $Script:AllEvents | Export-Csv -Path (Join-Path -Path $Config 'events.csv')
        $MainDataLoad.Invoke()
    }
    $ButtonRemoveClassSched.Text = "Delete Class Schedule"
    $ButtonRemoveClassSched.Enabled = $true
}

$QuarterlyReport = {
    $ButtonQuarterlyReport.Text = "Working...."
    $ButtonQuarterlyReport.Enabled = $false
    $UpdateFiltered.Invoke()
    $TempFile = Join-Path -Path $env:TEMP -ChildPath ("Utilization_Report_" + (Get-Date).ToString("dfff") + ".txt")
    $ht = @{
        InstructorEvents = $script:FilteredEvents
        Grouping         = "Quarterly"
        InstructorsAvailable = $NumericUpDownInstAvail.Value
    }
    Measure-Events @ht |
        Out-File -FilePath $TempFile
    Invoke-Item -Path $TempFile
    $ButtonQuarterlyReport.Text = "Quarterly Rollup"
    $ButtonQuarterlyReport.Enabled = $true
}
$MonthlyReport = {
    $ButtonMonthlyReport.Text = "Working...."
    $ButtonMonthlyReport.Enabled = $false
    $UpdateFiltered.Invoke()
    $TempFile = Join-Path -Path $env:TEMP -ChildPath ("Utilization_Report_" + (Get-Date).ToString("dfff") + ".txt")
    $ht = @{
        InstructorEvents = $script:FilteredEvents
        Grouping         = "Monthly"
        InstructorsAvailable = $NumericUpDownInstAvail.Value
    }
    Measure-Events @ht |
        Out-File -FilePath $TempFile
    Invoke-Item -Path $TempFile
    $ButtonMonthlyReport.Text = "Monthly Rollup"
    $ButtonMonthlyReport.Enabled = $true
}
$ImportSched = {
    #open new form. 
    . (Join-Path $PSScriptRoot 'NewSched.ps1')
}
$UpdateFiltered = {
    $SelectedInstructors = @(foreach ($row in $DataGridViewInstructors.SelectedRows) {
        $row.cells[0].value
    })
    if ($ComboBoxClassFilter.SelectedItem -and $ComboBoxCourseFilter.SelectedItem) {            
        [InstructorEvent[]]$script:FilteredEvents = @($AllEvents | 
            Where-Object {$_.Instructor -in $SelectedInstructors -and $_.start -ge $DateTimePickerStartFilter.Value -and $_.end -le $DateTimePickerEndFilter.Value -and $_.Course -like $ComboBoxCourseFilter.SelectedItem.ToString() -and $_.Class -like $ComboBoxClassFilter.SelectedItem.ToString()})
    } # if
} # UpdateFiltered
$OpenOutlookForm = {
    $defaultText = $ButtonOutlookSched.text
    $ButtonOutlookSched.text = "Working..."
    $ButtonOutlookSched.Enabled = $false
    $UpdateFiltered.Invoke()
    . (Join-Path $PSScriptRoot 'OutlookForm.ps1')
    $ButtonOutlookSched.text = $defaultText
    $ButtonOutlookSched.Enabled = $true
}

$InstrUtilizationLoaded = {
    $ComboBoxCourseFilter.SelectedItem = "*"
    $ComboBoxClassFilter.SelectedItem  = "*"
    $DataGridViewInstructors.SelectAll()    
} # InstrUtilizationLoaded

# Form Layout script
. (Join-Path $PSScriptRoot 'MainScript.designer.ps1')

# Getting Configuration Information
$Config = Get-Content -Path (Join-Path -Path $PSScriptRoot 'config.cfg')
$requiredfiles = "events.csv","NameAliases.csv","whitelist.csv"
$requiredfiles | 
    ForEach-Object {
        if (!(Test-Path -Path "$Config\$_")) {
            $msg = "Can't find required config files in $Config"
            $msg += "`nUpdate config.cfg with the appropriate path and ensure the following files are present:`n"
            $msg += $requiredfiles -join "`n"
            $caption = "Error"
            [System.Windows.Forms.MessageBox]::Show($msg, $caption, 0, 16)
            throw $msg
        }
    }

# Loading Data
$MainDataLoad = {
    [InstructorEvent[]]$Script:AllEvents = @(Import-Csv -Path (Join-Path -Path $Config 'events.csv'))
    [InstructorEvent[]]$Script:FilteredEvents = $AllEvents
    $Script:Instructors = @(Import-Csv -Path (Join-Path -Path $Config "whitelist.csv"))
    $Script:ClassesLoaded = @($AllEvents | 
        Sort-Object -Property Course, Class, Asof -Unique |
            Select-Object Course, Class, AsOf)
    $EarliestStart = $AllEvents |
        Sort-Object -Property Start |
            Select-Object -ExpandProperty Start -First 1
    $LatestStart   = $AllEvents |
        Sort-Object -Property Start -Descending |
            Select-Object -ExpandProperty Start -First 1
    $Script:Courses       = $AllEvents |
        Sort-Object -Property Course -Unique |
            Select-Object -ExpandProperty Course
    $Script:Classes       = $AllEvents |
        Sort-Object -Property Class -Unique |
            Select-Object -ExpandProperty Class

    $TotalEvents = ($AllEvents).count

    # Setting up main form with initial data
    #Checking if user has write privieges to data source file
    try {
        [system.io.file]::OpenWrite((Join-Path -Path $Config "events.csv")).close()
    }
    catch {
        $ButtonRemoveClassSched.Enabled = $false
        $ButtonImportSched.Enabled = $false
    }
    # Total Events Label
    $LabelTotalEvents.Text = "Total Events: {0:N0}" -f $TotalEvents

    # Loaded Classes Grid Column Headings
    $DataGridViewClassesLoaded.ColumnCount = 3
    $DataGridViewClassesLoaded.Columns[0].Name = "Course"
    $DataGridViewClassesLoaded.Columns[1].Name = "Class"
    $DataGridViewClassesLoaded.Columns[2].Name = "As Of"
    
    #remove any existing rows
    $DataGridViewClassesLoaded.ClearSelection()
    for ($i = 0; $i -lt $DataGridViewClassesLoaded.Rows.Count;) {
        $DataGridViewClassesLoaded.Rows.RemoveAt(0)
    }

    #create new rows
    $Script:ClassesLoaded | 
        Sort-Object -Property Course, Class |
            ForEach-Object {
                $DataGridViewClassesLoaded.Rows.Add($_.Course, $_.Class, $_.AsOf.ToString("d")) |  Out-Null
            } # Foreach-Object

    # Filtered Default Time Frame
    $DateTimePickerStartFilter.Value = $EarliestStart
    $DateTimePickerEndFilter.Value   = $LatestStart

    # Instructors
    $DataGridViewInstructors.ColumnCount = 2
    $DataGridViewInstructors.Columns[0].Name     = "Instructor"
    $DataGridViewInstructors.Columns[0].ReadOnly = $true
    $DataGridViewInstructors.Columns[1].Name     = "DOD Instr"
    $DataGridViewInstructors.Columns[1].ReadOnly = $true
    $Script:Instructors |
        ForEach-Object {$DataGridViewInstructors.Rows.Add($_.Name, $_.DOD) | Out-Null
            
        } # foreach-Object

    # Class Filter
    $ComboBoxClassFilter.Items.Add("*") | Out-Null
    $Script:Classes |
        ForEach-Object {
            $ComboBoxClassFilter.Items.Add($_) | Out-Null
        }

    #Course Filter
    $ComboBoxCourseFilter.Items.Add("*") | Out-Null
    $Script:Courses |
        ForEach-Object {
            $ComboBoxCourseFilter.Items.Add($_) | Out-Null
        }

} #MainDataLoad
$MainDataLoad.Invoke()
$FormInstructorUtilization.ShowDialog() 

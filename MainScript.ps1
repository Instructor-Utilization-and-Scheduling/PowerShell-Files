# Including what will be our module....
. (Join-Path $PSScriptRoot 'InstructorUtilizationModule.ps1')

Add-Type -AssemblyName System.Windows.Forms

$UpdateFiltered = {
    $SelectedInstructors = @(foreach ($row in $DataGridViewInstructors.SelectedRows) {
        $row.cells[0].value
    })
    if ($ComboBoxClassFilter.SelectedItem -and $ComboBoxCourseFilter.SelectedItem) {            
        $FilteredEvents = @($AllEvents | 
            Where-Object {$_.Instructor -in $SelectedInstructors -and $_.start -ge $DateTimePickerStartFilter.Value -and $_.end -le $DateTimePickerEndFilter.Value -and $_.Course -like $ComboBoxCourseFilter.SelectedItem.ToString() -and $_.Class -like $ComboBoxClassFilter.SelectedItem.ToString()})
        $LabelFilteredEvents.Text = "Filtered Events: {0:N0}" -f $FilteredEvents.count
    } # if
} # UpdateFiltered

$InstrUtilizationLoaded = {
    $ComboBoxCourseFilter.SelectedItem = "*"
    $ComboBoxClassFilter.SelectedItem  = "*"
    $DataGridViewInstructors.SelectAll()    
} # InstrUtilizationLoaded

# Form Layout script
. (Join-Path $PSScriptRoot 'MainScript.designer.ps1')

# Getting Configuration Information
$Config = Get-Content -Path (Join-Path -Path $PSScriptRoot 'config.cfg')

# Loading Data
[InstructorEvent[]]$AllEvents = @(Import-Csv -Path (Join-Path -Path $Config 'events.csv'))
$Instructors = @(Import-Csv -Path (Join-Path -Path $Config "whitelist.csv"))
$ClassesLoaded = @($AllEvents | 
    Sort-Object -Property Course, Class, Asof -Unique |
        Select-Object Course, Class, AsOf)
$EarliestStart = $AllEvents |
    Sort-Object -Property Start |
        Select-Object -ExpandProperty Start -First 1
$LatestStart   = $AllEvents |
    Sort-Object -Property Start -Descending |
        Select-Object -ExpandProperty Start -First 1
$Courses       = $AllEvents |
    Sort-Object -Property Course -Unique |
        Select-Object -ExpandProperty Course
$Classes       = $AllEvents |
    Sort-Object -Property Class -Unique |
        Select-Object -ExpandProperty Class

$TotalEvents = ($AllEvents).count

# Setting up main form with initial data
#Checking if user has write privieges to data source file
try {
    [io.file]::OpenWrite((Join-Path -Path $Config "events.csv")).close()
}
catch {
    $ButtonRemoveClassSched.Enabled = $false
    $ButtonImportSched.Enabled = $false
}
# Total Events Label
$LabelTotalEvents.Text = "Total Events: {0:N0}" -f $TotalEvents

# Loaded Classes Grid
$DataGridViewClassesLoaded.ColumnCount = 3
$DataGridViewClassesLoaded.Columns[0].Name = "Course"
$DataGridViewClassesLoaded.Columns[1].Name = "Class"
$DataGridViewClassesLoaded.Columns[2].Name = "As Of"
$ClassesLoaded | 
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
$Instructors |
    ForEach-Object {$DataGridViewInstructors.Rows.Add($_.Name, $_.DOD) | Out-Null
        
    } # foreach-Object

# Class Filter
$ComboBoxClassFilter.Items.Add("*") | Out-Null
$Classes |
    ForEach-Object {
        $ComboBoxClassFilter.Items.Add($_) | Out-Null
    }

#Course Filter
$ComboBoxCourseFilter.Items.Add("*") | Out-Null
$Courses |
    ForEach-Object {
        $ComboBoxCourseFilter.Items.Add($_) | Out-Null
    }
$FormInstructorUtilization.ShowDialog() 

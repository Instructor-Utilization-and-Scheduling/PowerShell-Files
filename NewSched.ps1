function ValidateSchedFile ($Path, $course, $class, [datetime]$asof){
    if (Test-Path -Path $Path) {
        #Check for exists
        $checkifoverwrite = @($script:ClassesLoaded | 
        Where-Object {$_.course -eq $course -and $_.class -eq $class})
        if ($checkifoverwrite.count -gt 0) {
            $msg     = "Looks like {0} is already loaded. Are you sure you want to overwrite it?" -f (($course, $class) -join ", ")
            $caption = "Are you sure?"
            $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
            $icon    = [System.Windows.Forms.MessageBoxIcon]::Question
            $result  = [System.Windows.Forms.MessageBox]::Show($msg, $caption, $buttons, $icon)
            if ($result = [System.Windows.Forms.DialogResult]::Yes) {
                return $True
            }
            else {return $false}
        } # if checkoverwrite
        else {return $True} #new file
    } #if valid file
    else {
        $msg = "{0} does not exist. Select a different file." -f $Path
        $caption = "Warning"
        [System.Windows.Forms.MessageBox]::Show($msg, $caption, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return $false
    } #else not a valid file
}
$ImportSchedule = {
    $ButtonImportSchedule.Text = "Working...."
    $ButtonImportSchedule.Enabled = $false
    $ht = @{
        Path   = $TextBoxSchedFile.Text
        Course = $ComboBoxNewSchedCourse.Text
        Class  = $ComboBoxNewSchedClass.Text
        AsOf   = $DateTimePickerNSAsOf.value
    } # hashtable definition
    if (ValidateSchedFile @ht) {
        [string[]]$ht.InstructorWhiteList = @(Import-Csv -Path "$($script:Config)\whitelist.csv" | Select-Object -ExpandProperty Name)
        $ht.AliasFile           = "$script:Config\NameAliases.csv"
        [InstructorEvent[]]$importedevents = @(Import-ExcelSched @ht)    
        $msg     = "Created {0} events. Please review the results. Are you sure you want to commit this data?" -f ($importedevents.count)
        $caption = "Are you sure?"   
        if ([System.Windows.Forms.MessageBox]::Show($msg, $caption, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq [System.Windows.Forms.DialogResult]::Yes) {
            #remove old sched from memory if exists
            [InstructorEvent[]]$script:AllEvents = $script:AllEvents |
                                                        Where-Object {!($_.course -eq $ht.course -and $_.class -eq $ht.class)}
            #append new sched in memory
            [InstructorEvent[]]$script:AllEvents += $importedevents
            #update csv file
            $script:AllEvents | Export-Csv -Path "$config\events.csv"
            
            #Update main form with new class data
            $script:MainDataLoad.Invoke()
            $FormNewSched.close()           
        } # if message box yes    
    } #if ValidateSchedFile
    $ButtonImportSchedule.Text = "Import Schedule"
    $ButtonImportSchedule.Enabled = $true
} #ImportSched

$NewSchedCancel = {
    $FormNewSched.close()
} #NewSchedCancel

$NewSchedSelectFile = {
    
    $openfiledialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog 
    $openfiledialog.InitialDirectory = $env:USERPROFILE
    $openfiledialog.Multiselect = $false
    $openfiledialog.Title = "Select Schedule to Import"
    $openfiledialog.Filter = "Excel Sched Files (*.xlsx)|*.xlsx"
    $openfiledialog.FilterIndex = 2
    $openfiledialog.ShowDialog()
    $TextBoxSchedFile.Text = $openfiledialog.FileName
} # NewSchedSelectFile

Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'NewSched.designer.ps1')

#load data
$Script:Classes |
    ForEach-Object {
        $ComboBoxNewSchedClass.Items.Add($_) | Out-Null
    }
$ComboBoxNewSchedClass.SelectedItem = $script:DataGridViewClassesLoaded.SelectedCells[1].value
$script:Courses |
    ForEach-Object {
        $ComboBoxNewSchedCourse.Items.Add($_) | Out-Null
    }
$ComboBoxNewSchedCourse.SelectedItem = $script:DataGridViewClassesLoaded.SelectedCells[0].value
$script:DataGridViewClassesLoaded.SelectAll()
$FormNewSched.ShowDialog()
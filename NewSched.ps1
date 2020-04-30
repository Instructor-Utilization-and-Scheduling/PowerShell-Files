$NewSchedCancel = {
    $FormNewSched.close()
}
$NewSchedSelectFile = {


    $openfiledialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog 
    $openfiledialog.InitialDirectory = $env:USERPROFILE
    $openfiledialog.Multiselect = $false
    $openfiledialog.Title = "Select Schedule to Import"
    if ($openfiledialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $checkifoverwrite = @($script:ClassesLoaded | 
        Where-Object {$_.course -eq $ComboBoxNewSchedCourse.SelectedItem -and $_.class -eq $ComboBoxNewSchedClass.SelectedItem})
        if ($checkifoverwrite.count -gt 0) {
            $msg = "Looks like {0} is already loaded. Are you sure you want to overwrite it?" -f (($ComboBoxNewSchedCourse.SelectedItem, $ComboBoxNewSchedClass.SelectedItem) -join ", ")
            $caption = "Are you sure?"
            $result = [System.Windows.Forms.MessageBox]::Show($msg, $caption, 4, 32)
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                [InstructorEvent[]]$updatedevents = @($script:AllEvents |
                    Where-Object {$_.course -ne $ComboBoxNewSchedCourse.SelectedItem -and $_.class -ne $ComboBoxNewSchedClass.SelectedItem})
                $updatedevents | Export-Csv -Path "$config\events.csv" -Force
                $script:AllEvents = $updatedevents    
            } # if overwrite yes
        } # If already exists
        $filepath = $openfiledialog.FileName
        if (Test-Path -Path $filepath){
            $ht = @{
                Path                = $filepath
                Course              = $ComboBoxNewSchedCourse.SelectedItem
                Class               = $ComboBoxNewSchedClass.SelectedItem
                AliasFile           = "$script:Config\NameAliases.csv"
                InstructorWhiteList = Import-Csv -Path "$($script:Config)\whitelist.csv"
                AsOfDate            = $DateTimePickerNSAsOf.value
            } #ht definition
            [InstructorEvent[]]$importedevents = @(Import-ExcelSched @ht)
            $msg = "Created {0} events. Please review the results. Are you sure you want to commit this data?" -f ($importedevents.count)
            $caption = "Are you sure?"
            if ([System.Windows.Forms.MessageBox]::Show($msg, $caption, 4, 32) -eq [System.Windows.Forms.DialogResult]::Yes) {
                $importedevents | Export-Csv -Path "$config\events.csv" -Append
                [InstructorEvent[]]$script:AllEvents += $importedevents
                $script:ClassesLoaded = @($script:AllEvents | 
                    Sort-Object -Property Course, Class, Asof -Unique |
                        Select-Object Course, Class, AsOf)
                0..($script:DataGridViewClassesLoaded.rows.Count) |
                    ForEach-Object { $script:DataGridViewClassesLoaded.rows.RemoveAt($_)}
                $ClassesLoaded |
                    ForEach-Object {
                        $script:DataGridViewClassesLoaded.Rows.Add($_.Course, $_.Class, $_.AsOf.ToString("d")) |  Out-Null
                    } #classes foreach-object                
            } # if message box yes
        } # if Test-Path
    } # if select schedule to import
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
$FormNewSched.ShowDialog()
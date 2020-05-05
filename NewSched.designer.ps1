$FormNewSched = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.Button]$ButtonNewSchedSelectFile = $null
[System.Windows.Forms.Button]$ButtonNewShedCancel = $null
[System.Windows.Forms.ComboBox]$ComboBoxNewSchedCourse = $null
[System.Windows.Forms.ComboBox]$ComboBoxNewSchedClass = $null
[System.Windows.Forms.Label]$LabelNewSchedCourse = $null
[System.Windows.Forms.Label]$LabelNewSchedClass = $null
[System.Windows.Forms.DateTimePicker]$DateTimePickerNSAsOf = $null
[System.Windows.Forms.Label]$LabelNSAsOf = $null
[System.Windows.Forms.TextBox]$TextBoxSchedFile = $null
[System.Windows.Forms.Button]$ButtonImportSchedule = $null
function InitializeComponent
{
$ButtonNewSchedSelectFile = (New-Object -TypeName System.Windows.Forms.Button)
$ButtonNewShedCancel = (New-Object -TypeName System.Windows.Forms.Button)
$ComboBoxNewSchedCourse = (New-Object -TypeName System.Windows.Forms.ComboBox)
$ComboBoxNewSchedClass = (New-Object -TypeName System.Windows.Forms.ComboBox)
$LabelNewSchedCourse = (New-Object -TypeName System.Windows.Forms.Label)
$LabelNewSchedClass = (New-Object -TypeName System.Windows.Forms.Label)
$DateTimePickerNSAsOf = (New-Object -TypeName System.Windows.Forms.DateTimePicker)
$LabelNSAsOf = (New-Object -TypeName System.Windows.Forms.Label)
$TextBoxSchedFile = (New-Object -TypeName System.Windows.Forms.TextBox)
$ButtonImportSchedule = (New-Object -TypeName System.Windows.Forms.Button)
$FormNewSched.SuspendLayout()
#
#ButtonNewSchedSelectFile
#
$ButtonNewSchedSelectFile.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$ButtonNewSchedSelectFile.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]32,[System.Int32]110))
$ButtonNewSchedSelectFile.Name = [System.String]'ButtonNewSchedSelectFile'
$ButtonNewSchedSelectFile.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]23))
$ButtonNewSchedSelectFile.TabIndex = [System.Int32]0
$ButtonNewSchedSelectFile.Text = [System.String]'Select File'
$ButtonNewSchedSelectFile.UseCompatibleTextRendering = $true
$ButtonNewSchedSelectFile.UseVisualStyleBackColor = $true
$ButtonNewSchedSelectFile.add_Click($NewSchedSelectFile)
#
#ButtonNewShedCancel
#
$ButtonNewShedCancel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$ButtonNewShedCancel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]270,[System.Int32]152))
$ButtonNewShedCancel.Name = [System.String]'ButtonNewShedCancel'
$ButtonNewShedCancel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]154,[System.Int32]53))
$ButtonNewShedCancel.TabIndex = [System.Int32]1
$ButtonNewShedCancel.Text = [System.String]'Cancel'
$ButtonNewShedCancel.UseCompatibleTextRendering = $true
$ButtonNewShedCancel.UseVisualStyleBackColor = $true
$ButtonNewShedCancel.add_Click($NewSchedCancel)
#
#ComboBoxNewSchedCourse
#
$ComboBoxNewSchedCourse.FormattingEnabled = $true
$ComboBoxNewSchedCourse.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]138,[System.Int32]31))
$ComboBoxNewSchedCourse.Name = [System.String]'ComboBoxNewSchedCourse'
$ComboBoxNewSchedCourse.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]286,[System.Int32]21))
$ComboBoxNewSchedCourse.TabIndex = [System.Int32]2
#
#ComboBoxNewSchedClass
#
$ComboBoxNewSchedClass.FormattingEnabled = $true
$ComboBoxNewSchedClass.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]138,[System.Int32]58))
$ComboBoxNewSchedClass.Name = [System.String]'ComboBoxNewSchedClass'
$ComboBoxNewSchedClass.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]286,[System.Int32]21))
$ComboBoxNewSchedClass.TabIndex = [System.Int32]3
#
#LabelNewSchedCourse
#
$LabelNewSchedCourse.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$LabelNewSchedCourse.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]32,[System.Int32]31))
$LabelNewSchedCourse.Name = [System.String]'LabelNewSchedCourse'
$LabelNewSchedCourse.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]23))
$LabelNewSchedCourse.TabIndex = [System.Int32]4
$LabelNewSchedCourse.Text = [System.String]'Course'
$LabelNewSchedCourse.UseCompatibleTextRendering = $true
#
#LabelNewSchedClass
#
$LabelNewSchedClass.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$LabelNewSchedClass.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]32,[System.Int32]58))
$LabelNewSchedClass.Name = [System.String]'LabelNewSchedClass'
$LabelNewSchedClass.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]23))
$LabelNewSchedClass.TabIndex = [System.Int32]5
$LabelNewSchedClass.Text = [System.String]'Class'
$LabelNewSchedClass.UseCompatibleTextRendering = $true
#
#DateTimePickerNSAsOf
#
$DateTimePickerNSAsOf.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]138,[System.Int32]85))
$DateTimePickerNSAsOf.Name = [System.String]'DateTimePickerNSAsOf'
$DateTimePickerNSAsOf.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]200,[System.Int32]21))
$DateTimePickerNSAsOf.TabIndex = [System.Int32]6
#
#LabelNSAsOf
#
$LabelNSAsOf.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$LabelNSAsOf.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]32,[System.Int32]85))
$LabelNSAsOf.Name = [System.String]'LabelNSAsOf'
$LabelNSAsOf.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]23))
$LabelNSAsOf.TabIndex = [System.Int32]7
$LabelNSAsOf.Text = [System.String]'As Of Date'
$LabelNSAsOf.UseCompatibleTextRendering = $true
#
#TextBoxSchedFile
#
$TextBoxSchedFile.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]138,[System.Int32]112))
$TextBoxSchedFile.Name = [System.String]'TextBoxSchedFile'
$TextBoxSchedFile.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]286,[System.Int32]21))
$TextBoxSchedFile.TabIndex = [System.Int32]8
#
#ButtonImportSchedule
#
$ButtonImportSchedule.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$ButtonImportSchedule.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]32,[System.Int32]152))
$ButtonImportSchedule.Name = [System.String]'ButtonImportSchedule'
$ButtonImportSchedule.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]154,[System.Int32]53))
$ButtonImportSchedule.TabIndex = [System.Int32]9
$ButtonImportSchedule.Text = [System.String]'Import Schedule'
$ButtonImportSchedule.UseCompatibleTextRendering = $true
$ButtonImportSchedule.UseVisualStyleBackColor = $true
$ButtonImportSchedule.add_Click($ImportSchedule)
#
#FormNewSched
#
$FormNewSched.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]460,[System.Int32]233))
$FormNewSched.Controls.Add($ButtonImportSchedule)
$FormNewSched.Controls.Add($TextBoxSchedFile)
$FormNewSched.Controls.Add($LabelNSAsOf)
$FormNewSched.Controls.Add($DateTimePickerNSAsOf)
$FormNewSched.Controls.Add($LabelNewSchedClass)
$FormNewSched.Controls.Add($LabelNewSchedCourse)
$FormNewSched.Controls.Add($ComboBoxNewSchedClass)
$FormNewSched.Controls.Add($ComboBoxNewSchedCourse)
$FormNewSched.Controls.Add($ButtonNewShedCancel)
$FormNewSched.Controls.Add($ButtonNewSchedSelectFile)
$FormNewSched.Text = [System.String]'Import New/Updated Schedule'
$FormNewSched.ResumeLayout($false)
$FormNewSched.PerformLayout()
Add-Member -InputObject $FormNewSched -Name base -Value $base -MemberType NoteProperty
Add-Member -InputObject $FormNewSched -Name ButtonNewSchedSelectFile -Value $ButtonNewSchedSelectFile -MemberType NoteProperty
Add-Member -InputObject $FormNewSched -Name ButtonNewShedCancel -Value $ButtonNewShedCancel -MemberType NoteProperty
Add-Member -InputObject $FormNewSched -Name ComboBoxNewSchedCourse -Value $ComboBoxNewSchedCourse -MemberType NoteProperty
Add-Member -InputObject $FormNewSched -Name ComboBoxNewSchedClass -Value $ComboBoxNewSchedClass -MemberType NoteProperty
Add-Member -InputObject $FormNewSched -Name LabelNewSchedCourse -Value $LabelNewSchedCourse -MemberType NoteProperty
Add-Member -InputObject $FormNewSched -Name LabelNewSchedClass -Value $LabelNewSchedClass -MemberType NoteProperty
Add-Member -InputObject $FormNewSched -Name DateTimePickerNSAsOf -Value $DateTimePickerNSAsOf -MemberType NoteProperty
Add-Member -InputObject $FormNewSched -Name LabelNSAsOf -Value $LabelNSAsOf -MemberType NoteProperty
Add-Member -InputObject $FormNewSched -Name TextBoxSchedFile -Value $TextBoxSchedFile -MemberType NoteProperty
Add-Member -InputObject $FormNewSched -Name ButtonImportSchedule -Value $ButtonImportSchedule -MemberType NoteProperty
}
. InitializeComponent

$FormOutlook = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.Label]$LabelOutlookNumEvents = $null
[System.Windows.Forms.ComboBox]$ComboBoxOutlookCal = $null
[System.Windows.Forms.Label]$LabelOutlookCal = $null
[System.Windows.Forms.Button]$ButtonOutlookCreate = $null
[System.Windows.Forms.Button]$ButtonOutlookCancel = $null
function InitializeComponent
{
$LabelOutlookNumEvents = (New-Object -TypeName System.Windows.Forms.Label)
$ComboBoxOutlookCal = (New-Object -TypeName System.Windows.Forms.ComboBox)
$LabelOutlookCal = (New-Object -TypeName System.Windows.Forms.Label)
$ButtonOutlookCreate = (New-Object -TypeName System.Windows.Forms.Button)
$ButtonOutlookCancel = (New-Object -TypeName System.Windows.Forms.Button)
$FormOutlook.SuspendLayout()
#
#LabelOutlookNumEvents
#
$LabelOutlookNumEvents.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$LabelOutlookNumEvents.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]23,[System.Int32]9))
$LabelOutlookNumEvents.Name = [System.String]'LabelOutlookNumEvents'
$LabelOutlookNumEvents.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]490,[System.Int32]23))
$LabelOutlookNumEvents.TabIndex = [System.Int32]0
$LabelOutlookNumEvents.Text = [System.String]'Number of Events: 0'
$LabelOutlookNumEvents.UseCompatibleTextRendering = $true
#
#ComboBoxOutlookCal
#
$ComboBoxOutlookCal.DisplayMember = [System.String]'name'
$ComboBoxOutlookCal.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$ComboBoxOutlookCal.FormattingEnabled = $true
$ComboBoxOutlookCal.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]157,[System.Int32]56))
$ComboBoxOutlookCal.Name = [System.String]'ComboBoxOutlookCal'
$ComboBoxOutlookCal.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]356,[System.Int32]21))
$ComboBoxOutlookCal.TabIndex = [System.Int32]1
$ComboBoxOutlookCal.ValueMember = [System.String]'fullfolderpath'
#
#LabelOutlookCal
#
$LabelOutlookCal.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$LabelOutlookCal.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]23,[System.Int32]56))
$LabelOutlookCal.Name = [System.String]'LabelOutlookCal'
$LabelOutlookCal.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]130,[System.Int32]23))
$LabelOutlookCal.TabIndex = [System.Int32]2
$LabelOutlookCal.Text = [System.String]'Outlook Calendar'
$LabelOutlookCal.UseCompatibleTextRendering = $true
#
#ButtonOutlookCreate
#
$ButtonOutlookCreate.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$ButtonOutlookCreate.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]23,[System.Int32]129))
$ButtonOutlookCreate.Name = [System.String]'ButtonOutlookCreate'
$ButtonOutlookCreate.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]141,[System.Int32]52))
$ButtonOutlookCreate.TabIndex = [System.Int32]3
$ButtonOutlookCreate.Text = [System.String]'Create Outlook Appointments'
$ButtonOutlookCreate.UseCompatibleTextRendering = $true
$ButtonOutlookCreate.UseVisualStyleBackColor = $true
$ButtonOutlookCreate.add_Click($GenerateAppts)
#
#ButtonOutlookCancel
#
$ButtonOutlookCancel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$ButtonOutlookCancel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]372,[System.Int32]129))
$ButtonOutlookCancel.Name = [System.String]'ButtonOutlookCancel'
$ButtonOutlookCancel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]141,[System.Int32]52))
$ButtonOutlookCancel.TabIndex = [System.Int32]4
$ButtonOutlookCancel.Text = [System.String]'CANCEL'
$ButtonOutlookCancel.UseCompatibleTextRendering = $true
$ButtonOutlookCancel.UseVisualStyleBackColor = $true
$ButtonOutlookCancel.add_Click($OutlookFormClose)
#
#FormOutlook
#
$FormOutlook.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]525,[System.Int32]208))
$FormOutlook.Controls.Add($ButtonOutlookCancel)
$FormOutlook.Controls.Add($ButtonOutlookCreate)
$FormOutlook.Controls.Add($LabelOutlookCal)
$FormOutlook.Controls.Add($ComboBoxOutlookCal)
$FormOutlook.Controls.Add($LabelOutlookNumEvents)
$FormOutlook.Text = [System.String]'Create Outlook Appts.'
$FormOutlook.ResumeLayout($false)
Add-Member -InputObject $FormOutlook -Name base -Value $base -MemberType NoteProperty
Add-Member -InputObject $FormOutlook -Name LabelOutlookNumEvents -Value $LabelOutlookNumEvents -MemberType NoteProperty
Add-Member -InputObject $FormOutlook -Name ComboBoxOutlookCal -Value $ComboBoxOutlookCal -MemberType NoteProperty
Add-Member -InputObject $FormOutlook -Name LabelOutlookCal -Value $LabelOutlookCal -MemberType NoteProperty
Add-Member -InputObject $FormOutlook -Name ButtonOutlookCreate -Value $ButtonOutlookCreate -MemberType NoteProperty
Add-Member -InputObject $FormOutlook -Name ButtonOutlookCancel -Value $ButtonOutlookCancel -MemberType NoteProperty
}
. InitializeComponent

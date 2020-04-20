$Form1 = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.Label]$Label1 = $null
[System.Windows.Forms.ListBox]$ListBox1 = $null
[System.Windows.Forms.DateTimePicker]$DateTimePicker1 = $null
function InitializeComponent
{
$Label1 = (New-Object -TypeName System.Windows.Forms.Label)
$ListBox1 = (New-Object -TypeName System.Windows.Forms.ListBox)
$DateTimePicker1 = (New-Object -TypeName System.Windows.Forms.DateTimePicker)
$Form1.SuspendLayout()
#
#Label1
#
$Label1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]52,[System.Int32]80))
$Label1.Name = [System.String]'Label1'
$Label1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]23))
$Label1.TabIndex = [System.Int32]0
$Label1.Text = [System.String]'Label1'
$Label1.UseCompatibleTextRendering = $true
$Label1.add_Click($Label1_Click)
#
#ListBox1
#
$ListBox1.FormattingEnabled = $true
$ListBox1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]69,[System.Int32]156))
$ListBox1.Name = [System.String]'ListBox1'
$ListBox1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]120,[System.Int32]95))
$ListBox1.TabIndex = [System.Int32]1
#
#DateTimePicker1
#
$DateTimePicker1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]175,[System.Int32]82))
$DateTimePicker1.Name = [System.String]'DateTimePicker1'
$DateTimePicker1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]200,[System.Int32]21))
$DateTimePicker1.TabIndex = [System.Int32]2
#
#Form1
#
$Form1.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]413,[System.Int32]372))
$Form1.Controls.Add($DateTimePicker1)
$Form1.Controls.Add($ListBox1)
$Form1.Controls.Add($Label1)
$Form1.Text = [System.String]'Form1'
$Form1.ResumeLayout($false)
Add-Member -InputObject $Form1 -Name base -Value $base -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name Label1 -Value $Label1 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name ListBox1 -Value $ListBox1 -MemberType NoteProperty
Add-Member -InputObject $Form1 -Name DateTimePicker1 -Value $DateTimePicker1 -MemberType NoteProperty
}
. InitializeComponent

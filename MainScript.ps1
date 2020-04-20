$Label1_Click = {
}
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'MainScript.designer.ps1')
$Form1.ShowDialog() 
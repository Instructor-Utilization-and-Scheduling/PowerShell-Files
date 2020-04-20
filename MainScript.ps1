Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'MainScript.designer.ps1')
$FormInstructorUtilization.ShowDialog() 
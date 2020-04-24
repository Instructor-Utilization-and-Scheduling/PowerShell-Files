# Including what will be our module....
. (Join-Path $PSScriptRoot 'InstructorUtilizationModule.ps1')
function TestingImport 
{
    [CmdletBinding()]
    param (
        [string[]]
        $filepath,

        [string]
        $InputPath
    )

    $schedules = @(foreach ($file in $filepath) {
        $FileObj = Get-Item -Path $file
        [PSCustomObject]@{
                path = $file
                course = (-split $FileObj.Name)[0]
                class  = (-split $FileObj.Name)[1]
                AliasFile = "$InputPath\Config\NameAliases.csv" 
                InstructorWhiteList = @(Import-Csv "$InputPath\Config\whitelist.csv" | Select-Object -ExpandProperty Name)             
            } # pscustomobject        
        } # foreach file
    ) # schedules array
    $schedules | Import-ExcelSched

} # function TestingImport


$InputPath = "C:\Users\micha\Documents\InputData\Config"
$OutputPath = "C:\Users\micha\Documents\OutputData"


<## Import Testing
$files = @(Get-ChildItem -Path "$InputPath\Schedules" -File -Recurse |
                Select-Object -ExpandProperty FullName
)
TestingImport -filepath $files -InputPath $InputPath | 
    Export-csv -Path "$OutputPath\events.csv" -Force
 
# Testing Report
[InstructorEvent[]]$events = Import-Csv -Path "$InputPath\events.csv"
$DODIntructors = Import-Csv -Path "$InputPath\whitelist.csv" |
    Where-Object {$_.DOD -eq "T"} |
        Select-Object -ExpandProperty Name
$DODEvents = $events | Where-Object {$_.Instructor -in $DODIntructors -and $_.start -ge (Get-Date "1-Oct-2019")}
$report = Measure-Events -InstructorEvents $DODEvents -Grouping "Quarterly"
$report | Out-File -FilePath "$OutputPath\Quarterly_Analysis.txt" -Force
$report = Measure-Events -Instructorevents $DODEvents -grouping "Monthly"
$report | Out-File -FilePath "$OutputPath\Monthly_Analysis.txt" -Force 
#>
#testing outlook
$ht = @{
    CalendarFolder = "\\michael.ralph72@gmail.com\Calendar (This computer only)\WorkGroup1 (This computer only)"
    Start          = Get-Date
    End            = (Get-Date).AddHours(2)
    Subject        = "Test Subject"
    Location       = "My Office"
    Category       = "Test basic"
    Body           = "This was a test"
}
New-OutlookEvent @ht
$myfolder = Get-OutlookCalendars | Where-Object {$_.fullfolderpath -eq $ht.CalendarFolder}
Remove-OutlookAppointment -OutlookCalendar $myfolder -Search "CWO / 20-01*"

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
                AliasFile = "$InputPath\NameAliases.csv" 
                InstructorWhiteList = @(Get-Content "$InputPath\whitelist.txt")             
            } # pscustomobject        
        } # foreach file
    ) # schedules array
    $schedules | Import-ExcelSched

} # function TestingImport


$InputPath = "C:\Users\micha\Documents\InputData"
$OutputPath = "C:\Users\micha\Documents\OutputData"
$CalendarFolder = "\\michael.ralph72@gmail.com\Calendar (This computer only)\WorkGroup1 (This computer only)"

# Import Testing
$files = @(Get-ChildItem -Path "$InputPath\Schedules" -File -Recurse |
                Select-Object -ExpandProperty FullName
)
TestingImport -filepath $files -InputPath $InputPath | 
    Export-csv -Path "$OutputPath\events.csv" -Force

# Testing Report
[InstructorEvent[]]$events = Import-Csv -Path "$OutputPath\events.csv"
$report = Measure-Events -InstructorEvents $events -Grouping "Quarterly"
$report | Out-File -FilePath "$OutputPath\Quarterly_Analysis.txt" -Force
[InstructorEvent[]]$events = Import-Csv -Path "$OutputPath\events.csv"
$report = Measure-Events -Instructorevents $events -grouping "Monthly"
$report | Out-File -FilePath "$OutputPath\Monthly_Analysis.txt" -Force 

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
# Testing Instructor Events
[InstructorEvent[]]$events = Import-Csv -Path "$OutputPath\events.csv"
$events | 
    Where-Object {$_.Instructor -eq "Mr. Ralph" -and $_.start -ge (Get-Date "1 Jan 2020")} |
        New-OutlookEvent -CalendarFolder $CalendarFolder

[InstructorEvent[]]$events = Import-Csv -Path "$OutputPath\events.csv"
    $events | 
        Where-Object {$_.Instructor -eq "Mr. Ralph" -and $_.start -ge (Get-Date "1 Jan 2020")} |
            Export-ICS -Path C:\Users\micha\Documents\OutputData\events.ics 

$events | 
    Where-Object {$_.Instructor -eq "Mr. Ralph" -and $_.start -ge (Get-Date "1 Jan 2020")} | 
        Remove-OutlookEvent -CalendarFolder "\\michael.ralph72@gmail.com\Calendar (This computer only)\WorkGroup1 (This computer only)"

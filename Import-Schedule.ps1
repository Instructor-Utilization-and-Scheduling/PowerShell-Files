# Define InstructorEvent Class
class InstructorEvent
{
   [string]   $Instructor
   [datetime] $start
   [datetime] $end
   [string]   $course
   [string]   $class
   [string]   $lesson
   [string]   $room
   [string]   $CYWeek
   [string]   $FYWeek
   [double]   $duration
   
   [ValidateSet("Primary", "Secondary", "Support")]
   [string] $Role
   
   hidden [int] $FY 

   [double] CalculateDuration()
   {
       return ($this.end - $this.start).TotalHours.ToString("F2")
   } # CaluculateDuration

   [string] CYWeekOfYear()
   {
        $startdate = $this.start.AddDays(-$this.start.DayOfWeek.value__)
        $enddate   = $startdate.AddDays(6)
        return "CY{0}-{1} / {2} - {3}" -f $this.start.Year, 
                                            (Get-Date -Date $this.start -UFormat %U), 
                                            $startdate.ToString("d-MMM-yyyy"), 
                                            $enddate.ToString("d-MMM-yyyy")

   } # CYWeekOfYear

   [string] FYWeekOfYear()
   {   
        $FYWeekBegin = Get-Date -Year ($this.FY - 1) -Month 10 -Day 1
        $FYWeekBegin = $FYWeekBegin.AddDays(-$FYWeekBegin.DayOfWeek.value__)
        $startdate = $FYWeekBegin   
        foreach ($wk in 1..52) {
            $startdate = $FYWeekBegin.AddDays(($wk - 1) * 7)
            if ($startdate.AddDays(7) -gt $this.start) { BREAK }
        } #foreach
        $enddate   = $startdate.AddDays(6)

       return "FY{0}-{1} / {2} - {3}" -f $this.FY, 
                                        $Wk, 
                                        $startdate.ToString("d-MMM-yyyy"), 
                                        $enddate.ToString("d-MMM-yyyy")
   } # FYWeekOfYear
     
   # Constructor
   InstructorEvent ([string]   $Instructor, 
          [datetime] $start,
          [datetime] $end,
          [string]   $class,
          [string]   $course,
          [string]   $room,
          [string]   $lesson,
          [string]   $role          
          ) # Constructor parameters
   {
       If ($end -le $start){Write-Error -Category InvalidData -Message "End must be after start!"}
       If ($role -notin "Primary", "Secondary", "Support") {Write-Error -Category InvalidData -Message "Role must be Primary, Secondary or Support"}
       $this.Instructor = $Instructor
       $this.start      = $start
       $this.end        = $end
       $this.class      = $class
       $this.course     = $course
       $this.room       = $room
       $this.lesson     = $lesson
       $this.Role       = $role    
       $this.duration   = $this.CalculateDuration()
       $this.FY         = switch ($this.start.Month) {
                            {$_ -le 9} {$this.start.Year}
                            Default {$this.start.Year + 1 }
                        } # switch
       $this.CYWeek     = $this.CYWeekOfYear()
       $this.FYWeek     = $this.FYWeekOfYear()  

   } #Constructor Definition

   #Empty Constructor
   InstructorEvent()
   {$this.Instructor = ""}
} #Class Defition

#This function is used to call the constructor so we can use a hashtable to create objects (splatting)
function New-SchedEvent {
    param (
        [string]   $Instructor, 
        [datetime] $start,
        [datetime] $end,
        [string]   $class,
        [string]   $course,
        [string]   $room,
        [string]   $lesson,
        [string]   $role          
          ) # param
    [InstructorEvent]::new($Instructor, $start, $end, $class, $course, $room, $lesson, $role)
    
} #function New-SchedEvent
<#
.Synopsis
    Imports an Excel Schedule and returns an array of InstructorEvent objects.
.DESCRIPTION
    Long description
.EXAMPLE
    Import-ExcelSched -Path '.\CVAH\CVAH 20-04.xlsx' -Course "CVAH" -Class "20-04"
.EXAMPLE
    $schedules = @(
        [PSCustomObject]@{
            path   = '.\Schedules\CVAH\CVAH 20-04.xlsx'
            course = "CVAH"
            class  = "20-04"
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CVAH\CVAH 20-05.xlsx'
            course = "CVAH"
            class  = "20-05"
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CVAH\CVAH 20-06.xlsx'
            course = "CVAH"
            class  = "20-06"
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CVAH\CVAH 20-08.xlsx'
            course = "CVAH"
            class  = "20-08"
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CWO\CWO 20-06 Schedule.xlsx'
            course = "CWO"
            class  = "20-06"
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CWO\CWO 20-08 Schedule.xlsx'
            course = "CWO"
            class  = "20-08"
        }
     )
    $schedules | Import-ExcelSched
#>
function Import-ExcelSched
{
    [CmdletBinding()]
    [OutputType([InstructorEvent])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]
        $Path,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]
        $Course,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]
        $Class
    ) # Param
    Begin
    {
        $objExcel = New-Object -ComObject Excel.application
    } # Begin
    Process
    {
        If (!(Test-Path $Path)) {
            Write-Error -Category ObjectNotFound -Message "$Path does not exist!"
            Return $null
        } #if Test-Path
        Write-Progress -Activity "Importing Schedule $Path" -CurrentOperation "Creating Excel Objects" -PercentComplete -1
        Try {
            $xlsfile = Get-Item -Path $Path 
            $objWorkbook = $objExcel.Workbooks.Open($xlsfile.FullName)
            $objExcel.Visible = $false
            $objWorksheet = $objWorkbook.Sheets.item(1)
        } #Try create com objects
        Catch {
            Write-Error -Category OpenError -Message "Can't create required Excel objects"
            Return $null
        } #catch open errors
        $progress = 5
        Write-Progress -Activity "Importing Schedule" -CurrentOperation "Parsing Days" -PercentComplete $progress
        $Day = $objWorksheet.Cells.Find('Day:')
        If (!$Day) {
            Write-Error -Category InvalidType -Message "$Path is not properly formatted. Can't find Day:"
            Return $null
        } #if can't find "Day:"
        $BeginAddress = $Day.Address(0,0,1,1)
        $Address = $BeginAddress
        while ($true)
        {
            $Date = $objWorksheet.Cells.Cells($Day.Row, $day.Column + 2).text
            Try {$Date = Get-Date -Date $Date}
            Catch { Write-Error -Category InvalidData -Message "Can't convert $Date to date ($Address)"; BREAK} # Can't convert date probably at end of sheet on template section.
            $DayHT = @{Day = $Date.Day; Month = $Date.Month; Year = $Date.Year}
            Write-Progress -Activity "Importing Schedule $Path" -CurrentOperation "Processing $($Date.ToString('d-MMM-yy'))" -PercentComplete $progress
            $NextDay = $objWorksheet.Cells.FindNext($Day)
            $NextAddress = $NextDay.Address(0,0,1,1)
            If ($NextAddress -eq $BeginAddress) { BREAK }  #End of schedule (Relies on Templates at bottom of worksheet)
            foreach ($row in ($Day.Row + 1)..($NextDay.Row - 1)) {
                If ($objWorksheet.Cells.Cells($row, 2).Text -in "", "X") {Continue} #blank row or maint day
                If ($objWorksheet.Cells.Cells($row, 1).Text -in "Rm", "Location") {Continue} #header row
                $startHour, $startMin = $objWorksheet.Cells.Cells($row, 2).Text -split ":"
                $endHour, $endMin     = $objWorksheet.Cells.Cells($row, 3).Text -split ":"

                $Eventht            = @{}
                $Eventht.room       = $objWorksheet.Cells.Cells($row, 1).Text
                Try {
                    $Eventht.start      = Get-Date @DayHT -Hour $startHour -Minute $startMin -Second 0
                    $Eventht.end        = Get-Date @DayHT -Hour $endHour   -Minute $endMin   -Second 0
                }
                Catch {
                    Write-Error -Category InvalidData -Message "Can't get times ($Address)"
                    Continue
                }
                $Eventht.lesson     = ($objWorksheet.Cells.Cells($row, 5).Text + " / " + $objWorksheet.Cells.Cells($row, 6).Text) -replace "\n"," - "
                $Eventht.class      = $class
                $Eventht.course     = $Course
                $Eventht.Role       = "Primary"
                $Eventht.Instructor = ($objWorksheet.Cells.Cells($row, 8).Text).Trim()
                If ($Eventht.Instructor -ne "") {New-SchedEvent @Eventht} # Return object from function
                $Eventht.Role       = "Secondary"
                $Eventht.Instructor = ($objWorksheet.Cells.Cells($row, 13).Text).Trim()
                If ($Eventht.Instructor -ne "") {New-SchedEvent @Eventht} # Return object from function
                foreach ($SupportInstructor in ($objWorksheet.Cells.Cells($Row,12).Text -split "[,]|[\n]")) {
                    $Eventht.Role       = "Support"
                    $Eventht.Instructor = $SupportInstructor.Trim()
                    If ($Eventht.Instructor -ne "") {New-SchedEvent @Eventht} # Return object from function
                } #foreach support instructor
            }            
            $Day = $NextDay
            $Address = $Day.Address(0,0,1,1)
            $progress += 2
            If ($Address -eq $BeginAddress) { BREAK } #End of schedule   
        }# Main while loop for every day in the schedule.
        #cleaning up
        Write-Progress -Activity "Importing Schedule $Path" -CurrentOperation "Cleaning Up" -PercentComplete 99
        $objWorksheet = $null
        $objWorkbook.Close($false)
        $objWorkbook = $null
       
    } #process
    End
    {
        $objExcel.Quit()
        $objExcel = $null
        Get-Process -Name EXCEL | Where-Object {$_.MainWindowHandle -eq 0} | Stop-Process
    } # End
} # function Import-ExcelSched

function ImportTesting {
    $schedules = @(
        [PSCustomObject]@{
            path   = '.\Schedules\CVAH\CVAH 20-04.xlsx'
            course = "CVAH"
            class  = "20-04"
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CVAH\CVAH 20-05.xlsx'
            course = "CVAH"
            class  = "20-05"
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CVAH\CVAH 20-06.xlsx'
            course = "CVAH"
            class  = "20-06"
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CVAH\CVAH 20-08.xlsx'
            course = "CVAH"
            class  = "20-08"
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CWO\CWO 20-06 Schedule.xlsx'
            course = "CWO"
            class  = "20-06"
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CWO\CWO 20-08 Schedule.xlsx'
            course = "CWO"
            class  = "20-08"
        }
     )
         $events = $schedules | Import-ExcelSched
         $events | Out-GridView
         $events | Export-Csv .\events.csv -Force
}
function Measure-Events {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [InstructorEvent[]]$events,

        [ValidateSet("CYWeek","FYWeek")]
        [string]
        $WkGrouping = "FYWeek",

        [double]
        $UtilizationRate = .37
    ) # param block

    Begin {
        $TotalWeeks = ($events | 
            Sort-Object -Property $WkGrouping -Unique |
                Measure-Object).count

        $startdate = ($events | 
            Sort-Object -Property start -Unique | 
                Select-Object -First 1 -ExpandProperty start).ToString("d-MMM-yyyy")

        $enddate = ($events | 
            Sort-Object -Property start -Unique | 
                Select-Object -Last 1 -ExpandProperty start).ToString("d-MMM-yyyy")

        "Report Covering {0} - {1}" -f $startdate, $enddate
        "`nClasses Loaded: {0}" -f (($events |
                                        Sort-Object -Property "course", "class" -Unique |
                                            Select-Object -Property @{n="classstring";e={"{0}-{1}" -f $_.course, $_.class}} |
                                                Select-Object -ExpandProperty classstring) -join ", ")
        "Utilization Rate: {0:P2}" -f $UtilizationRate
        "Weekly: {0:N2} hrs / This Report: {1:N2} hrs" -f ($UtilizationRate * 40), ($UtilizationRate * (40 * $TotalWeeks))
        "`nWeekly Summary:"
        "-" * 100

    } # Begin
    process {
        $events |
            Sort-Object -Property $WkGrouping, Instructor |
                Group-Object -Property $WkGrouping, Instructor |
                    Select-Object -Property @{n=$WkGrouping;e={($_.Name -split ", ")[0]}},
                                            @{n="Instructor";e={($_.Name -split ", ")[1]}}, * |
                        Format-Table -GroupBy $WkGrouping -AutoSize -Property "Instructor",
                                            @{n="Primary_Hours"
                                                e={[double]($_.Group | Where-Object Role -eq "Primary" | Measure-Object -Sum duration).sum}
                                                f="N2"},
                                            @{n="Secondary_Hours"
                                                e={[double]($_.Group | Where-Object Role -eq "Secondary" | Measure-Object -Sum duration).sum}
                                                f="N2"},
                                            @{n="Support_Hours"
                                                e={[double]($_.Group | Where-Object Role -eq "Support" | Measure-Object -Sum duration).sum}
                                                f="N2"},
                                            @{n="Total_Hours"
                                                e={[double]($_.Group | Measure-Object -Sum duration).sum}
                                                f="N2"},
                                            @{n="Utilization"
                                                e={[double]($_.Group | Measure-Object -Sum duration).sum / (40 * $UtilizationRate)}
                                                f="P2"}
    } # Process

    End {

        "Totals {0} - {1}:" -f $startdate, $enddate
        "-" * 100
        $events |
            Sort-Object -Property Instructor |
                Group-Object -Property Instructor |
                Sort-Object -Property @{e={[double]($_.Group | Measure-Object -Sum duration).sum}} -descending | 
                    Format-Table -AutoSize -Property "Name",
                                            @{n="Primary_Hours"
                                                e={[double]($_.Group | Where-Object Role -eq "Primary" | Measure-Object -Sum duration).sum}
                                                f="N2"},
                                            @{n="Secondary_Hours"
                                                e={[double]($_.Group | Where-Object Role -eq "Secondary" | Measure-Object -Sum duration).sum}
                                                f="N2"},
                                            @{n="Support_Hours"
                                                e={[double]($_.Group | Where-Object Role -eq "Support" | Measure-Object -Sum duration).sum}
                                                f="N2"},
                                            @{n="Total_Hours"
                                                e={[double]($_.Group | Measure-Object -Sum duration).sum}
                                                f="N2"},
                                            @{n="Utilization"
                                                e={[double]($_.Group | Measure-Object -Sum duration).sum / (($TotalWeeks * 40) * $UtilizationRate)}
                                                f="P2"} 

    } # End
} # function Measure-Events

function MeasureTesting()
{
    [InstructorEvent[]]$events = Import-Csv .\events.csv
    Measure-Events -events $events
}

MeasureTesting | Out-File -FilePath .\Analysis.txt -Force

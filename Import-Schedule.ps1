# Define Event Class
class Event
{
   [string]   $Instructor
   [datetime] $start
   [datetime] $end
   [string]   $course
   [string]   $class
   [string]   $lesson
   [string]   $room

   [ValidateSet("Primary", "Secondary", "Support")]
   [string] $Role

   # Constructor
   Event ()
   {
       $this.Instructor = ""
   }
   # Constructor
   Event ([string]   $Instructor, 
          [datetime] $start,
          [datetime] $end,
          [string]   $class,
          [string]   $course,
          [string]   $room,
          [string]   $lesson,
          [string]   $role          
          )
   {
       If ($end -le $start){Write-Error -Category InvalidData -Message "End must be after start!"}
       If ($role -notin "Primary", "Secondary", "Support") {Write-Error -Category InvalidData -Message "Role must be Primary, Secondary or Support"}
       $this.Instructor = $Instructor
       $this.start      = $start
       $this.end        = $end
       $this.class      = $class
       $this.room       = $room
       $this.lesson     = $lesson
       $this.Role       = $role      
   }
}

<#
.Synopsis
   Imports an Excel Schedule and returns an array of Event objects.
.DESCRIPTION
   Long description
.EXAMPLE
   Import-ExcelSched -Path '.\CVAH\CVAH 20-04.xlsx' -Course "CVAH" -Class "20-04"
.EXAMPLE
   "20-02.xlsx", "20-04.xlsx" | Import-ExcelSched
#>
function Import-ExcelSched
{
    [CmdletBinding()]
    [OutputType([Event])]
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
    )
    Begin
    {
        $objExcel = New-Object -ComObject Excel.application
    }
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
                If ($Eventht.Instructor -ne "") {New-Object -TypeName event -Property $Eventht} # Return object from function
                $Eventht.Role       = "Secondary"
                $Eventht.Instructor = ($objWorksheet.Cells.Cells($row, 13).Text).Trim()
                If ($Eventht.Instructor -ne "") {New-Object -TypeName event -Property $Eventht} # Return object from function
                foreach ($SupportInstructor in ($objWorksheet.Cells.Cells($Row,12).Text -split "[,]|[\n]")) {
                    $Eventht.Role       = "Support"
                    $Eventht.Instructor = $SupportInstructor.Trim()
                    If ($Eventht.Instructor -ne "") {New-Object -TypeName event -Property $Eventht} # Return object from function
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
    }
}

function DevTesting {
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



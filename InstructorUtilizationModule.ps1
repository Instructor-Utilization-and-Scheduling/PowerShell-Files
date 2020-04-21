# Define InstructorEvent Class
class InstructorEvent
{
   [string]   $Instructor
   [datetime] $start
   [datetime] $end
   [datetime] $AsOf
   [string]   $course
   [string]   $class
   [string]   $lesson
   [string]   $room
   [string]   $CYWeek
   [string]   $FYWeek
   [double]   $duration
   
   [ValidateSet("Primary", "Secondary", "Support", "Secondary/Support")]
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

   [InstructorEvent] Copy()
   {
        return [InstructorEvent]::new($this.Instructor, 
                                    $this.start, 
                                    $this.end, 
                                    $this.AsOf, 
                                    $this.class, 
                                    $this.course, 
                                    $this.room, 
                                    $this.lesson, 
                                    $this.role)
   } # Copy
     
   # Constructor
   InstructorEvent ([string]   $Instructor, 
          [datetime] $start,
          [datetime] $end,
          [datetime] $AsOf,
          [string]   $class,
          [string]   $course,
          [string]   $room,
          [string]   $lesson,
          [string]   $role          
          ) # Constructor parameters
   {
       If ($end -le $start){Write-Error -Category InvalidData -Message "End must be after start!"}
       If ($role -notin "Primary", "Secondary", "Support", "Secondary/Support") {Write-Error -Category InvalidData -Message "Role must be Primary, Secondary or Support"}
       $this.Instructor = $Instructor
       $this.start      = $start
       $this.end        = $end
       $this.AsOf       = $AsOf
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

   } #Constructor Definition

   #Empty Constructor for automation
   InstructorEvent()
   {$this.Instructor = ""}
} #Class Defition

#This function is used to call the constructor so we can use a hashtable to create objects (splatting)
function New-InstructorEvent 
{
    param (
        [string]   $Instructor, 
        [datetime] $start,
        [datetime] $end,
        [datetime] $AsOf,
        [string]   $class,
        [string]   $course,
        [string]   $room,
        [string]   $lesson,
        [string]   $role          
          ) # param
    [InstructorEvent]::new($Instructor, $start, $end, $AsOf, $class, $course, $room, $lesson, $role)
    
} #function New-InstructorEvent

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
            AliasFile = "C:\Users\micha\Documents\InputData\NameAliases.csv" 
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CVAH\CVAH 20-05.xlsx'
            course = "CVAH"
            class  = "20-05"
            AliasFile = "C:\Users\micha\Documents\InputData\NameAliases.csv" 
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CVAH\CVAH 20-06.xlsx'
            course = "CVAH"
            class  = "20-06"
            AliasFile = "C:\Users\micha\Documents\InputData\NameAliases.csv" 
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CVAH\CVAH 20-08.xlsx'
            course = "CVAH"
            class  = "20-08"
            AliasFile = "C:\Users\micha\Documents\InputData\NameAliases.csv" 
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CWO\CWO 20-06 Schedule.xlsx'
            course = "CWO"
            class  = "20-06"
            AliasFile = "C:\Users\micha\Documents\InputData\NameAliases.csv" 
        },
        [PSCustomObject]@{
            path   = '.\Schedules\CWO\CWO 20-08 Schedule.xlsx'
            course = "CWO"
            class  = "20-08"
            AliasFile = "C:\Users\micha\Documents\InputData\NameAliases.csv" 
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
        # Path to excel file
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]
        $Path,

        # Course name of the schedule imported
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]
        $Course,

        # Class name of the schedule imported
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]
        $Class,

        # Path to csv file of any Instructor name aliases. This csv will have an alias and name column.
        # If the function discovers a name in the schedule that is an alias identified in this csv,
        # it will replace the alias with the actual name per the csv.
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $AliasFile,

        # checks instructor names against this array of instructor names. If not  in this array,
        # event is still created/returned but identified in the results file.
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [string[]]
        $InstructorWhiteList,

        # Date the schedule was publised.
        #Make this mandatory
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [datetime]
        $AsOfDate = (Get-Date)
    ) # Param

    Begin
    {
        # Excel ComObject used to import the excel workbook
        $objExcel = New-Object -ComObject Excel.application
        
        # File the function will write the results of the import to.
        $ResultsFile = "$($env:TEMP)\SchedImportResults.txt"
        # Remove results file from previous imports and initialize file.
        If (Test-Path -Path $ResultsFile){Remove-Item -Path $ResultsFile}
    } # Begin

    Process
    {        
        "`nResults from importing {0}:" -f $Path  |
            Out-File -FilePath $ResultsFile -Append
        
            # Checking to see if the AliasFile is a valid argument, then importing it.
        if ($AliasFile){
            if (Test-Path -Path $AliasFile){
                $alias = Import-Csv -Path $AliasFile | Sort-Object -Property Alias -Unique
            }
            else {
                "Alias file: $AliasFile is not valid." |
                    Out-File -FilePath $ResultsFile -Append
            }
        }

        # Checking to see if the schedule is a valid file. If not, we can't do anything so writing an error
        # and terminating function.
        If (!(Test-Path -Path $Path)) {
            Write-Error -Category ObjectNotFound -Message "$Path does not exist!"
            Return $null
        } #if Test-Path
        
        # Creating hashtable that will be used for splatting the Write-Progress command throughout.
        $Status = @{
            Activity         = "Importing Schedule $Path"
            CurrentOperation = "Creating Excel Objects"
            PercentComplete  = -1
        } # Status hashtable definition
        # Update progress bar
        Write-Progress @Status

        Try {
            # The open method requires the full name (not relative path), creating a file object to force the full path.
            $xlsfile = Get-Item -Path $Path 
            $objWorkbook = $objExcel.Workbooks.Open($xlsfile.FullName)
            # For some unknown reason the open method changes the Visible property to $true, so changing it back
            $objExcel.Visible = $false
            # Assumes the schedule is the first sheet in the workbook.
            $objWorksheet = $objWorkbook.Sheets.item(1)
        } #Try create com objects
        Catch {
            # If we can't create the required Excel objects, we can't do anything so generating an error and 
            # terminating the
            Write-Error -Category OpenError -Message "Can't create required Excel objects"
            Return $null
        } #catch open errors

        # Updating progress bar
        $status.PercentComplete = 5
        $Status.CurrentOperation = "Parsing Days"
        Write-Progress @Status

        # Looking for "Day:" on the worksheet as it is our reference to find the start date.
        $Day = $objWorksheet.Cells.Find('Day:')
        # if we can't find "Day:" anywhere, nothing to do. Generating error and terminating function.
        If (!$Day) {
            Write-Error -Category InvalidType -Message "$Path is not properly formatted. Can't find Day:"
            Return $null
        } #if can't find "Day:"

        # Stores the absolute reference to the first cell containing "Day:"
        $BeginAddress = $Day.Address(0,0,1,1)
        # Starting with the first address used for the while loop.
        $Address = $BeginAddress

        # Infinite loop will BREAK when encountered in syntax.
        $EventsCreated = 0
        $lastday       = $false
        while ($true)
        {
            # Finding the next "Day:" in the schedule so we know where the current day ends.
            $NextDay = $objWorksheet.Cells.FindNext($Day)
            $NextAddress = $NextDay.Address(0,0,1,1)

            # If the next "Day:" is the beginning then we know we are on the last day in the schedule.
            # In this case we need to approximate the number of row in the last day. We will conservatively use 10.
            If ($NextAddress -eq $BeginAddress) { 
                $NextDay = $objWorksheet.Cells.Cells($Day.Row + 10, $Day.Column)
                $NextAddress = $NextDay.Address(0,0,1,1)
                $lastday = $true
                Write-Verbose -Message "Last Day"
            }  # If on last day.
             
            # The actual date is offset 2 colums from "Date:". Trim will remove and leading/trailing spaces.
            $Date = ($objWorksheet.Cells.Cells($Day.Row, $day.Column + 2).text).trim()
            # Making sure the value can be converted to a datetime object. Moving to next if so...
            Try {$Date = Get-Date -Date $Date}
            Catch { "$Address : Can't convert `"$Date`" value to a datetime object" | 
                        Out-File -FilePath $ResultsFile -Append
                    $Day = $NextDay
                    $Address = $Day.Address(0,0,1,1)
                    if ($lastday) { BREAK }
                    else { CONTINUE }
            } # Catch
            
            # Creating a hashtable used for splatting the Get-Date function
            $DayHT = @{Day = $Date.Day; Month = $Date.Month; Year = $Date.Year}
            # Updating progress bar
            $status.CurrentOperation = "Processing $($Date.ToString('d-MMM-yy'))"
            if ($status.PercentComplete -le 98) {
                $status.PercentComplete += 2
                Write-Progress @Status
            }

            # Now we go through each row in the current day
            foreach ($row in ($Day.Row + 1)..($NextDay.Row - 1)) {
                If ($objWorksheet.Cells.Cells($row, 2).Text -in "", "X") { CONTINUE } #blank row or maint day
                If ($objWorksheet.Cells.Cells($row, 1).Text -in "Rm", "Location", "Room") { CONTINUE } #header row
                
                # Creating hashtable for splatting New-SchedEvent function.
                $Eventht = @{}
                $Eventht.AsOf = $AsOfDate

                # Getting the start and end time of the event.
                $startHour, $startMin = $objWorksheet.Cells.Cells($row, 2).Text -split ":"
                $endHour, $endMin     = $objWorksheet.Cells.Cells($row, 3).Text -split ":"
                Try {
                    $Eventht.start    = Get-Date @DayHT -Hour $startHour -Minute $startMin -Second 0
                    $Eventht.end      = Get-Date @DayHT -Hour $endHour   -Minute $endMin   -Second 0
                }
                Catch {
                    "$Address : Can't convert date/time `"{0}`" - `"{1}`"" -f $objWorksheet.Cells.Cells($row, 2).Text, 
                                                                      $objWorksheet.Cells.Cells($row, 3).Text  |
                        Out-File -FilePath $ResultsFile -Append
                    CONTINUE
                }
                
                $Eventht.room       = $objWorksheet.Cells.Cells($row, 1).Text
                $Eventht.lesson     = ($objWorksheet.Cells.Cells($row, 5).Text + " / " + $objWorksheet.Cells.Cells($row, 6).Text) -replace "\n"," - "
                $Eventht.class      = $class
                $Eventht.course     = $Course
                $Eventht.Role       = "Primary"
                
                $Primary = @(($objWorksheet.Cells.Cells($row, 8).Text -split "[/]|[,]").Trim())
                foreach ($Instructor in $Primary) {
                    if ($Instructor -eq "" -or $Instructor -eq " ") { CONTINUE }
                    $Eventht.Instructor = $Instructor
                    # Checking if schedule used an alias name, if so converting to real name
                    If ($Instructor -in ($alias).alias){
                        $Eventht.Instructor = $alias | 
                            Where-Object {$_.alias -eq $Eventht.Instructor} | 
                                Select-Object -ExpandProperty Name
                    } # if alias
                    if ($InstructorWhiteList -and $eventht.instructor -notin $InstructorWhiteList) {
                        $eventht.instructor + " is not a valid instructor name. " + $Address |
                            Out-File -FilePath $ResultsFile -Append -Force
                    }                    
                    New-InstructorEvent @Eventht # Return event object
                    $EventsCreated++
                } # foreach Primary Instructor

                # Working on Secondary instructor. We'll hold off on creating the event until we know the instructor
                # is not also in the support role.
                $Secondary = ($objWorksheet.Cells.Cells($row, 13).Text).Trim()
                #Checking to see if alias used
                If ($Secondary -in ($alias.alias)){
                    $Secondary = $alias | 
                        Where-Object {$_.alias -eq $Secondary} | 
                            Select-Object -ExpandProperty Name
                } # if alias
                
                # Getting a list of all the support instructors
                $MIR = @($objWorksheet.Cells.Cells($Row,12).Text -split "[,]|[\n]|[/]").trim() | 
                            Where-Object {$_ -ne ""}
                
                # if the Secondary is not also a support instructor, creating an event for Secondary
                If ($Secondary -notin $MIR -and $Secondary -ne "" -and $Secondary -ne "ISSO") {
                    $Eventht.role       = "Secondary"
                    $Eventht.instructor = $Secondary
                    if ($InstructorWhiteList -and $eventht.instructor -notin $InstructorWhiteList) {
                        $eventht.instructor + " is not a valid instructor name. " + $Address |
                            Out-File -FilePath $ResultsFile -Append -Force
                    }   
                    New-InstructorEvent @Eventht # Return object from function
                    $EventsCreated++
                } # If just secondary

                # Iterating through all the Support Instructors and creating an event for each
                foreach ($SupportInstructor in $MIR) {
                    If ($SupportInstructor -in ($alias).alias){
                        $SupportInstructor = $alias | 
                            Where-Object {$_.alias -eq $SupportInstructor} | 
                                Select-Object -ExpandProperty Name
                    } # If alias

                    # Checking if support instructor is also the secondary
                    If ($SupportInstructor -eq $Secondary){
                        $Eventht.role       = "Secondary/Support"
                        $Eventht.Instructor = $Secondary
                        if ($InstructorWhiteList -and $eventht.instructor -notin $InstructorWhiteList) {
                            $eventht.instructor + " is not a valid instructor name. " + $Address |
                                Out-File -FilePath $ResultsFile -Append -Force
                        }   
                        New-InstructorEvent @Eventht # Return object from function
                        $EventsCreated++
                        CONTINUE
                    } # if multi-role

                    # Checking for outliers and creating Support Instructor event
                    $Eventht.role       = "Support"
                    $Eventht.Instructor = $SupportInstructor
                    if ($SupportInstructor -like "*Evaluator*" -or $SupportInstructor -like "DOM*" -or $SupportInstructor -like "CCV*") {
                        CONTINUE 
                    } # if
                    if ($InstructorWhiteList -and $eventht.instructor -notin $InstructorWhiteList) {
                        $eventht.instructor + " is not a valid instructor name. " + $Address |
                            Out-File -FilePath $ResultsFile -Append -Force
                    }   
                    New-InstructorEvent @Eventht # Return object from function
                    $EventsCreated++
                } # foreach support instructor
            } # foreach row in the day

            # Moving to the next day in the loop           
            $Day = $NextDay
            $Address = $Day.Address(0,0,1,1)

            if ($lastday) { BREAK } #End of schedule   
        } # Main while loop for every day in the schedule.

        # cleaning up
        $Status.CurrentOperation = "Cleaning Up"
        $Status.PercentComplete = 99
        Write-Progress @Status
        $Address      = $null
        $NextAddress  = $null
        $BeginAddress = $null
        $objWorksheet = $null
        $objWorkbook.Close($false)
        $objWorkbook  = $null

        "`nEvents created: {0}`n{1}" -f $EventsCreated, ("-" * 60) |
            Out-File -FilePath $ResultsFile -Append
       
    } # process

    End
    {
        # more cleanup
        $objExcel.Quit()
        $objExcel = $null
        Get-Process -Name EXCEL | 
            Where-Object {$_.MainWindowHandle -eq 0} | 
                Stop-Process
        
        # open Import Results file
        Invoke-Item -Path $ResultsFile
    } # End
} # function Import-ExcelSched

<#
.Synopsis
    This function takes an array of InstructorEvent objects and creates a report.
.DESCRIPTION
    Long description
.EXAMPLE
    [InstructorEvent[]]$events = Import-Csv -Path C:\Users\micha\Documents\OutputData\events.csv
    $report = Measure-Events -events $events
#>
function Measure-Events 
{
    [CmdletBinding()]
    param (
        # array of events to report on. Filter the events as needed prior to calling.
        [Parameter(Mandatory=$true)]
        [InstructorEvent[]]$InstructorEvents,

        # Set how to group the summary
        [Parameter(Mandatory=$true)]
        [ValidateSet("Monthly","Quarterly")]
        [string]
        $Grouping,

        # Rate that instructors should be utilized in the classroom.
        [double]
        $UtilizationRate = .37,

        # Total hours available to work per year.
        [int]
        $AnnualCapacity = 1752

    ) # param block

    Process {
        
        # Make a deep copy since we will be modifying the objects
        $events = @($InstructorEvents | ForEach-Object {$_.copy()})

        # Finding the earliest event in the array
        [datetime]$firsteventdate = $events | 
            Sort-Object -Property start -Unique | 
                Select-Object -First 1 -ExpandProperty start

        # Finding the last event in the array
        [datetime]$lasteventdate = $events | 
            Sort-Object -Property start -Unique | 
                Select-Object -Last 1 -ExpandProperty start
        
        # Setting the events per the grouping.
        switch ($Grouping) {
            "Monthly"   {
                $events | Add-Member -MemberType ScriptProperty -Name "Grouping" -Value {$this.start.ToString("MMM-yyyy")}
                # Need to start the report at the beginning of the month and end at the end of the month
                $ReportStart = Get-Date -Year $firsteventdate.Year -Month $firsteventdate.Month -Day 1
                $ReportEnd   = Get-Date -Year $lasteventdate.Year  -Month $lasteventdate.Month  -Day 1
                $ReportEnd   = $ReportEnd.AddMonths(1).AddDays(-1)
              } # Monthly Grouping
            "Quarterly" {
                $events | Add-Member -MemberType ScriptProperty -Name "Grouping" -Value {
                                "{0}-Q{1}" -f ($this.start.year), ([math]::Ceiling($this.start.month / 3))}              
                # Need to start the report at the beginning of the quarter and end at the end of the quarter
                $ReportStart = Get-Date -Year $firsteventdate.Year -Month ([math]::ceiling($firsteventdate.Month / 3) * 3 - 2) -Day 1
                $ReportEnd   = Get-Date -Year $lasteventdate.Year  -Month ([math]::ceiling($lasteventdate.Month / 3) * 3 - 2)  -Day 1
                $ReportEnd   = $ReportEnd.AddMonths(3).AddDays(-1)
            } # Quarterly Grouping
        } #switch

        # Writing report header information
        "Report Covering {0:d} - {1:d}" -f $ReportStart, $ReportEnd
        "`nClasses Loaded: {0}" -f (($events |
                                        Sort-Object -Property "course", "class" -Unique |
                                            Select-Object -Property @{n="classstring";e={"{0}-{1}" -f $_.course, $_.class}} |
                                                Select-Object -ExpandProperty classstring) -join ", ")
        "`nClassroom Utilization Rate Used for Calculations: {0:P0}" -f $UtilizationRate
        "Annual Capacity: {0:N0} hours" -f $AnnualCapacity
        "Annual Classroom Utilization: {0:N0} hours" -f ($AnnualCapacity*$UtilizationRate)
        "`n{0}Summary:" -f $Grouping
        "-" * 100
        # Calculating report totals
        $totaldays = ($ReportEnd - $ReportStart).TotalDays + 1
        $DailyCapacity = $AnnualCapacity / 365
        $TotalCapacity = $totaldays * $DailyCapacity
        $totalutilization = $TotalCapacity * $UtilizationRate

        # Looping the group summary
        foreach ($gp in ($events | 
                    Sort-Object -Property start |
                        Select-Object -ExpandProperty Grouping -Unique)) {
            $GroupEvents = $events | 
                                Where-Object {$_.Grouping -eq $gp}

            $firstgroupdate = $GroupEvents | 
                    Sort-Object -Property start | 
                        Select-Object -ExpandProperty start -First 1

            switch ($Grouping) {
                "Monthly"   { 
                    # Need to start at the beginning of the month and end at the end of the month
                    $startgroupdate = Get-Date -Year $firstgroupdate.Year -Month $firstgroupdate.Month -Day 1
                    $endgroupdate   = $startgroupdate.AddMonths(1).AddDays(-1)
                } # Monthly Grouping
                "Quarterly" {
                    # Need to start at the beginning of the quarter and end at the end of the quarter
                    $startgroupdate = Get-Date -Year $firstgroupdate.Year -Month ([math]::ceiling($firstgroupdate.Month / 3) * 3 - 2) -Day 1
                    $endgroupdate   = $startgroupdate.AddMonths(3).AddDays(-1)
                } # Quarterly Grouping
            } # Switch
           
            # Calculating group totals
            $gpdays  = ($endgroupdate - $startgroupdate).TotalDays + 1
            $GpCapacity = $gpdays * $DailyCapacity
            $CRUtilization = $GpCapacity * $UtilizationRate
            "{0}" -f $gp
            "Days: {0:N0}`tCapacity: {1:N2} hours`tClassroom Utilization: {2:N2} hours" -f $gpdays, $GpCapacity, $CRUtilization
            $GroupEvents |
                Group-Object -Property Instructor |
                    Sort-Object -Property @{e={($_.Group | Measure-Object -Sum duration).sum}} -Descending |
                    Format-Table -AutoSize -Property @{n="Instructor";e={$_.Name}},
                        @{n="Primary_Hours"
                            e={[double]($_.Group | Where-Object Role -eq "Primary" | Measure-Object -Sum duration).sum}
                            f="N2"},
                        @{n="Secondary_Hours"
                            e={[double]($_.Group | Where-Object Role -eq "Secondary" | Measure-Object -Sum duration).sum}
                            f="N2"},
                        @{n="Support_Hours"
                            e={[double]($_.Group | Where-Object Role -eq "Support" | Measure-Object -Sum duration).sum}
                            f="N2"},
                        @{n="Secondary/Support_Hours"
                            e={[double]($_.Group | Where-Object Role -eq "Secondary/Support" | Measure-Object -Sum duration).sum}
                            f="N2"},
                        @{n="Total_Hours"
                            e={[double]($_.Group | Measure-Object -Sum duration).sum}
                            f="N2"},
                        @{n="CR Utilization Rate"
                            e={[double]($_.Group | Measure-Object -Sum duration).sum / ($GpCapacity * $UtilizationRate)}
                            f="P2"},
                        @{n="Capacity Rate"
                            e={[double]($_.Group | Measure-Object -Sum duration).sum / ($GpCapacity)}
                            f="P2"}
        } #foreach Group

        # Final total rollup
        "`nReport Rollup {0:d} - {1:d}" -f $ReportStart, $ReportEnd
        "Total Days: {0:N0}`tTotal Capacity: {1:N2} hours`tClassroom Utilization Hours: {2:N2}" -f $totaldays, $TotalCapacity, $totalutilization
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
                                            @{n="Secondary/Support_Hours"
                                                e={[double]($_.Group | Where-Object Role -eq "Secondary/Support" | Measure-Object -Sum duration).sum}
                                                f="N2"},
                                            @{n="Total_Hours"
                                                e={[double]($_.Group | Measure-Object -Sum duration).sum}
                                                f="N2"},
                                            @{n="CR Utilization Rate"
                                                e={[double]($_.Group | Measure-Object -Sum duration).sum / ($TotalCapacity * $UtilizationRate)}
                                                f="P2"}, 
                                            @{n="Capacity Rate"
                                                e={[double]($_.Group | Measure-Object -Sum duration).sum / ($TotalCapacity)}
                                                f="P2"} 

    } # Process
} # function Measure-Events

function Export-ICS
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [InstructorEvent]
        $InstructorEvent,

        [Parameter(Mandatory=$true)]
        [string]
        $Path
    )
    Begin {

        # file header information
        $Calendar = "BEGIN:VCALENDAR`r`nVERSION:2.0`r`nPRODID:-//INSTRUCTOR_UTILIZATION_SCRIPT//EN"

    } #Begin
    Process {
        $Subject = "{0} {1} `r`n {2} / {3} `r`n / {4}" -f $InstructorEvent.course,
                                                $InstructorEvent.class,
                                                $InstructorEvent.lesson,
                                                $InstructorEvent.Role,
                                                $InstructorEvent.AsOf.ToString("d")
        $DateStamp = $InstructorEvent.AsOf.ToUniversalTime().ToString("yyyyMMddTHHmmssZ")
        $start     = $InstructorEvent.start.ToUniversalTime().ToString("yyyyMMddTHHmmssZ")
        $end       = $InstructorEvent.end.ToUniversalTime().ToString("yyyyMMddTHHmmssZ")
        $Calendar += "`r`nBEGIN:VEVENT"
        $Calendar += "`r`nSUMMARY:$Subject"
        $Calendar += "`r`nUID:$([guid]::NewGuid())"
        $Calendar += "`r`nSEQUENCE:0"
        $Calendar += "`r`nSTATUS:CONFIRMED"
        $Calendar += "`r`nTRANSP:TRANSPARENT"
        $Calendar += "`r`nDTSTART:$start"
        $Calendar += "`r`nDTEND:$end"
        $Calendar += "`r`nDTSTAMP:$DateStamp"
        $Calendar += "`r`nCATEGORIES:$($InstructorEvent.Role)"
        $Calendar += "`r`nLOCATION:$($InstructorEvent.room)"
        $Calendar += "`r`nDESCRIPTION:$subject"
        $Calendar += "`r`nEND:VEVENT"
    } # Process
    End {
        $Calendar += "`r`nEND:VCALENDAR"
        Set-Content -Path $Path -Value $Calendar -Force
    } # End
} # Function export-ics

function Remove-OutlookEvent 
{
    [CmdletBinding()]
    param (
        # Full path of calendar folder to use. If doesn't exist, it will create it under the default folder.
        [Parameter(Mandatory=$true)]
        [string]
        $CalendarFolder,

        # Instructor Event to use for outlook calendar event
        [Parameter(Mandatory=$true, 
                    ParameterSetName="Instructor Event",
                    ValueFromPipeline=$true) ]
        [InstructorEvent]
        $InstructorEvent        
    )
    
    begin {
        # Determine if outlook was running prior to the function call. If not we'll close the application when done.
        if (Get-Process -name Outlook -ErrorAction SilentlyContinue){$OutlookRunning = $true}
        try {
            # Create the outlook application object
            $outlook = New-Object -ComObject Outlook.application                    
        }
        catch {
            Write-Error "Unable to create outlook objects"
            return 0
        }     
        # Adds the Outlook interop assembly
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
       
        # The next line adds a type for enumeration. This makes the code more readable.
        # For example without this you would need to understand all the enumeration values
        # for the different types like olAppointmentitem = 1...
        $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
        
        # We need this namespace to enumerate the outlook calendar folders
        $namespace = $outlook.GetNameSpace("MAPI")
        # Gets the default calendar folders
        $DefaultCalFolder = $namespace.GetDefaultFolder($olFolders::olFolderCalendar)
        # Gets all the folders underneath the default calendar folder
        $Calendars = @($DefaultCalFolder.folders) + $DefaultCalFolder  
        # Get the Calendar object associated with the CalendarFolder parameter
        $Calendar = $Calendars | Where-Object {$_.fullfolderpath -eq $CalendarFolder}
        # If the calendar doesn't exist, return error
        if (!$Calendar) { 
            Write-Error "Can't find Outlook calendar"
            return $null
        }      
        $NumOfDeletedItems = 0
    }    
    process {
        $Subject = "{0} / {1} / {2} / {3} [Current_As_Of: {4}]" -f $InstructorEvent.course,
                                                                    $InstructorEvent.class,
                                                                    $InstructorEvent.lesson,
                                                                    $InstructorEvent.Role,
                                                                    $InstructorEvent.AsOf.ToString("d")
        $Calendar.items |
            Where-Object {$_.Subject -eq $Subject} |
                ForEach-Object {$_.Delete();$NumOfDeletedItems++}
    }    
    end {
        # If Outlook was not running prior to the function call, quit the application
        if (!$OutlookRunning){ $outlook.quit() }    

        # Return how many items deleted
        $NumOfDeletedItems        
    }
} # function remove-outlookevent

function New-OutlookEvent 
{
    [CmdletBinding()]
    param (

        # Full path of calendar folder to use. If doesn't exist, it will create it under the default folder.
        [Parameter(Mandatory=$true)]
        [string]
        $CalendarFolder,

        # Instructor Event to use for outlook calendar event
        [Parameter(Mandatory=$true, 
                    ParameterSetName="Instructor Event",
                    ValueFromPipeline=$true) ]
        [InstructorEvent]
        $InstructorEvent,

        # Subject
        [Parameter(Mandatory=$true, ParameterSetName="Regular Event")]
        [string]
        $Subject,

        # Start
        [Parameter(Mandatory=$true, ParameterSetName="Regular Event")]
        [datetime]
        $start,

        # End
        [Parameter(Mandatory=$true, ParameterSetName="Regular Event")]
        [datetime]
        $end,

        # Reminder in minutes
        [int]
        $reminder = 15,

        # Body of event
        [Parameter(Mandatory=$true, ParameterSetName="Regular Event")]
        [string]
        $body,

        # Location of event
        [Parameter(Mandatory=$true, ParameterSetName="Regular Event")]
        [string]
        $location,

        # Category of event
        [Parameter(ParameterSetName="Regular Event")]
        [string]
        $Category       
    )
    
    begin {
        # Determine if outlook was running prior to the function call. If not we'll close the application when done.
        if (Get-Process -name Outlook -ErrorAction SilentlyContinue){$OutlookRunning = $true}
        try {
            # Create the outlook application object
            $outlook = New-Object -ComObject Outlook.application                    
        }
        catch {
            Write-Error "Unable to create outlook objects"
            return 0
        }     
        # Adds the Outlook interop assembly
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
       
        # The next 4 lines adds types for enumeration. This makes the code more readable.
        # For example without this you would need to understand all the enumeration values
        # for the different types like olAppointmentitem = 1...
        $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
        $olItems   = "Microsoft.Office.Interop.Outlook.olItemType" -as [type]
        $olClose   = "Microsoft.Office.Interop.Outlook.olInspectorClose" -as [type] 
        $olColors  = "Microsoft.Office.Interop.Outlook.olCategoryColor" -as [type]  
        
        # We need this namespace to enumerate the outlook calendar folders
        $namespace = $outlook.GetNameSpace("MAPI")
        # Gets the default calendar folders
        $DefaultCalFolder = $namespace.GetDefaultFolder($olFolders::olFolderCalendar)
        # Gets all the folders underneath the default calendar folder
        $Calendars = @($DefaultCalFolder.folders) + $DefaultCalFolder  
        # Get the Calendar object associated with the CalendarFolder parameter
        $Calendar = $Calendars | Where-Object {$_.fullfolderpath -eq $CalendarFolder}
        # If the calendar doesn't exist, create it.
        if (!$Calendar) { 
            #create calendar and add to array of calendars
            $Calendar = $DefaultCalFolder.folders.Add($CalendarFolder, $olFolders::olFolderCalendar)
        }
        # Get the categories already loaded
        $categories = $namespace.categories | Select-Object -ExpandProperty Name
        # Create categories if needed
        if ($categories -notcontains "Primary") {
            $namespace.categories.Add("Primary", $olColors::olCategoryColorRed) | Out-Null
        }
        if ($categories -notcontains "Secondary") {
            $namespace.categories.Add("Secondary", $olColors::olCategoryColorYellow) | Out-Null
        }
        if ($categories -notcontains "Secondary/Support") {
            $namespace.categories.Add("Secondary/Support", $olColors::olCategoryColorDarkYellow) | Out-Null
        }
        if ($categories -notcontains "Support") {
            $namespace.categories.Add("Support", $olColors::olCategoryColorGreen) | Out-Null
        }

        $TotalEventsCreated = 0
    } #Begin
    process {
        # Set values for instructor event
        if ($InstructorEvent) {
            $Subject = "{0} / {1} / {2} / {3} [Current_As_Of: {4}]" -f $InstructorEvent.course,
                                                                             $InstructorEvent.class,
                                                                             $InstructorEvent.lesson,
                                                                             $InstructorEvent.Role,
                                                                             $InstructorEvent.AsOf.ToString("d")
            $Location = $InstructorEvent.room
            $start    = $InstructorEvent.start
            $end      = $InstructorEvent.end
            $body     = "{0}`nEvent created using PowerShell script." -f $InstructorEvent.Role
            $category = $InstructorEvent.Role
        }
        # Create outlook schedule event object
        $appt = $Calendar.items.add($olItems::olAppointmentItem) 
        $appt.start      = $start
        $appt.end        = $end
        $appt.Subject    = $Subject
        $appt.Location   = $location
        $appt.categories = $category
        $appt.body       = $body
        $appt.close($olClose::olSave)  
        $TotalEventsCreated++     
    }    
    end {
        # If Outlook was not running prior to the function call, quit the application
        if (!$OutlookRunning){ $outlook.quit() }    
        $TotalEventsCreated    
    }
}
<# # Uncomment this when actual module .psm1 is created.
Export-ModuleMember -Function New-InstructorEvent,
                              New-OutlookEvent,
                              Remove-OutlookEvent,
                              Export-ICS,
                              Measure-Events,
                              Import-ExcelSched,
                              New-InstructorEvent #>



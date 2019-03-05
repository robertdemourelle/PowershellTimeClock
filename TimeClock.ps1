Function Time() 
{
    $currentHour = Get-Date -Format hh
    $currentHour = [int]$currentHour
    $currentMinute = Get-Date -Format mm
    $currentMinute = [int]$currentMinute
    $AMorPM = Get-Date -Format tt
    
    #Round to the closest 15 minute mark 
    if($currentMinute -le 07){
        $currentMinute = "00"
    } ElseIf($currentMinute -gt 07) {
        if($currentMinute -le 22){
            $currentMinute = "15"
        } ElseIf($currentMinute -gt 22) {
            if($currentMinute -le 37){
                $currentMinute = "30"
            } ElseIf($currentMinute -gt 37) {
                if($currentMinute -le 52){
                    $currentMinute = "45"
                } Else {
                    $currentMinute ="00"
                    if($currentHour -eq 12){
                        $currentHour = [string]"01"
                    }
                    Else{
                        $currentHour = $currentHour + 1
                        $currentHour = [string]$currentHour
                    }
                }
            }
        }
    }
    $currentMinute = [string]$currentMinute
    $currentHour = [string]$currentHour
    
    $currentTime = $currentHour+':'+$currentMinute + $AMorPM
    
    return $currentTime
}
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$frmTimeClock                    = New-Object system.Windows.Forms.Form
$frmTimeClock.ClientSize         = '400,231'
$frmTimeClock.text               = "Time Clock"
$frmTimeClock.TopMost            = $false

$btnClockIn                      = New-Object system.Windows.Forms.Button
$btnClockIn.text                 = "Clock In"
$btnClockIn.width                = 163
$btnClockIn.height               = 30
$btnClockIn.location             = New-Object System.Drawing.Point(208,16)
$btnClockIn.Font                 = 'Microsoft Sans Serif,10'

$btnLunchOut                     = New-Object system.Windows.Forms.Button
$btnLunchOut.text                = "Lunch Out"
$btnLunchOut.width               = 164
$btnLunchOut.height              = 30
$btnLunchOut.location            = New-Object System.Drawing.Point(208,56)
$btnLunchOut.Font                = 'Microsoft Sans Serif,10'

$btnLunchIn                      = New-Object system.Windows.Forms.Button
$btnLunchIn.text                 = "Lunch In"
$btnLunchIn.width                = 162
$btnLunchIn.height               = 30
$btnLunchIn.location             = New-Object System.Drawing.Point(208,96)
$btnLunchIn.Font                 = 'Microsoft Sans Serif,10'

$btnClockOut                     = New-Object system.Windows.Forms.Button
$btnClockOut.text                = "Clock Out"
$btnClockOut.width               = 162
$btnClockOut.height              = 30
$btnClockOut.location            = New-Object System.Drawing.Point(208,136)
$btnClockOut.Font                = 'Microsoft Sans Serif,10'

$btnNewWeek                      = New-Object system.Windows.Forms.Button
$btnNewWeek.text                 = "New Week"
$btnNewWeek.width                = 162
$btnNewWeek.height               = 30
$btnNewWeek.location             = New-Object System.Drawing.Point(208,176)
$btnNewWeek.Font                 = 'Microsoft Sans Serif,10'

$lblDate                         = New-Object system.Windows.Forms.Label
$lblDate.text                    = "Date:"
$lblDate.AutoSize                = $true
$lblDate.width                   = 25
$lblDate.height                  = 10
$lblDate.location                = New-Object System.Drawing.Point(21,12)
$lblDate.Font                    = 'Microsoft Sans Serif,10'

$lblCurrentDate                  = New-Object system.Windows.Forms.Label
$lblCurrentDate.text             = Get-Date
$lblCurrentDate.AutoSize         = $true
$lblCurrentDate.width            = 25
$lblCurrentDate.height           = 10
$lblCurrentDate.location         = New-Object System.Drawing.Point(21,36)
$lblCurrentDate.Font             = 'Microsoft Sans Serif,10'

$lblTime                         = New-Object system.Windows.Forms.Label
$lblTime.text                    = "Input Time:"
$lblTime.AutoSize                = $true
$lblTime.width                   = 25
$lblTime.height                  = 10
$lblTime.location                = New-Object System.Drawing.Point(21,76)
$lblTime.Font                    = 'Microsoft Sans Serif,10'

$TEMP = Time
$lblCurrentTime                  = New-Object system.Windows.Forms.Label
$lblCurrentTime.text             = Time
$lblCurrentTime.AutoSize         = $true
$lblCurrentTime.width            = 25
$lblCurrentTime.height           = 10
$lblCurrentTime.location         = New-Object System.Drawing.Point(21,100)
$lblCurrentTime.Font             = 'Microsoft Sans Serif,10'

$frmTimeClock.controls.AddRange(@($btnClockIn,$btnLunchOut,$btnLunchIn,$btnClockOut,$btnNewWeek,$lblDate,$lblCurrentDate,$lblTime,$lblCurrentTime))

#region gui events {
#endregion events }

#endregion GUI }


#Define button clicks
#***********************************************
$btnClockIn.Add_Click({ClockIn})
$btnLunchOut.Add_Click({LunchOut})
$btnLunchIn.Add_Click({LunchIn})
$btnClockOut.Add_Click({ClockOut})
$btnNewWeek.Add_Click({CreateNewTimesheet})

#***********************************************
$Date = Get-Date
Function TimeSheet($x)
{       
    $nameTimesheet = Get-ChildItem C:\Users\rdemour\Documents\TimeSheets\Current
    $Time = Time
    $excel_file_path = "C:\Users\rdemour\Documents\TimeSheets\Current\$($nameTimesheet)"
    ## Instantiate the COM object
    $excel = New-Object -ComObject Excel.Application
    #Open Excel
    $workbook = $Excel.Workbooks.Open($excel_file_path)
    $wksht = $Excel.WorkSheets.item(1)
      
    #Find which row current day is
    $row = 0
    $DayOfWeek = (Get-Date).DayOfWeek
    switch($DayOfWeek){
        Monday{
            $row = 3
        }
        Tuesday{
            $row = 4
        }
        Wednesday{
            $row = 5
        }
        Thursday{
            $row = 6
        }
        Friday{
            $row = 7
        }   
    }
    
    #Input time based on Clock In(1), Lunch Out(2), Lunch In(3), Clock Out(4)
    switch($x){
        1{
            $wksht.Cells.Item($row,2) = $Time
        }
        2{
            $wksht.Cells.Item($row,3) = $Time
        }
        3{
            $wksht.Cells.Item($row,4) = $Time
        }
        4{
            $wksht.Cells.Item($row,6) = $Time
        }
    }
    ##AutoFit Cells
    $usedRange = $wksht.UsedRange						
    $usedRange.EntireColumn.AutoFit() | Out-Null
    
    #Save and exit excel
    $workbook.Save()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    Stop-Process -Name EXCEL -Force
    Stop-Process -Id $PID
}

Function CreateNewTimesheet()
{
    #If 'Current' directory is not empty, move current spreadsheet to 'Old' directory and create new timesheet in 'Current'
    $nameTimesheet = ""
    $nameTimesheet = Get-ChildItem C:\Users\rdemour\Documents\TimeSheets\Current | Measure-Object
    if($nameTimesheet.count -ne 0){
        #Move current spreadsheet to 'Old' directory
        $currentSS = Get-ChildItem C:\Users\rdemour\Documents\TimeSheets\Current
        Move-Item -Path "C:\Users\rdemour\Documents\TimeSheets\Current\$($currentSS)" -Destination "C:\Users\rdemour\Documents\TimeSheets\Old"
        
        #Create new timesheet
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

        $title = 'New TimeSheet'
        $msg   = 'Enter dates of new spreadsheet:'

        $nameOfTimesheet = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
        
        $xl=New-Object -ComObject "Excel.Application" 
     
        $wb=$xl.Workbooks.Add()
        $ws=$wb.ActiveSheet
        $ws.Cells.Item(1,2) = 'In'
        $ws.Cells.Item(1,3) = 'Lunch'
        $ws.Cells.Item(1,4) = 'In'
        $ws.Cells.Item(1,5) = 'Transfer'
        $ws.Cells.Item(1,6) = 'Out'
        $ws.Cells.Item(2,1) = 'Sun'
        $ws.Cells.Item(3,1) = 'Mon'
        $ws.Cells.Item(4,1) = 'Tue'
        $ws.Cells.Item(5,1) = 'Wed'
        $ws.Cells.Item(6,1) = 'Thu'
        $ws.Cells.Item(7,1) = 'Fri'
        $ws.Cells.Item(8,1) = 'Sat'
        $wb.SaveAs('C:\Users\rdemour\Documents\TimeSheets\Current\' + $nameofTimesheet + '.xlsx.')
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
        Stop-Process -Name EXCEL -Force
    }
}


Function ClockIn() 
{
    TimeSheet(1)
}
Function LunchOut()
{
    TimeSheet(2)
}
Function LunchIn()
{
    TimeSheet(3)
}
Function ClockOut()
{
    TimeSheet(4)
}


[void]$frmTimeClock.ShowDialog()
# Get-OLCalendarItem -start "6/25/2019 12:30:22" -end "12/25/2019 12:30:22"

param (
[String]$mode="0"
)

#Requires -Version 3.0
Function Get-OLCalendarItem {

[CmdletBinding()]
Param (
$start = $(Get-Date) ,
$end   = $((Get-date).AddMonths(12))
)

# Lade Outlook
[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Outlook") | out-null
$Outlook = new-object -comobject outlook.application
# Lade Outlook Ordner
$OlFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
$Namespace = $Outlook.GetNameSpace("MAPI")
$Calendar = $namespace.GetDefaultFolder($olFolders::olFolderCalendar)

# Für jeden Termin in Outlook...
Write-Host "Schreibe Termine in CSV Datei..."
ForEach ($citem in ($Calendar.Items | sort start)) {

    If ($citem.start -ge $start -and $citem.start -LE $end) { 

        $Start_n = $(Get-Date -UFormat "%m/%d/%Y %R:%S" $Citem.Start) | Out-String
        $End_n = $(Get-Date -UFormat "%m/%d/%Y %R:%S" $Citem.End) | Out-String

        $StartDate,$StartTime = $Start_n.split(' ')
        $EndDate,$EndTime = $End_n.split(' ')
        $StartTime = $StartTime -replace "`n","" -replace "`r",""
        $EndTime = $EndTime -replace "`n","" -replace "`r",""
        $Location = $($Citem.Location -replace "`n","" -replace "`r","")

        $O = "$($Citem.Subject);$StartDate;$StartTime;$EndDate;$EndTime;$($citem.AllDayEvent);$($Citem.Body);$Location"
        $O = $O -replace "`n","" -replace "`r",""
    }
##

# Ausgabe in Terminal

$O | Out-File -FilePath .\events.csv -Append -Encoding 'utf8'


} #Ende von foreach
} #Ende von Funktion
 
Function Set-OLCalendarItem {

    # load the required .NET types
    Add-Type -AssemblyName 'Microsoft.Office.Interop.Outlook'
    
    # access Outlook object model
    $outlook = New-Object -ComObject outlook.application

    # connect to the appropriate location
    $namespace = $outlook.GetNameSpace('MAPI')
    $Calendar = [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar
    $folder = $namespace.getDefaultFolder($Calendar)
    $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
    $outlook = new-object -comobject outlook.application
    $folder = $namespace.getDefaultFolder($olFolders::olFolderCalendar)
    $Appointments = $folder.Items
    $Appointments.IncludeRecurrences = $true
    $Appointments.Sort("[Start]")

    $CSV = Import-Csv -Path .\events.csv -Delimiter ";" -encoding UTF8
    ForEach ($line in $CSV){
        $Subject = $($line.Subject)

        [datetime][String] $StartDate = $($line.StartDate)
        [datetime][String] $EndDate = $($line.EndDate)

        $AllDayEvent = $($line.AllDayEvent)
        $Description = $($line.Description)
        $Location = $($line.Location)

        $ol = New-Object -ComObject Outlook.Application
        $meeting = $ol.CreateItem('olAppointmentItem')
        $meeting.Subject = $Subject
        $meeting.Body = $Description
        $meeting.Location = $Location
        $meeting.ReminderSet = $true
        $meeting.Importance = 1
        $meeting.AllDayEvent = $AllDayEvent
        $meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
        $meeting.ReminderMinutesBeforeStart = 15
        $meeting.Start = $StartDate 
        $meeting.End = $EndDate

        $a = $StartDate
        $b = $StartDate.AddDays(1).AddSeconds(-1)

        $checkstart = $($Appointments | Where-object { $_.start -ge $a -AND $_.start -le $b -AND $_subject -ne $Subject } | Select-Object -Property Subject, Start | foreach {$_.Start} | Out-String)
        if (([string]::IsNullOrEmpty($checkstart))) {
                write-host "Event $Subject am $StartDate erstellt"
                $meeting.Send()
        }else{
                write-host "Termin $Subject am $StartDate ist bereits vorhanden!"
        }
    }
Remove-Item -path "events.csv"
}

if($mode -eq "1"){
    Get-OLCalendarItem
}
elseif($mode -eq "2"){
    Set-OLCalendarItem
}
else{
    Write-Host "Kein Parameter übergeben!"
}

Write-Host "Fertig!"
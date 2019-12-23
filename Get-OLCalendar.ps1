# Get-OLCalendarItem -start "6/25/2019 12:30:22" -end "12/25/2019 12:30:22"

#Requires -Version 3.0
Function Get-OLCalendarItem {

[CmdletBinding()]
Param (
$start = $(Get-Date) ,
$end   = $((Get-date).AddMonths(12))
)

Write-Verbose "Returning objects between: $($start.tostring()) and $($end.tostring())"
# Load Outlook interop and Outlook iteslf
[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Outlook") | out-null
$Outlook = new-object -comobject outlook.application
# Get OL default folders
$OlFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
$Namespace = $Outlook.GetNameSpace("MAPI")
$Calendar = $namespace.GetDefaultFolder($olFolders::olFolderCalendar)
Write-Verbose "There are $($calendar.items.count) items in the calender in total"

# Für jeden Termin in Outlook...
ForEach ($citem in ($Calendar.Items | sort start)) {
#Write-Verbose "Processing [$($citem.Subject)]  Starting: [$($Citem.Start)]"

If ($citem.start -ge $start -and $citem.start -LE $end) { 
#Write-Output $Citem

$Start_n = $(Get-Date -UFormat "%m/%d/%Y %R:%S" $Citem.Start) | Out-String
$End_n = $(Get-Date -UFormat "%m/%d/%Y %R:%S" $Citem.End) | Out-String

$StartDate,$StartTime = $Start_n.split(' ')
$EndDate,$EndTime = $End_n.split(' ')
$StartTime = $StartTime -replace "`n","" -replace "`r",""
$EndTime = $EndTime -replace "`n","" -replace "`r",""

$O = "$($Citem.Subject),$StartDate,$StartTime,$EndDate,$EndTime,$($citem.AllDayEvent),$($Citem.Body),$($Citem.Location)"

}
##

$CSVHeader = "Subject,Start Date,Start Time,End Date,End Time,All Day Event,Description,Location"

# Ausgabe in Terminal
Write-output $O
$O | Out-File -FilePath .\events.csv -Append

} #Ende von foreach
} #Ende von Funktion

Set-Alias GCALI Get-OLCalendarItem 


Get-OLCalendarItem

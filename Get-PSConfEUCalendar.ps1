function Create-ICS
{
    [CmdletBinding()]
    Param
    (
            [string]$StartTime = (Get-Date -Format "yyyyMMddTHHmmssZ"),

        [Parameter(Mandatory = $true)]
            [string]$EndTime,

        [Parameter(Mandatory = $true)]
            [string]$Location,

        [Parameter(Mandatory = $true)]
            [string]$Summary,

        [Parameter(Mandatory = $true)]
            [string]$Description,

            [string]$Filepath = "File.ics",

            [string]$Timezone = "(GMT+01.00) Amsterdam / Berlin / Bern / Rome / Stockholm / Vienna"
    )

"BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//PowerShell/handcal//NONSGML v1.0//EN
BEGIN:VTIMEZONE
TZID:Romance Standard Time
BEGIN:STANDARD
DTSTART:16011028T030000
RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=10
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:16010325T020000
RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=3
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
END:DAYLIGHT" | Out-File -FilePath $Filepath -Append -Encoding default

"BEGIN:VEVENT
UID:201704170T172345Z-AF23B2@psconf.eu
DTSTAMP:201704170T172345Z
DTSTART;TZID="Romance Standard Time":20170505T151500
DTEND;TZID="Romance Standard Time":20170505T161500
SEQUENCE:1
SUMMARY:ChatOps with PowerShell - Matthew Hodgkins
LOCATION:12+14
DESCRIPTION:Written a ton of awesome PowerShell scripts but find it hard getting your end users or your support people to use them? Let's solve this with ChatOps. We will have a quick introduction to ChatOps and then dive into installing a Hubot (a chat bot) on Windows. After learning the basics of Hubot, we will run through some code examples so you can easily bring your scripts to users through chat. Finally we will touch on what security and authentication methods you can use when getting into ChatOps with PowerShell.
BEGIN:VALARM
TRIGGER:-PT15M
ACTION:DISPLAY
DESCRIPTION:Reminder
END:VALARM
END:VEVENT" | Out-File -FilePath $Filepath -Append -Encoding default


"END:VCALENDAR" | Out-File -FilePath $Filepath -Append -Encoding default


}

### The Duration for the Event in Minutes
$duration_minutes=120

### Location for the Event
$location="Somewhere"

### Summary of the event
$summary="TEST-EVENT"

### Description
$description="JUST A TEST EVENT!"

### Filepath for the Output
$file="C:\test.ics"

### Create Startdate as formated String
$start_datum=(Get-Date -Format "yyyyMMddTHHmmssZ")

### Create Enddate as formated String
$end_datum=[datetime]((get-date).AddMinutes($($duration_minutes)))
$end_datum=(Get-Date $end_datum -Format "yyyyMMddTHHmmssZ")

### Timezone for the Event
$time_zone = "(GMT+01.00) Amsterdam / Berlin / Bern / Rome / Stockholm / Vienna"

Create-ICS -Filepath $file -StartTime $start_datum -EndTime $end_datum -Location "$($location)" -Summary "$($summary)" -Description "$($description)" -Timezone $time_zone
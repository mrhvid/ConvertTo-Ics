<#
.Synopsis
   Convert PSConf.EU 2017 Agenda to .ics format
.DESCRIPTION
   This module was created for the http://www.powertheshell.com/agendacompetition/ 
   It pulls down the json file at powershell.love using the oneliner 
   (iwr powershell.love -UseB).Content.SubString(1) | ConvertFrom-Json | %{$_}  
   and converts the output to a .ics file that can be used by e.g. Outlook. 
.EXAMPLE
   (iwr powershell.love -UseB).Content.SubString(1) | ConvertFrom-Json | %{$_} | ConvertTo-Ics | Set-Content -Path Test.ics -Encoding Default

   Convert full Agenda to .ics
.EXAMPLE
   (iwr powershell.beer -UseB).Content.SubString(1) | ConvertFrom-Json | %{$_} | ogv -PassThru | ConvertTo-Ics | Set-Content -Path Test.ics -Encoding Default

   Display agenda in Out-GridView and convert selected events to Test.ics
.EXAMPLE
   (iwr powershell.love -UseB).Content.SubString(1) | ConvertFrom-Json | %{$_} | ogv -PassThru | ConvertTo-Ics -Reminder 12 | Set-Content -Path Test.ics -Encoding Default

   Same as previous example. But with 12 min reminder enabled on all events.
.NOTES
   This function was written in a hurry and dosn't comply 100% with https://icalendar.org/validator.html

   The GitHub repository is located here https://github.com/mrhvid/PSConfEU_ConvertTo-Ics
   Ideeas are apreciated. (I intend to update the module to be more general after the conference)

   - Jonas Sommer Nielsen

#>
function ConvertTo-Ics 
{
    [CmdletBinding(PositionalBinding=$false,
                  ConfirmImpact='Low')]
    Param
    (
        # Start Time of event
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("StartTime")]
        [datetime]
        $DTStart,

        # End Time of event
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("EndTime")]
        [datetime]
        $DTEnd,

        # Summary / subject of event
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("Title")]
        [string]
        $Summary,

        # Location of event
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("Room","Track")]
        [string]
        $Location,

        # Speakers
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("Speaker")]
        [string]
        $SpeakerList,

        # Description of event
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("abstract")]
        [string]
        $Description,

        # Reminder Number of minutes
        [Parameter(Mandatory=$false)]
        [int]
        $Reminder
    )

    Begin
    {
@"
BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//PowerShell/handcal//NONSGML v1.0//EN
"@ | Write-Output
    }
    Process
    {
        function Add-LineFold ([String]$Text) {
            # Simpel implementation to comply with https://icalendar.org/iCalendar-RFC-5545/3-1-content-lines.html
            $x = 60
            while($x -lt $text.Length) {
                $text = $text.Insert($x, "@@`n ")
                $x = $x + 60
            }
            $text
        }

        $DTStart = "{0:yyyyMMddTHHmmss}" -f $DTStart
        $DTEnd = "{0:yyyyMMddTHHmmss}" -f $DTEnd
        $Summary = Add-LineFold -Text "SUMMARY:$Summary - $SpeakerList"
        $SpeakerList = Add-LineFold -Text $SpeakerList
        $Location = Add-LineFold -Text "LOCATION:$Location"
        $Description = $Description.Replace('`n', '\n')  # Not working https://stackoverflow.com/questions/666929/encoding-newlines-in-ical-files
        $Description = Add-LineFold -Text "DESCRIPTION:$Description"
        if($Reminder) {
            $Alarm = '
BEGIN:VALARM
TRIGGER:-PT{0}M
ACTION:DISPLAY
DESCRIPTION:Reminder
END:VALARM' -f $Reminder
        }

@"
BEGIN:VEVENT
UID:201704170T172345Z-AF23B2@psconf.eu
DTSTAMP:201704170T172345
DTSTART:$DTStart
DTEND:$DTEnd
SEQUENCE:1
$Summary
$Location
$Description $Alarm
END:VEVENT
"@ | Write-Output


    }
    End
    {
        "END:VCALENDAR" | Write-Output
    }


}

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
   (iwr powershell.love -UseB).Content.SubString(1) | ConvertFrom-Json | %{$_} | ogv -PassThru | ConvertTo-Ics | Set-Content -Path Test.ics -Encoding Default

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
        [Alias("Room")]
        [string]
        $Location,

        # Description of event
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]       
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

        $DTStart = "{0:yyyyMMddTHHmmss}" -f $DTStart
        $DTEnd = "{0:yyyyMMddTHHmmss}" -f $DTEnd
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
SUMMARY:$Summary
LOCATION:$Location
DESCRIPTION:$Description $Alarm
END:VEVENT
"@ | Write-Output


    }
    End
    {
        "END:VCALENDAR" | Write-Output
    }


}

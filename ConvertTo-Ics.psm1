<#
.Synopsis
   Convert PSConf.EU 2017 Agenda to .ics format
.DESCRIPTION
   This module was created for the http://www.powertheshell.com/agendacompetition/ 
   It pulls down the json file at powershell.love using the oneliner 
   (iwr powershell.love -UseB).Content.SubString(1) | ConvertFrom-Json | %{$_}  
   and converts the output to a .ics file that can be used by e.g. Outlook. 

.EXAMPLE
    $(IRM powershell.beer -UseB) | ConvertTo-ics

    Convert the .json file and output to terminal
.EXAMPLE
    $(IRM powershell.beer -UseB) | ConvertTo-ics -Reminder 12 

    Convert the .json file and output to terminal add a 12 min reminder to all events. 
.EXAMPLE
    $(IRM powershell.beer -UseB) | ConvertTo-ics | Set-Content -Path Test.ics -Encoding Default

    Convert the .json file and save to Test.ics
.EXAMPLE
    $(IRM powershell.beer -UseB) | ogv -PassThru | ConvertTo-ics

    Use Out-GridView to filter what events to convert
.EXAMPLE
    $(IRM powershell.beer -UseB) | ogv -PassThru | ConvertTo-ics | Set-Content -Path Test.ics -Encoding Default

    Use Out-GridView to filter what events to convert save to Test.ics
.EXAMPLE
    $(IRM powershell.beer -UseB) | Where Category -Like '*pester*' | ConvertTo-ics

    Filter by Category all containing "pester" and convert
.NOTES
   This function was written in a hurry and dosn't comply 100% with https://icalendar.org/validator.html

   The GitHub repository is located here https://github.com/mrhvid/ConvertTo-Ics
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
        [Alias("StartTime", 'Starts')]
        [datetime]
        $DTStart,

        # End Time of event
        [Parameter(ParameterSetName='EndTime',
                   Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("EndTime",'Ends')]
        [datetime]
        $DTEnd,

        # Duration (event duration in minutes)
        [Parameter(ParameterSetName='Duration',
                   Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [int]
        $Duration,

        # Summary / subject of event
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias("Title",'Name')]
        [string]
        $Summary,

        # Location of event
        [Parameter(Mandatory=$false,
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

        # UID of event - Global uniqui identifier see: 5.3.  UID Property in https://tools.ietf.org/rfc/rfc7986.txt (This is a bad test implementation)
        [Parameter(Mandatory=$false,
                    ValueFromPipelineByPropertyName=$true)]
        [Alias("ID")]
        [string]
        $UID = $null,

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
                $text = $text.Insert($x, "`n ")
                $x = $x + 60
            }
            $text
        }

        if(!$PSBoundParameters.ContainsKey('UID')) {
            $NonAlfabeticChars = '\W'
            # UID of event - Global uniqui identifier see: 5.3.  UID Property in https://tools.ietf.org/rfc/rfc7986.txt (This is a bad test implementation)
            $UID = Add-LineFold -Text (($Summary -Replace($NonAlfabeticChars,'')) + '@psconf.eu')
        }

        if($Duration) {
            $DTEnd = $DTStart.AddMinutes($Duration)  
        }
        $Start = "{0:yyyyMMddTHHmmss}" -f $DTStart
        $End = "{0:yyyyMMddTHHmmss}" -f $DTEnd
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
UID:$UID
DTSTAMP:201704170T172345
DTSTART:$Start
DTEND:$End
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

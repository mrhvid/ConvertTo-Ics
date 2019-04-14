# PSConfEU_ConvertTo-Ics

Inspired by [PSConf EU 2017 Agenda Competition](http://www.powertheshell.com/agendacompetition/)

## Installation
`Find-Module ConvertTo-Ics | Install-Module`

(Works out of box with PowerShell 5.0 e.g. Windows 10)
You can find the module on the  [PowerShell Gallery](https://www.powershellgallery.com/packages/ConvertTo-Ics)


## Examples 

### Oneliner Install module, save all tracks to desktop (2019)
`inmo ConvertTo-Ics -Sc CurrentUser -Force;$(irm powershell.fun -UseB)| Where Name -NE "" |Group Track|%{$_.group|ConvertTo-Ics|Set-Content "$home\Desktop\PSConf$($_.name).ics" -Encoding Default};explorer "/select,$home\Desktop\PSConf$($_.name).ics"`


## Last years version. I haven't tested with 2019 data. 
### Output selected events (select and click OK bottom right) 
`PS C:\temp> (iwr powershell.love -UseB).Content.SubString(1) | ConvertFrom-Json | %{$_} | ogv -PassThru | ConvertTo-Ics | Set-Content -Path c:\temp\FirstTest.ics -Encoding Default`

### Output first 3 events
`PS C:\temp> (iwr powershell.love -UseB).Content.SubString(1) | ConvertFrom-Json | %{$_} | select -First 3 | ConvertTo-Ics | Set-Content -Path c:\temp\SecondTest.ics -Encoding Default`

### Output All events 13 min reminder
`PS C:\temp> (iwr powershell.love -UseB).Content.SubString(1) | ConvertFrom-Json | %{$_} | ConvertTo-Ics -Reminder 13 | Set-Content -Path c:\temp\ThirdTest.ics -Encoding Default`

Outputs all events to ThirdTest.ics

It's testing and not 100% complient with https://icalendar.org/validator.html But Outlook seems to be ok with the .ics created. 

Ideas are more than welcome :) 

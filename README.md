# PSConfEU_ConvertTo-Ics

Inspired by http://www.powertheshell.com/agendacompetition/


PS C:\temp> (iwr powershell.love -UseB).Content.SubString(1) | ConvertFrom-Json | %{$_} | select -First 3 | ConvertTo-Ics | Set-Content -Path c:\temp\FirstTest.ics -Encoding UTF8


Unfortunatley the output doesn't seem to work at the moment according to Outlook and
https://icalendar.org/validator.html
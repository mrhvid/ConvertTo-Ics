# PSConfEU_ConvertTo-Ics
Test


PS C:\temp> (iwr powershell.love -UseB).Content.SubString(1) | ConvertFrom-Json | %{$_} | select -First 3 | ConvertTo-Ics | Set-Content -Path c:\temp\FirstTest.ics -Encoding UTF8

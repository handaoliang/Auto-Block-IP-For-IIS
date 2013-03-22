@echo off
setlocal
set PATH=c:\msys\1.0\bin;C:\WINDOWS\system32;%PATH%

sh.exe -c "tail.exe -n 200000 /C/inetpub/logs/LogFiles/W3SVC1/u_ex$(echo """" | awk '{now = systime()-8*3600; print strftime(""%%y%%m%%d"", now)}').log | grep -i '/member/register' | awk '{now=systime()-8*3600-60*5; now=strftime(""%%H:%%M:%%S"", now); print($2>now?$9"" ""$8:"""")}' | grep '.' | sort | uniq -c | sort -rn -k1 | head -n30 | CScript.exe //nologo /D/iapps/AutoBlockIP/AutoBlockIP.vbs"

endlocal
@echo on

@echo off
setlocal
set PATH=c:\msys\1.0\bin;C:\WINDOWS\system32;%PATH%
rem sh.exe -c "echo """" | awk '{now=systime()-8*3600-60*5; now=strftime(""%%H:%%M:%%S"", now); print(now)}'"
rem sh.exe -c "tail.exe -n 200000 /D/u_ex121129.log |grep -i '/member/register'| awk '{now=systime()-8*3600-3600; now=strftime(""%%H:%%M:%%S"", now); print($2>now?$9:"""")}' | grep '.' | sort | uniq -c | sort -rn -k1 | head -n30"
rem | CScript.exe //nologo AutoDenyIP.vbs

rem sh.exe -c "echo ""200 113.69.46.196"" | CScript.exe //nologo AutoDenyIP.vbs"

rem sh.exe -c "tail.exe -n 200000 /C/inetpub/logs/LogFiles/W3SVC1/u_ex$(echo """" | awk '{now = systime()-8*3600; print strftime(""%%y%%m%%d"", now)}').log|grep -i '/member/register'| awk '{now=systime()-8*3600-60*5; now=strftime(""%%H:%%M:%%S"", now); print($2>now?$9:"""")}' | grep '.' | sort | uniq -c | sort -rn -k1 | head -n30 | CScript.exe //nologo /D/iapps/AutoBlockIP/AutoBlockIP.vbs"

sh.exe -c "tail.exe -n 200000 /C/inetpub/logs/LogFiles/W3SVC1/u_ex$(echo """" | awk '{now = systime()-8*3600; print strftime(""%%y%%m%%d"", now)}').log|grep -i '/member/register'| awk '{now=systime()-8*3600-60*30; now=strftime(""%%H:%%M:%%S"", now); print($2>now?$9"" ""$8:"""")}' | grep '.' | sort | uniq -c | sort -rn -k1 | head -n30"

rem sh.exe -c "tail.exe -n 200000 /C/inetpub/logs/LogFiles/W3SVC1/u_ex$(echo """" | awk '{now = systime()-8*3600; print strftime(""%%y%%m%%d"", now)}').log|grep -i '/member/register'| awk '{now=systime()-8*3600-60*5; now=strftime(""%%H:%%M:%%S"", now); print($2>now?$9"" ""$8:"""")}' | grep '.' | sort | uniq -c | sort -rn -k1 | head -n30"
endlocal
@echo on

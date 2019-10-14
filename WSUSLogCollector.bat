@echo off

set yyyy=%date:~0,4%
set mm=%date:~5,2%
set dd=%date:~8,2%
 
set time2=%time: =0%
 
set hh=%time2:~0,2%
set mn=%time2:~3,2%
set ss=%time2:~6,2%
 
set _TEMPDIR=%UserProfile%\desktop\WSUSLogCollector-%ComputerName%-%yyyy%%mm%%dd%%hh%%mn%%ss%

md %_TEMPDIR%

wevtutil epl System %_TEMPDIR%\System.evtx
wevtutil epl Application %_TEMPDIR%\Application.evtx
wevtutil epl Security %_TEMPDIR%\Security.evtx
wevtutil epl Setup %_TEMPDIR%\Setup.evtx
wevtutil epl Microsoft-Windows-Bits-Client/Operational %_TEMPDIR%\Bits-Client_Operational.evtx

wmic qfe list /Format:Table > %_TEMPDIR%\QFE.log
msinfo32 /nfo %_TEMPDIR%\msinfo32.nfo
GPRESULT /H %_TEMPDIR%\GPReport.html

reg export "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Update Services" %_TEMPDIR%\registry_WSUS.txt

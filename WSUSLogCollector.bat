@echo off

set yyyy=%date:~0,4%
set mm=%date:~5,2%
set dd=%date:~8,2%
 
set time2=%time: =0%
 
set hh=%time2:~0,2%
set mn=%time2:~3,2%
set ss=%time2:~6,2%
 
set _TEMPDIR=%SystemDrive%\WSUSLogs-%ComputerName%-%yyyy%%mm%%dd%%hh%%mn%%ss%

mkdir %_TEMPDIR%

wevtutil epl System %_TEMPDIR%\System.evtx
wevtutil epl Application %_TEMPDIR%\Application.evtx
wevtutil epl Security %_TEMPDIR%\Security.evtx
wevtutil epl Setup %_TEMPDIR%\Setup.evtx
wevtutil epl Microsoft-Windows-Bits-Client/Operational %_TEMPDIR%\Bits-Client_Operational.evtx

msinfo32 /nfo %_TEMPDIR%\msinfo32.nfo
GPRESULT /H %_TEMPDIR%\GPReport.html
wmic qfe list /Format:Table > %_TEMPDIR%\QFE.log
powershell -command "Get-WindowsFeature" > %_TEMPDIR%\WindowsFeature.txt

bitsadmin /list /allusers /verbose > %_TEMPDIR%\bitsadmin.log
ipconfig /all > %_TEMPDIR%\ipconfig.txt
netsh winhttp show proxy > %_TEMPDIR%\winhttp.txt
copy %windir%\System32\drivers\etc\hosts %_TEMPDIR%\hosts.txt
certutil -store root > %_TEMPDIR%\certs.txt 2>&1

mkdir %_TEMPDIR%\WSUS
reg export "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Update Services" %_TEMPDIR%\WSUS\registry_WSUS.txt
copy "%ProgramFiles%\Update Services\LogFiles\*.log" %_TEMPDIR%\WSUS
copy "%ProgramFiles%\Update Services\LogFiles\*.old" %_TEMPDIR%\WSUS

set _WSUS_SCRIPT=%TEMP%\WSUSGetInfo.ps1

echo # WSUSGetInfo.ps1 > %_WSUS_SCRIPT%
echo [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") >> %_WSUS_SCRIPT%
echo $WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer() >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "- WSUS  WebServiceUrl, Name, Version, PortNumber, ServerName, UseSecureConnection, ServerProtocolVersion`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo $WSUS ^| Select-Object WebServiceUrl, Name, Version, PortNumber, ServerName, UseSecureConnection, ServerProtocolVersion >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "- WSUS.GetStatus()`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo $WSUS.GetStatus() >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "- WSUS.GetConfiguration()`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo $WSUS.GetConfiguration() >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "- WSUS.GetSubscription()`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo $WSUS.GetSubscription() >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "- WSUS.GetEmailNotificationConfiguration()`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo $WSUS.GetEmailNotificationConfiguration() >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "- WSUS.GetDownstreamServers()`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo $WSUS.GetDownstreamServers() >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "- WSUS.GetDatabaseConfiguration()`r`n" >> %_WSUS_SCRIPT%
echo Write-Host "=================================`r`n" >> %_WSUS_SCRIPT%
echo $WSUS.GetDatabaseConfiguration() >> %_WSUS_SCRIPT%

powershell -ExecutionPolicy Bypass -Command %_WSUS_SCRIPT% > %_TEMPDIR%\WSUS\WSUSinfo.log
del %_WSUS_SCRIPT%

set _WSUSContent_SCRIPT=%TEMP%\WSUSGetContentInfo.ps1

echo # WSUSGetContentInfo.ps1 > %_WSUSContent_SCRIPT%
echo [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") >> %_WSUSContent_SCRIPT%
echo $WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer() >> %_WSUSContent_SCRIPT%
echo Get-ChildItem ($WSUS.GetConfiguration()).LocalContentCachePath -Recurse >> %_WSUSContent_SCRIPT%

powershell -ExecutionPolicy Bypass -Command %_WSUSContent_SCRIPT%> %_TEMPDIR%\WSUS\WSUSContentinfo.log
del %_WSUSContent_SCRIPT%

mkdir %_TEMPDIR%\IIS
copy %SystemRoot%\System32\inetsrv\config\applicationHost.config %_TEMPDIR%\IIS\applicationHost.config
copy "%ProgramFiles%\Update Services\WebSevices\ApiRemoting30\Web.config" %_TEMPDIR%\IIS\ApiRemoting30_Web.config
copy "%ProgramFiles%\Update Services\WebSevices\ClientWebService\Web.config" %_TEMPDIR%\IIS\ClientWebService_Web.config
copy "%ProgramFiles%\Update Services\WebSevices\DssAuthWebService\Web.config" %_TEMPDIR%\IIS\DssAuthWebService_Web.config
copy "%ProgramFiles%\Update Services\WebSevices\ReportingWebService\Web.config" %_TEMPDIR%\IIS\ReportingWebService_Web.config
copy "%ProgramFiles%\Update Services\WebSevices\ServerSyncWebService\Web.config" %_TEMPDIR%\IIS\ServerSyncWebService_Web.config
copy "%ProgramFiles%\Update Services\WebSevices\SimpleAuthWebService\Web.config" %_TEMPDIR%\IIS\SimpleAuthWebService_Web.config
robocopy %SystemRoot%\System32\LogFiles\HTTPERR\ %_TEMPDIR%\IIS\ /MAXAGE:7
robocopy %SystemDrive%\inetpub\logs\LogFiles %_TEMPDIR%\IIS\ /MAXAGE:3 /s

mkdir %_TEMPDIR%\


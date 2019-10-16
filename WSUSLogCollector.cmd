@echo off

set yyyy=%date:~0,4%
set mm=%date:~5,2%
set dd=%date:~8,2%
 
set time2=%time: =0%
 
set hh=%time2:~0,2%
set mn=%time2:~3,2%
set ss=%time2:~6,2%
 
set _TEMPDIR=%SystemDrive%\WSUSLogs-%ComputerName%-%yyyy%%mm%%dd%%hh%%mn%%ss%

echo **********************************************************************************
echo - イベント ログを取得します。(1/7)
echo **********************************************************************************

mkdir %_TEMPDIR%

wevtutil epl System %_TEMPDIR%\System.evtx
wevtutil epl Application %_TEMPDIR%\Application.evtx
wevtutil epl Security %_TEMPDIR%\Security.evtx
wevtutil epl Setup %_TEMPDIR%\Setup.evtx
wevtutil epl Microsoft-Windows-Bits-Client/Operational %_TEMPDIR%\Bits-Client_Operational.evtx

echo **********************************************************************************
echo - システム関連情報を取得します。(2/7)
echo **********************************************************************************

msinfo32 /nfo %_TEMPDIR%\msinfo32.nfo
gpresult /H %_TEMPDIR%\gpresult.html
wmic qfe list /Format:Table > %_TEMPDIR%\QFE.log
powershell -command "Get-WindowsFeature" > %_TEMPDIR%\WindowsFeature.txt

echo **********************************************************************************
echo - ネットワークおよび証明書関連情報を取得します。(3/7)
echo **********************************************************************************

bitsadmin /list /allusers /verbose > %_TEMPDIR%\bitsadmin.log
ipconfig /all > %_TEMPDIR%\ipconfig.txt
netsh winhttp show proxy > %_TEMPDIR%\winhttp.txt
copy %windir%\System32\drivers\etc\hosts %_TEMPDIR%\hosts.txt
certutil -store root > %_TEMPDIR%\certs.txt 2>&1

echo **********************************************************************************
echo - WSUS 関連情報を取得します。(4/7)
echo **********************************************************************************

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

echo **********************************************************************************
echo - IIS 関連情報を取得します。(5/7)
echo **********************************************************************************

mkdir %_TEMPDIR%\IIS
copy %SystemRoot%\System32\inetsrv\config\applicationHost.config %_TEMPDIR%\IIS\applicationHost.config
copy "%ProgramFiles%\Update Services\WebServices\ApiRemoting30\Web.config" %_TEMPDIR%\IIS\ApiRemoting30_Web.config
copy "%ProgramFiles%\Update Services\WebServices\ClientWebService\Web.config" %_TEMPDIR%\IIS\ClientWebService_Web.config
copy "%ProgramFiles%\Update Services\WebServices\DssAuthWebService\Web.config" %_TEMPDIR%\IIS\DssAuthWebService_Web.config
copy "%ProgramFiles%\Update Services\WebServices\ReportingWebService\Web.config" %_TEMPDIR%\IIS\ReportingWebService_Web.config
copy "%ProgramFiles%\Update Services\WebServices\ServerSyncWebService\Web.config" %_TEMPDIR%\IIS\ServerSyncWebService_Web.config
copy "%ProgramFiles%\Update Services\WebServices\SimpleAuthWebService\Web.config" %_TEMPDIR%\IIS\SimpleAuthWebService_Web.config
robocopy %SystemRoot%\System32\LogFiles\HTTPERR\ %_TEMPDIR%\IIS\ /MAXAGE:7
robocopy %SystemDrive%\inetpub\logs\LogFiles %_TEMPDIR%\IIS\ /MAXAGE:3 /s

echo **********************************************************************************
echo - データベース 関連情報を取得します。(6/7)
echo **********************************************************************************

mkdir %_TEMPDIR%\Database
mkdir %_TEMPDIR%\Database\WID
copy %SystemRoot%\WID\Log\*ERROR*.log %_TEMPDIR%\Database\WID\
robocopy "%ProgramFiles%\Microsoft SQL Server" %_TEMPDIR%\Database\ *ERRORLOG* /s

echo **********************************************************************************
echo - 取得した情報を圧縮します。(7/7)
echo **********************************************************************************

call :ZIPLOGS_ALT "%_TEMPDIR%" "%_TEMPDIR%.zip"

if not exist "%_TEMPDIR%.zip" (
    echo.
    echo フォルダの圧縮に失敗しました。お手数おかけしますが %_TEMPDIR% を圧縮して弊社宛てに送付ください。
    echo.
) else (
    rd /s %_TEMPDIR%
    echo.
    echo 情報の取得が完了しました。%_TEMPDIR%.zip を弊社宛てに送付ください。
    echo.
)

:ZIPLOGS_ALT
setlocal
set _SOURCEFOLDER=%~1
set _ZIPFILEPATH=%~2

set _ZIP_SCRIPT=%TEMP%\ZipFile.ps1

echo # ZipFile.ps1 > %_ZIP_SCRIPT%

echo $zipHeader=[char]80 + [char]75 + [char]5 + [char]6 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 >> %_ZIP_SCRIPT%

echo $sourceFolder = "%_SOURCEFOLDER%" >> %_ZIP_SCRIPT%
echo $folderSize = Get-ChildItem $sourceFolder ^| Measure-Object -Property length -sum >> %_ZIP_SCRIPT%
echo $zipFilePath = "%_ZIPFILEPATH%" >> %_ZIP_SCRIPT%

echo Add-Content $zipFilePath -value $zipHeader >> %_ZIP_SCRIPT%
echo $explorerShell = New-Object -com 'Shell.Application' >> %_ZIP_SCRIPT%
echo $sendToZip = $explorerShell.Namespace($zipFilePath.ToString()).CopyHere($sourceFolder.ToString()) >> %_ZIP_SCRIPT%

echo $size = $folderSize.Sum * 0.05 >> %_ZIP_SCRIPT%
echo $timeout = 60 >> %_ZIP_SCRIPT%
echo $progress = "." >> %_ZIP_SCRIPT%
echo do >> %_ZIP_SCRIPT%
echo { >> %_ZIP_SCRIPT%
echo     $item = Get-Item $zipFilePath >> %_ZIP_SCRIPT%
echo     write-host $progress >> %_ZIP_SCRIPT%
echo     Start-Sleep 1 >> %_ZIP_SCRIPT%
echo     $progress = $progress + "." >> %_ZIP_SCRIPT%
echo     $timeout = $timeout - 1 >> %_ZIP_SCRIPT%
echo } while ($item.Length -lt $size -and $timeout -gt 0) >> %_ZIP_SCRIPT%

powershell -ExecutionPolicy Bypass -Command %_ZIP_SCRIPT%
del %_ZIP_SCRIPT%
endlocal
goto :EOF

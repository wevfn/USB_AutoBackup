
@echo off
chcp 936 >nul 2>nul
if "%~1"=="acp" xcopy "%~dp0*" "%usbp_dir%" /e /d /y /i /g /h /c&exit
if "%~1"=="bcp" xcopy "%usbp_dir%*" "%~dp0%~n0\" /e /d /y /i /g /h /c&exit
%1 start "" mshta vbscript:createobject("shell.application").shellexecute("""%~0""","::",,"runas",1)(window.close)&&exit
title USB BACKUP
set errorlevel=
set ert=if errorlevel 1 echo;^&pause^&exit
if "%~p0"=="\" goto way
copy \\n
%ert%
:way
findstr "acp" "%public%\USBbp.vbs" >nul 2>&1 &&set un=c&&echo; USB -^> %usbp_dir%
findstr "bcp" "%public%\USBbp.vbs" >nul 2>&1 &&set un=c&&echo; USB ^<- %usbp_dir%
echo;
echo;A=USB-^>PC
echo;B=USB^<-PC
echo;C=Off backup
echo;
set az=acp&set mde=-^^^>
choice /c ab%un%
if %errorlevel%==2 set az=bcp&set mde=^^^<-
if %errorlevel%==3 schtasks /delete /f /tn USB_Insert&schtasks /delete /f /tn USB_Scheduled&del "%public%\USBbp.vbs"&echo;&pause&exit
set "psd="(new-object -com 'shell.application').browseforfolder(0,'USB %mde% ?',0,0).self.path""
for /f "usebackq delims=" %%I in (`powershell %psd%`) do set "pc=%%I"
cd /d "%pc%"
%ert%
cd..
if not "%pc%"=="%cd%" set pc=%pc%\
if %az%==acp set pc=%pc%%~n0\
setx /m USBP_DIR "%pc%\"
%ert%
for /f "skip=1 delims=." %%a in ('wmic os get version') do (
if %%a gtr 6 set ent=Partition/Diagnostic -mo "*[System[(EventID='1006')]]"&&goto new
)
taskkill /f /t /im taskeng.exe >nul 2>nul
set ent=DriverFrameworks-UserMode/Operational -mo "*[System[(EventID='2101')]]"
:new
echo;
schtasks -create -tn "USB_Insert" -tr "%public%\USBbp.vbs" -rl highest -f -delay 0000:05 -sc onevent -ec Microsoft-Windows-%ent%
schtasks -create -tn "USB_Scheduled" -tr "%public%\USBbp.vbs" -rl highest /f /sc hourly /mo 2
%ert%
echo strPath = "%~n0.bat" >"%public%\USBbp.vbs"
(
echo Set objFSO = CreateObject("Scripting.FileSystemObject"^)
echo Set colDrives = objFSO.Drives
echo For Each objDrive in colDrives
echo If objDrive.IsReady Then
echo strDrive = objDrive.DriveLetter ^& ":"
echo If objFSO.FileExists(strDrive ^& "\" ^& strPath^) Then
echo WScript.CreateObject("WScript.Shell"^).Run """" ^& strDrive ^& "\" ^& strPath ^& """" ^& " %az%" ,0
echo End If
echo End If
echo Next
) >>"%public%\USBbp.vbs"
if %az%==acp set src=%~dp0&set tgt=%pc%
if %az%==bcp set src=%pc%&set tgt=%~dp0%~n0\
echo;
echo; %src% -^> %tgt%
xcopy "%src%*" "%tgt%" /e /d /y /i /g /h /c /w
%ert%
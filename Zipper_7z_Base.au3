#Include <File.au3>
#include <Date.au3>
#include <Array.au3>

;;;;;;;;;;;;;;;;;;;;;;;;;
;1. Create empty Zip Archive
;2. Create/Check Network Share
;3. Copy Files to Zip
;4. Create Protocol
;5. Copy Protocol toZip
;19.05.2017: Umstellung auf 7z-cmd Version
;19.07.2017: Kommentare ergänzt
;;;;;;;;;;;;;;;;;;;;;;;;;

;Variablendefinition
Dim $Zip,$sName,$sFirstString,$FreeDrive
Dim $sPCSPath
$LogFolder="D:\Backups\PCULogs"
$PCU_Log=$LogFolder&"\PCU"
$CTM_Log1=$LogFolder&"\CTM\CTM_CH1"
$CTM_Log2=$LogFolder&"\CTM\CTM_CH2"
$CTM_alarms=$LogFolder&"\CTM\alarms"
$PCS_Log=$LogFolder&"\PCS"
$Sprint_Log=$LogFolder&"\Sprint"
$Sprint_Log_lokal=$LogFolder&"\Sprint\lokal"
$Rel_Log=$LogFolder&"\RelReport"
$ProcessLog=$LogFolder&"\ProcessData"
$Archive_Path="D:\Backups\log_archive_" & _NowDate() & ".7z"
Global $t_time, $t_hours, $t_mins, $t_secs
Opt("TrayAutoPause", 0)
HotKeySet("{ESC}","Terminate")
TrayTip("Zipper.exe","Zipper running...",5)


;COM Object Error Handle
$oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")




If @error=1 Then
   Sleep(10000)
   Exit
EndIf

;Zeitstring für Archivname
$thandle=TimerInit()
If FileExists($Archive_Path) Then
   $aTime=StringSplit(_NowTime(5),":/")
   $Archive_Path="D:\Backups\log_archive_" & _NowDate() & "_" & $aTime[1] & "_" & $aTime[2] & "_" & $aTime[3] & ".7z"
EndIf
$dFindPath=@HomeDrive
Global $depth=2



;Create Collection Folder
If FileExists($LogFolder) Then DirRemove($LogFolder,1)
If Not DirCreate($LogFolder) Then ZipperLog("Couldn't create " & $LogFolder)
If Not DirCreate($PCU_Log) Then ZipperLog("Couldn't create " & $PCU_Log)
If Not DirCreate($CTM_Log1) Then ZipperLog("Couldn't create " & $CTM_Log1)
If Not DirCreate($CTM_Log2) Then ZipperLog("Couldn't create " & $CTM_Log2)
If Not DirCreate($CTM_alarms) Then ZipperLog("Couldn't create " & $CTM_alarms)
If Not DirCreate($PCS_Log) Then ZipperLog("Couldn't create " & $PCS_Log)
If Not DirCreate($Sprint_Log) Then ZipperLog("Couldn't create " & $Sprint_Log)
If Not DirCreate($Sprint_Log_lokal) Then ZipperLog("Couldn't create " & $Sprint_Log_lokal)
If Not DirCreate($Rel_Log) Then ZipperLog("Couldn't create " & $Rel_Log)
If Not DirCreate($ProcessLog) Then ZipperLog("Couldn't create " & $ProcessLog)
;Pfad verarbeiten
$aRootPath=StringSplit($dFindPath,"\",1)

;get ini Parameters
$siniPath=@ScriptDir&"\zipper.ini"
If Not FileExists($siniPath)  Then
   IniWrite($siniPath,"Path to PCS Online","Path","")
   IniWrite($siniPath,"Don't Skip Crash","Crash","0")
   IniWrite($siniPath,"Don't Skip RENFM","RENFM","0")
   IniWrite($siniPath,"Get ProcessData.log","ProcessData","0")
EndIf


TrayTip("Zipper.exe","Copying PCU Files...",5)

;Get/find (unmapped) network Drive
$FreeDrive_c=_GetFreeDriveLetter("\\192.168.214.241\c$")
$FreeDrive_d=_GetFreeDriveLetter("\\192.168.214.241\d$")

;Get PCU Logs
If Not GetFromPCU($FreeDrive_c,$FreeDrive_d ) Then
   ZipperLog("Couldn't connect to the PCU50")
EndIf

TrayTip("Zipper.exe","Copying RS Files...",5)
;Get Local P+ Logs
$path5="C:\P+\ProductSupportErrorLog.txt"
If FileExists($path5) Then
   FileCopy ( $path5,$Sprint_Log_lokal ,0)
Else
   ZipperLog("No local RS Log File in Path: "&$path5)
EndIf

If Not GetFromRS() Then
   ZipperLog("No external Renishaw PC")
EndIf

TrayTip("Zipper.exe","Copying Rel Report...",5)
if Not RelReport($Rel_Log) Then
   ZipperLog("Couldn't create Relaibility Records")
EndIf


TrayTip("Zipper.exe","Copying PCS Files...",5)
If Not GetFromPCS($PCS_Log) Then
   ZipperLog("Couldn't write PCS Logfiles")
EndIf


;Create Index File
$hcontent=FileOpen($LogFolder&"\content.txt",10)
$archive_content = _FileListToArrayRec($LogFolder,"*|*content.txt",1,1,1,1)
For $i=1 to $archive_content[0]
FileWrite($hcontent,$archive_content[$i]& @CRLF)
Next
 _TicksToTime(Int(TimerDiff($thandle)), $t_hours, $t_mins, $t_secs)
$t_time = StringFormat("%02i:%02i:%02i", $t_hours, $t_mins, $t_secs)
FileWrite($hcontent,"Dauer: "&$t_time& @CRLF)
FileClose($hcontent)


;Copy Folder Content to Archive
TrayTip("Zipper.exe","Zipping Files...",5)
$bZip=ZipCreate($Archive_Path,$LogFolder)
if Not $bZip Then
   ZipperLog("Func ZipCreate nicht erfolgreich")
EndIf



;Delete Backup Folder
If FileExists($Archive_Path) Then
   MsgBox(0,"Successfull","New Log Archive in "& $Archive_Path,10)
   DirRemove($LogFolder,1)
Else
   MsgBox(0,"Unknown Error. Couldn't create the Zip Archive in " & $Archive_Path,10)
EndIf

;MsgBox(0,"","Main Schleife durchlaufen")
Exit













Func GetFromPCU($share_c,$share_d)

$path1=$share_c & "\ProgramData\Siemens\MotionControl\user\sinumerik\hmi\log\"
$path2=$share_c & "\ProgramData\Siemens\MotionControl\oem\sinumerik\hmi\CTM\"
$path3=$share_c & "\P+\ProductSupportErrorLog.txt"
$pathCH1=$share_c & "\ProgramData\Siemens\MotionControl\oem\sinumerik\hmi\CTM_CH1\"
$pathCH2=$share_c & "\ProgramData\Siemens\MotionControl\oem\sinumerik\hmi\CTM_CH2\"
$pathREN=$share_d & "\RenMF\"
$pathProcData=$share_c & "\ProgramData\Siemens\MotionControl\addon\sinumerik\hmi\cfg\script\"

If (DriveMapGet($share_c)=="") Then
   local $shareFRG_c=DriveMapAdd($share_c,"\\192.168.214.241\c$",0,@ComputerName&"\auduser","SUNRISE")
   If @error=3 Then $shareFRG_c=1
Else
   local $shareFRG_c=1
EndIf


If $shareFRG_c Then

   If FileExists($path1&"alarm_log\alarmlog.txt") Then  FileCopy( $path1&"alarm_log\alarmlog.txt",$PCU_Log,1)
   If (IniRead($siniPath,"Don't Skip Crash","Crash","0")=="1") Then
	  If FileExists($path1&"hmi\") Then
		 TrayTip("Zipper.exe","Copying dmp Files...",5)
		 local $source=$path1&"hmi\slsmhmihost*.dmp"
		 local $destination=$PCU_Log
		 Runwait(@ComSpec & " /c " & "xcopy " & '"' & $source & '"' & ' "' & $destination & '"' & " /C /Y /J","",@SW_HIDE)
	  EndIf
	  If FileExists($path1&"hmi\") Then
		 TrayTip("Zipper.exe","Copying Crash Logs...",5)
		 local $source=$path1&"hmi\slsmhmihost*.log"
		 local $destination=$PCU_Log
		 Runwait(@ComSpec & " /c " & "xcopy " & '"' & $source & '"' & ' "' & $destination & '"' & " /C /Y /J","",@SW_HIDE)
	  EndIf
	  If FileExists($path1&"hmi\run_hmi.log") Then FileCopy( $path1&"hmi\run_hmi.log",$PCU_Log,1)
	  If FileExists($path1&"hmi\blackbox_run_hmi.asl") Then
		 TrayTip("Zipper.exe","Copying Blackbox...",5)
		 local $source=$path1&"hmi\blackbox_run_hmi.asl"
		 local $destination=$PCU_Log
		 Runwait(@ComSpec & " /c " & "xcopy " & '"' & $source & '"' & ' "' & $destination & '"' & " /C /Y /J","",@SW_HIDE)
	  EndIf
   EndIf
   If (IniRead($siniPath,"Get ProcessData.log","ProcessData","0")=="1") Then FileCopy($pathProcData& "ProcessData.log”",$ProcessLog,1)
   If FileExists($path2) Then FileCopy( $path2&"C*.txt",$CTM_Log1,1)
   If FileExists($path2) Then FileCopy( $path2&"C*.ini",$CTM_Log1,1)
   If FileExists($path2) Then DirCopy( $path2&"alarm",$CTM_alarms,1)
   If FileExists($pathCH1) Then FileCopy( $pathCH1&"C*.txt",$CTM_Log1,1)
   If FileExists($pathCH1) Then FileCopy( $pathCH1&"C*.ini",$CTM_Log1,1)
   If FileExists($pathCH2) Then FileCopy( $pathCH2&"C*.txt",$CTM_Log2,1)
   If FileExists($pathCH2) Then FileCopy( $pathCH2&"C*.ini",$CTM_Log2,1)
   If FileExists($path3) Then
	  FileCopy( $path3,$Sprint_Log_lokal,0)
   Else
	  ZipperLog("No Renishaw LogFile on PCU: " & $path3)
   EndIf
Else
ZipperLog("Couldn't connect to the network share " & $share_c & " . Please check the connection and retry. Error number: "&@error)
EndIf



If (DriveMapGet($share_d)=="") Then
   local $shareFRG_d=DriveMapAdd($share_d,"\\192.168.214.241\d$",0,@ComputerName&"\auduser","SUNRISE")
   If @error=3 Then $shareFRG_d=1
Else
   local $shareFRG_d=1
EndIf


If (IniRead($siniPath,"Don't Skip RENFM","RENFM","0")=="1") Then
If $shareFRG_d Then
   If FileExists($pathREN) Then
		 local $source=$pathREN&"*.renmf"
		 local $destination=$Sprint_Log_lokal
		 Runwait(@ComSpec & " /c " & "xcopy " & '"' & $source & '"' & ' "' & $destination & '"' & " /C /Y /J","",@SW_HIDE)
		 return 1
	  EndIf
Else
   ZipperLog("Couldn't connect to the network share " & $share_d & " . Please check the connection and retry. Error number: "&@error)
   return 0
EndIf
Endif
return 1
EndFunc



Func Filename($fullPath)
Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
Local $aPathSplit = _PathSplit($fullPath, $sDrive, $sDir, $sFileName, $sExtension)
Return $sFilename
EndFunc

Func ZipperLog($sMsg)
   Local $hZipLog=FileOpen($LogFolder & "\Zipperlog.txt",1)
   FileWrite($hZipLog,$sMsg & @CRLF)
   FileClose($hZipLog)
   Return 1
EndFunc


;;insert string for vElement
Func _GetFreeDriveLetter($sShare)
Local $aArray = DriveGetDrive($DT_NETWORK)
Local $skip=0
For $vElement in $aArray
   If (DriveMapGet($vElement)==$sShare) Then
	  ZipperLog("First Loop. For " & $sShare & " Drive on "&$vElement)
	  return $vElement
	  $skip=1
	EndIf
 Next
For $x = 68 To 90
        If DriveMapGet(Chr($x)&':')==$sShare Then
		   return (Chr($x)&':')
		   ZipperLog("Second Loop. Drive on "&Chr($x)&":")
		EndIf

Next
If Not $skip Then
	For $x = 68 To 90
	  If DriveStatus(Chr($x) & ':\') = 'INVALID' Then Return(Chr($x) & ':' )
		Next
EndIf
EndFunc



Func GetFromPCS($LogDir)

Local $sPCSPath,$iDiff,$aTime,$sLatestFile
Local $maxDiff=9999999
$path6=""
$path7=""
$path8=""

;Get PCS Path
;Aus Aufgabenverwaltung
If IniRead( $siniPath, "Path to PCS Online", "Path", "" ) == "" Then
   $path6="C:\PCS\PCS_Online\"
Else
   $path6=IniRead( $siniPath, "Path to PCS Online Folder", "Path", "" )
EndIf

$path7="C:\OEM\PCS_Online 020\"
$path8="C:\OEM\PCS Online\"

If FileExists($path6) Then
   $sPCSPath=$path6
ElseIf FileExists($path7) Then
   $sPCSPath=$path7
ElseIf FileExists($path8) Then
   $sPCSPath=$path8
Else
   ZipperLog("Es konnte keiner der lokalen PCS Pfade gefunden werden.")
   ;MsgBox(0,"Warning","Couldn't use the standard PCS PATH. Please enter a valid Path into: " & $siniPath)
   return 0
EndIf


;Get xml Files
If Not FileCopy($sPCSPath & "N30.xml",$PCS_Log,0) Then ZipperLog("N30.xml konnte nicht kopiert werden")
If Not FileCopy($sPCSPath & "settings.xml",$PCS_Log,0) Then ZipperLog("settings.xml konnte nicht kopiert werden")


;Get Log Files
$aFS=_FileListToArray($sPCSPath& "\logs\","*.log",1,1)

For $i=1 To $aFS[0] Step 1
   $aTime=FileGetTime($aFS[$i])
   $iDiff=_DateDiff("D",$aTime[0] & "/" & $aTime[1]  & "/" & $aTime[2],_NowCalc())

   If $iDiff <= $maxDiff Then
	  $maxDiff=$iDiff
	  $sLatestFile=$aTime[0] & $aTime[1] & $aTime[2] & "*"
   EndIf

Next

$aFS=_FileListToArray($sPCSPath& "\logs\",$sLatestFile,1,1)

For $j=1 To $aFS[0]
   FileCopy ( $aFS[$j],$LogDir ,0)
Next

Return 1
   EndFunc


Func GetFromRS()

;Get RS IP
Local $RSShare=0
If Ping("192.168.214.231",500) > 0 Then
   $RSShare="192.168.214.231"
ElseIf Ping("192.168.214.232",500) > 0 Then
   $RSShare="192.168.214.232"
ElseIf Ping("192.168.214.233",500) > 0 Then
   $RSShare="192.168.214.233"
Else
   return 0
EndIf

$passString_c="\\"&$RSSHare&"\c$"
Local $freeLetter_c=_GetFreeDriveLetter($passString_c)

$FRGRSSHare_c=DriveMapAdd($freeLetter_c,"\\"&$RSShare&"\c$",0,@ComputerName&"\auduser","SUNRISE")
if @error=3 Then $FRGRSShare_c=1
If $FRGRSShare_c Then

If FileExists($freeLetter_c&"\P+\ProductSupportErrorLog.txt") Then
	  FileCopy($freeLetter_c&"\P+\ProductSupportErrorLog.txt",$Sprint_Log,0)
   Else
	  ZipperLog("No ProductSupportErrorLog.txt on "& $passString_c)
   EndIf
EndIf


If (IniRead($siniPath,"Don't Skip RENFM","RENFM","0")=="1") Then
   $passString_d="\\"&$RSSHare&"\d$"
   Local $freeLetter_d=_GetFreeDriveLetter($passString_d)

   $FRGRSSHare_d=DriveMapAdd($freeLetter_d,"\\"&$RSShare&"\d$",0,@ComputerName&"\auduser","SUNRISE")
   if @error=3 Then $FRGRSShare_d=1
   If $FRGRSShare_d Then
	  If FileExists($freeLetter_d&"\RenMF\") Then
			local $source=$freeLetter_d&"\RenMF\*.renmf"
			local $destination=$Sprint_Log
			Runwait(@ComSpec & " /c " & "xcopy " & '"' & $source & '"' & ' "' & $destination & '"' & " /C /Y /J","",@SW_HIDE)
			return 1
	  Else
		 ZipperLog("No *renmf-Files in "& $passString_d)
	  EndIf
   EndIf
   return 0
EndIf
return 1

EndFunc



Func ZipCreate($ZipArchive,$AddFile)
$aFS=_FileListToArrayRec($AddFile,"*",1,1)
local $iMax=$aFS[0]
$line=""
local $7z=@ScriptDir&"\7z\7za.exe"
local $command = $7z & " a -t7z " & $ZipArchive & " " & $AddFile

$pid=RunWait(@ComSpec & " /c " & $command, @ScriptDir, @SW_HIDE,$STDERR_MERGED)
sleep(100)

If ProcessExists($pid)>0 Then return 0

If FileExists($ZipArchive) Then
   return 1
Else
   return 0
EndIf
EndFunc



Func Terminate()
Exit
EndFunc


Func RelReport($sPath)

$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$colItems = ""
$hFile=FileOpen($sPath & "\ReliabilityRecords.txt",10)
$strComputer = "localhost"
$Output=""
$Output = $Output & "Computer: " & $strComputer  & @CRLF
$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_ReliabilityRecords", "WQL", _
                                          $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

If IsObj($colItems) then
   For $objItem In $colItems
      $Output = $Output & "ComputerName: " & $objItem.ComputerName & @CRLF
      $Output = $Output & "EventIdentifier: " & $objItem.EventIdentifier & @CRLF
      $strInsertionStrings = $objItem.InsertionStrings(0)
      $Output = $Output & "InsertionStrings: " & $strInsertionStrings & @CRLF
      $Output = $Output & "Logfile: " & $objItem.Logfile & @CRLF
      $Output = $Output & "Message: " & $objItem.Message & @CRLF
      $Output = $Output & "ProductName: " & $objItem.ProductName & @CRLF
      $Output = $Output & "RecordNumber: " & $objItem.RecordNumber & @CRLF
      $Output = $Output & "SourceName: " & $objItem.SourceName & @CRLF
      $Output = $Output & "TimeGenerated: " & WMIDateStringToDate($objItem.TimeGenerated) & @CRLF
      $Output = $Output & "User: " & $objItem.User & @CRLF & @CRLF & @CRLF
   Next
   FileWrite($hFile, $Output )
Else
   ZipperLog("No WMI Objects Found for class: Win32_ReliabilityRecords")
   FileClose($hFile)
   return 0
Endif
FileClose($hFile)
return 1
EndFunc

Func WMIDateStringToDate($dtmDate)

	Return (StringMid($dtmDate, 5, 2) & "/" & _
	StringMid($dtmDate, 7, 2) & "/" & StringLeft($dtmDate, 4) _
	& " " & StringMid($dtmDate, 9, 2) & ":" & StringMid($dtmDate, 11, 2) & ":" & StringMid($dtmDate,13, 2))
EndFunc


Func MyErrFunc()
    $HexNumber = Hex($oMyError.number, 8)
    ZipperLog("COM Error Test. We intercepted a COM Error !")
    ZipperLog("err.description is: " & @TAB & $oMyError.description)
    ZipperLog("err.windescription:" & @TAB & $oMyError.windescription)
    ZipperLog("err.number is: " & @TAB & $HexNumber)
    ZipperLog("err.lastdllerror is: " & @TAB & $oMyError.lastdllerror)
    ZipperLog("err.scriptline is: " & @TAB & $oMyError.scriptline)
    ZipperLog("err.source is: " & @TAB & $oMyError.source)
    ZipperLog("err.helpfile is: " & @TAB & $oMyError.helpfile)
    ZipperLog("err.helpcontext is: " & @TAB & $oMyError.helpcontext)
    SetError(1) ; to check for after this function returns
EndFunc   ;==>MyErrFunc
#AutoIt3Wrapper_icon=C:\Users\rissr\Desktop\Script_Grab\Zipper\zip_gold.ico
#Include <File.au3>
#include <Date.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <Security.au3>


;;;;;;;;;;;;;;;;;;;;;;;;;
;1. Create empty Zip Archive
;2. Create/Check Network Share
;3. Copy Files to Zip
;4. Create Protocol
;5. Copy Protocol toZip
;6. GUI ersetzt ini Datei
;7. Domain Abfrage
;8. GUI NSI.jpg in hidden folder
;9. Comment Function
;;;;;;;;;;;;;;;;;;;;;;;;;


;GUI Init
$With_HMILogs=0
$With_Hang=0
$With_Rel=0
$With_CTMLog1=0
$With_CTMLog2=0
$With_CTMAlarms=0
$With_PCS=0
$With_PCS_custom_Path=0
$With_Sprint=0
$With_ProcessLog=0
$new_PCS_Path=0
$With_Comment=0



;Variablendefinition
Dim $Zip,$sName,$sFirstString,$FreeDrive
Dim $sPCSPath
$LogFolder="\PCULogs"
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
$thandle=TimerInit()
Global $t_time, $t_hours, $t_mins, $t_secs
Opt("TrayAutoPause", 0)
HotKeySet("{ESC}","Terminate")
TrayTip("Zipper.exe","Zipper running...",5)
$Comment_String=""

;COM Object Error Handle
$oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

;Get Configuration from GUI
Pack_GUI()


;Check for extented Filename
$Archive_Path=ArchiveName()


;Create the Folder Structure in D:\Backups\
Create_Folder_Struct()


;Status Information
TrayTip("Zipper.exe","Copying PCU Files...",5)

;find unmapped/existing network Drive
$FreeDrive_c=_GetFreeDriveLetter("\\192.168.214.241\c$")
$FreeDrive_d=_GetFreeDriveLetter("\\192.168.214.241\d$")


;Get PCU Logs
If BitAND(($FreeDrive_c<>"0"),($FreeDrive_d<>"0")) Then
   If Not GetFromPCU($FreeDrive_c,$FreeDrive_d ) Then
   ZipperLog("Couldn't connect to the PCU50")
   EndIf
EndIf



;Get Sprint logs
If $With_Sprint Then
   ;Status Information
   TrayTip("Zipper.exe","Copying RS Files...",5)

   ;Get Local P+ Logs
   $path_RS_local="C:\P+\ProductSupportErrorLog.txt"
   If FileExists($path_RS_local) Then
	  FileCopy ( $path_RS_local,$Sprint_Log_lokal ,0)
   Else
	  ZipperLog("No local RS Log File in Path: "&$path_RS_local)
   EndIf

   If Not GetFromRS() Then
	  ZipperLog("No external Renishaw PC")
   EndIf
EndIf


If $With_Rel Then
   TrayTip("Zipper.exe","Copying Rel Report...",5)
   If Not RelReport($Rel_Log) Then
	  ZipperLog("Couldn't create Relaibility Records")
   EndIf
EndIf


If $With_PCS Then
   TrayTip("Zipper.exe","Copying PCS Files...",5)
   If Not GetFromPCS($PCS_Log) Then
	  ZipperLog("Couldn't write PCS Logfiles")
   EndIf
EndIf


;txt with Content of Zip Folder
Create_Index_File()


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


Exit


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;FUNCTIONS:
;Pack_GUI
;ArchiveName
;Get From PCU
;Filename
;Zipperlog
;FreeDrive
;Get From PCS
;Get From RS
;ZipCreate
;Reliability Report
;Create Index File
;Terminate
;Error Func for COM Objects
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



Func Pack_GUI()

	$GUIBKCOLOR=0xFFFFFF
    Local $hGUI = GUICreate("Choose your package", 350, 350,-1,-1,0x80880000)
    GUISetBkColor($GUIBKCOLOR,$hGUI)
	DirCreate(@ScriptDir&"\Temp")
	FileInstall("C:\Users\rissr\Desktop\Script_Grab\Zipper\NSI.jpg",@ScriptDir&"\Temp\NSI.jpg",1)
    FileSetAttrib(@ScriptDir&"\Temp","+H",1)

	; Create a checkbox control.
    Local $idCheckbox_PCU_Crash = GUICtrlCreateCheckbox("HMI Crash- und Alarmlogs", 10, 10, 240, 25)
	Local $idCheckbox_PCU_Hang = GUICtrlCreateCheckbox("PCU Hang-Dumps", 40,40,185, 25)
	Local $idCheckbox_PCU_Rel = GUICtrlCreateCheckbox("PCU50 Reliability Report", 40, 70, 185, 25)
	Local $idCheckbox_CTM1 = GUICtrlCreateCheckbox("CTM Config Files CH1", 10, 100, 185, 25)
	Local $idCheckbox_CTM2 = GUICtrlCreateCheckbox("CTM Config Files CH2", 10, 130, 185, 25)
	Local $idCheckbox_CTM_ALM = GUICtrlCreateCheckbox("CTM Alarmlog", 10, 160, 185, 25)
	Local $idCheckbox_PCS = GUICtrlCreateCheckbox("PCS Settings and Logfiles", 10, 190, 185, 25)
	Local $idCheckbox_PCS_Path = GUICtrlCreateCheckbox("Use Custom PCS Path", 10, 220, 185, 25)
	Local $idCheckbox_RS = GUICtrlCreateCheckbox("Renishaw Logfiles", 10, 250, 185, 25)
	Local $idCheckbox_ProcessLog = GUICtrlCreateCheckbox("HMI ProcessLog", 10, 280, 185, 25)
    Local $idClose = GUICtrlCreateButton("Set", 220, 310, 120, 25)
    local $idLogo=GUICtrlCreatePic(@ScriptDir&"\Temp\NSI.jpg",250,10,80,95)
	local $idCheckbox_Com=GUICtrlCreateCheckbox("Add a comment", 10,310,185, 25)

    ; Display the GUI.
    GUISetState(@SW_SHOW, $hGUI)

    ; Loop until the user exits.
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idClose
                ExitLoop
            Case $idCheckbox_PCU_Crash
                $With_HMILogs=_IsChecked($idCheckbox_PCU_Crash)
			Case $idCheckbox_PCU_Hang
                $With_Hang=_IsChecked($idCheckbox_PCU_Hang)
				if $With_Hang Then MsgBox(0,"Warning","PCU Hang Dumps will collect up to 1,5GB of data. Make sure that every other network process is finished beforehand")
			Case $idCheckbox_PCU_Rel
                $With_Rel=_IsChecked($idCheckbox_PCU_Rel)
			Case $idCheckbox_CTM1
                $With_CTMLog1=_IsChecked($idCheckbox_CTM1)
			Case $idCheckbox_CTM2
                $With_CTMLog2=_IsChecked($idCheckbox_CTM2)
			Case $idCheckbox_CTM_ALM
                $With_CTMAlarms=_IsChecked($idCheckbox_CTM_ALM)
			Case $idCheckbox_PCS
                $With_PCS=_IsChecked($idCheckbox_PCS)
			Case $idCheckbox_PCS_Path
                $With_PCS_custom_Path=_IsChecked($idCheckbox_PCS_Path)
				 If $With_PCS_custom_Path Then $new_PCS_Path=InputBox ( "Input Prompt", "Enter the custom PCS Path" , "C:\PCS\PCS_Online\" )
			Case $idCheckbox_RS
                $With_Sprint=_IsChecked($idCheckbox_RS)
			Case $idCheckbox_ProcessLog
                $With_ProcessLog=_IsChecked($idCheckbox_ProcessLog)
			 Case $idCheckbox_Com
				$With_Comment=_IsChecked($idCheckbox_Com)
				If $With_comment Then $Comment_String=InputBox ( "error description", "Enter here" , "" )
        EndSwitch
    WEnd

    ; Delete the previous GUI and all controls.
    GUIDelete($hGUI)
	DirRemove(@ScriptDir&"\Temp\",1)
If StringLen($new_PCS_Path)<7 Then $new_PCS_Path="C:\PCS\PCS_Online\"
EndFunc

Func _IsChecked($idControlID)
    Return BitAND(GUICtrlRead($idControlID),$GUI_CHECKED)= $GUI_CHECKED
EndFunc



Func ArchiveName()
;Zeitstring für Archivname
;If FileExists($Archive_Path) Then
   $aTime=StringSplit(_NowTime(5),":/")
   $aDate=StringSplit(_NowDate(),"/.")
   $Archive_Path="D:\Backups\log_archive_" & $aDate[1] & "_" & $aDate[2] & "_" & $aDate[3] & "_" & $aTime[1] & "_" & $aTime[2] & "_" & $aTime[3] & ".7z"
   return $Archive_Path
;Else
;   return $Archive_Path
;EndIf
EndFunc






Func Create_Folder_Struct()
If FileExists($LogFolder) Then DirRemove($LogFolder,1)
If Not DirCreate($LogFolder) Then ZipperLog("Couldn't create " & $LogFolder)
If BitOR($With_HMILogs,$With_Hang) Then
   If Not DirCreate($PCU_Log) Then ZipperLog("Couldn't create " & $PCU_Log)
EndIf
If $With_CTMLog1 Then
   If Not DirCreate($CTM_Log1) Then ZipperLog("Couldn't create " & $CTM_Log1)
EndIf
If $With_CTMLog2 Then
   If Not DirCreate($CTM_Log2) Then ZipperLog("Couldn't create " & $CTM_Log2)
EndIf
If $With_CTMAlarms Then
   If Not DirCreate($CTM_alarms) Then ZipperLog("Couldn't create " & $CTM_alarms)
EndIf
If $With_PCS Then
   If Not DirCreate($PCS_Log) Then ZipperLog("Couldn't create " & $PCS_Log)
EndIf
If $With_Sprint Then
   If Not DirCreate($Sprint_Log) Then ZipperLog("Couldn't create " & $Sprint_Log)
   If Not DirCreate($Sprint_Log_lokal) Then ZipperLog("Couldn't create " & $Sprint_Log_lokal)
EndIf
If $With_Rel Then
   If Not DirCreate($Rel_Log) Then ZipperLog("Couldn't create " & $Rel_Log)
EndIf
If $With_ProcessLog Then
   If Not DirCreate($ProcessLog) Then ZipperLog("Couldn't create " & $ProcessLog)
EndIf
EndFunc






Func GetFromPCU($share_c,$share_d)

$user_hmi="\ProgramData\Siemens\MotionControl\user\sinumerik\hmi\"
$oem_hmi="\ProgramData\Siemens\MotionControl\oem\sinumerik\hmi\"
$path1=$share_c & $user_hmi & "log\"
$path2=$share_c & $oem_hmi & "CTM\"
$path3=$share_c & "\P+\ProductSupportErrorLog.txt"
$pathCH1=$share_c & $oem_hmi & "CTM_CH1\"
$pathCH2=$share_c & $oem_hmi & "CTM_CH2\"
$pathREN=$share_d & "\RenMF\"
$pathProcData=$share_c & "\ProgramData\Siemens\MotionControl\addon\sinumerik\hmi\cfg\script\"


;Get Files from C$
   ;Get HMI Logs
   If $With_HMILogs Then
	  If FileExists($path1&"alarm_log\alarmlog.txt") Then  FileCopy( $path1&"alarm_log\alarmlog.txt",$PCU_Log,1)

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

   ;Get Hang Dumps
   If $With_Hang Then
	  If FileExists($path1&"hmi\") Then
		 TrayTip("Zipper.exe","Copying dmp Files...",5)
		 local $source=$path1&"hmi\slsmhmihost*.dmp"
		 local $destination=$PCU_Log
		 Runwait(@ComSpec & " /c " & "xcopy " & '"' & $source & '"' & ' "' & $destination & '"' & " /C /Y /J","",@SW_HIDE)
	  EndIf
   EndIf

   ;Get ProcessLog (PWC020 only)
   If $With_ProcessLog Then FileCopy($pathProcData& "ProcessData.log”",$ProcessLog,1)

   ;Get CTM/CTM CH1
   If $With_CTMLog1 Then
	  If FileExists($path2) Then FileCopy( $path2&"C*.txt",$CTM_Log1,1)
	  If FileExists($path2) Then FileCopy( $path2&"C*.ini",$CTM_Log1,1)
	  If FileExists($pathCH1) Then FileCopy( $pathCH1&"C*.txt",$CTM_Log1,1)
	  If FileExists($pathCH1) Then FileCopy( $pathCH1&"C*.ini",$CTM_Log1,1)
   EndIf

   ;Get CTM CH2
   If $With_CTMLog2 Then
	  If FileExists($pathCH2) Then FileCopy( $pathCH2&"C*.txt",$CTM_Log2,1)
	  If FileExists($pathCH2) Then FileCopy( $pathCH2&"C*.ini",$CTM_Log2,1)
   EndIf

   ;Get CTM Alarmlog
   If $With_CTMAlarms Then
	  If FileExists($path2) Then DirCopy( $path2&"alarm",$CTM_alarms,1)
   EndIf

  ;Get local Sprint Logs
  If $With_Sprint Then
	  If FileExists($path3) Then
		 FileCopy( $path3,$Sprint_Log_lokal,0)
	  Else
		 ZipperLog("No Renishaw LogFile on PCU: " & $path3)
	  EndIf
   EndIf





;Get Sprint Logs from D$
If $With_Sprint Then
   If FileExists($pathREN) Then
		 local $source=$pathREN&"*.renmf"
		 local $destination=$Sprint_Log_lokal
		 Runwait(@ComSpec & " /c " & "xcopy " & '"' & $source & '"' & ' "' & $destination & '"' & " /C /Y /J","",@SW_HIDE)
		 return 1
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



Func _GetFreeDriveLetter($sShare)
;Connect Networkdrive

Local $userDomain=""
Local $skip=0
Local $RetVal=""
local $shareFRG=0

;Read existing Drives from Reg Key
For $x = 90 To 68 Step -1
   $sVar = DriveMapGet(Chr($x)&":")
   If $sVar=$sShare Then
	  $RetVal=Chr($x)&":"
	  $skip=1
	  $shareFRG=1
	  ZipperLog("Drive already exists. For " & $sShare & " Drive on "& Chr($x))
   EndIf
Next

;find unused Driveletter
If $skip=0 Then
   For $x = 90 To 68 Step -1
	  If DriveStatus(Chr($x) & ":\") == "INVALID" Then
		 $RetVal=(Chr($x) & ":" )
		 ExitLoop
	  EndIf
   Next
EndIf

;Get the Domain
Local $aArrayOfData = _Security__LookupAccountName(@UserName)
If IsArray($aArrayOfData) Then
    $userDomain=$aArrayOfData[1]
 Else
	$userDomain=@ComputerName
 EndIf


;Connect Part
If $skip=0 Then
   $shareFRG=DriveMapAdd($RetVal,$sShare,0,$userDomain&"\auduser","SUNRISE")
EndIf

;return the right drive letter
If $shareFRG Then
   return $RetVal
Else
	  return 0
EndIf
EndFunc



Func GetFromPCS($LogDir)

Local $sPCSPath,$iDiff,$aTime,$sLatestFile
Local $maxDiff=9999999
$path6=""
$path7=""
$path8=""

;Get PCS Path
If $With_PCS_custom_Path Then
   $path6=$new_PCS_Path
Else
   $path6="C:\PCS\PCS_Online\"
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

$passString_c="\\"&$RSShare&"\c$"
Local $freeLetter_c=_GetFreeDriveLetter($passString_c)

If ($freeLetter_c<>"0") Then
   If FileExists($freeLetter_c&"\P+\ProductSupportErrorLog.txt") Then
		 FileCopy($freeLetter_c&"\P+\ProductSupportErrorLog.txt",$Sprint_Log,0)
	  Else
		 ZipperLog("No ProductSupportErrorLog.txt on "& $passString_c)
   EndIf
EndIf





$passString_d="\\"&$RSShare&"\d$"
Local $freeLetter_d=_GetFreeDriveLetter($passString_d)


If ($freeLetter_d<>"0") Then
   If FileExists($freeLetter_d&"\RenMF\") Then
		 local $source=$freeLetter_d&"\RenMF\*.renmf"
		 local $destination=$Sprint_Log
		 Runwait(@ComSpec & " /c " & "xcopy " & '"' & $source & '"' & ' "' & $destination & '"' & " /C /Y /J","",@SW_HIDE)
		 return 1
   Else
		 ZipperLog("No *renmf-Files in "& $passString_d)
		 return 0
   EndIf
EndIf
EndFunc



Func ZipCreate($ZipArchive,$AddFile)
$aFS=_FileListToArrayRec($AddFile,"*",1,1)
local $iMax=$aFS[0]
$line=""
local $7z=@ScriptDir&"\7z\7za.exe"
If Not FileExists(@ScriptDir & "\7z\7za.exe") Then
   DirCreate(@ScriptDir&"\7z\")
   FileInstall("C:\Users\rissr\Desktop\Script_Grab\Zipper\7z\7za.exe", $7z,1)
EndIf
local $command = $7z & " a -w -t7z " & $ZipArchive & " " & $AddFile

$pid=RunWait(@ComSpec & " /c " & $command, @ScriptDir, @SW_HIDE,$STDERR_MERGED)
sleep(100)

If ProcessExists($pid)>0 Then return 0

If FileExists($ZipArchive) Then
   DirRemove(@ScriptDir&"\7z\",1)
   return 1
Else
   return 0
EndIf
EndFunc



Func Create_Index_File()

;Break comment string down
$Comment_String=StringReplace($Comment_String,".","."&@CRLF)

;Create Index File
local $hcontent=FileOpen($LogFolder&"\content.txt",10)
FileWrite($hcontent,":-------------Operators description--------------:" & @CRLF)
FileWrite($hcontent,$Comment_String & @CRLF)
FileWrite($hcontent,":-----------------Collected Files----------------:" & @CRLF)
local $archive_content = _FileListToArrayRec($LogFolder,"*|*content.txt",1,1,1,1)
For $i=1 to $archive_content[0]
FileWrite($hcontent,$archive_content[$i]& @CRLF)
Next
 _TicksToTime(Int(TimerDiff($thandle)), $t_hours, $t_mins, $t_secs)
local $t_time = StringFormat("%02i:%02i:%02i", $t_hours, $t_mins, $t_secs)
FileWrite($hcontent,"Dauer: "&$t_time& @CRLF)
FileWrite($hcontent,":---------------Choosen Parameters---------------:" & @CRLF)
FileWrite($hcontent,"With HMI Log: "& Binary($With_HMILogs) & @CRLF)
FileWrite($hcontent,"With Hang Dumps: "& Binary($With_Hang) & @CRLF)
FileWrite($hcontent,"With Rel: "& Binary($With_Rel) & @CRLF)
FileWrite($hcontent,"With CTM CH1: "& Binary($With_CTMLog1) & @CRLF)
FileWrite($hcontent,"With CTM CH2: "& Binary($With_CTMLog2) & @CRLF)
FileWrite($hcontent,"With CTM Alarms: "& Binary($With_CTMAlarms) & @CRLF)
FileWrite($hcontent,"With PCS: "& Binary($With_PCS) & @CRLF)
FileWrite($hcontent,"With Sprint: "& Binary($With_Sprint) & @CRLF)
FileWrite($hcontent,"With Process Log: "& Binary($With_ProcessLog) & @CRLF)
FileWrite($hcontent,";------------------------------------------------:" & @CRLF)
FileClose($hcontent)

EndFunc




Func Terminate()
Exit
EndFunc


Func RelReport($sPath)

$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$colItems = ""
$hFile=FileOpen($sPath & "\ReliabilityRecords.txt",10)
$strComputer = "192.168.214.241"
$Output=""
$Output = $Output & "Computer: " & $strComputer  & @CRLF
$WbemAuthenticationLevelPktPrivacy = 6
$strNamespace = "root\cimv2"
$strUser = "auduser"
$strPassword = "SUNRISE"
$objWbemLocator = ObjCreate("WbemScripting.SWbemLocator")
$objWMIService = $objwbemLocator.ConnectServer($strComputer,$strNamespace,$strUser,$strPassword)
$objWMIService.Security_.authenticationLevel = $WbemAuthenticationLevelPktPrivacy
;$objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & $strComputer & "\root\CIMV2")
$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_ReliabilityRecords", "WQL",$wbemFlagReturnImmediately + $wbemFlagForwardOnly)


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
EndFunc
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=.\Support\Icons\uniextract_exe.ico
#AutoIt3Wrapper_Outfile=.\UniExtract.exe
#AutoIt3Wrapper_Res_Comment=Compiled with AutoIt http://www.autoitscript.com/
#AutoIt3Wrapper_Res_Description=Universal Extractor
#AutoIt3Wrapper_Res_Fileversion=2.0.0
#AutoIt3Wrapper_Res_LegalCopyright=GNU General Public License v2
#AutoIt3Wrapper_Res_Field=Author|Jared Breland <jbreland@legroom.net>
#AutoIt3Wrapper_Res_Field=Homepage|http://www.legroom.net/software
#AutoIt3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_AU3Check_Parameters=-w 4 -w 5
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;
; ---------------------------------------------------------------------------e
;
; Universal Extractor v2.0.0
; Author:	Jared Breland <jbreland@legroom.net>, Version 2.0.0 by Bioruebe
; Homepage:	http://www.legroom.net/mysoft
; Language:	AutoIt v3.3.10.2
; License:	GNU General Public License v2 (http://www.gnu.org/copyleft/gpl.html)
;
; Very Basic Script Function:
;	Use Unix File Tool and TrID to determine filetype
;	Use Exeinfo PE and PEiD to identify executable filetypes
;	Extract known archive types
;
; ----------------------------------------------------------------------------

; Setup environment
#include <APIConstants.au3>
#include <Array.au3>
#include <ComboConstants.au3>
#include <Constants.au3>
#include <Crypt.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <GDIPlus.au3>
#include <GUIConstantsEx.au3>
#include <GuiComboBox.au3>
#include <GUIEdit.au3>
#include <GuiListBox.au3>
#include <INet.au3>
#include <Math.au3>
#include <Misc.au3>
#include <ProgressConstants.au3>
#include <SQLite.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <WinAPI.au3>
#include <WinAPIShPath.au3>
#include <WindowsConstants.au3>
#include "HexDump.au3"
#include "Pie.au3"

Const $name = "Universal Extractor"
Const $version = "2.0.0 RC 1"
Const $codename = '"Back from the grave"'
Const $title = $name & " " & $version
Const $website = "https://www.legroom.net/software/uniextract"
Const $website2 = "https://bioruebe.com/uniextract"
Const $websiteGithub = "https://github.com/Bioruebe/UniExtract2"
Const $sUpdateURL = "https://update.bioruebe.com/uniextract/data/"
Const $sLegacyUpdateURL = "https://update.bioruebe.com/uniextract/update2.php"
Const $sGetLinkURL = "https://update.bioruebe.com/uniextract/geturl.php?q="
Const $sSupportURL = "https://support.bioruebe.com/uniextract/upload.php"
Const $sStatsURL = "https://stat.bioruebe.com/uniextract/stats.php?a="
Const $bindir = @ScriptDir & "\bin\"
Const $langdir = @ScriptDir & "\lang\"
Const $defdir = @ScriptDir & "\def\"
Const $sUpdater = @ScriptDir & '\UniExtractUpdater.exe'
Const $sUpdaterNoAdmin = @ScriptDir & '\UniExtractUpdater_NoAdmin.exe'
Const $sEnglishLangFile = @ScriptDir & '\English.ini'
Const $notAsciiPattern = "(?i)(?m)^[\w\Q @!§$%&/\()=?,.-:+~'²³{[]}*#ß°^âëöäüîêôûïáéíóúàèìòù\E]+$"
;~ Const $cmd = @ComSpec & ' /d /k '
Const $cmd = (FileExists(@ComSpec)? @ComSpec: @WindowsDir & '\system32\cmd.exe') & ' /d /c '
Const $OPTION_KEEP = 0, $OPTION_DELETE = 1, $OPTION_ASK = 2, $OPTION_MOVE = 2
Const $HISTORY_FILE = "File History", $HISTORY_DIR = "Directory History"
Const $PACKER_UPX = "UPX", $PACKER_ASPACK = "Aspack"
Const $RESULT_UNKNOWN = 0, $RESULT_SUCCESS = 1, $RESULT_FAILED = 2, $RESULT_CANCELED = 3
Const $UNICODE_NONE = 0, $UNICODE_MOVE = 1, $UNICODE_COPY = 2
Const $UPDATE_ALL = 0, $UPDATE_HELPER = 1, $UPDATE_MAIN = 2
Const $UPDATEMSG_PROMPT = 0, $UPDATEMSG_SILENT = 1, $UPDATEMSG_FOUND_ONLY = 2
Const $STATUS_SYNTAX = "syntax", $STATUS_FILEINFO = "fileinfo", $STATUS_UNKNOWNEXE = "unknownexe", $STATUS_UNKNOWNEXT = "unknownext", _
	  $STATUS_INVALIDFILE = "invalidfile", $STATUS_INVALIDDIR = "invaliddir", $STATUS_NOTPACKED = "notpacked", $STATUS_BATCH = "batch", _
	  $STATUS_NOTSUPPORTED = "notsupported", $STATUS_MISSINGEXE = "missingexe", $STATUS_TIMEOUT = "timeout", $STATUS_PASSWORD = "password", _
	  $STATUS_MISSINGDEF = "missingdef", $STATUS_MOVEFAILED = "movefailed", $STATUS_FAILED = "failed", $STATUS_SUCCESS = "success", _
	  $STATUS_SILENT = "silent"


Opt("GUIOnEventMode", 1)
Opt("TrayOnEventMode", 1)
Opt("TrayMenuMode", 1 + 2)
Opt("TrayIconDebug", 1)

; Preferences
Global $settingsdir = @AppDataDir & "\Bioruebe\UniExtract"
Global $batchEnabled = 0
Global $language = ""
Global $history = 1
Global $appendext = 0
Global $warnexecute = 1
Global $freeSpaceCheck = 1
Global $NoBox = 0
Global $bHideStatusBoxIfFullscreen = 1
Global $OpenOutDir = 0
Global $iDeleteOrigFile = $OPTION_KEEP
Global $Timeout = 60000 ; milliseconds
Global $updateinterval = 1 ; days
Global $lastupdate = "2010/12/05" ; last official version
Global $addassocenabled = 0
Global $addassocallusers = 0
Global $addassoc = ""
Global $ID = ""
Global $FB_ask = 0
Global $Opt_ConsoleOutput = 0
Global $Log = 0
Global $CheckGame = 1
Global $bSendStats = 1
Global $iCleanup = $OPTION_MOVE
Global $KeepOutdir = 0
Global $KeepOpen = 0
Global $silentmode = 0
Global $extract = 1
Global $checkUnicode = 1
Global $bExtractVideo = 1
Global $StoreGUIPosition = 0
Global $iTopmost = 0
Global $posx = -1, $posy = -1
Global $trayX = -1, $trayY = -1

; Global variables
Dim $file, $filename, $filenamefull, $filedir, $fileext, $initoutdir, $outdir, $filetype = "", $initdirsize
Dim $prompt, $return, $Output, $hMutex
Dim $About, $Type, $win7, $silent, $iUnicodeMode = False, $reg64 = ""
Dim $debug = "", $guimain = False, $success = $RESULT_UNKNOWN, $TBgui = 0, $isofile = 0, $exStyle = -1, $sArcTypeOverride = 0
Dim $test, $test7z, $testzip, $testie, $testinno
Dim $innofailed, $arjfailed, $7zfailed, $zipfailed, $iefailed, $isfailed, $isofailed, $tridfailed, $gamefailed, $unpackfailed, $exefailed
Dim $oldpath, $oldoutdir, $sUnicodeName, $createdir
Dim $FS_GUI = False, $idTrayStatusExt, $BatchBut
Dim $isexe = False, $Message, $run = 0, $runtitle, $DeleteOrigFileOpt[3]
Dim $gaDropFiles[1], $queueArray[0], $aTridDefinitions[0][0], $aFileDefinitions[0][0]

; Check if OS is 64 bit version
If @OSArch == "X64" Or @OSArch == "IA64" Then
	Global $OSArch = "x64"
	Global $reg64 = 64
Else
	Global $OSArch = "x86"
EndIf
Global Const $archdir = $bindir & $OSArch & "\"

; Extractors
Const $7z = Quote($archdir & '7z.exe', True) ;x64									;15.05
Const $7zsplit = "7ZSplit.exe" 														;0.2
Const $ace = $bindir & "xace.exe" 													;2.6
Const $alz = "unalz.exe" 															;0.64
Const $arj = "arj.exe" 																;3.10
Const $aspack = "AspackDie.exe" 													;1.4.1
Const $bcm = Quote($archdir & "bcm.exe", True) ;x64									;1.00
Const $daa = "daa2iso.exe" 															;0.1.7e
Const $enigma = "EnigmaVBUnpacker.exe"												;0.44
Const $ethornell = "ethornell.exe" 													;unknown
Const $exeinfope = Quote($bindir & "exeinfope.exe")									;0.0.3.7
Const $filetool = Quote($bindir & "file\bin\file.exe", True)						;5.03
Const $freearc = "unarc.exe"														;0.666
Const $fsb = "fsbext.exe" 															;0.3.3
Const $gcf = $archdir & "GCFScape.exe" ;x64											;1.8.2
Const $hlp = "helpdeco.exe" 														;2.1
Const $img = "EXTRNT.EXE" 															;2.10
Const $inno = "innounp.exe" 														;0.45
Const $is6cab = "i6comp.exe" 														;0.2
Const $isxunp = "IsXunpack.exe" 													;0.99
Const $isz = "unisz.exe"															;?
Const $kgb = 'kgb\kgb2_console.exe'													;1.2.1.24
Const $lit = "clit.exe" 															;1.8
Const $lzo = "lzop.exe" 															;1.03
Const $lzx = "unlzx.exe" 															;1.21
Const $mht = "extractMHT.exe" 														;1.0
Const $mole = "demoleition.exe"														;0.5
Const $msi_msix = "MsiX.exe" 														;1.0
Const $msi_jsmsix = "jsMSIx.exe" 													;1.11.0704
Const $msi_lessmsi = Quote($bindir & 'lessmsi\lessmsi.exe', True)					;1.4
Const $nbh = "NBHextract.exe" 														;1.0
Const $pea = Quote($bindir & "pea.exe") 											;0.53/1.0
Const $peid = Quote($bindir & "peid.exe")											;0.95   2012/04/24
Const $quickbms = Quote($bindir & "quickbms.exe", True)								;0.6.4
Const $rai = "RAIU.EXE" 															;0.1a
Const $rar = "unrar.exe" 															;5.50
Const $rpa = "unrpa.exe"															;1.5.2
Const $sfark = "sfarkxtc.exe"														;3.0
Const $sit = Quote($bindir & "Expander.exe")										;6.0
Const $sqlite = "sqlite3.exe"														;3.10.2
Const $stix = "stix_d.exe" 															;2001/06/13
Const $swf = "swfextract.exe" 														;0.9.1
Const $trid = "trid.exe" 															;2.10	2012/05/06
Const $ttarch = "ttarchext.exe"														;0.2.4
Const $uharc = "UNUHARC06.EXE" 														;0.6b
Const $uharc04 = "UHARC04.EXE" 														;0.4
Const $uharc02 = "UHARC02.EXE" 														;0.2
Const $uif = "uif2iso.exe" 															;0.1.7c
Const $unity = "disunity.bat" 														;0.3.2
Const $unshield = "unshield.exe" 													;0.5
Const $upx = "upx.exe" 																;3.08w
Const $wise_ewise = "e_wise_w.exe" 													;2002/07/01
Const $wise_wun = "wun.exe" 														;0.90A
Const $wix = Quote($bindir & "dark\dark.exe", True)									;3.10.3.3007
Const $zip = "unzip.exe" 															;6.00
Const $zpaq = _IsWinXP("zpaqxp.exe", Quote($archdir & "zpaq.exe", True)) ;x64		;7.07
Const $zoo = "unzoo.exe" 															;4.5

; Plugins
Const $bms = "BMS.bms"
Const $dbx = "dbxplug.wcx"
Const $gaup = "gaup_pro.wcx"
Const $ie = "InstExpl.wcx"
Const $iso = "Iso.wcx"
Const $mht_plug = "MhtUnPack.wcx"
Const $msi_plug = "msi.wcx"
Const $sis = "PDunSIS.wcx"

; Other
Const $mtee = Quote($bindir & "mtee.exe")
Const $wtee = Quote($bindir & "wtee.exe")
Const $tee = @OSVersion = "WIN_10"? $wtee: $mtee
Const $mediainfo = $bindir & "MediaInfo.dll"										; 0.7.72
Const $xor = "xor.exe"

; Not included binaries
Const $arc_conv = "arc_conv.exe"
Const $bootimg = "bootimg.exe"
Const $ci = "ci-extractor.exe"
Const $crage = Quote($bindir & "crass-0.4.14.0\crage.exe", True)
Const $dcp = "dcp_unpacker.exe"
Const $dgca = "dgcac.exe"
Const $ffmpeg = Quote($archdir & "ffmpeg.exe", True)	;x64
Const $iscab = "iscab.exe"
Const $is5cab = "i5comp.exe"
Const $mpq = "mpq.wcx" & $reg64
Const $rgss3 = Quote($bindir & "RPGDecrypter.exe")
Const $sim = "sim_unpacker.exe"
Const $thinstall = Quote($bindir & "Extractor.exe", True)
Const $unreal = "umodel.exe"

; Define registry keys
Global Const $reg = "HKCU" & $reg64 & "\Software\UniExtract"
Global Const $regcurrent = "HKCU" & $reg64 & "\Software\Classes\*\shell\"
Global Const $regall = "HKCR" & $reg64 & "\*\shell\"
Global $reguser = $regcurrent

; Define context menu commands
; On top to make remove via command line parameter possible
; shell	| commandline parameter | translation
Global $CM_Shells[5][3] = [ _
	['uniextract_files', '', 'EXTRACT_FILES'], _
	['uniextract_here', ' .', 'EXTRACT_HERE'], _
	['uniextract_sub', ' /sub', 'EXTRACT_SUB'], _
	['uniextract_last', ' /last', 'EXTRACT_LAST'], _
	['uniextract_scan', ' /scan', 'SCAN_FILE'] _
]

ReadPrefs()

Cout("Starting " & $name & " " & $version)

; Set environment options for Windows NT, automatically reverted on exit
;~ Cout("Setting path environment variable: " & EnvSet("PATH", $archdir & ';' & $bindir & ';' & EnvGet("PATH")))

ParseCommandLine()

; Create tray menu items
$Tray_Statusbox = TrayCreateItem(t('PREFS_HIDE_STATUS_LABEL'))
If $NoBox Then TrayItemSetState(-1, $TRAY_CHECKED)
TrayCreateItem("")
$Tray_Exit = TrayCreateItem(t('MENU_FILE_QUIT_LABEL'))

TrayItemSetOnEvent($Tray_Statusbox, "Tray_Statusbox")
TrayItemSetOnEvent($Tray_Exit, "Tray_Exit")
TraySetToolTip($name)
TraySetClick(8)

; If no file passed, display GUI to select file and set options
If $prompt Then
	; Make sure a language file exists
	If Not FileExists($sEnglishLangFile) And Not FileExists($langdir) Then
		If MsgBox(48+4, $title, "No language file found." & @CRLF & @CRLF & "Do you want Universal Extractor to download all missing files?") Then _
			CheckUpdate(True, False, $UPDATE_HELPER)
	EndIf

	; Check if Universal Extractor is started the first time
	If $ID = "" Or StringIsSpace($ID) Then
		$ID = StringRight(String(_Crypt_EncryptData(Random(10000, 1000000), @ComputerName & Random(10000, 1000000), $CALG_AES_256)), 25)
		Cout("Created User ID: " & $ID)
		SavePref("ID", $ID)
		GUI_FirstStart()
		While $FS_GUI
			Sleep(250)
		WEnd
	EndIf

	CheckUpdate(True, True)
	CreateGUI()

	While 1
		If Not $guimain Then ExitLoop
		Sleep(100)
	WEnd
EndIf

; Prevent multiple instances to avoid errors
; Only necessary when extraction starts
; Do not do this in StartExtraction, the function can be called twice
$hMutex = _Singleton($name & " " & $version, 1)
If $hMutex = 0 And $extract Then
	AddToBatch()
	terminate($STATUS_SILENT, '', '')
EndIf

StartExtraction()

; -------------------------- Begin Custom Functions ---------------------------

; Start extraction process
Func StartExtraction()
	Cout("------------------------------------------------------------")
	$iUnicodeMode = False

	FilenameParse($file)

	; Collect file information, for log/feedback only
	Local $return = Round(FileGetSize($file) / 1048576, 2)
	Cout("File size: " & ($return < 1? Round(FileGetSize($file) / 1024, 2) & " KB": $return & " MB"))
	Cout("Created " & FileGetTime($file, 1, 1) & ", modified " & FileGetTime($file, 0, 1))

	; Set full output directory
	If $outdir = '/sub' Then
		$outdir = $initoutdir
	ElseIf $outdir = '/last' Then
		$outdir = GetLastOutdir()
	ElseIf StringMid($outdir, 2, 1) <> ":" Then
		If StringLeft($outdir, 1) == '\' And StringMid($outdir, 2, 1) <> '\' Then
			$outdir = StringLeft($filedir, 2) & $outdir
		ElseIf StringLeft($outdir, 2) <> '\\' Then
			$outdir = _PathFull($filedir & '\' & $outdir)
		EndIf
	EndIf
	Cout("Output directory: " & $outdir)

	; Update history
	If $history Then
		WriteHist($HISTORY_FILE, $file)
		WriteHist($HISTORY_DIR, $outdir)
	EndIf

	; Set filename as tray icon tooltip and event handler
	TraySetToolTip($filenamefull)
	TraySetOnEvent($TRAY_EVENT_PRIMARYUP, "Tray_ShowHide")

	MoveInputFileIfNecessary()

	; Reset variables
	$isexe = False
	$exefailed = False
	$tridfailed = False
	$innofailed = False
	$arjfailed = False
	$7zfailed = False
	$zipfailed = False
	$iefailed = False
	$isfailed = False
	$gamefailed = False
	$unpackfailed = False
	$testinno = False
	$test7z = False
	$testzip = False
	$testie = False
	$filetype = ""

	; If an extractor is specified via command line parameter, we simply use that without scanning
	If $sArcTypeOverride Then Return extract($sArcTypeOverride, $sArcTypeOverride & " " & t('TERM_FILE'))

	; Extract contents from known file types

	; UniExtract uses four methods of detection (in order):
	; 1. File extensions for special cases
	; 2. Binary file analysis of files using TrID if file extension is not .exe
	; 3. Binary file analysis of PE (executable) files using Exeinfo PE
	; 4. Extra analysis using PeID if executable is not recognized by Exeinfo PE
	; 5. Binary file analysis of files using TrID
	; 6. File extensions

	; First, check for file extensions that require special actions
	InitialCheckExt()

	; If file is an .exe, scan with Exeinfo PE and PEiD
	If $fileext = "exe" Or $fileext = "dll" Then IsExe()

	; Scan file with TrID, if file is not an .exe
	filescan($file, $extract)

	; Display file information and terminate if scan only mode
	If Not $extract Then
		MediaFileScan($file)
		terminate($STATUS_FILEINFO, "", "")
	EndIf

	; Else perform additional extraction methods
	CheckIso()
	CheckGame()

	; Use file extension if signature not recognized
	CheckExt()

	check7z()

	; Cannot determine filetype, all checks failed - abort
	_DeleteTrayMessageBox()
	terminate($STATUS_UNKNOWNEXT, $file, "")
EndFunc

; Extract if exe file detected
Func IsExe()
	If $exefailed Then Return
	Cout("File seems to be executable")

	; Just for fun
	If $file = @ScriptFullPath Then
		$filetype = $name
		terminate($STATUS_NOTPACKED, $file, "")
	EndIf

	; Check executable using Exeinfo PE
	advexescan()

	; Check executable using PEiD
	exescan($file, 'ext', $extract) ; Userdb is much faster, so do that first
	exescan($file, 'hard', $extract)

	If Not $extract Then Return

	; Perform additional tests if necessary
	If $testinno And Not $innofailed Then checkInno()
	If $testzip Then checkZip()
	If $testie And Not $iefailed Then checkIE()
	If $test7z Then check7z()

	If Not $iefailed Then checkIE()

	CheckGame()

	; Make sure TrID doesn't call IsExe again
	$exefailed = True

	; Scan using TrID
	filescan($file)

	; Exit with unknown file type
	terminate($STATUS_UNKNOWNEXE, $file, $filetype)
EndFunc

; Parse filename
Func FilenameParse($f)
	$file = _PathFull($f)
	$filedir = StringLeft($f, StringInStr($f, '\', 0, -1) - 1)
	$filename = StringTrimLeft($f, StringInStr($f, '\', 0, -1))
	If StringInStr($filename, '.') Then
		$fileext = StringTrimLeft($filename, StringInStr($filename, '.', 0, -1))
		$filename = StringTrimRight($filename, StringLen($fileext) + 1)
		$initoutdir = $filedir & '\' & StringReplace($filename, ".", "_")
	Else
		$fileext = ''
		$initoutdir = $filedir & '\' & $filename & '_' & t('TERM_UNPACKED')
	EndIf
	$filenamefull = $filename & "." & $fileext
;~ 	Cout("FilenameParse: " & @CRLF & "Raw input: " & $f & @CRLF & "FileName: " & $filename & @CRLF & "FileExt: " & $fileext & @CRLF & "FileDir: " & $filedir & @CRLF & "InitOutDir: " & $initoutdir)
EndFunc

; Parse string for environmental variables and return expanded output
Func EnvParse($string)
	$arr = StringRegExp($string, "%.*%", 2)
	For $i = 0 To UBound($arr) - 1
		$string = StringReplace($string, $arr[$i], EnvGet(StringReplace($arr[$i], "%", "")))
	Next
	Return $string
EndFunc   ;==>EnvParse

; Translate text
Func t($t, $aVars = 0, $lang = $language)
	$return = IniRead($lang = 'English'? $sEnglishLangFile: $langdir & '\' & $lang & '.ini', 'UniExtract', $t, '')
	If $return == '' Then
		Cout("Translation not found for term " & $t)
		$return = IniRead($sEnglishLangFile, 'UniExtract', $t, '')
		If $return = '' Then
			Cout("Warning: term " & $t & " is not defined")
			Return $t
		EndIf
	EndIf

	If Not StringInStr($return, "%") Then Return $return
	$return = StringReplace($return, '%name', $name)
	$return = StringReplace($return, '%n', @CRLF)
	$return = StringReplace($return, '%t', @TAB)

	If $aVars == 0 Then Return $return

	If IsArray($aVars) Then
		For $i = 0 To UBound($aVars) - 1
			$return = StringReplace($return, '%' & $i+1, $aVars[$i])
		Next
	Else
		$return = StringReplace($return, '%1', $aVars)
	EndIf

	Return $return
EndFunc

; Parse command line
Func ParseCommandLine()
	If $cmdline[0] = 0 Then
		$prompt = 1
		Return
	EndIf

	Cout("Command line parameters: " & _ArrayToString($cmdline, " ", 1))

	If _ArraySearch($cmdline, "/silent") > -1 Then $silentmode = True

	If $cmdline[1] = "/prefs" Then
		GUI_Prefs()
		While $guiprefs
			Sleep(250)
		WEnd
		terminate($STATUS_SILENT)

	ElseIf $cmdline[1] = "/help" Or $cmdline[1] = "/?" Or $cmdline[1] = "-h" Or $cmdline[1] = "/h" Or $cmdline[1] = "-?" Or $cmdline[1] = "--help" Then
		terminate($STATUS_SYNTAX, "", $cmdline[0] > 1)

	ElseIf $cmdline[1] = "/afterupdate" Then
		_AfterUpdate()

	ElseIf $cmdline[1] = "/update" Then
		CheckUpdate()
		terminate($STATUS_SILENT)

	ElseIf $cmdline[1] = "/updatehelper" Then
		CheckUpdate($UPDATEMSG_SILENT, False, $UPDATE_HELPER)
		$prompt = 1

	ElseIf $cmdline[1] = "/plugins" Then
		$prompt = 1
		GUI_Plugins()

	ElseIf $cmdline[1] = "/remove" Then
		; Completely delete registry entries, used by uninstaller
		_IsWin7()
		GUI_ContextMenu_remove()
		GUI_ContextMenu_fileassoc(0)
		terminate($STATUS_SILENT)

	ElseIf $cmdline[1] = "/batchclear" Then
		GUI_Batch_Clear()
		terminate($STATUS_SILENT)

	Else
		If Not FileExists($cmdline[1]) Then	terminate($STATUS_INVALIDFILE, $cmdline[1], "")

		$file = $cmdline[1]
		If $cmdline[0] > 1 Then
			; Scan only
			If $cmdline[2] = "/scan" Then
				$extract = False
				$Log = False
			Else ; Outdir specified
				$outdir = $cmdline[2]
				; When executed from context menu, opening the outdir is not wanted
				$OpenOutDir = 0
			EndIf

			If $cmdline[0] > 2 And StringLeft($cmdline[3], 6) = "/type=" Then
				$sArcTypeOverride = StringTrimLeft($cmdline[3], 6)
				If StringLen($sArcTypeOverride) < 1 Then
					; TODO: Display type select GUI
				EndIf
			EndIf
		Else
			$prompt = 1
		EndIf

		If _ArraySearch($cmdline, "/batch") > -1 Then
			AddToBatch()
			terminate($STATUS_SILENT, '', '')
		EndIf
	EndIf
EndFunc

; Read complete preferences
Func ReadPrefs()
	; Select ini file
	Local Const $globalIni = @ScriptDir & "\UniExtract.ini"
	Local Const $userIni = $settingsdir & "\UniExtract.ini"

	If FileExists($userIni) Then
		Cout("Using current user's settings")
	Else
		; Test file permissions, e.g. when UniExtract is in program files directory,
		; user settings are stored in %appdata% due to permission issues
		If CanAccess($globalIni) Then
			Cout("Using global settings")
			Global $settingsdir = @ScriptDir
		Else
			Cout("Cannot write to " & $globalIni & ", using %appdata%")
			FileCopy($globalIni, $userIni, 8)
		EndIf
	EndIf

	; Setup paths
	Global $prefs = $settingsdir & "\UniExtract.ini"
	Global $batchQueue = $settingsdir & "\batch.queue"
	Global $logdir = $settingsdir & "\log\"
	Global $userDefDir = $settingsdir & "\def\"
	Global $aDefDirs[] = [$userDefDir, $defdir]
	Global $fileScanLogFile = $logdir & "filescan.txt"
	Global Const $sPasswordFile = $settingsdir & "\passwords.txt"

	LoadPref("consoleoutput", $Opt_ConsoleOutput)
	LoadPref("language", $language, False)
	LoadPref("batchqueue", $batchQueue, False)
	If $batchQueue Then $batchQueue = _PathFull($batchQueue, $settingsdir)
	LoadPref("filescanlogfile", $fileScanLogFile, False)
	If Not @error Then $fileScanLogFile = _PathFull($fileScanLogFile, $settingsdir)
	LoadPref("batchenabled", $batchEnabled, 0)
	LoadPref("history", $history)
	LoadPref("appendext", $appendext)
	LoadPref("warnexecute", $warnexecute)
	LoadPref("nostatusbox", $NoBox)
	If Not $NoBox Then LoadPref("hidestatusboxiffullscreen", $bHideStatusBoxIfFullscreen)
	LoadPref("openfolderafterextr", $OpenOutDir)
	LoadPref("deletesourcefile", $iDeleteOrigFile)
	LoadPref("freespacecheck", $freeSpaceCheck)

	LoadPref($STATUS_TIMEOUT, $Timeout)
	$Timeout *= 1000
	If $Timeout < 10000 Then $Timeout = 60000

	LoadPref("keepoutputdir", $KeepOutdir)
	LoadPref("keepopen", $KeepOpen)
	LoadPref("feedbackprompt", $FB_ask)
	LoadPref("log", $Log)
	LoadPref("checkgame", $CheckGame)
	LoadPref("sendstats", $bSendStats)
	LoadPref("extract", $extract)
	LoadPref("unicodecheck", $checkUnicode)
	LoadPref("extractvideotrack", $bExtractVideo)
	LoadPref("silentmode", $silentmode)
	LoadPref("storeguiposition", $StoreGUIPosition)

	If $StoreGUIPosition Then
		LoadPref("posx", $posx)
		LoadPref("posy", $posy)
	EndIf

	LoadPref("statusposx", $trayX)
	LoadPref("statusposy", $trayY)
	LoadPref("addassocenabled", $addassocenabled)
	LoadPref("addassoc", $addassoc, False)
	LoadPref("addassocallusers", $addassocallusers)
	LoadPref("topmost", $iTopmost)
	If $iTopmost Then $iTopmost = $WS_EX_TOPMOST

	LoadPref("updateinterval", $updateinterval)
	If $updateinterval < 1 Then $updateinterval = 1
	LoadPref("lastupdate", $lastupdate, False)
	LoadPref("ID", $ID, False)

	If Not HasTranslation($language) Then
		$language = _WinAPI_GetLocaleInfo(_WinAPI_GetSystemDefaultUILanguage(), $LOCALE_SENGLANGUAGE)
		If Not HasTranslation($language) Then $language = _GetOSLanguage()
		If Not HasTranslation($language) Then $language = "English"
		Cout("Language set to " & $language)
		SavePref('language', $language)
	EndIf

	Cout("Program directory: " & @ScriptDir)
	Cout("Finished loading preferences from file " & $prefs)
EndFunc

; Write complete preferences
Func WritePrefs()
	Cout("Saving preferences")
	SavePref('history', $history)
	SavePref('language', $language)
	SavePref('appendext', $appendext)
	SavePref('warnexecute', $warnexecute)
	SavePref('nostatusbox', $NoBox)
	SavePref("hidestatusboxiffullscreen", $bHideStatusBoxIfFullscreen)
	SavePref('openfolderafterextr', $OpenOutDir)
	SavePref('deletesourcefile', $iDeleteOrigFile)
	SavePref('freespacecheck', $freeSpaceCheck)
	SavePref('unicodecheck', $checkUnicode)
	SavePref('feedbackprompt', $FB_ask)
	SavePref('consoleoutput', $Opt_ConsoleOutput)
	SavePref('log', $Log)
	SavePref('checkgame', $CheckGame)
	SavePref('sendstats', $bSendStats)
	SavePref("extractvideotrack", $bExtractVideo)
	SavePref('storeguiposition', $StoreGUIPosition)
	SavePref('timeout', $Timeout / 1000)
	SavePref('updateinterval', $updateinterval)
	SavePref("topmost", Number($iTopmost > 0))
EndFunc

; Save single preference
Func SavePref($name, $value)
	IniWrite($prefs, "UniExtract Preferences", $name, $value)
	Cout("Saving: " & $name & " = " & $value)
EndFunc

; Load single preference
Func LoadPref($name, ByRef $value, $int = True)
	Local $return = IniRead($prefs, "UniExtract Preferences", $name, "#Error#")
	If @error Or $return = "#Error#" Then
		Cout("Failed to read option " & $name)
		SavePref($name, $value)
		Return SetError(1, "", -1)
	EndIf

	If $int Then
		$value = Int($return)
	Else
		$value = $return
	EndIf

	Cout("Option: " & $name & " = " & $value)
EndFunc

; Read history
Func ReadHist($sSection)
	Local $items

	; Read from .ini file
	For $i = 0 To 9
		$value = IniRead($prefs, $sSection, $i, "")
		If $value <> "" Then $items &= '|' & $value
	Next

	Return StringTrimLeft($items, 1)
EndFunc

; Write history
Func WriteHist($sSection, $new)
	$histarr = StringSplit(ReadHist($sSection), '|')
	IniWrite($prefs, $sSection, "0", $new)
	If $histarr[1] == "" Then Return
	For $i = 1 To $histarr[0]
		If $i > 9 Then ExitLoop
		If $histarr[$i] = $new Then
			IniDelete($prefs, $sSection, String($i))
			ContinueLoop
		EndIf
		IniWrite($prefs, $sSection, String($i), $histarr[$i])
	Next
EndFunc

; Read last used directory from history and terminate if an error occurs
Func GetLastOutdir()
	$return = IniRead($prefs, $HISTORY_DIR, "0", -1)
	If $return <> -1 Then Return $return

	MsgBox(48, $title, t('NO_HISTORY', CreateArray($file, StringReplace(t('PREFS_HISTORY_LABEL'), "&", ""))))
	terminate($STATUS_SILENT)
EndFunc

; Scan file using TrID
Func filescan($f, $analyze = 1)
	If $tridfailed Then Return

	; Scan file using unix file tool
	advfilescan($f)

	_CreateTrayMessageBox(t('SCANNING_FILE', "TrID"))
	Cout("Starting filescan using TrID")

	If $extract Then
		Local $return = ""
		$hDll = DllOpen($bindir & "TrIDLib.dll")
		DllCall($hDll, "int", "TrID_LoadDefsPack", "str", $bindir)
		DllCall($hDll, "int", "TrID_SubmitFileA", "str", $f)
		DllCall($hDll, "int", "TrID_Analyze")

		Local $aReturn = DllCall($hDll, "int", "TrID_GetInfo", "int", 1, "int", 0, "str", $return)
		If $aReturn[0] = 0 Then
			Cout("Unknown filetype!")
			Return _DeleteTrayMessageBox()
		EndIf

		For $i = 1 To $aReturn[0]
			$aReturn = DllCall($hDll, "int", "TrID_GetInfo", "int", 2, "int", $i, "str", $return)
			$filetype &= $aReturn[3] & @CRLF
			If $analyze Then tridcompare($aReturn[3])
		Next

		If $appendext Then
			$aReturn = DllCall($hDll, "int", "TrID_GetInfo", "int", 3, "int", 1, "str", $return)
			$aReturn[3] = StringLower($aReturn[3])
			Local $ret = $filedir & "\" & $filename & "." & $aReturn[3]
			If $ret <> $file Then
				Cout("Changing file extension from ." & $fileext & " to ." & $aReturn[3])
				If FileMove($file, $ret) Then FilenameParse($file)
			EndIf
		EndIf

	Else ; Run TrID and fetch output to include additional information about the file type
		$return = StringSplit(FetchStdout($trid & ' "' & $f & '"' & ($appendext ? " -ce" : "") & ($analyze ? "" : " -v"), $filedir, @SW_HIDE, 0, True, False), @CRLF)
		Local $filetype_curr = ""
		For $i = 1 To UBound($return) - 1
			If StringInStr($return[$i], "%") Or (Not $analyze And (StringInStr($return[$i], "Related URL") Or StringInStr($return[$i], "Remarks"))) Then _
				$filetype_curr &= $return[$i] & @CRLF
		Next

		If $filetype_curr <> "" Then
			$filetype &= $filetype_curr
			If $analyze Then tridcompare($filetype_curr)
		EndIf
	EndIf

	$filetype &= @CRLF
	$tridfailed = True
EndFunc

; Additional file scan using unix file tool
Func advfilescan($f)
	Local $filetype_curr = ""

	_CreateTrayMessageBox(t('SCANNING_FILE', "Unix File Tool"))

	Cout("Start filescan using unix file tool")
	$filetype_curr = StringReplace(StringReplace(FetchStdout($filetool & ' "' & $f & '"', $filedir, @SW_HIDE), $f & "; ", ""), @CRLF, "")
	If $filetype_curr And $filetype_curr <> "data" Then $filetype &= $filetype_curr & @CRLF & @CRLF

	_DeleteTrayMessageBox()

	If Not $extract Then
		; Text files often lead to wrong detection, so renaming them is not a good idea
		If $appendext And (StringInStr($filetype_curr, "text", 0) Or StringInStr($filetype_curr, "ASCII", 0)) Then $appendext = False
		Return
	EndIf

	filecompare($filetype_curr)
EndFunc

; Compare unix file tool's return to supported file types
Func filecompare($filetype_curr)
	Select
		Case StringInStr($filetype_curr, "7 zip archive data") Or StringInStr($filetype_curr, "7-zip archive data")
			extract("7z", '7-Zip ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "RAR archive data")
			extract("rar", 'RAR ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "lzip compressed data")
			extract("lz", "LZIP " & t('TERM_COMPRESSED') & " " & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "Zip archive data") And Not StringInStr($filetype_curr, "7")
			extract("zip", 'ZIP ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "StuffIt Archive")
			extract("sit", 'StuffIt ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "UHarc archive data", 0)
			extract("uha", 'UHARC ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "Symbian installation file", 0)
			extract('qbms', 'SymbianOS ' & t('TERM_INSTALLER'), $sis)
		Case StringInStr($filetype_curr, "Zoo archive data", 0)
			extract("zoo", 'ZOO ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "MS Outlook Express DBX file", 0)
			extract('qbms', 'Outlook Express ' & t('TERM_ARCHIVE'), $dbx)
		Case StringInStr($filetype_curr, "bzip2 compressed data", 0)
			extract("7z", 'bzip2 ' & t('TERM_COMPRESSED'), "bz2")
		Case StringInStr($filetype_curr, "ASCII cpio archive", 0)
			extract("7z", 'CPIO ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "gzip compressed", 0)
			extract("7z", 'gzip ' & t('TERM_COMPRESSED'), "gz")
		Case StringInStr($filetype_curr, "LZX compressed archive", 0)
			extract("lzx", 'LZX ' & t('TERM_COMPRESSED'))
		Case StringInStr($filetype_curr, "ARJ archive", 0)
			extract("7z", 'ARJ ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "POSIX tar archive", 0)
			extract("tar", 'Tar ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "LHa", 0) And StringInStr($filetype_curr, "archive data", 0)
			extract("7z", 'LZH ' & t('TERM_COMPRESSED'))
		Case StringInStr($filetype_curr, "Macromedia Flash data", 0)
			extract("swf", 'Shockwave Flash ' & t('TERM_CONTAINER'))
		Case StringInStr($filetype_curr, "PowerISO Direct-Access-Archive", 0)
			extract("daa", 'DAA/GBI ' & t('TERM_IMAGE'))
		Case StringInStr($filetype_curr, "sfArk compressed Soundfont")
			extract("sfark", 'sfArk ' & t('TERM_COMPRESSED'))
		Case StringInStr($filetype_curr, "SQLite", 0)
			extract("sqlite", 'SQLite ' & t('TERM_FILE'))
		Case StringInStr($filetype_curr, "XZ compressed data")
			extract("7z", 'XZ ' & t('TERM_COMPRESSED'), "xz")
		Case StringInStr($filetype_curr, "MS Windows HtmlHelp Data")
			extract("chm", 'Compiled HTML ' & t('TERM_HELP'))
		Case StringInStr($filetype_curr, "MoPaQ", 0)
			HasPlugin($mpq)
			extract("qbms", 'MPQ ' & t('TERM_ARCHIVE'), $mpq)
		Case (StringInStr($filetype_curr, "RIFF", 0) And Not StringInStr($filetype_curr, "WAVE audio", 0)) Or _
			 StringInStr($filetype_curr, "MPEG v", 0) Or StringInStr($filetype_curr, "MPEG sequence") Or _
			 StringInStr($filetype_curr, "Microsoft ASF") Or StringInStr($filetype_curr, "GIF image") Or _
			 StringInStr($filetype_curr, "PNG image") Or StringInStr($filetype_curr, "MNG video")
			extract("video", t('TERM_VIDEO') & ' ' & t('TERM_FILE'))
		Case StringInStr($filetype_curr, "AAC,")
			extract("audio", 'AAC ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))
		Case StringInStr($filetype_curr, "FLAC audio")
			extract("audio", 'FLAC ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))
		Case StringInStr($filetype_curr, "Ogg data, Vorbis audio")
			extract("audio", 'OGG Vorbis ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))
		Case StringInStr($filetype_curr, "Audio file", 0) Or StringInStr($filetype_curr, "Dolby Digital stream", 0)
			extract("audio", t('TERM_AUDIO') & ' ' & t('TERM_FILE'))
		Case StringInStr($filetype_curr, "ISO", 0) And StringInStr($filetype_curr, "filesystem", 0)
			CheckIso()
		Case Else
			UserDefCompare($aFileDefinitions, $filetype_curr, "File")
	EndSelect

	; Not extractable filetypes
	If StringInStr($filetype_curr, "CDF V2 document") Then Return

	If (StringInStr($filetype_curr, "text", 0) And (StringInStr($filetype_curr, "CRLF", 0) Or _
	  StringInStr($filetype_curr, "long lines", 0) Or StringInStr($filetype_curr, "ASCII", 0)) Or _
	  StringInStr($filetype_curr, "batch file") Or StringInStr($filetype_curr, "XML") Or _
	  StringInStr($filetype_curr, "HTML") Or StringInStr($filetype_curr, "source", 0) Or _
	  StringInStr($filetype_curr, "Rich ")) Or _
	  StringInStr($filetype_curr, "image", 0) Or StringInStr($filetype_curr, "icon resource", 0) Or _
	 (StringInStr($filetype_curr, "bitmap", 0) And Not StringInStr($filetype_curr, "MGR bitmap")) Or _
	  StringInStr($filetype_curr, "WAVE audio", 0) Or StringInStr($filetype_curr, "boot sector;") Or _
	  StringInStr($filetype_curr, "shortcut", 0) Or StringInStr($filetype_curr, "empty") Then _
		terminate($STATUS_NOTPACKED, $file, "")
EndFunc

; Compare TrID's return to supported file types
Func tridcompare($filetype_curr)
	Cout("--> " & $filetype_curr)
	Select
		Case StringInStr($filetype_curr, "7-Zip compressed archive", 0)
			extract("7z", '7-Zip ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "ACE compressed archive", 0) _
			 Or StringInStr($filetype_curr, "ACE Self-Extracting Archive", 0)
			extract("ace", 'ACE ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Android boot image")
			extract("bootimg", ' Android boot ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "ALZip compressed archive")
			CheckAlz()

		Case StringInStr($filetype_curr, "LZIP compressed archive")
			extract("lz", "LZIP " & t('TERM_COMPRESSED'))

		Case StringInStr($filetype_curr, "FreeArc compressed archive", 0)
			extract("freearc", 'FreeArc ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "ARJ compressed archive", 0)
			extract("7z", 'ARJ ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "BCM compressed file")
			extract("bcm", 'BCM ' & t('TERM_COMPRESSED'))

		Case StringInStr($filetype_curr, "bzip2 compressed archive", 0)
			extract("7z", 'bzip2 ' & t('TERM_COMPRESSED'), "bz2")

		Case StringInStr($filetype_curr, "Broken Age package", 0)
			CheckGame(False)

		Case StringInStr($filetype_curr, "Microsoft Cabinet Archive", 0) Or StringInStr($filetype_curr, "IncrediMail letter/ecard", 0)
			extract("cab", 'Microsoft CAB ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Magic ISO Universal Image Format", 0)
			extract("uif", 'UIF ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "CDImage", 0) Or StringInStr($filetype_curr, "null bytes", 0)
			CheckIso()
			check7z()

		Case StringInStr($filetype_curr, "Compiled HTML Help File", 0)
			extract("chm", 'Compiled HTML ' & t('TERM_HELP'))

		Case StringInStr($filetype_curr, "CPIO Archive", 0)
			extract("7z", 'CPIO ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "PowerISO Direct-Access-Archive", 0) Or StringInStr($filetype_curr, "gBurner Image", 0)
			extract("daa", 'DAA/GBI ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "Debian Linux Package", 0)
			extract("7z", 'Debian ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Wintermute Engine data", 0)
			extract("dcp", 'Wintermute Engine ' & t('TERM_GAME') & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "DGCA Digital G Codec Archiver", 0)
			extract("dgca", 'DGCA ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Disk Image (Macintosh)", 0)
			extract("7z", 'Macintosh ' & t('TERM_DISK') & ' ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "FMOD Sample Bank Format")
			extract("fsb", 'FMOD ' & t('TERM_CONTAINER'))

		Case StringInStr($filetype_curr, "Gentee Installer executable", 0) Or StringInStr($filetype_curr, "Installer VISE executable", 0) Or _
			 StringInStr($filetype_curr, "Setup Factory", 0)
			checkIE()

		Case StringInStr($filetype_curr, "GZipped", 0)
			extract("7z", 'gzip ' & t('TERM_COMPRESSED'), "gz")

		Case StringInStr($filetype_curr, "Windows Help File", 0)
			extract("hlp", 'Windows ' & t('TERM_HELP'))

		Case StringInStr($filetype_curr, "Generic PC disk image", 0)
			CheckIso()
			check7z()
			extract("img", 'Floppy ' & t('TERM_DISK') & ' ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "Inno Setup installer", 0)
			checkInno()

		Case StringInStr($filetype_curr, "InstallShield archive", 0)
			extract("is3arc", 'InstallShield 3.x ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "InstallShield compressed archive", 0)
			extract("iscab", 'InstallShield CAB ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "ISo Zipped format", 0)
			extract("isz", 'Zipped ISO ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "KiriKiri Adventure Game System Package", 0)
			extract("arc_conv", 'KiriKiri Adventure Game System ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "KGB archive", 0)
			extract("kgb", 'KGB ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "LHARC/LZARK compressed archive", 0)
			extract("7z", 'LZH ' & t('TERM_COMPRESSED'))

		Case StringInStr($filetype_curr, "Livemaker Engine main game executable", 0)
			extract("crage", 'Livemaker ' & t('TERM_GAME') & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "lzop compressed", 0)
			extract("lzo", 'LZO ' & t('TERM_COMPRESSED'))

		Case StringInStr($filetype_curr, "LZX Amiga compressed archive", 0)
			extract("lzx", 'LZX ' & t('TERM_COMPRESSED'))

		Case StringInStr($filetype_curr, "Microsoft Internet Explorer Web Archive") Or StringInStr($filetype_curr, "MIME HTML archive format")
			extract("mht", 'MHTML ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Microsoft Windows Installer merge module", 0)
			extract("msm", 'Windows Installer (MSM) ' & t('TERM_MERGE_MODULE'))

		Case StringInStr($filetype_curr, "Microsoft Windows Installer", 0)
			extract("msi", 'Windows Installer (MSI) ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Microsoft Windows Installer patch", 0)
			extract("msp", 'Windows Installer (MSP) ' & t('TERM_PATCH'))

		Case StringInStr($filetype_curr, "MPQ Archive - Blizzard game data", 0)
			HasPlugin($mpq)
			extract("qbms", 'MPQ ' & t('TERM_ARCHIVE'), $mpq)

		Case StringInStr($filetype_curr, "HTC NBH ROM Image", 0)
			extract("nbh", 'NBH ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "Outlook Express E-mail folder", 0)
			extract('qbms', 'Outlook Express ' & t('TERM_ARCHIVE'), $dbx)

		Case StringInStr($filetype_curr, "PEA compressed archive", 0)
			extract("pea", 'Pea ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "RAR Archive", 0)
			extract("rar", 'RAR ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "RPG Maker VX Ace", 0)
			extract("rgss3", "RPG Maker VX Ace " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "NScripter archive, version 1", 0)
			extract("arc_conv", "NScripter " & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "RPG Maker", 0)
			extract("arc_conv", "RPG Maker " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Smile Game Builder", 0)
			extract("sgb", "Smile Game Builder " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Telltale Games ressource archive", 0)
			extract("ttarch", "Telltale " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Wolf RPG Editor", 0)
			extract("arc_conv", "Wolf RPG Editor " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "YU-RIS Script Engine", 0)
			extract("arc_conv", "YU-RIS Script Engine " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "ClsFileLink", 0) Or StringInStr($filetype_curr, "ERISA archive file", 0)
			extract("arc_conv", t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Ethornell", 0)
			extract("ethornell", "Ethornell Engine " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Reflexive Arcade installer wrapper", 0)
			extract("inno", 'Reflexive Arcade ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Ren'Py data file", 0)
			extract("rpa", "Ren'Py " & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "RPM Linux Package", 0)
			extract("7z", 'RPM ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "sfArk compressed SoundFont")
			extract("sfark", 'sfArk ' & t('TERM_COMPRESSED'))

		Case StringInStr($filetype_curr, "StuffIT SIT compressed archive", 0)
			extract("sit", 'StuffIt ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "SymbianOS Installer", 0)
			extract('qbms', 'SymbianOS ' & t('TERM_INSTALLER'), $sis)

		Case StringInStr($filetype_curr, "Macromedia Flash Player", 0)
			extract("swf", 'Shockwave Flash ' & t('TERM_CONTAINER'))

		Case StringInStr($filetype_curr, "TAR - Tape ARchive", 0)
			extract("tar", 'Tar ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "UHARC compressed archive", 0)
			extract("uha", 'UHARC ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Unity Engine Asset file")
			extract("unity", 'Unity Engine Asset ' & t('TERM_FILE'))

		Case StringInStr($filetype_curr, "Unreal Package")
			extract("unreal", 'Unreal Engine ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Base64 Encoded file", 0)
			extract("uu", 'Base64 ' & t('TERM_ENCODED'))

		Case StringInStr($filetype_curr, "Quoted-Printable Encoded file", 0)
			extract("uu", 'Quoted-Printable ' & t('TERM_ENCODED'))

		Case StringInStr($filetype_curr, "UUencoded file", 0) Or StringInStr($filetype_curr, "XXencoded file", 0)
			extract("uu", 'UUencoded ' & t('TERM_ENCODED'))

		Case StringInStr($filetype_curr, "yEnc Encoded file", 0)
			extract("uu", 'yEnc ' & t('TERM_ENCODED'))

		Case StringInStr($filetype_curr, "Valve package", 0)
			extract("gcf", 'Valve ' & $fileext & " " & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Windows Imaging Format", 0)
			extract("7z", 'WIM ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "Wise Installer Executable", 0)
			extract("wise", 'Wise Installer ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "UNIX Compressed", 0)
			extract("7z", 'LZW ' & t('TERM_COMPRESSED'), "Z")

		Case StringInStr($filetype_curr, "xz container", 0)
			extract("7z", 'XZ ' & t('TERM_COMPRESSED'), "xz")

		Case StringInStr($filetype_curr, "ZIP compressed archive", 0) Or _
			 StringInStr($filetype_curr, "Winzip Win32 self-extracting archive", 0)
			extract("zip", 'ZIP ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Zip Self-Extracting archive", 0)
			checkInno()

		Case StringInStr($filetype_curr, "ZOO compressed archive", 0)
			extract("zoo", 'ZOO ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "ZPAQ compressed archive", 0)
			extract("zpaq", 'ZPAQ ' & t('TERM_ARCHIVE'))

		; Forced to bottom of list due to false positives
		Case StringInStr($filetype_curr, "LZMA compressed archive", 0)
			check7z()

		Case StringInStr($filetype_curr, "Enigma Virtual Box virtualized executable")
			extract("enigma", 'Enigma Virtual Box ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "InstallShield setup", 0)
			;extract("isexe", 'InstallShield ' & t('TERM_INSTALLER'))
			checkInstallShield()

		Case  StringInStr($filetype_curr, "MP3 audio", 0)
			extract("audio", 'MP3 ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))

		Case StringInStr($filetype_curr, "FLAC lossless", 0)
			extract("audio", 'FLAC ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))

		Case StringInStr($filetype_curr, "Windows Media (generic)", 0)
			extract("audio", 'Windows Media ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))

		Case StringInStr($filetype_curr, "OGG Vorbis audio", 0)
			extract("audio", 'OGG Vorbis ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))

		Case StringInStr($filetype_curr, "Smacker movie/video")
			extract("video_convert", t('TERM_VIDEO') & ' ' & t('TERM_FILE'))

		Case StringInStr($filetype_curr, "Video") Or StringInStr($filetype_curr, "QuickTime Movie") Or _
			 StringInStr($filetype_curr, "Matroska") Or StringInStr($filetype_curr, "Material Exchange Format") Or _
			 StringInStr($filetype_curr, "Windows Media (generic)") Or StringInStr($filetype_curr, "GIF animated")
			extract("video", t('TERM_VIDEO') & ' ' & t('TERM_FILE'))

		Case StringInStr($filetype_curr, "null bytes") Or (StringInStr($filetype_curr, "phpMyAdmin SQL dump"))
			terminate($STATUS_NOTPACKED, $file, "")

		; Not supported filetypes
		Case StringInStr($filetype_curr, "Long Range ZIP") Or StringInStr($filetype_curr, "Kremlin Encrypted File")
			terminate($STATUS_NOTSUPPORTED, $file, "")

		; Check for .exe file, only when fileext not .exe
		Case Not $isexe And (StringInStr($filetype_curr, 'Executable', 0) Or StringInStr($filetype_curr, '(.EXE)', 1))
			IsExe()

		Case Else
			UserDefCompare($aTridDefinitions, $filetype_curr, "Trid")
	EndSelect
EndFunc

; Compare file type with definitions stored in def/registry.ini
Func UserDefCompare(ByRef $aDefinitions, $filetype_curr, $sSection)
	For $dir In $aDefDirs
		If UBound($aDefinitions) == 0 Then
			$aDefinitions = IniReadSection($dir & "registry.ini", $sSection)
			If @error Then
				Cout("Could not load custom " & $sSection & " definitions from " & $dir)
				ContinueLoop
			EndIf
			Cout("Loaded " & $sSection & " definitions from " & $dir)
		EndIf

		For $i = 1 To $aDefinitions[0][0]
			If (StringInStr($filetype_curr, $aDefinitions[$i][1])) Then extract($aDefinitions[$i][0])
		Next
	Next
EndFunc

; Scan .exe file using PEiD
Func exescan($f, $scantype, $analyze = 1)
	Local $filetype_curr = "", $bHasRegKey = True
	Local Const $key = "HKCU\Software\PEiD"

	Cout("Start filescan using PEiD")
	_CreateTrayMessageBox(t('SCANNING_EXE', "PEiD (" & $scantype & ")"))

	; Backup existing PEiD options
	Local $exsig = RegRead($key, "ExSig")
	If @error Then $bHasRegKey = False
	Local $loadplugins = RegRead($key, "LoadPlugins")
	Local $stayontop = RegRead($key, "StayOnTop")

	; Set PEiD options
	RegWrite($key, "ExSig", "REG_DWORD", 1)
	RegWrite($key, "LoadPlugins", "REG_DWORD", 0)
	RegWrite($key, "StayOnTop", "REG_DWORD", 0)

	; Analyze file
	Run($peid & ' -' & $scantype & ' "' & $f & '"', $bindir, @SW_HIDE)
	WinWait("PEiD v")
	$TimerStart = TimerInit()
	While ($filetype_curr = "") Or ($filetype_curr = "Scanning...")
		Sleep(100)
		$filetype_curr = ControlGetText("PEiD v", "", "Edit2")
		$TimerDiff = TimerDiff($TimerStart)
		If $TimerDiff > $Timeout Then ExitLoop
	WEnd
	WinClose("PEiD v")

	If $filetype_curr <> "" Then $filetype &= $filetype_curr & @CRLF & @CRLF
	Cout($filetype_curr)

	; Restore previous PEiD options
	If $bHasRegKey Then
		RegWrite($key, "ExSig", "REG_DWORD", $exsig)
		RegWrite($key, "LoadPlugins", "REG_DWORD", $loadplugins)
		RegWrite($key, "StayOnTop", "REG_DWORD", $stayontop)
	Else
		RegDelete($key)
	EndIf

	_DeleteTrayMessageBox()

	; Return filetype without matching if specified
	If Not $analyze Then Return $filetype_curr

	; Match known patterns
	Select
		; ExeInfo cannot detect big files, so PEiD is used as a fallback here
		Case StringInStr($filetype_curr, "Enigma Virtual Box")
			extract("enigma", 'Enigma Virtual Box ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "ARJ SFX", 0)
			extract("7z", t('TERM_SFX') & ' ARJ ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Borland Delphi", 0) And Not StringInStr($filetype_curr, "RAR SFX", 0)
			$testinno = True
			$testzip = True

		Case StringInStr($filetype_curr, "Gentee Installer", 0)
			checkIE()

		Case StringInStr($filetype_curr, "Inno Setup", 0)
			checkInno()

		Case StringInStr($filetype_curr, "Installer VISE", 0)
			extract("ie", 'Installer VISE ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "InstallShield", 0)
			If Not $isfailed Then extract("isexe", 'InstallShield ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "KGB SFX", 0)
			extract("kgb", t('TERM_SFX') & ' KGB ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Microsoft Visual C++", 0) And Not StringInStr($filetype_curr, "SPx Method", 0) And Not StringInStr($filetype_curr, "Custom", 0) And Not StringInStr($filetype_curr, "7.0", 0)
			$test7z = True
			$testie = True

		Case StringInStr($filetype_curr, "Microsoft Visual C++ 7.0", 0) And StringInStr($filetype_curr, "Custom", 0) And Not StringInStr($filetype_curr, "Hotfix", 0)
			extract("vssfx", 'Visual C++ ' & t('TERM_SFX') & ' ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Microsoft Visual C++ 6.0", 0) And StringInStr($filetype_curr, "Custom", 0)
			extract("vssfxpath", 'Visual C++ ' & t('TERM_SFX') & '' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Netopsystems FEAD Optimizer", 0)
			extract("fead", 'Netopsystems FEAD ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Nullsoft PiMP SFX", 0)
			checkNSIS()

		Case StringInStr($filetype_curr, "PEtite", 1)
			If Not checkArj() Then extract("ace", t('TERM_SFX') & ' ACE ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "RAR SFX", 0)
			extract("rar", t('TERM_SFX') & ' RAR ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Reflexive Arcade Installer", 0)
			extract("inno", 'Reflexive Arcade ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "RoboForm Installer", 0)
			extract("robo", 'RoboForm ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Setup Factory 6.x", 0)
			extract("ie", 'Setup Factory ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "SPx Method", 0) Or StringInStr($filetype_curr, "CAB SFX", 0)
			extract("cab", t('TERM_SFX') & ' Microsoft CAB ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "SuperDAT", 0)
			extract("superdat", 'McAfee SuperDAT ' & t('TERM_UPDATER'))

		Case StringInStr($filetype_curr, "Wise", 0) Or StringInStr($filetype_curr, "PEncrypt 4.0", 0)
			extract("wise", 'Wise Installer ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "ZIP SFX", 0)
			extract("zip", t('TERM_SFX') & ' ZIP ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "upx", 0)
			unpack($PACKER_UPX)

		Case StringInStr($filetype_curr, "aspack", 0)
			unpack($PACKER_ASPACK)

		Case StringInStr($filetype_curr, "Unable to open file", 0)
			$isexe = False

	EndSelect
EndFunc

; Scan file using Exeinfo PE
Func advexescan()
	Local $filetype_curr = ""

	Cout("Start filescan using Exeinfo PE")
	_CreateTrayMessageBox(t('SCANNING_EXE', "Exeinfo PE"))

	; Analyze file
	If $extract Then ; Use log command line for best speed
		Local Const $LogFile = $logdir & "exeinfo.log"
		RunWait($exeinfope & ' "' & $file & '*" /sx /log:"' & $LogFile & '"', $bindir, @SW_HIDE)
		$filetype_curr = _FileRead($LogFile, True)
	Else ; Run and read GUI fields to get additional information on how to extract for scan only mode
		$aReturn = OpenExeInfo()
		$TimerStart = TimerInit()

		While ($filetype_curr = "") Or StringInStr($filetype_curr, "File too big") Or StringInStr($filetype_curr, "Antivirus may slow")
			Sleep(200)
			$filetype_curr = ControlGetText($aReturn[0], "", "TEdit6")
			$TimerDiff = TimerDiff($TimerStart)
			If $TimerDiff > $Timeout Then ExitLoop
		WEnd

		$filetype_curr &= @CRLF & @CRLF & ControlGetText($aReturn[0], "", "TEdit5")

		CloseExeInfo($aReturn)
	EndIf

	_DeleteTrayMessageBox()

	$filetype_curr = StringReplace($filetype_curr, $filenamefull & "   - ", "")

	; Return if not .exe file
	If StringInStr($filetype_curr, "NOT EXE") Then Return

	; Return if file is too big
	If StringInStr($filetype_curr, "Skipped") Then Return

	If $filetype_curr <> "" Then $filetype &= $filetype_curr & @CRLF & @CRLF

	; Otherwise continue exe file procedure
	$isexe = True

	; Return filetype without matching if specified
	If Not $extract Then Return $filetype_curr

	; Match known patterns
	Select
		Case StringInStr($filetype_curr, "WinAce / SFX Factory", 0)
			extract("ace", t('TERM_SFX') & ' ACE ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Advanced Installer", 0)
			extract("ai", 'Advanced Installer ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "FreeArc", 0)
			extract("freearc", 'FreeArc ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "CreateInstall", 0)
			extract("ci", 'CreateInstall ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Excelsior Installer", 0)
			extract("ei", 'Excelsior Installer ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Ghost Installer Studio", 0)
			extract("ghost", 'Ghost Installer Studio ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Gentee Installer", 0) Or StringInStr($filetype_curr, "Installer VISE", 0) Or _
			 StringInStr($filetype_curr, "Setup Factory 6.x", 0)
			checkIE()

		Case StringInStr($filetype_curr, "Inno Setup", 0)
			checkInno()

		; Needs to be before InstallShield
		Case StringInStr($filetype_curr, "InstallAware", 0)
			extract("7z", 'InstallAware ' & t('TERM_INSTALLER') & ' ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "InstallShield", 0)
			If Not $isfailed Then extract("isexe", 'InstallShield ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "KGB SFX", 0)
			extract("kgb", t('TERM_SFX') & ' KGB ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Microsoft Visual C++ 7.0", 0) And StringInStr($filetype_curr, "Custom", 0) And Not StringInStr($filetype_curr, "Hotfix", 0)
			extract("vssfx", 'Visual C++ ' & t('TERM_SFX') & ' ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Microsoft Visual C++ 6.0", 0) And StringInStr($filetype_curr, "Custom", 0)
			extract("vssfxpath", 'Visual C++ ' & t('TERM_SFX') & '' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "www.molebox.com")
			extract("mole", 'Mole Box ' & t('TERM_CONTAINER'))

		Case StringInStr($filetype_curr, "Netopsystems FEAD Optimizer", 0)
			extract("fead", 'Netopsystems FEAD ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Nullsoft", 0);
			checkNSIS()

		Case StringInStr($filetype_curr, "RAR SFX", 0)
			extract("rar", t('TERM_SFX') & ' RAR ' & t('TERM_ARCHIVE'));

		Case StringInStr($filetype_curr, "Reflexive Arcade Installer", 0)
			extract("inno", 'Reflexive Arcade ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "RoboForm Installer", 0)
			extract("robo", 'RoboForm ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "WiX Installer")
			extract("wix", 'WiX ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Microsoft SFX CAB", 0) And StringInStr($filetype_curr, "rename file *.exe as *.cab", 0)
			Prompt(1, "FILE_COPY", $file, 1)
			Local $return = _TempFile($filedir, $filename & "_", ".cab")
			FileCopy($file, $return)
			$file = $return
			Global $iDeleteOrigFile = $OPTION_DELETE
			check7z()

		Case StringInStr($filetype_curr, "Smart Install Maker")
			extract("sim", 'Smart Install Maker ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "SPx Method", 0) Or StringInStr($filetype_curr, "Microsoft SFX CAB", 0)
			extract("cab", t('TERM_SFX') & ' Microsoft CAB ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "SuperDAT", 0)
			extract("superdat", 'McAfee SuperDAT ' & t('TERM_UPDATER'))

		Case StringInStr($filetype_curr, "Overlay :  SWF flash object ver", 0)
			extract("swfexe", 'Shockwave Flash ' & t('TERM_CONTAINER'))

		Case StringInStr($filetype_curr, "VMware ThinApp", 0) Or StringInStr($filetype_curr, "Thinstall", 0) Or StringInStr($filetype_curr, "ThinyApp Packager", 0)
			extract("thinstall", "ThinApp/Thinstall" & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Wise", 0) Or StringInStr($filetype_curr, "PEncrypt 4.0", 0)
			extract("wise", 'Wise Installer ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "ZIP SFX", 0)
			extract("zip", t('TERM_SFX') & ' ZIP ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Borland Delphi", 0) And Not StringInStr($filetype_curr, "RAR SFX", 0)
			$testinno = True
			$testzip = True

		Case StringInStr($filetype_curr, "Enigma Virtual Box")
			extract("enigma", 'Enigma Virtual Box ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, ".dmg  Mac OS", 0)
			extract("7z", 'DMG ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "upx", 0)
			unpack($PACKER_UPX)

		Case StringInStr($filetype_curr, "aspack", 0)
			unpack($PACKER_ASPACK)

			; not supported filetypes
		Case StringInStr($filetype_curr, "Autoit") Or StringInStr($filetype_curr, "Not packed , try") Or _
			 StringInStr($filetype_curr, "Microsoft Visual C# / Basic.NET")
			terminate($STATUS_NOTPACKED, $file, "")

		Case StringInStr($filetype_curr, "Astrum InstallWizard") Or StringInStr($filetype_curr, "clickteam") Or _
			 StringInStr($filetype_curr, "BitRock InstallBuilder") Or StringInStr($filetype_curr, "NE <- Windows 16bit")
			terminate($STATUS_NOTSUPPORTED, $file, "")

		; Terminate if file cannot be unpacked
		Case StringInStr($filetype_curr, "Not packed") And Not StringInStr($filetype_curr, "Microsoft Visual C++")
			terminate($STATUS_NOTPACKED, $file, "")

		; Needs to be at the end, otherwise files might not be recognized
		Case StringInStr($filetype_curr, "Microsoft Visual C++", 0) And Not StringInStr($filetype_curr, "SPx Method", 0) And Not StringInStr($filetype_curr, "Custom", 0) And Not StringInStr($filetype_curr, "7.0", 0)
			$test7z = True
			$testie = True
	EndSelect

	Cout("No matches for known Exeinfo PE types")
EndFunc

; Open ExeInfo PE and return an array containing initial registry values
Func OpenExeInfo($f = $file)
	Local $aReturn[12]

	; Backup existing Exeinfo PE options
	; WinTitle, reg key, reg key existed, backup values
	$aReturn[0] = "Exeinfo PE"
	$aReturn[1] = "HKCU\Software\ExEi-pe"
	$aReturn[2] = True
	$aReturn[3] = RegRead($aReturn[1], "ExeError")
	If @error Then $aReturn[2] = False
	$aReturn[4] = RegRead($aReturn[1], "Scan")
	$aReturn[5] = RegRead($aReturn[1], "AllwaysOnTop")
	$aReturn[6] = RegRead($aReturn[1], "Skin")
	$aReturn[7] = RegRead($aReturn[1], "Shell_integr")
	$aReturn[8] = RegRead($aReturn[1], "Log")
	$aReturn[9] = RegRead($aReturn[1], "Big_GUI")
	$aReturn[10] = RegRead($aReturn[1], "Lang")
	$aReturn[11] = RegRead($aReturn[1], "closeExEi_whenExtRun")

	; Set Exeinfo PE options
	RegWrite($aReturn[1], "ExeError", "REG_DWORD", 1)
	RegWrite($aReturn[1], "Scan", "REG_DWORD", 1)
	RegWrite($aReturn[1], "AllwaysOnTop", "REG_DWORD", 0)
	RegWrite($aReturn[1], "Skin", "REG_DWORD", 0xFFFFFFFF)
	RegWrite($aReturn[1], "Shell_integr", "REG_DWORD", 0)
	RegWrite($aReturn[1], "Log", "REG_DWORD", 0xFFFFFFFF)
	RegWrite($aReturn[1], "Big_GUI", "REG_DWORD", 0)
	RegWrite($aReturn[1], "Lang", "REG_DWORD", 0xFFFFFFFF)
	RegWrite($aReturn[1], "closeExEi_whenExtRun", "REG_DWORD", 0)

	; Execute and hide
	Run($exeinfope & ' "' & $f & '"', $bindir, @SW_MINIMIZE)
	WinWait($aReturn[0])
	WinSetState($aReturn[0], "", @SW_HIDE)

	Return $aReturn
EndFunc

; Close ExeInfo PE and restore registry settings
Func CloseExeInfo($aReturn)
	WinClose($aReturn[0])

	; Restore previous Exeinfo PE options
	If $aReturn[2] Then
		RegWrite($aReturn[1], "ExeError", "REG_DWORD", $aReturn[3])
		RegWrite($aReturn[1], "Scan", "REG_DWORD", $aReturn[4])
		RegWrite($aReturn[1], "AllwaysOnTop", "REG_DWORD", $aReturn[5])
		RegWrite($aReturn[1], "Skin", "REG_DWORD", $aReturn[6])
		RegWrite($aReturn[1], "Shell_integr", "REG_DWORD", $aReturn[7])
		RegWrite($aReturn[1], "Log", "REG_DWORD", $aReturn[8])
		RegWrite($aReturn[1], "Big_GUI", "REG_DWORD", $aReturn[9])
		RegWrite($aReturn[1], "Lang", "REG_DWORD", $aReturn[10])
		RegWrite($aReturn[1], "closeExEi_whenExtRun", "REG_DWORD", $aReturn[11])
	Else
		RegDelete($aReturn[1])
	EndIf
EndFunc

; Scan file using MediaInfo dll, only used in scan only mode
Func MediaFileScan($f)
	Local $filetype_curr = ""
	Cout("Start filescan using MediaInfo dll")
	_CreateTrayMessageBox(t('SCANNING_FILE', "MediaInfo"))

	HasPlugin($mediainfo)
	$hDll = DllOpen($mediainfo)
	If $hDll == -1 Then Return Cout("Failed to load " & $mediainfo)
	$hMI = DllCall($hDll, "ptr", "MediaInfo_New")

	DllCall($hDll, "int", "MediaInfo_Open", "ptr", $hMI[0], "wstr", $f)
	$return = DllCall($hDll, "wstr", "MediaInfo_Inform", "ptr", $hMI[0], "int", 0)

	$hMI = DllCall($hDll, "none", "MediaInfo_Delete", "ptr", $hMI[0])
	DllClose($hDll)

	Cout($return[0])

	; Return if file is not a media file
	$return = StringSplit($return[0], @CRLF, 2)
	If UBound($return) < 10 Then Return _DeleteTrayMessageBox()

	; Format returned string to align in message box
	For $i in $return
		$return = StringSplit($i, " : ", 2+1)
		If @error Then
			If Not StringIsSpace($i) Then $filetype_curr &= @CRLF & "[" & $i & "]" & @CRLF
			ContinueLoop
		EndIf
		$sType = StringStripWS($return[0], 4+2+1)
		$iLen = StringLen($sType)
		$filetype_curr &= $sType & _StringRepeat(@TAB, 3 - Floor($iLen / 10)) & (($iLen > 20 And $iLen < 25)? @TAB: "") & StringStripWS($return[1], 4+2+1) & @CRLF
	Next

	$filetype &= $filetype_curr & @CRLF & @CRLF
	_DeleteTrayMessageBox()
EndFunc

; Determine if 7-zip can extract the file
Func check7z()
	If $7zfailed Then Return
	Cout("Testing 7zip")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' 7-Zip')
	$return = FetchStdout($7z & ' l "' & $file & '"', $filedir, @SW_HIDE)

	If StringInStr($return, "Listing archive:") And Not (StringInStr($return, "Can not open the file as ") Or _
	StringInStr($return, "There are data after the end of archive")) Then
		; Failsafe in case TrID misidentifies MS SFX
		If StringInStr($return, "_sfx_manifest_") Then
			_DeleteTrayMessageBox()
			extract("hotfix", 'Microsoft ' & t('TERM_HOTFIX'))
		EndIf
		_DeleteTrayMessageBox()
		If $fileext = "exe" Then
			extract("7z", '7-Zip ' & t('TERM_INSTALLER') & ' ' & t('TERM_PACKAGE'))
		ElseIf $fileext = "iso" Then
			extract("7z", 'Iso ' & t('TERM_IMAGE'))
		ElseIf $fileext = "xz" Then
			extract("7z", 'XZ ' & t('TERM_COMPRESSED'), "xz")
		ElseIf $fileext = "z" Then
			extract("7z", 'LZW ' & t('TERM_COMPRESSED'), "Z")
		Else
			extract("7z", '7-Zip ' & t('TERM_ARCHIVE'))
		EndIf
	EndIf

	_DeleteTrayMessageBox()
	$7zfailed = True
	Return False
EndFunc

; Determine if file is ALZip archive
Func CheckAlz()
	Cout("Testing ALZ")

	_CreateTrayMessageBox(t('TERM_TESTING') & ' ALZ ' & t('TERM_ARCHIVE'))
	$return = FetchStdout($alz & ' -l "' & $file & '"', $filedir, @SW_HIDE)

	If StringInStr($return, "Listing archive:") And Not (StringInStr($return, "corrupted file") Or StringInStr($return, "file open error")) Then
		_DeleteTrayMessageBox()
		extract("alz", 'ALZ ' & t('TERM_ARCHIVE'))
	EndIf

	Return False
EndFunc   ;==>CheckAlz

; Determine if file is self-extracting ARJ archive
Func checkArj()
	If $arjfailed Then Return False
	Cout("Testing ARJ")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' SFX ARJ ' & t('TERM_ARCHIVE'))
	$return = FetchStdout($arj & ' l "' & $file & '"', $filedir, @SW_HIDE)

	If StringInStr($return, "Archive created:", 0) Then
		_DeleteTrayMessageBox()
		extract("7z", t('TERM_SFX') & ' ARJ ' & t('TERM_ARCHIVE'))
	EndIf

	_DeleteTrayMessageBox()
	$arjfailed = True
	Return False
EndFunc

; Determine if folder contains .bin files
Func checkBin()
	Cout("Searching additional .bin files to be extracted")

	$NSISbin = FileFindFirstFile($filedir & "\data*.bin")
	If $NSISbin == -1 Then Return
	If Prompt(64 + 1, "NSIS_BINFILES", CreateArray($file, $filenamefull), 0) Then
		While 1
			$file = $filedir & "\" & FileFindNextFile($NSISbin)
			If @error Then ExitLoop
			;FilenameParse($file)
			extract("7z", ".bin " & t('TERM_ARCHIVE'), "", True, True)
		WEnd
	EndIf
	FileClose($NSISbin)
	terminate($STATUS_SUCCESS, "", "NSIS")
EndFunc

; Determine if file is supported game archive
Func CheckGame($bUseGaup = True)
	If (Not $CheckGame) Or $gamefailed Then Return

	Cout("Testing Game archive")

	_CreateTrayMessageBox(t('TERM_TESTING') & ' ' & t('TERM_GAME') & t('TERM_PACKAGE'))

	If $bUseGaup Then
		; Check GAUP first
		$return = FetchStdout($quickbms & ' -l "' & $bindir & $gaup & '" "' & $file & '"', $filedir, @SW_HIDE, -1)

		If StringInStr($return, "Target directory:", 0) Or StringInStr($return, "0 files found", 0) Or StringInStr($return, "Error", 0) _
		Or StringInStr($return, "exception occured", 0) Or StringInStr($return, "not supported", 0) Or $return == "" Then

		Else
			_DeleteTrayMessageBox()
			extract("qbms", t('TERM_GAME') & t('TERM_PACKAGE'), $gaup, False, True)
		EndIf
	EndIf

	$gamefailed = True

	If $silentmode Then
		Cout("INFO: File may be extractable via BMS script, but user input is needed. Disable silent mode to try this method.")
		_DeleteTrayMessageBox()
		Return False
	EndIf

	; Check if game specific bms script is available
	_SQLite_Startup()
	If @error Then Return Cout("[ERROR] SQLite startup failed with code " & @error)

	_SQLite_Open($bindir & "BMS.db", $SQLITE_OPEN_READONLY)
	If @error Then Return Cout("[ERROR] Failed to open BMS database")

	Local $gameformat, $iRows, $iColumns, $game

	_SQLite_GetTable(-1, "SELECT n.Name FROM Names n, Scripts s, Extensions e WHERE s.SID = e.EID AND s.SID = n.NID AND e.Extension= '" _
			 & StringLower($fileext) & "' ORDER BY n.Name", $gameformat, $iRows, $iColumns)
 	_ArrayDelete($gameformat, 1)

	If $gameformat[0] > 1 Then
		_ArrayDelete($gameformat, 0)
		_ArraySort($gameformat)
		$game = GameSelect(_ArrayToString($gameformat), t('METHOD_GAME_NOGAME'))

		If $game Then
			Cout("SELECT s.Script FROM Scripts s, Names n WHERE s.SID = n.NID AND Name = '" & $game & "'")
			_SQLite_GetTable(-1, "SELECT s.Script FROM Scripts s, Names n WHERE s.SID = n.NID AND Name = '" & _
								 $game & "'", $gameformat, $iRows, $iColumns)
			;_ArrayDisplay($gameformat)
			; Write script to file and execute it
			$bmsScript = FileOpen($bindir & $bms, 2)
			FileWrite($bmsScript, $gameformat[2])
			FileClose($bmsScript)
			$return = FetchStdout($quickbms & ' -l "' & $bindir & $bms & '" "' & $file & '"', $filedir, @SW_HIDE, -1)
			;_ArrayDisplay($gameformat)
			If Not StringInStr($return, "0 files found") And Not StringInStr($return, "Error") And Not StringInStr($return, "invalid") _
			And Not StringInStr($return, "expeceted") And $return <> "" Then
				_SQLite_Close()
				_SQLite_Shutdown()
				extract('qbms', $game & " " & t('TERM_PACKAGE'), $bms)
			EndIf
		EndIf
	EndIf

	_SQLite_Close()
	_SQLite_Shutdown()

	_DeleteTrayMessageBox()
	Return False
EndFunc

; Determine if InstallExplorer can extract the file
Func checkIE()
	Cout("Testing InstallExplorer")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' InstallExplorer ' & t('TERM_INSTALLER'))
	$return = FetchStdout($quickbms & ' -l "' & $bindir & $ie & '" "' & $file & '"', $filedir, @SW_HIDE)
	_DeleteTrayMessageBox()

	If StringInStr($return, "Target directory:", 0) Or StringInStr($return, "0 files found", 0) Or StringInStr($return, "Error", 0) _
	Or StringInStr($return, "exception occured", 0) Or StringInStr($return, "not supported", 0) Or StringInStr($return, "crash occurred", 0) _
	Or $return == "" Then
		$iefailed = True
		Return False
	EndIf

	extract("qbms", 'InstallExplorer ' & t('TERM_INSTALLER'), $ie)
EndFunc

; Determine if file is Inno Setup installer
Func checkInno()
	Cout("Testing Inno Setup")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' Inno Setup ' & t('TERM_INSTALLER'))

	$return = FetchStdout($inno & ' "' & $file & '"', $filedir, @SW_HIDE)
	If (StringInStr($return, "Version detected:", 0) And Not (StringInStr($return, "error", 0))) _
			Or (StringInStr($return, "Signature detected:", 0) _
			And Not StringInStr($return, "not a supported version", 0)) Then
		_DeleteTrayMessageBox()
		extract("inno", 'Inno Setup ' & t('TERM_INSTALLER'))
	EndIf

	_DeleteTrayMessageBox()
	$innofailed = True
	checkIE()
	Return False
EndFunc   ;==>checkInno

; Determine if file is really an InstallShield installer (not false positive)
Func checkInstallShield()
	; InstallShield testing handled by extract function
	Cout("Testing InstallShield")
	extract("isexe", 'InstallShield ' & t('TERM_INSTALLER'))
	Return False
EndFunc   ;==>checkInstallShield

; Determine if file is CD/DVD image
Func CheckIso()
	If $isofailed Then Return False
	Cout("Testing image file")
	_CreateTrayMessageBox(t('TERM_TESTING') & " " & t('TERM_IMAGE') & " " & t('TERM_FILE'))
	If $fileext = "cue" And FileExists(StringTrimRight($file, 3) & "bin") Then $file = StringTrimRight($file, 3) & "bin"

	$return = FetchStdout($quickbms & ' -l "' & $bindir & $iso & '" "' & $file & '"', $filedir, @SW_HIDE, -1)
	_DeleteTrayMessageBox()
	If StringInStr($return, "Target directory:", 0) Or StringInStr($return, "0 files found", 0) Or StringInStr($return, "Error", 0) _
			Or StringInStr($return, "exception occured", 0) Or StringInStr($return, "not supported", 0) Or $return == "" Then
		$isofailed = True
		Return False
	Else
		Return extract("qbms", t('TERM_IMAGE') & " " & t('TERM_FILE'), $iso, $isofile, $isofile)
	EndIf
EndFunc

; Determine if file is NSIS installer
Func checkNSIS()
	Cout("Testing NSIS")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' NSIS ' & t('TERM_INSTALLER'))

	$return = FetchStdout($7z & ' l "' & $file & '"', $filedir, @SW_HIDE)
	If StringInStr($return, "Listing archive:") And Not StringInStr($return, "Can not open the file as") Then _
		extract("NSIS", 'NSIS ' & t('TERM_INSTALLER'))

	_DeleteTrayMessageBox()
	checkIE()
	Return False
EndFunc

; Determine if file is self-extracting Zip archive
Func checkZip()
	if $zipfailed Then Return
	Cout("Testing Zip")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' SFX ZIP ' & t('TERM_ARCHIVE'))
	$return = FetchStdout($zip & ' -l "' & $file & '"', $filedir, @SW_HIDE)
	If Not StringInStr($return, "signature not found", 0) Then
		_DeleteTrayMessageBox()
		extract("zip", t('TERM_SFX') & ' ZIP ' & t('TERM_ARCHIVE'))
	EndIf

	_DeleteTrayMessageBox()
	$zipfailed = True
	Return False
EndFunc

; If detection fails, try to determine file type by extension
Func CheckExt()
	For $dir In $aDefDirs
		Local $aDefinitions = IniReadSection($defdir & "registry.ini", "Extensions")
		If @error Then
			Cout("Could not load definition registry from " & $dir)
			ContinueLoop
		EndIf

		For $i = 0 To $aDefinitions[0][0]
			$aReturn = StringSplit($aDefinitions[$i][0], ",")
			For $j = 1 To $aReturn[0]
				If StringCompare($fileext, StringStripWS($aReturn[$j], 8)) == 0 Then extract($aDefinitions[$i][1])
			Next
		Next
	Next
#cs
	Switch $fileext
		Case "1", "lib"
			extract("is3arc", 'InstallShield 3.x ' & t('TERM_ARCHIVE'))

		Case "ace"
			extract("ace", 'ACE ' & t('TERM_ARCHIVE'))

		Case "assets"
			extract("unity", 'Unity Engine Asset ' & t('TERM_FILE'))

		Case "b64"
			extract("uu", 'Base64 ' & t('TERM_ENCODED'))

		Case "chm"
			extract("chm", 'Compiled HTML ' & t('TERM_HELP'))

		Case "cpio"
			extract("7z", 'CPIO ' & t('TERM_ARCHIVE'))

		Case "dbx"
			extract('qbms', 'Outlook Express ' & t('TERM_ARCHIVE'), $dbx)

		Case "deb"
			extract("7z", 'Debian ' & t('TERM_PACKAGE'))

		Case "hlp"
			extract("hlp", 'Windows ' & t('TERM_HELP'))

		Case "imf"
			extract("cab", 'IncrediMail ' & t('TERM_ECARD'))

		Case "img"
			extract("img", 'Floppy ' & t('TERM_DISK') & ' ' & t('TERM_IMAGE'))

		Case "kgb", "kge"
			extract("kgb", 'KGB ' & t('TERM_ARCHIVE'))

		Case "lzh", "lha"
			extract("7z", 'LZH ' & t('TERM_COMPRESSED'))

		Case "lzo"
			extract("lzo", 'LZO ' & t('TERM_COMPRESSED'))

		Case "lzx"
			extract("lzx", 'LZX ' & t('TERM_COMPRESSED'))

		Case "msm"
			extract("msm", 'Windows Installer (MSM) ' & t('TERM_MERGE_MODULE'))

		Case "pea"
			extract("pea", 'Pea ' & t('TERM_ARCHIVE'))

		Case "rar", "001", "cbr"
			extract("rar", 'RAR ' & t('TERM_ARCHIVE'))

		Case "rpm"
			extract("7z", 'RPM ' & t('TERM_PACKAGE'))

		Case "sis"
			extract('qbms', 'SymbianOS ' & t('TERM_INSTALLER'), $sis)

		Case "sit"
			extract("sit", 'StuffIt ' & t('TERM_ARCHIVE'))

		Case "tar"
			extract("tar", 'Tar ' & t('TERM_ARCHIVE'))

		Case "uha"
			extract("uha", 'UHARC ' & t('TERM_ARCHIVE'))

		Case "uif"
			extract("uif", 'UIF ' & t('TERM_IMAGE'))

		Case "vpk", "gcf", "ncf", "wad", "xzp"
			extract("gcf", 'Valve ' & $fileext & " " & t('TERM_PACKAGE'))

		Case "wim"
			extract("7z", 'WIM ' & t('TERM_IMAGE'))

		Case "yenc", "ntx"
			extract("uu", 'yEnc ' & t('TERM_ENCODED'))

		Case "zip", "cbz", "jar", "xpi", "wz"
			extract("zip", 'ZIP ' & t('TERM_ARCHIVE'))

		Case "zoo"
			extract("zoo", 'ZOO ' & t('TERM_ARCHIVE'))

		Case "rgss", "rgss2a"
			extract("arc_conv", "RPG Maker " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case "rgss3a"
			extract("rgss3", "RPG Maker VX Ace " & t('TERM_GAME') & t('TERM_ARCHIVE'))
	EndSwitch
#ce
EndFunc

; Perform special actions for some file types
Func InitialCheckExt()
	If Not $extract Then Return
	; Compound compressed files that require multiple actions
	Switch $fileext
		Case "ipk", "tbz2", "tgz", "tz", "tlz", "txz"
			extract("ctar", 'Compressed Tar ' & t('TERM_ARCHIVE'))
		; image files - TrID is not always reliable with these formats
		Case "bin", "cue", "cdi", "mdf"
			CheckIso()
		Case "dmg"
			extract("7z", 'DMG ' & t('TERM_IMAGE'))
		Case "iso", "nrg"
			check7z()
			CheckIso()
	EndSwitch
EndFunc

; Check for unicode characters in path
Func MoveInputFileIfNecessary()
	$bIsUnc = _WinAPI_PathIsUNC($file)

	Local $new = 0
	If $checkUnicode And Not StringRegExp($file, $notAsciiPattern, 0) Then
		Cout("File name seems to be unicode")

		If StringRegExp($filedir, $notAsciiPattern, 0) Then
			; Directory is ASCII, only rename file
			$new = _TempFile($filedir, "Unicode_", $fileext)
		Else
			Cout("Path seems to be unicode")
			If Not StringRegExp(@TempDir, $notAsciiPattern, 0) Then Return Cout("Temp directory contains unicode characters, aborting")
			$new = StringRegExp($filename, $notAsciiPattern, 0)? @TempDir & "\" & $filenamefull: _TempFile(@TempDir, "Unicode_", $fileext)
		EndIf
	EndIf

	If $new == 0 Then
		If Not $bIsUnc Then Return
		$new = @TempDir & "\" & $filenamefull
	EndIf

	; Multipart archive, TODO: move all parts
	If StringRegExp($file, ".*part\d+\.rar", 0) Or StringRegExp($fileext, "\d{3}", 0) Then Return Cout("File seems to be multipart archive, not moving")

	HasFreeSpace($new, 2)

	Cout('Renaming "' & $filenamefull & '" to "' & $new & '"')
	_CreateTrayMessageBox(t('MOVING_FILE') & @CRLF & $new)

	If StringLeft($file, 1) = StringLeft($new, 1) Then
		If Not FileMove($file, $new) Then Return terminate($STATUS_MOVEFAILED, $new)
		$iUnicodeMode = $UNICODE_MOVE
	Else
		If Not FileCopy($file, $new) Then Return terminate($STATUS_MOVEFAILED, $new)
		$iUnicodeMode = $UNICODE_COPY
	EndIf
	Cout("Unicode file mode: " & $iUnicodeMode)

	$oldpath = $file
	$sUnicodeName = $filename
	$oldoutdir = $outdir
	FilenameParse($new)
	$outdir = $initoutdir
EndFunc

; Extract from known archive format
Func extract($arctype, $arcdisp = 0, $additionalParameters = "", $returnSuccess = False, $returnFail = False)
	Local $dirmtime = -1
	$success = $RESULT_UNKNOWN
	Cout("Starting " & $arctype & " extraction")
	If $arcdisp == 0 Then $arcdisp = "." & $fileext & " " & t('TERM_FILE')
	_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & $arcdisp)

	; Create subdirectory
	If StringRight($outdir, 1) = '\' Then $outdir = StringTrimRight($outdir, 1)
	If FileExists($outdir) Then
		If Not StringInStr(FileGetAttrib($outdir), "D") Then terminate($STATUS_INVALIDDIR, $outdir, "")
		$dirmtime = FileGetTime($outdir, 0, 1)
		If @error Then $dirmtime = -1
		If Not CanAccess($outdir) Then terminate($STATUS_INVALIDDIR, $outdir, "")
	Else
		If Not DirCreate($outdir) Then terminate($STATUS_INVALIDDIR, $outdir, "")
		$createdir = True
	EndIf

	HasFreeSpace()

	$initdirsize = _DirGetSize($outdir)
	$tempoutdir = TempDir($outdir, 7)

	; Extract archive based on filetype
	Switch $arctype
		Case "7z"
			Local $sPassword = _FindArchivePassword($7z & ' l -p -slt "' & $file & '"', $7z & ' t -p"%PASSWORD%" "' & $file & '"', "Encrypted = +", "Wrong password?", 0, "Everything is Ok")
			_Run($7z & ' x ' & ($sPassword == 0? '"': '-p"' & $sPassword & '" "') & $file & '"', $outdir, @SW_HIDE, True, True, True, True)
			If @extended Then terminate($STATUS_PASSWORD, $file, $arcdisp)

			If FileExists($outdir & '\.text') Then
				; Generic .exe extraction should not be considered successful
				$success = $RESULT_FAILED
			ElseIf StringInStr($filetype, 'RPM Linux Package', 0) Then
				; Extract inner CPIO for RPMs
				If FileExists($outdir & '\' & $filename & '.cpio') Then
					_Run($7z & ' x "' & $outdir & '\' & $filename & '.cpio"', $outdir)
					FileDelete($outdir & '\' & $filename & '.cpio')
				EndIf
			ElseIf StringInStr($filetype, 'Debian Linux Package', 0) Then
				; Extract inner tarball for DEBs
				If FileExists($outdir & '\data.tar') Then
					_Run($7z & ' x "' & $outdir & '\data.tar"', $outdir)
					FileDelete($outdir & '\data.tar')
				EndIf
			ElseIf $additionalParameters == "xz" Or $additionalParameters == "gz" Or $additionalParameters == "Z" Then
				If FileExists($outdir & '\' & $filename) And StringTrimLeft($filename, StringInStr($filename, '.', 0, -1)) = "tar" Then
					_Run($7z & ' x "' & $outdir & '\' & $filename & '"', $outdir)
					FileDelete($outdir & '\' & $filename)
				EndIf
			ElseIf $additionalParameters == "bz2" Then
				If FileExists($outdir & '\' & $filename) Then
					_Run($7z & ' x "' & $outdir & '\' & $filename & '"', $outdir)
					FileDelete($outdir & '\' & $filename)
				EndIf
			EndIf

		Case "ace"
			$success = $RESULT_SUCCESS
			Opt("WinTitleMatchMode", 3)
			$pid = Run($ace & ' -x "' & $file & '" "' & $outdir & '"', $filedir)
			ProcessWait($pid, $Timeout)
			While ProcessExists($pid)
				If WinExists("Error") Then
					ProcessClose($pid)
					$success = $RESULT_FAILED
					ExitLoop
				EndIf
				Sleep(50)
			WEnd

		Case "ai"
			Warn_Execute($file & ' /extract:"' & $outdir & '"')
			; ShellExecute is needed here to display UAC prompt, fails with Run()
			ShellExecute($file, ' /extract:"' & $outdir & '"', $outdir)
			ProcessWait($filenamefull, $Timeout)
			ProcessWaitClose($filenamefull, $Timeout)

		Case "alz"
			_Run($alz & ' -d "' & $outdir & '" "' & $file & '"', $outdir, True, True, False)

		Case "arc_conv"
			If Not HasPlugin($arc_conv, $returnFail) Then Return

			$run = Run(Cout(_MakeCommand($arc_conv & ' "' & $file & '"', True)), $outdir, @SW_HIDE)
			If @error Then Return

			Local $hWnd = WinWait("arc_conv", "", $Timeout)
			If $hWnd == 0 Then terminate($STATUS_TIMEOUT, $file, $arcdisp)
			Local $current = "", $last = ""
			; Hide not possible as window text has to be read
			WinSetState("arc_conv", "", @SW_MINIMIZE)
			While WinExists("arc_conv")
				$current = WinGetText("arc_conv")
				If $current <> $last Then
					GUICtrlSetData($idTrayStatusExt, StringInStr($current, "/")? $current: (t('TERM_FILE') & " #" & $current))
					$last = $current
					Sleep(10)
				EndIf
			WEnd
			$run = 0
			MoveFiles($file & "~", $outdir, True, "", True)

		Case "audio"
			HasFFMPEG()
			_Run($cmd & $ffmpeg & ' -i "' & $file & '" "' & ($iUnicodeMode? $sUnicodeName: $filename) & '.wav"', $outdir, @SW_HIDE)

		Case "bcm"
			_Run($bcm & ' -d "' & $file & '" "' & $outdir & '\' & ($iUnicodeMode? $sUnicodeName: $filename) & '"', $filedir, @SW_HIDE, True, True, False)

		Case "bootimg" ;Test
			HasPlugin($bootimg)
			$ret = $outdir & "\" & $bootimg
			$ret2 = $outdir & '\boot.img'
			FileCopy($bindir & $bootimg, $outdir)
			FileMove($file, $ret2)
			_Run($cmd & '"' & $ret & ' --unpack-bootimg"', $outdir, @SW_MINIMIZE, False, False)
			FileMove($ret2, $file)
			FileDelete($ret)

		Case "cab"
			If StringInStr($filetype, 'Type 1', 0) Then
				RunWait(Warn_Execute(Quote($file & '" /q /x:"' & $outdir), $outdir)
			Else
				check7z()
			EndIf

		Case "chm"
			_Run($7z & ' x "' & $file & '"', $outdir)
			Local $aReturn[2] = ['#*', '$*']
			Cleanup($aReturn)
			$handle = FileFindFirstFile($outdir & '\*')
			If $handle <> -1 Then
				$dir = FileFindNextFile($handle)
				Do
					$char = StringLeft($dir, 1)
					If $char = '#' Or $char = '$' Then Cleanup($outdir & "\" & $dir)
					$dir = FileFindNextFile($handle)
				Until @error
			EndIf
			FileClose($handle)

		Case "ci"
			HasPlugin($ci)
			$return = @TempDir & "\ci.txt"
			$ret = FileOpen($return, 8+2)
			FileWrite($ret, "1" & @LF & $file & @LF & $outdir & @LF & "3" & @LF & "1")
			FileClose($ret)
			_Run($ci & ' ' & $return, $outdir, @SW_SHOW, False, False, False)
			FileDelete($return)
			terminate($STATUS_SILENT, "", "")

		Case "crage"
			HasPlugin($crage)
			_Run($crage & ' -v -p "' & $file & '" -o "' & $outdir & '"', $bindir & "crass-0.4.14.0", @SW_SHOW, True, False, False)

		Case "ctar"
			$oldfiles = ReturnFiles($outdir)

			; Decompress archive with 7-zip
			_Run($7z & ' x "' & $file & '"', $outdir)

			; Check for new files
			$handle = FileFindFirstFile($outdir & "\*")
			If Not @error Then
				While 1
					$fname = FileFindNextFile($handle)
					If @error Then ExitLoop
					If Not StringInStr($oldfiles, $fname) Then
						; Check for supported archive format
						$return = FetchStdout($7z & ' l "' & $outdir & '\' & $fname & '"', $outdir, @SW_HIDE)
						If StringInStr($return, "Listing archive:", 0) Then
							_Run($7z & ' x "' & $outdir & '\' & $fname & '"', $outdir, @SW_HIDE)
							FileDelete($outdir & '\' & $fname)
						EndIf
					EndIf
				WEnd
			EndIf
			FileClose($handle)

		Case "dgca"
			HasPlugin($dgca)
			Local $sPassword = _FindArchivePassword($dgca & ' e "' & $file & '"', $dgca & ' l -p%PASSWORD% "' & $file & '"', "Archive encrypted.", 0, -2, "-------------------------")
			_Run($dgca & ' e ' & ($sPassword == 0? '"': '-p' & $sPassword & ' "') & $file & '" "' & $outdir & '"', $outdir, @SW_HIDE, True, True, False, False)
			If @extended Then terminate($STATUS_PASSWORD, $file, $arcdisp)

		Case "daa"
			; Prompt user to continue
			_DeleteTrayMessageBox()
			Prompt(32 + 4, 'CONVERT_CDROM', 'DAA/GBI', 1)

			; Begin conversion to .iso format
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'DAA ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 1)')
			$isofile = $filedir & '\' & $filename & '.iso'
			_Run($daa & ' "' & $file & '" "' & $isofile & '"', $filedir)

		Case "dcp"
			HasPlugin($dcp)
			_Run($dcp & ' "' & $file & '"', $outdir)

		Case "ei"
			Warn_Execute($file & ' /batch /no-reg /no-postinstall /dest "' & $outdir & '"')
			ShellExecuteWait($file, '/batch /no-reg /no-postinstall /dest "' & $outdir & '"', $outdir)

		Case "ethornell" ; Test
			_Run($ethornell & ' "' & $file & '" "' & $outdir & '"', $outdir)

		Case "enigma"
			_Run($enigma & ' /nogui "' & $file & '"', $outdir, @SW_HIDE, True, False, False)

			; Move files
			FileMove($filedir & "\" & $filename & "_unpacked.exe", $outdir & "\" & $filename & "_unpacked.exe")
			DirMove($filedir & "\%DEFAULT FOLDER%", $outdir & "\%DEFAULT FOLDER%")
			; TODO: move other folders

			; Read log file
			FileDelete($outdir & "\!unpacker.log")
			$ret = $filedir & "\!unpacker.log"
			$return = Cout(FileRead($ret))
			FileDelete($ret)

			; Success evaluation
			If StringInStr($return, 'Expected section name ".enigma2"') Then
				$success = $RESULT_FAILED
			ElseIf StringInStr($return, '[+] Finished!') Then
				$success = $RESULT_SUCCESS
			EndIf

		Case "fead"
			RunWait(Warn_Execute($file & ' /s -nos_ne -nos_o"' & $tempoutdir & '"'), $filedir)
			FileSetAttrib($tempoutdir & '*', '-R', 1)
			MoveFiles($tempoutdir, $outdir, False, "", True)
			DirRemove($tempoutdir)

		Case "freearc"
			_Run($freearc & ' x -dp"' & $outdir & '" "' & $file & '"', $filedir, @SW_HIDE, True, True, False, False)

		Case "fsb"
			_Run($fsb & ' -d "' & $outdir & '" "' & $file & '"', $filedir, @SW_MINIMIZE, True, True, False)

		Case "gcf"
			Prompt(48 + 1, 'PACKAGE_EXPLORER', $file, 1)
			Run($gcf & ' "' & $file & '"')
			terminate($STATUS_SILENT, "", "")

		Case "ghost"
			$ret = $outdir & "\" & $filename & ".exe"
			Cout("Moving file to " & $ret)
			FileMove($file, $ret)

			$aReturn = OpenExeInfo($ret)

			MouseMove(0, 0, 0)
			ControlClick($aReturn[0], "", "[CLASS:TBitBtn; INSTANCE:15]")
			ControlSend($aReturn[0], "", "[CLASS:TBitBtn; INSTANCE:15]", "{DOWN}{DOWN}{RIGHT}{ENTER}")

			$TimerStart = TimerInit()
			$return = ""

			While Not StringInStr($return, "file saved")
				Sleep(200)
				$return = ControlGetText($aReturn[0], "", "TEdit5")
				If TimerDiff($TimerStart) > $Timeout Then ExitLoop
			WEnd

			CloseExeInfo($aReturn)

			Cout("Moving file back")
			FileMove($ret, $filedir & "\")

			$ret2 = $ret & "-ovl"
			Cout($ret2)
			If FileExists($ret2) Then
				Cout("Overlay extracted successfully, xor-ing")
				_Run($xor & ' "' & $ret2 & '" "' & $outdir & '\' & $filename & '.cab" 0x8D')
				FileDelete($ret2)
			Else
				Cout("Failed to extract overlay")
				$success = $RESULT_FAILED
			EndIf

		Case "hlp"
			_Run($hlp & ' "' & $file & '"', $outdir)
			If _DirGetSize($outdir, $initdirsize + 1) > $initdirsize Then
				DirCreate($tempoutdir)
				_Run($hlp & ' /r /n "' & $file & '"', $tempoutdir)
				FileMove($tempoutdir & $filename & '.rtf', $outdir & '\' & $filename & '_' & t('TERM_RECONSTRUCTED') & '.rtf')
				DirRemove($tempoutdir, 1)
			EndIf

		; failsafe in case TrID misidentifies MS SFX hotfixes
		Case "hotfix"
			Cout("Executing: " & Warn_Execute(Quote($file & '" /q /x:"' & $outdir)))
			ShellExecuteWait($file, '/q /x:' & Quote($outdir), $outdir)

		Case "img" ; Test
			_Run($img & ' -x "' & $file & '"', $outdir, True, True, False)

		Case "inno"
			If StringInStr($filetype, "Reflexive Arcade", 0) Then
				DirCreate($tempoutdir)
				$ret = $tempoutdir & $filename & '.exe"'
				_Run($rai & ' "' & $file & '" "' & $ret, $filedir)
				_Run($inno & ' -x -m -a "' & $ret, $outdir)
				FileDelete($tempoutdir & $filename & '.exe')
				DirRemove($tempoutdir)
			Else
				_Run($inno & ' -x -m -a "' & $file & '"', $outdir)
			EndIf

			; Inno setup files can contain multiple versions of files, they are named ',1', ',2',... after extraction
			; rename the first file(s), so extracted programs do not fail with not found exceptions
			; This is a convenience function, so the user does not have to rename them manually
			$return = $outdir & "\{app}\"
			$aReturn = _FileListToArrayRec($return, "*,1.*", 1, 1)
			If Not @error Then
				For $i = 1 To $aReturn[0]
					$ret = StringReplace($aReturn[$i], ",1", "", -1)
					Cout("Renaming " & $return & $aReturn[$i] & " to " & $return & $ret)
					FileMove($return & $aReturn[$i], $return & $ret)
				Next
			EndIf

			; (Re)move ',2' files and install_script.iss
			$aReturn = _FileListToArrayRec($return, "*,2.*", 1, 1, 0, 2)
			_ArrayDelete($aReturn, 0)
			Cleanup($aReturn)

			; Change output directory structure
			Local $aReturn = ["embedded", "{tmp}", "{commonappdata}", "{cf}", "{cf32}", "{{userappdata}}", "{{userdocs}}"]
			Cleanup($aReturn)
			MoveFiles($outdir & "\{app}", $outdir, True, '', True)
			; TODO: {syswow64}, {sys} - move files to outdir as dlls might be needed by the program?

			Cleanup("install_script.iss")
			Cleanup("setup.iss")

		Case "is3arc" ; Test
			$aReturn = ['InstallShield 3.x ' & t('TERM_ARCHIVE'), t('METHOD_EXTRACTION_RADIO', 'STIX'), t('METHOD_EXTRACTION_RADIO', 'unshield')]
			$choice = MethodSelect($aReturn, $arcdisp)

			If $choice == 1 Then ; STIX
				_Run($stix & ' ' & FileGetShortName($file) & ' ' & FileGetShortName($outdir), $filedir)
			Else ; Unshield
				_Run($unshield & ' -d "' & $outdir & '" x "' & $file & '"', $outdir)
			EndIf

		Case "iscab" ; Test
			Local $aReturn = ['InstallShield Cabinet ' & t('TERM_ARCHIVE'), t('METHOD_EXTRACTION_RADIO', 'is6comp'), t('METHOD_EXTRACTION_RADIO', 'is5comp'), t('METHOD_EXTRACTION_RADIO', 'iscab')]
			$choice = MethodSelect($aReturn, $arcdisp)

			Switch $choice
				Case 1
					; List contents of archive
					$return = FetchStdout($is6cab & ' l "' & $file & '"', $filedir, @SW_HIDE)
					$return = _StringBetween(StringRight($return, 22), " ", " file(s) total")
					If Not @error Then $return = Number(StringStripWS($return[0], 8))

					; If successful, extract contents of InstallShield cabs file-by-file
					If $return > 0 Then
						RunWait(MakeCommand($is6cab & ' x "' & $file & '"', True), $outdir, @SW_MINIMIZE)
					Else
						; Otherwise, attempt to extract with unshield
						_Run($unshield & ' -d "' & $outdir & '" x "' & $file & '"', $outdir)
					EndIf
				Case 2
					HasPlugin($is5cab)
					RunWait($is5cab & ' x "' & $file & '"', $outdir, @SW_MINIMIZE)
				Case 3
					HasPlugin($iscab)
					RunWait($iscab & ' "' & $file & '" -i"files.ini" -lx', $outdir, @SW_HIDE)
					RunWait($iscab & ' "' & $file & '" -i"files.ini" -x', $outdir, @SW_MINIMIZE)
					FileDelete($outdir & "\files.ini")
			EndSwitch

		Case "isexe"
			exescan($file, 'ext', 0)
			If StringInStr($filetype, "3.x", 0) Then
				; Extract 3.x SFX installer using stix
				_Run($stix & ' ' & FileGetShortName($file) & ' ' & FileGetShortName($outdir), $filedir)
			Else
				Local $aReturn = ["InstallShield " & t('TERM_INSTALLER'), t('METHOD_EXTRACTION_RADIO', 'isxunpack'), t('METHOD_SWITCH_RADIO', 'InstallShield /b'), t('METHOD_NOTIS_RADIO')]
				$choice = MethodSelect($aReturn, $arcdisp)

				; User-specified false positive; return for additional analysis
				Switch $choice
					; Extract using isxunpack
					Case 1
						FileMove($file, $outdir)
						Run(_MakeCommand($isxunp & ' "' & $outdir & '\' & $filename & '.' & $fileext & '"', True), $outdir)
						WinWait(@ComSpec)
						WinActivate(@ComSpec)
						Send("{ENTER}")
						ProcessWaitClose($isxunp)
						FileMove($outdir & '\' & $filename & '.' & $fileext, $filedir)

					; Try to extract MSI using cache switch
					Case 2
						; Run installer and wait for temp files to be copied
						_CreateTrayMessageBox(t('INIT_WAIT'))
						DirCreate($tempoutdir)
						ShellExecute($file, '/b"' & $tempoutdir, $filedir)

						; TODO: Rewrite
						; Wait for matching windows for up to 30 seconds (60 * .5)
						Opt("WinTitleMatchMode", 4)
						Local $success
						For $i = 1 To $Timeout / 500
							If WinExists("classname=MsiDialogCloseClass") Then
								; Search temp directory for MSI support and copy to tempoutdir
								$msihandle = FileFindFirstFile($tempoutdir & "*.msi")
								If Not @error Then
									While 1
										$msiname = FileFindNextFile($msihandle)
										If @error Then ExitLoop
										$tsearch = FileSearch(@TempDir & "\" & $msiname)
										If @error Then ContinueLoop

										$isdir = StringLeft($tsearch[1], StringInStr($tsearch[1], '\', 0, -1) - 1)
										$ishandle = FileFindFirstFile($isdir & "\*")
										$fname = FileFindNextFile($ishandle)
										Do
											If $fname <> $msiname Then
												FileCopy($isdir & "\" & $fname, $tempoutdir)
											EndIf
											$fname = FileFindNextFile($ishandle)
										Until @error
										FileClose($ishandle)
									WEnd
									FileClose($msihandle)
								EndIf

								; Move files to outdir
								_DeleteTrayMessageBox()
								Prompt(64, 'INIT_COMPLETE', 0)
								MoveFiles($tempoutdir, $outdir, False, "", True)
								$success = $RESULT_SUCCESS
								ExitLoop
							EndIf

							Sleep(500)
						Next
						$run = 0

						; Not a supported installer
						If $success <> $RESULT_SUCCESS Then
							_DeleteTrayMessageBox()
							Prompt(64, 'INIT_COMPLETE', 0)
						EndIf

					; Not InstallShield
					Case 4
						$isfailed = True
						Return False
				EndSwitch
			EndIf

		Case "isz"
			_DeleteTrayMessageBox()
			Prompt(32 + 4, 'CONVERT_CDROM', 'ISZ', 1)
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'ISZ ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 1)')

			$return = $tempoutdir & $filename & '.' & $fileext
			FileMove($file, $return, 8)
			_Run($isz & ' "' & $return & '"', $tempoutdir, True, True)
			$isofile = $tempoutdir & $filename & '.iso'
			FileMove($return, $file)

		Case "kgb"
			_Run($kgb & ' "' & $file & '"', $outdir, @SW_SHOW, True, False, False)

		Case "lzo"
			_Run($lzo & ' -d -p"' & $outdir & '" "' & $file & '"', $filedir)

		Case "lzx"
			_Run($lzx & ' -x "' & $file & '"', $outdir)

		Case "mht"
			Local $aReturn = ['MHTML ' & t('TERM_ARCHIVE'), t('METHOD_EXTRACTION_RADIO', 'ExtractMHT'), t('METHOD_EXTRACTION_RADIO', 'MhtUnPack')]
			$choice = MethodSelect($aReturn, $arcdisp)

			; Extract using ExtractMHT
			If $choice == 1 Then
				_Run($mht & ' "' & $file & '" "' & $outdir & '"', $outdir, @SW_MINIMIZE, False, False, False)
			Else
				extract('qbms', $arcdisp, $mht_plug)
			EndIf

		Case "mole"
			_Run($mole & ' /nogui "' & $file & '"', $outdir, @SW_HIDE, True, False, False)

			; Move files
			FileMove($filedir & "\" & $filename & "_unpacked.exe", $outdir & "\" & $filename & "_unpacked.exe")
			MoveFiles($filedir & "\_extracted", $outdir, False, '', True)

			; Read log file
			FileDelete($outdir & "\!unpacker.log")
			$ret = $filedir & "\!unpacker.log"
			$return = Cout(FileRead($ret))
			FileDelete($ret)

			; Success evaluation
			If StringInStr($return, '[x] Not a Molebox or unknown version') Then
				$success = $RESULT_FAILED
			ElseIf StringInStr($return, '[i] Finished! Have a nice day!') Then
				$success = $RESULT_SUCCESS
			EndIf

		Case "msi"
			If HasNetFramework(4) Then
				_Run($msi_lessmsi & ' x "' & $file & '" "' & $outdir & '\"', $outdir, @SW_SHOW, True, True, True)
				MoveFiles($outdir & "\SourceDir", $outdir, False, '', True)
			Else
				Local $aReturn = ['MSI ' & t('TERM_INSTALLER'), t('METHOD_EXTRACTION_RADIO', 'jsMSI Unpacker'), t('METHOD_EXTRACTION_RADIO', 'MsiX'), t('METHOD_EXTRACTION_RADIO', 'MSI TC Packer'), t('METHOD_ADMIN_RADIO', 'MSI')]
				$choice = MethodSelect($aReturn, $arcdisp)

				Switch $choice
					Case 1 ; jsMSI Unpacker
						_Run($msi_jsmsix & ' "' & $file & '"|"' & $outdir & '"', $filedir, @SW_SHOW, False, False)
						_FileRead($outdir & "\MSI Unpack.log", True)

					Case 2 ; MsiX
						Local $appendargs = $appendext? '/ext': ''
						_Run($msi_msix & ' "' & $file & '" /out "' & $outdir & '" ' & $appendargs, $filedir)

					Case 3 ; MSI Total Commander plugin
						extract('qbms', $arcdisp, $msi_plug, True)

						; Extract files from extracted CABs
						$cabfiles = FileSearch($tempoutdir)
						For $i = 1 To $cabfiles[0]
							filescan($cabfiles[$i], 0)
							If StringInStr($filetype, "Microsoft Cabinet Archive", 0) Then
								_Run($cmd & $7z & ' x "' & $cabfiles[$i] & '"', $outdir)
								FileDelete($cabfiles[$i])
							EndIf
						Next

						; Append missing file extensions
						If $appendext Then AppendExtensions($tempoutdir)

						; Move files to output directory and remove tempdir
						MoveFiles($tempoutdir, $outdir, False, "", True)

					Case 4 ; Administrative install
						RunWait(Warn_Execute('msiexec.exe /a "' & $file & '" /qb TARGETDIR="' & $outdir & '"'), $filedir, @SW_SHOW)
				EndSwitch
			EndIf

		Case "msm" ; Test
			; Due to the appendext argument, a definition file cannot be used here
			_Run($msi_msix & ' "' & $file & '" /out "' & $outdir & '" ' & $appendext? '/ext': '', $filedir)

		Case "msp"
			Local $aReturn = ['MSP ' & t('TERM_PACKAGE'), t('METHOD_EXTRACTION_RADIO', 'MSI TC Packer'), t('METHOD_EXTRACTION_RADIO', 'MsiX'), t('METHOD_EXTRACTION_RADIO', '7-Zip')]
			$choice = MethodSelect($aReturn, $arcdisp)

			DirCreate($tempoutdir)
			Switch $choice
				Case 1 ; TC MSI
					extract('qbms', $arcdisp, $msi_plug, True)
				Case 2 ; MsiX
					_Run($msi_msix & ' "' & $file & '" /out "' & $tempoutdir & '"', $filedir)
				Case 3 ; 7-Zip
					_Run($7z & ' x "' & $file & '"', $outdir)
			EndSwitch

			; Regardless of method, extract files from extracted CABs
			$cabfiles = FileSearch($tempoutdir)
			For $i = 1 To $cabfiles[0]
				filescan($cabfiles[$i], 0)
				If StringInStr($filetype, "Microsoft Cabinet Archive", 0) Then
					_Run($7z & ' x "' & $cabfiles[$i] & '"', $outdir)
					FileDelete($cabfiles[$i])
				EndIf
			Next

			; Append missing file extensions
			If $appendext Then AppendExtensions($tempoutdir)

			; Move files to output directory and remove tempdir
			MoveFiles($tempoutdir, $outdir, False, "", True)

		Case "nbh" ; Test
			RunWait(_MakeCommand($nbh, True) & ' "' & $file & '"', $outdir)

		Case "NSIS"
			; Rename duplicates and extract
			_Run($7z & ' x -aou' & ' "' & $file & '"', $outdir)

			If $success == $RESULT_FAILED Then checkIE()

			Local $aReturn[] = ["[NSIS].nsi", "[LICENSE].*", "$PLUGINSDIR", "$TEMP", "uninstall.exe", "[LICENSE]"]
			Cleanup($aReturn)

			; Determine if there are .bin files in filedir
			checkBin()

		Case "pea"
			DirCreate($tempoutdir)
			Local $pid = Run($pea & ' UNPEA "' & $file & '" "' & $tempoutdir & '" RESETDATE SETATTR EXTRACT2DIR INTERACTIVE', $filedir)
			While ProcessExists($pid)
				$return = ControlGetText(_WinGetByPID($pid), '', 'Button1')
				If StringLeft($return, 4) = 'Done' Then ProcessClose($pid)
				Sleep(10)
			WEnd
			MoveFiles($tempoutdir, $outdir, False, "", True)

		Case "qbms"
			_Run($quickbms & ' "' & $bindir & $additionalParameters & '" "' & $file & '" "' & $outdir & '"', $outdir, @SW_MINIMIZE, True, False)
			If FileExists($bindir & $bms) Then FileDelete($bindir & $bms)

			If $additionalParameters == $ie Then
				Local $aReturn[] = ["[NSIS].nsi", "[LICENSE].*", "$PLUGINSDIR", "$TEMP", "uninstall.exe", "[LICENSE]"]
				Cleanup($aReturn)
			EndIf

		Case "rar"
			Local $sPassword = _FindArchivePassword($rar & ' lt -p- "' & $file & '"', $rar & ' t -p"%PASSWORD%" "' & $file & '"', "encrypted", 0, 0)
			_Run($rar & ' x ' & ($sPassword == 0? '"': '-p"' & $sPassword & '" "') & $file & '"', $outdir, @SW_SHOW)
			If @extended Then terminate($STATUS_PASSWORD, $file, $arcdisp)

		Case "rgss3"
			HasPlugin($rgss3)
			Run($rgss3, $outdir, @SW_HIDE)
			Local $handle = WinWait("RPGMaker Decrypter", "", $Timeout)
			If $handle = 0 Then terminate($STATUS_TIMEOUT, $file, $arcdisp)
			ControlSetText($handle, "", "Edit1", $file)
			ControlSetText($handle, "", "Edit2", $outdir)
			ControlClick($handle, "", "Button4")
			Local $handle2 = WinWait("[CLASS:#32770]", "Extraction completed")
			WinClose($handle2)
			WinClose($handle)
			Local $return = _FileRead($outdir & "\extract.log", True)
			If @error Then terminate($STATUS_FAILED, $file, $arcdisp)
			$success = $RESULT_SUCCESS

		Case "robo"
			RunWait(Warn_Execute($file & ' /unpack="' & $outdir & '"'), $filedir)

		Case "rpa"
			_Run($rpa & ' -m -v -p "' & $outdir & '" "' & $file & '"', @ScriptDir, True, True, True)

		Case "sfark"
			_Run($sfark & ' "' & $file & '" "' & $outdir & '\' & $filename & '.sf2"', $filedir, @SW_SHOW)

		Case "sgb"
			$return = _Run($7z & ' x "' & $file & '"', $outdir, @SW_HIDE)
			; Effect.sgr cannot be extracted by 7zip, that should not be considered failed extraction as all other files are OK
			If StringInStr($return, "Headers Error : Effect.sgr") And StringInStr($return, "Sub items Errors: 1") Then $success = $RESULT_SUCCESS

		Case "sim"
			HasPlugin($sim)
			_Run($sim & ' "' & $file & '" "' & $outdir & '"', $outdir, @SW_SHOW)

			; Extract cab files
			$ret = $outdir & "\data.cab"
			_Run($7z & ' x "' & $ret & '"', $outdir, @SW_SHOW, True, True, True)
			Cleanup($ret, $OPTION_DELETE)

			; Restore original file names and folder structure
			$ret = _FileRead($outdir & "\installer.config", False, 16)
			If @error Then
				$success = $RESULT_FAILED
			Else
				Local $j = 0, $aReturn = StringSplit($ret, "00402426253034", 1)
				For $i = 1 To $aReturn[0]
					If StringLeft($aReturn[$i], 2) <> "5C" Then ContinueLoop

					$ret = StringRight($aReturn[$i], 4)
					$bNext = $ret = "0031" Or $ret = "0032"
					If Not $bNext And $i <> $aReturn[0] Then ContinueLoop

					; Get file/directory name
					$ret = StringReplace(StringSplit($aReturn[$i], "0031", 1)[1], "0032", "")
					$ret = $outdir & BinaryToString("0x" & $ret)

					; Rename files/create sub directory
					$return = $outdir & "\" & $j
					If FileExists($return) Then
						FileMove($return, $ret, 9)
					Else
						DirCreate($ret)
					EndIf

					; Update next file name variable
					If $bNext Then $j += 1
				Next
			EndIf

			Local $aReturn = ["runtime.cab", "installer.config"]
			Cleanup($aReturn)

		Case "sit"
			FileMove($file, $tempoutdir, 8)
			$return = $tempoutdir & $filename & '.' & $fileext
			_Run($sit & ' "' & $return & '"', $tempoutdir, @SW_SHOW, False, False, False)
			FileMove($return, $file)
			MoveFiles($tempoutdir, $outdir, True, "", True)

		Case "sqlite"
			$return = FetchStdout($sqlite & ' "' & $file & '" .dump ', $filedir, @SW_HIDE, 0, False)
			$handle = FileOpen($outdir & '\' & $filename & '.sql', 8+2)
			FileWrite($handle, $return)
			FileClose($handle)

		Case "superdat"
			RunWait(Warn_Execute($file & ' /e "' & $outdir & '"'), $outdir)
			_FileRead($filedir & '\SuperDAT.log', True)

		Case "swf"
			; Run swfextract to get list of contents
			$return = StringSplit(FetchStdout($swf & ' "' & $file & '"', $filedir, @SW_HIDE), @CRLF)
;~ 			_ArrayDisplay($return)
			For $i = 2 To $return[0]
				$line = $return[$i]
				; Extract files
				If StringInStr($line, "MP3 Soundstream") Then
					_Run($swf & ' -m "' & $file & '"', $outdir, @SW_HIDE, True, True, False, False)
					If FileExists($outdir & "\output.mp3") Then FileMove($outdir & "\output.mp3", $outdir & "\MP3 Soundstream\soundstream.mp3", 8 + 1)
				ElseIf $line <> "" Then
					$swf_arr = StringSplit(StringRegExpReplace(StringStripWS($line, 8), '(?i)\[(-\w)\]\d+(.+):(.*?)\)', "$1,$2,"), ",")
;~ 					_ArrayDisplay($swf_arr)
					$j = 3
					Do
						;Cout("$j = " & $j & @TAB & $swf_arr[$j])
						$swf_obj = StringInStr($swf_arr[$j], "-")
						If $swf_obj Then
							For $k = StringMid($swf_arr[$j], 1, $swf_obj - 1) To StringMid($swf_arr[$j], $swf_obj + 1)
								_ArrayAdd($swf_arr, $k)
							Next
							$swf_arr[0] = UBound($swf_arr) - 1
;~ 							_ArrayDisplay($swf_arr)
						Else
							; Progress indicator
							GUICtrlSetData($idTrayStatusExt, $swf_arr[2] & ": " & $j & "/" & $swf_arr[0] + 1)

							; Set output file name
							$swf_arr[$j] = StringStripWS($swf_arr[$j], 1)
							$fname = $swf_arr[$j]

							If $swf_arr[2] = "Sounds" Or $swf_arr[2] = "Embedded MP3s" Then
								$fname &= ".mp3"
							ElseIf $swf_arr[2] = "PNGs" Then
								$fname &= ".png"
							ElseIf $swf_arr[2] = "JPEGs" Then
								$fname &= ".jpg"
							Else
								$fname &= ".swf"
							EndIf

							_Run($swf & " " & $swf_arr[1] & " " & $swf_arr[$j] & ' -o ' & $fname & ' "' & $file & '"', $outdir, @SW_HIDE, True, True, -1, False)
;~							_ArrayDisplay($swf_arr)

							FileMove($outdir & "\" & $fname, $outdir & "\" & $swf_arr[2] & "\", 8 + 1)
						EndIf
						$j += 1
					Until $j = $swf_arr[0] + 1
				EndIf
			Next

		Case "swfexe"
			$ret = $outdir & "\" & $filenamefull
			Cout("Moving file to " & $ret)
			FileMove($file, $ret)

			$aReturn = OpenExeInfo($ret)

			MouseMove(0, 0, 0)
			ControlClick($aReturn[0], "", "[CLASS:TBitBtn; INSTANCE:16]")
			ControlSend($aReturn[0], "", "[CLASS:TBitBtn; INSTANCE:16]", "{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{ENTER}")

			Local $handle = WinWait("[TITLE:s; CLASS:TSViewer]", "", $Timeout)
			Local $hControl = ControlGetHandle($handle, "", "TListBox1")

			$TimerStart = TimerInit()
			$return = -1

			While $return < 0
				Sleep(200)
				$return = _GUICtrlListBox_FindString($hControl, "--- End of file ---")
				If TimerDiff($TimerStart) > $Timeout Then ExitLoop
			WEnd

			CloseExeInfo($aReturn)

			Cout("Moving file back")
			FileMove($ret, $filedir & "\")

		Case "tar"
			_Run($7z & ' x "' & $file & '"', $outdir)

			If $fileext <> "tar" Then
				_Run($7z & ' x "' & $outdir & '\' & $filename & '.tar"', $outdir)
				FileDelete($outdir & '\' & $filename & '.tar')
			EndIf

		Case "thinstall"
			HasPlugin($thinstall)

			$pid = Run(Warn_Execute($file), $filedir)
			Do
				Sleep(100)
			Until ProcessExists($pid)
			Sleep(1000)
			Run($thinstall)
			WinWait("h4sh3m Virtual Apps Dependency Extractor")
			WinActivate("h4sh3m Virtual Apps Dependency Extractor")
			ControlSetText("h4sh3m Virtual Apps Dependency Extractor", "", "TEdit1", $pid)
			ControlClick("h4sh3m Virtual Apps Dependency Extractor", "", "TBitBtn3")
			WinWait("h4sh3m Virtual App's Extractor", "", 60)
			WinActivate("h4sh3m Virtual App's Extractor")
			ControlSetText("h4sh3m Virtual App's Extractor", "", "TEdit1", $outdir)
			ControlClick("h4sh3m Virtual App's Extractor", "", "TBitBtn1")
			WinWait("Done")
			ControlClick("Done", "", "Button1")
			WinClose("h4sh3m Virtual Apps Dependency Extractor")
			Sleep(1000)
			ProcessClose($pid)

		Case "ttarch" ; Test
			; Get all supported games
			$aReturn = _StringBetween(FetchStdout($ttarch, @ScriptDir, @SW_HIDE), "Games", "Examples")
			If @error Then terminate($STATUS_FAILED, $file, $arcdisp)
			$gameformat = StringRegExp($aReturn[0], "\d+ +(.+)", 3)
;~ 			_ArrayDisplay($gameformat)

			; Display game select GUI
			$game = GameSelect(_ArrayToString($gameformat), t('METHOD_GAME_NOGAME'))
			If $game Then
				$game = _ArraySearch($gameformat, $game)
				If $game > -1 Then _Run($ttarch & ' -m ' & $game & ' "' & $file & '" "' & $outdir & '"', $outdir, @SW_HIDE)
			Else ; Delete outdir and return
				$returnFail = True
			EndIf

		Case "uha"
			_Run($uharc & ' x -t"' & $outdir & '" "' & $file & '"', $outdir)
			If Not $success And _DirGetSize($outdir, $initdirsize + 1) <= $initdirsize Then
				_Run($uharc04 & ' x -t"' & $outdir & '" "' & $file & '"', $outdir)
				If Not $success And _DirGetSize($outdir, $initdirsize + 1) <= $initdirsize Then _
					_Run($uharc02 & ' x -t' & FileGetShortName($outdir) & ' ' & FileGetShortName($file), $outdir)
			EndIf

		Case "uif"
			_DeleteTrayMessageBox()
			Prompt(32 + 4, 'CONVERT_CDROM', 'UIF', 1)

			; Begin conversion, format selected by uif2iso
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'UIF ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 1)')
			_Run($uif & ' "' & $file & '" "' & $outdir & "\" & $filename & '"', $filedir, True, True, True)
			$handle = FileFindFirstFile($outdir & "\" & $filename & ".*")
			$isofile = $outdir & "\" & FileFindNextFile($handle)
			FileClose($handle)

		Case "unity" ; Test
			IsJavaInstalled()
			_Run($unity & ' extract "' & $file & '"', $filedir, @SW_MINIMIZE, True, True, True, False)

		Case "unreal"
			HasPlugin($unreal)
			; Currently extracts files from all packages in folder, not only the selected one!
			_Run($unreal & ' -export -all -sounds -3rdparty -path="' & $filedir & '" -out="' & $outdir & '" *', $outdir, @SW_MINIMIZE, True, True, False)

		Case "video"
			HasFFMPEG()

			; Collect information about number of tracks
			$command = $ffmpeg & ' -i "' & $file & '"'
			$return = FetchStdout($command, $outdir, @SW_HIDE)

			; Terminate if file could not be read by FFmpeg
			If StringInStr($return, "Invalid data found when processing input") Or Not StringInStr($return, "Stream") Then terminate($STATUS_FAILED, $file, $arcdisp)

			; Otherwise, extract all tracks
			$Streams = StringSplit($return, "Stream", 1)
			;_ArrayDisplay($Streams)
			Local $iVideo = 0, $iAudio = 0
			For $i = 2 To $Streams[0]
				$Streams[$i] = StringRegExpReplace($Streams[$i], "(?i)(?s).*?#(\d:\d)(.*?): (\w+): (\w+).*", "$3,$4,$1,$2")
;~ 				_ArrayDisplay($Streams)
				$StreamType = StringSplit($Streams[$i], ",")
;~ 				_ArrayDisplay($StreamType)
				If $StreamType[1] == "Video" Then
					; Split gif files into single images
					If $StreamType[2] == "gif" Or $StreamType[2] == "apng" Or $StreamType[2] == "webp" Then
						_Run($command & ' "' & ($iUnicodeMode? $sUnicodeName: $filename) & '%05d.png"', $outdir, @SW_HIDE, True, False)
						$iVideo += 1
						ContinueLoop
					EndIf

					If Not $bExtractVideo Then ContinueLoop
					$iVideo += 1
					If $StreamType[2] == "h264" Then
						_Run($command & ' -vcodec copy -an -bsf:v h264_mp4toannexb -map ' & $StreamType[3] & ' "' & ($iUnicodeMode? $sUnicodeName: $filename) & "_" & t('TERM_VIDEO') & StringFormat("_%02s", $iVideo) & $StreamType[4] & "." & $StreamType[2] & '"', $outdir, @SW_HIDE, True, False)
					Else
						; Special cases
						If StringInStr($StreamType[2], "wmv") Then
							$StreamType[2] = "wmv" ;wmv3
						ElseIf StringInStr($StreamType[2], "mpeg") Then
							$StreamType[2] = "mpeg" ;mpeg1video
						ElseIf StringInStr($StreamType[2], "v8") Then
							$StreamType[2] = "webm"
						EndIf
						_Run($command & ' -vcodec copy -an -map ' & $StreamType[3] & ' "' & ($iUnicodeMode? $sUnicodeName: $filename) & "_" & t('TERM_VIDEO') & StringFormat("_%02s", $iVideo) & $StreamType[4] & "." & $StreamType[2] & '"', $outdir, @SW_HIDE, True, True, False)
					EndIf
				ElseIf $StreamType[1] == "Audio" Then
					$iAudio += 1
					; Special cases
					If StringInStr($StreamType[2], "wma") Then
						$StreamType[2] = "wma" ;wmav2
					ElseIf StringInStr($StreamType[2], "vorbis") Then
						$StreamType[2] = "ogg"
					ElseIf StringInStr($StreamType[2], "pcm") Then
						$StreamType[2] = "wav"
					EndIf

					_Run($command & ' -acodec copy -vn -map ' & $StreamType[3] & ' "' & ($iUnicodeMode? $sUnicodeName: $filename) & "_" & t('TERM_AUDIO') & StringFormat("_%02s", $iAudio) & $StreamType[4] & "." & $StreamType[2] & '"', $outdir, @SW_HIDE, True, True, False)
				Else
					Cout("Unknown stream type: " & $StreamType[1])
				EndIf
			Next
			If $iVideo + $iAudio < 1 Then terminate($STATUS_NOTPACKED, $file, $arcdisp)

		Case "video_convert"
			HasFFMPEG()

			_Run($ffmpeg & ' -i "' & $file & '" "' & ($iUnicodeMode? $sUnicodeName: $filename) & '.mp4"', $outdir, @SW_HIDE, True, True, False)

		Case "vssfx"
			FileMove($file, $outdir)
			RunWait(Warn_Execute($outdir & '\' & $filename & '.' & $fileext & ' /extract'), $outdir)
			FileMove($outdir & '\' & $filename & '.' & $fileext, $filedir)

		Case "vssfxpath"
			RunWait(Warn_Execute($file & ' /extract:"' & $outdir & '" /quiet'), $outdir)

		Case "wise"
			Local $aReturn = ['Wise ' & t('TERM_INSTALLER'), t('METHOD_UNPACKER_RADIO', 'E_Wise'), t('METHOD_UNPACKER_RADIO', 'WUN'), t('METHOD_SWITCH_RADIO', 'Wise Installer /x'), t('METHOD_EXTRACTION_RADIO', 'Wise MSI'), t('METHOD_EXTRACTION_RADIO', 'Unzip')]
			$choice = MethodSelect($aReturn, $arcdisp)

			Switch $choice
				; Extract with E_WISE
				Case 1
					_Run($wise_ewise & ' "' & $file & '" "' & $outdir & '"', $filedir)
					If _DirGetSize($outdir, $initdirsize + 1) > $initdirsize Then
						RunWait($cmd & '00000000.BAT', $outdir, @SW_HIDE)
						FileDelete($outdir & '\00000000.BAT')
					EndIf

				; Extract with WUN
				Case 2
					RunWait(_MakeCommand($wise_wun, True) & ' "' & $filename & '" "' & $tempoutdir & '"', $filedir)
					Local $removetemp = 1
					LoadPref("removetemp", $removetemp)
					If $removetemp Then
						FileDelete($tempoutdir & "INST0*")
						FileDelete($tempoutdir & "WISE0*")
					Else
						FileMove($tempoutdir & "INST0*", $outdir)
						FileMove($tempoutdir & "WISE0*", $outdir)
					EndIf
					MoveFiles($tempoutdir, $outdir, False, "", True)

				; Extract using the /x switch
				Case 3
					RunWait(Warn_Execute($file & ' /x ' & $outdir), $filedir)

				; Attempt to extract MSI
				Case 4
					; Prompt to continue
					_DeleteTrayMessageBox()
					Prompt(48 + 4, 'WISE_MSI_PROMPT', 1)

					; First, check for any files that are already in extraction dir
					_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & $arcdisp)
					$oldfiles = ReturnFiles(@CommonFilesDir & "\Wise Installation Wizard")

					; Run installer
					Opt("WinTitleMatchMode", 3)
					$pid = Run(Warn_Execute($file & ' /?'), $filedir)
					While 1
						Sleep(10)
						If WinExists("Windows Installer") Then
							WinSetState("Windows Installer", '', @SW_HIDE)
							ExitLoop
						Else
							If Not ProcessExists($pid) Then ExitLoop
						EndIf
					WEnd

					; Move new files
					MoveFiles(@CommonFilesDir & "\Wise Installation Wizard", $outdir, 0, $oldfiles, True)
					WinClose("Windows Installer")

				; Extract using unzip, falling back to 7-Zip
				Case 5
					_Run($zip & ' -x "' & $file & '"', $outdir)
					If $success == $RESULT_FAILED Then _Run($7z & ' x "' & $file & '"', $outdir)
			EndSwitch

			; Append missing file extensions
			If $appendext Then AppendExtensions($outdir)

		Case "wix"
			If Not HasNetFramework(4) Then terminate($STATUS_MISSINGEXE, $file, ".Net Framework 4")
			_Run($wix & ' -x "' & $outdir & '" "' & $file & '"', $outdir, @SW_MINIMIZE, True, True, False)

		Case "zip"
			Local $sPassword = _FindArchivePassword($7z & ' l -p -slt "' & $file & '"', $7z & ' t -p"%PASSWORD%" "' & $file & '"', "Encrypted = +", "Wrong password?", 0, "Everything is Ok")
			_Run($7z & ' x ' & ($sPassword == 0? '"': '-p"' & $sPassword & '" "') & $file & '"', $outdir, @SW_HIDE, True, True, True, True)
			If @extended Then terminate($STATUS_PASSWORD, $file, $arcdisp)
			If $success <> $RESULT_SUCCESS Then _Run($zip & ' -x "' & $file & '"', $outdir, @SW_MINIMIZE, True, False)

		Case "zoo"
			FileMove($file, $tempoutdir, 8)
			_Run($zoo & ' -x ' & $filename & '.' & $fileext, $tempoutdir, @SW_HIDE)
			FileMove($tempoutdir & $filename & '.' & $fileext, $file)
			MoveFiles($tempoutdir, $outdir, False, "", True)

		Case "zpaq"
			; ZPaq uses a different executable for Windows XP, so a definition file cannot be used
			_Run($zpaq & ' x "' & $file & '" -to "' & $outdir & '"', $outdir, @SW_SHOW, True, True, False)

		Case Else
			pluginExtract($arctype)
			If @error Then Cout("Unknown arctype: " & $arctype & ". Feature not implemented!")
	EndSwitch

	Opt("WinTitleMatchMode", 1)
	_DeleteTrayMessageBox()


	If $isofile And FileExists($isofile) And Not ($returnSuccess Or $returnFail) Then
		; Extract converted iso file
		If FileGetSize($isofile) > 0 Then
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & StringUpper($arctype) & ' ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 2)')
			$file = $isofile
			If Not CheckIso() Then _Run($7z & ' x "' & $isofile & '"', $outdir)
			If $success <> $RESULT_FAILED Or Prompt(16 + 4, 'CONVERT_CDROM_STAGE2_FAILED', '', 0) = 0 Then
				$initdirsize -= FileGetSize($isofile)
				FileDelete($isofile)
			EndIf
		Else
			Prompt(16, 'CONVERT_CDROM_STAGE1_FAILED', '', 0)
			If FileExists($isofile) Then FileDelete($isofile)
			check7z()
			$success = $RESULT_FAILED
		EndIf
	EndIf


	; -----Success evaluation----- ;

	; Exit if success returned by _Run function
	Cout("Extraction finished, success: " & $success)
	If FileExists($tempoutdir) Then DirRemove($tempoutdir, 1)

	Switch $success
		Case $RESULT_SUCCESS
			; Special actions for 7zip extraction
			If $arctype == "7z" And ($fileext = "exe" Or StringInStr($filetype, "SFX")) Then
				; Check if sfx archive and extract sfx script using 7ZSplit if possible
				_CreateTrayMessageBox(t('SCANNING_FILE', "7z SFX Archives splitter"))
				Cout("Trying to extract sfx script")
				Run(_MakeCommand($7zsplit & ' "' & $file & '"'), $outdir, @SW_HIDE)
				WinWait("7z SFX Archives splitter")
				;WinActivate("7z SFX Archives splitter")
				ControlClick("7z SFX Archives splitter", "", "Button8")
				ControlClick("7z SFX Archives splitter", "", "Button1")
				$TimerStart = TimerInit()
				Do
					Sleep(100)
					$TimerDiff = TimerDiff($TimerStart)
					If $TimerDiff > $Timeout Then ExitLoop
				Until FileExists($filedir & "\" & $filename & ".txt") Or WinExists("7z SFX Archives splitter error")
				; Force close all messages
				ProcessClose("7ZSplit.exe")

				; Move sfx script to outdir
				If FileExists($filedir & "\" & $filename & ".txt") Then FileMove($filedir & "\" & $filename & ".txt", $outdir & "\sfx_script_" & $filename & ".txt")
				_DeleteTrayMessageBox()

				; Check generic .exe ressource extraction
				If FileExists($outdir & "\[0]") Then
					; Try to extract unpacked file (skip file extension checks)
					If Prompt(48 + 1, "UNPACK_GENERIC_ZIP", CreateArray($file, "7Zip", $outdir & "\[0]"), 0) Then
						Cout("Trying to extract unpacked file [0]")
						$file = $outdir & "\[0]"
						$outdir = $file & "\"
						FilenameParse($file)
						Return StartExtraction(False)
					Else	; Try to find correct file extensions for unpacked files
						Cout("Trying to find correct file extensions for unpacked file [0]")
						AppendExtensions($outdir)
					EndIf
				EndIf
			EndIf
		Case $RESULT_FAILED

		Case $RESULT_CANCELED

		Case $RESULT_UNKNOWN
			; Otherwise, check directory size
			If ($initdirsize > -1 And _DirGetSize($outdir, $initdirsize + 1) <= $initdirsize) Or (FileGetTime($outdir, 0, 1) == $dirmtime) Then
				If $arctype = "ace" And $fileext = "exe" Then Return False
				$success = $RESULT_FAILED
			EndIf
	EndSwitch

	If $success = $RESULT_FAILED Then Return $returnFail? 0: terminate($STATUS_FAILED, $file, $arcdisp)
	Return $returnSuccess? 1: terminate($STATUS_SUCCESS, "", $arctype)
EndFunc

Func pluginExtract($sPlugin)
	Cout("Starting custom " & $sPlugin & " extraction")

	Local $ret = $userDefDir & $sPlugin & ".ini"
	If Not FileExists($ret) Then
		$ret = $defdir & $sPlugin & ".ini"
		If Not FileExists($ret) Then terminate($STATUS_MISSINGDEF, $ret, '')
	EndIf

	Local Const $ret2 = "Extract"
	Local $sBinary = IniRead($ret, $ret2, "executable", $sPlugin)
	HasPlugin($sBinary)

	; Set status box
	If Not $NoBox Then
		Local $aReturn = StringSplit(IniRead($ret, "Display", "displayTypeTranslation", ""), ",")
		Local $arcdisp = t('EXTRACTING') & @CRLF & IniRead($ret, "Display", "display", $sPlugin)
		If StringLen($aReturn[1]) > 0 And Not StringIsSpace($aReturn[1]) Then
			For $i = 1 To $aReturn[0]
				$arcdisp &= " " & t(StringStripWS($aReturn[$i], 8))
			Next
		EndIf
		_CreateTrayMessageBox($arcdisp)
	EndIf

	Local $sParameters = " " & IniRead($ret, $ret2, "parameters", "")
	Local $sWorkingdir = IniRead($ret, $ret2, "workingdir", "")
	If $sWorkingdir = "" Then $sWorkingdir = $outdir

	$sParameters = ReplacePlaceholders($sParameters)
	$sWorkingdir = ReplacePlaceholders($sWorkingdir)

	_Run($sBinary & $sParameters, $sWorkingdir, Number(IniRead($ret, $ret2, "hide", 0)) == 1? @SW_HIDE: @SW_MINIMIZE, Number(IniRead($ret, $ret2, "useCmd", 1)) == 1, Number(IniRead($ret, $ret2, "log", 1)) == 1, Number(IniRead($ret, $ret2, "patternSearch", 0)) == 1, Number(IniRead($ret, $ret2, "initialShow", 1)) == 1)
EndFunc

; Replace % placeholders with variable contents
Func ReplacePlaceholders($sString, $bQuote = True)
	If Not StringInStr($sString, "%") Then Return $sString
	$sString = StringReplace($sString, "%filename", $filename)
	$sString = StringReplace($sString, "%fileext", $fileext)
	$sString = StringReplace($sString, "%filedir", $filedir)
	$sString = StringReplace($sString, "%file", $bQuote? Quote($file): $file)
	$sString = StringReplace($sString, "%outdir", $bQuote? Quote($outdir): $outdir)
	Return $sString
EndFunc

; Encapsulate a string with "
Func Quote($sString, $bDouble = False)
	Return ($bDouble? '""': '"') & $sString & '"'
EndFunc

; Unpack packed executable
Func unpack($packer)
	If $unpackfailed Then Return
	$unpackfailed = True ; Don't display the prompt multiple times
	$ret = $filedir & "\" & $filename & "_" & t('TERM_UNPACKED') & "." & $fileext

	; Prompt to continue
	If Not Prompt(32 + 4, 'UNPACK_PROMPT', CreateArray($packer, $ret), 0) Then Return

	; Unpack file
	Switch $packer
		Case $PACKER_UPX
			_Run($upx & ' -d -k "' & $file & '"', $filedir)
			$tempext = StringTrimRight($fileext, 1) & '~'
			If FileExists($filedir & "\" & $filename & "." & $tempext) Then
				FileMove($file, $ret)
				FileMove($filedir & "\" & $filename & "." & $tempext, $file)
			EndIf
		Case $PACKER_ASPACK
			RunWait($aspack & ' "' & $file & '" "' & $filedir & '\' & $filename & '_' & t('TERM_UNPACKED') & _
					$fileext & '" /NO_PROMPT', $filedir)
	EndSwitch

	; Success evaluation
	If FileExists($ret) Then
		; Prompt if unpacked file should be scanned
		If Prompt(32 + 4, 'UNPACK_AGAIN', CreateArray($file, $filename & '_' & t('TERM_UNPACKED') & "." & $fileext), 0) Then
			$file = $ret
			$outdir = $filedir & "\" & $filename & "_" & t('TERM_UNPACKED') & "\"
			StartExtraction()
		Else
			terminate($STATUS_SUCCESS, "", $packer)
		EndIf
	Else
		Prompt(16, 'UNPACK_FAILED', $file, 1)
	EndIf
EndFunc

; Perform outdir cleanup: move/delete given files according to $iCleanup setting
Func Cleanup($aFiles, $iMode = $iCleanup, $dir = 0)
	If Not $iMode Then Return
	If Not IsArray($aFiles) Then
		If Not FileExists($aFiles) And Not FileExists($outdir & "\" & $aFiles) Then Return SetError(1, 0, 0)
		$return = $aFiles
		Dim $aFiles = [$return]
	EndIf

	If $iMode = $OPTION_MOVE And $dir == 0 Then $dir = $outdir & "\" & t('DIR_ADDITIONAL_FILES')
;~ 	Cout("Cleanup - " & _ArrayToString($aFiles))

	For $sFile In $aFiles
		If Not StringInStr($sFile, $outdir) Then $sFile = $outdir & "\" & $sFile
		If Not FileExists($sFile) Then ContinueLoop
		$bIsFolder = StringInStr(FileGetAttrib($sFile), "D")
		If $iMode = $OPTION_DELETE Then
			Cout("Cleanup: Deleting " & $sFile)
			If $bIsFolder Then
				DirRemove($sFile, 1)
			Else
				FileDelete($sFile)
			EndIf
		Else
			If Not FileExists($dir) Then DirCreate($dir)
			Cout("Cleanup: Moving " & $sFile & " to " & $dir)
			If $bIsFolder Then
				DirMove($sFile, $dir, 1)
			Else
				FileMove($sFile, $dir, 1)
			EndIf
		EndIf
	Next
EndFunc

; Check write permissions for specified file or folder
Func CanAccess($sPath)
	Local $bIsTemp = False

	Cout("Checking permissions for path " & $sPath)
	$bExists = FileExists($sPath)
	If StringInStr(FileGetAttrib($sPath), "D") Or ($bExists = False And StringRight($sPath, 1)) Then
		$bIsTemp = True
		$sPath = _TempFile($sPath)
	Else
		If Not $bExists Then Return SetError(1, 0, False)
	EndIf

	$handle = FileOpen($sPath, 1)
	If $handle == -1 Then
		Cout("Access denied")
		Return False
	EndIf
	FileClose($handle)
	If $bIsTemp Then FileDelete($sPath)
	Return True
EndFunc

; Terminate if specified plugin was not found
Func HasPlugin($plugin, $returnFail = False)
	$plugin = StringReplace($plugin, '"', '')
	Cout("Searching for plugin " & $plugin)
	If FileExists($plugin) Or (_WinAPI_PathIsRelative($plugin) And (FileExists(_PathFull($plugin, $bindir)) Or FileExists(_PathFull($plugin, $archdir)))) Then Return True
	Cout("Plugin not found")
	If $returnFail Then Return False
	terminate($STATUS_MISSINGEXE, $file, $plugin)
EndFunc

; Search for translation file for given language and return result
Func HasTranslation($language)
	If $language = "English" Then Return True
	$return = FileExists($langdir & $language & ".ini")
	If Not $return Then Cout("Language file for " & $language & " does not exist")
	Return $return
EndFunc

; Check if enough free space is available
Func HasFreeSpace($sPath = $outdir, $fModifier = 1)
	While Not StringInStr(FileGetAttrib($sPath), "D")
		$pos = StringInStr($sPath, "\", 0, -1)
		If $pos < 1 Then ExitLoop
		$sPath = StringLeft($sPath, $pos - 1)
	WEnd

	Local $freeSpace = Round(DriveSpaceFree($sPath), 2)
	Local $fileSize = Round(FileGetSize($file) / 1048576, 2) * $fModifier; MB

	If $freeSpace < $fileSize Then
		Local $diff = Round(Abs($freeSpace - $fileSize), 2)
		Cout("Not enough free space available: " & $freeSpace & " MB, needed: " & $fileSize & " MB, difference: " & $diff & " MB.")
		Local $Msg = t('NO_FREE_SPACE', CreateArray(StringLeft($sPath, 1), $freeSpace, $fileSize, $diff))
		If $silentmode Then terminate($STATUS_FAILED, '', $Msg)

		$return = MsgBox($iTopmost + 48 + 2, $name, $Msg)
		If $return = 4 Then ; Retry
			Return HasFreeSpace($sPath)
		ElseIf $return = 3 Then ; Cancel
			If $createdir Then DirRemove($outdir, 0)
			terminate($STATUS_SILENT, '', '')
		EndIf
	EndIf
EndFunc

; Search for FFMPEG and prompt to download it if not found
Func HasFFMPEG()
	If HasPlugin($ffmpeg, True) Then Return
	Prompt(48 + 4, 'FFMPEG_NEEDED', CreateArray($file, "https://ffmpeg.org/legal.html"), 1)
	GetFFmpeg()
	If @error Then terminate($STATUS_SILENT, "", "")
EndFunc

; Create a temporary directory which did not exist before
Func TempDir($sDir, $iLen)
	Do
		Local $return = ""
		While StringLen($return) < $iLen
			$return &= Chr(Random(97, 122, 1))
		WEnd
		$return = $sDir & "\" & $return & "\"
	Until Not FileExists($return)
	Return $return
EndFunc

; Return list of files and directories in directory as a pipe-delimited string
Func ReturnFiles($dir)
	Local $handle, $files, $fname
	$handle = FileFindFirstFile($dir & "\*")
	If Not @error Then
		While 1
			$fname = FileFindNextFile($handle)
			If @error Then ExitLoop
			$files &= $fname & '|'
		WEnd
		$files = StringTrimRight($files, 1)
		FileClose($handle)
	Else
		SetError(1)
		Return
	EndIf
	Return $files
EndFunc   ;==>ReturnFiles

; Move all files and subdirectories from one directory to another
; $force is an integer that specifies whether or not to replace existing files
; $omit is a string that includes files to be excluded from move
Func MoveFiles($source, $dest, $force = False, $omit = '', $removeSourceDir = False)
	Local $handle, $fname
	Cout("Moving files from " & $source & " to " & $dest)
	DirCreate($dest)

	$handle = FileFindFirstFile($source & "\*")
	If @error Then Return SetError(1)
	While 1
		$fname = FileFindNextFile($handle)
		If @error Then ExitLoop
		If StringInStr($omit, $fname) Then ContinueLoop

		If StringInStr(FileGetAttrib($source & '\' & $fname), 'D') Then
			DirMove($source & '\' & $fname, $dest, 1)
		Else
			FileMove($source & '\' & $fname, $dest, $force)
		EndIf
	WEnd
	FileClose($handle)

	If $removeSourceDir Then Return DirRemove($source, ($omit = ""? 1: 0))
EndFunc

; Append missing file extensions using TrID
Func AppendExtensions($dir)
	Local $files = FileSearch($dir)
	If $files[1] = '' Then Return
	For $i = 1 To $files[0]
		If StringInStr(FileGetAttrib($files[$i]), 'D') Then ContinueLoop
		$filename = StringTrimLeft($files[$i], StringInStr($files[$i], '\', 0, -1))
		If StringInStr($filename, '.') And Not (StringLeft($filename, 7) = 'Binary.' Or StringRight($filename, 4) = '.bin') Then ContinueLoop
		RunWait($cmd & $trid & ' "' & $files[$i] & '" -ae', $dir, @SW_HIDE)
	Next
EndFunc

; Recursively search for given pattern
; code by w0uter (http://www.autoitscript.com/forum/index.php?showtopic=16421)
Func FileSearch($s_Mask = '', $i_Recurse = 1)
	Local $s_Buf = ''
	Local $s_Command = Cout(@ComSpec & ' /c dir /B ' & ($i_Recurse? '/S ': '') & Quote($s_Mask))
	$i_Pid = Run($s_Command, @WorkingDir, @SW_HIDE, 2 + 4)
	While Not @error
		$s_Buf &= StdoutRead($i_Pid)
	WEnd
	$s_Buf = StringSplit(StringTrimRight($s_Buf, 2), @CRLF, 1)
	ProcessClose($i_Pid)
	If UBound($s_Buf) = 2 And $s_Buf[1] = '' Then SetError(1)
	Return $s_Buf
EndFunc

; Open file and return contents
Func _FileRead($f, $delete = False, $iFlag = 0)
	$handle = FileOpen($f, $iFlag)
	If $handle = -1 Then Return SetError(1, 0, "")
	$return = FileRead($handle)
	FileClose($handle)
	Cout($return)
	If $delete Then FileDelete($f)
	Return $return
EndFunc

; Handle program termination with appropriate error message
Func terminate($status, $fname = '', $sFileType = '')
	Local $LogFile = 0, $exitcode = 0, $shortStatus = ($status = $STATUS_SUCCESS)? $sFileType: $status

	; Rename unicode file
	If $iUnicodeMode Then
		Cout("Renaming unicode file")
		_CreateTrayMessageBox(t('MOVING_FILE') & @CRLF & $oldoutdir)
		If $iUnicodeMode = $UNICODE_MOVE Then
			FileMove($file, $oldpath, 1)
		Else
			FileRecycle($file)
		EndIf
		Cout("Moving extracted files: " & DirMove($outdir, $oldoutdir))
		$fname = $oldpath
		$file = $oldpath
	EndIf

	_DeleteTrayMessageBox()
	Cout("Terminating - Status: " & $status)

	; When multiple files are selected and executed via command line, they are added to batch queue, but the working instance uses in-memory data.
	; So we need to look for changes in the batch queue file, so batch mode could be enabled if necessary.
	If Not $silentmode And GetBatchQueue() Then $silentmode = True

	; Create log file if enabled in options
	If $Log And Not ($status = $STATUS_SILENT Or $status = $STATUS_SYNTAX Or $status = $STATUS_FILEINFO Or $status = $STATUS_NOTPACKED Or $status = $STATUS_BATCH) Or ($status = $STATUS_FILEINFO And $silentmode) Then _
		$LogFile = CreateLog($shortStatus)

	; Save statistics
	IniWrite($prefs, "Statistics", $status, Number(IniRead($prefs, "Statistics", $status, 0)) + 1)
	If $status = $STATUS_SUCCESS Then IniWrite($prefs, "Statistics", $sFileType, Number(IniRead($prefs, "Statistics", $sFileType, 0)) + 1)

	; Remove singleton
	If $hMutex <> 0 Then DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hMutex)

	Switch $status
		; Display usage information and exit
		Case $STATUS_SYNTAX
			$syntax = t('HELP_SUMMARY')
			$syntax &= t('HELP_SYNTAX', @ScriptName)
			$syntax &= t('HELP_ARGUMENTS')
			$syntax &= t('HELP_HELP', "/help")
			$syntax &= t('HELP_PREFS', "/prefs")
			$syntax &= t('HELP_REMOVE', "/remove")
			$syntax &= t('HELP_CLEAR', "/batchclear")
			$syntax &= t('HELP_UPDATE', "/update")
			$syntax &= t('HELP_FILENAME')
			$syntax &= t('HELP_DESTINATION')
			$syntax &= t('HELP_SCAN', "/scan")
			$syntax &= t('HELP_SILENT', "/silent")
			$syntax &= t('HELP_BATCH', "/batch")
			$syntax &= t('HELP_SUB', "/sub")
			$syntax &= t('HELP_LAST', "/last")
			$syntax &= t('HELP_EXAMPLE1')
			$syntax &= t('HELP_EXAMPLE2', @ScriptName)
			$syntax &= t('HELP_NOARGS')
			MsgBox($iTopmost + 32, $title, $syntax, $sFileType? 15: 0)

		; Display file type information and exit
		Case $STATUS_FILEINFO
			If StringStripWS($filetype, 4+2+1) == "" Then
				$exitcode = 4
				$filetype = t('UNKNOWN_EXT', CreateArray($file, ""))
			EndIf
			If $silentmode Then ; Save info to result file if in silent mode
				$handle = FileOpen($fileScanLogFile, 8 + 1)
				FileWrite($handle, $file & @CRLF & @CRLF & $filetype & @CRLF & "------------------------------------------------------------" & @CRLF)
				FileClose($handle)
			Else
				MsgBox($iTopmost + 64, $title, $filetype)
			EndIf

		; Display error information and exit
		Case $STATUS_UNKNOWNEXE
			If Not $silentmode And Prompt(256 + 16 + 4, 'CANNOT_EXTRACT', CreateArray($file, $filetype), 0) Then Run($exeinfope & ' "' & $file & '"', $filedir)
			$exitcode = 3
		Case $STATUS_UNKNOWNEXT
			Prompt(16, 'UNKNOWN_EXT', CreateArray($file, $filetype), 0)
			$exitcode = 4
		Case $STATUS_INVALIDFILE
			Prompt(16, 'INVALID_FILE', $fname, 0)
			$exitcode = 5
		Case $STATUS_INVALIDDIR
			Prompt(16, 'INVALID_DIR', $fname, 0)
			$exitcode = 5
		Case $STATUS_NOTPACKED
			Prompt(48, 'NOT_PACKED', CreateArray($file, $filetype), 0)
			$exitcode = 6
		Case $STATUS_NOTSUPPORTED
			Prompt(16, 'NOT_SUPPORTED', CreateArray($file, $filetype), 0)
			$exitcode = 7
		Case $STATUS_MISSINGEXE
			Prompt(48, 'MISSING_EXE', CreateArray($file, $sFileType))
			$exitcode = 8
		Case $STATUS_TIMEOUT
			Prompt(48, 'EXTRACT_TIMEOUT', $file)
			$exitcode = 9
		Case $STATUS_PASSWORD
			$exitcode = 10
			Prompt(48, 'WRONG_PASSWORD', CreateArray($file, StringReplace(t('MENU_EDIT_LABEL'), "&", "")))
		Case $STATUS_MISSINGDEF
			$exitcode = 11
			Prompt(48, 'MISSING_DEFINITION', CreateArray($file, $fname))
		Case $STATUS_MOVEFAILED
			$exitcode = 12
			Prompt(48, 'MOVE_FAILED', CreateArray($file, $fname))

			; Display failed attempt information and exit
		Case $STATUS_FAILED
			If Not $silentmode And Prompt(256 + 16 + 4, 'EXTRACT_FAILED', CreateArray($file, $sFileType), 0) Then ShellExecute($LogFile? $LogFile: CreateLog($status))
			$exitcode = 1

			; Exit successfully
		Case $STATUS_SUCCESS
			If $iDeleteOrigFile = $OPTION_DELETE Then
				FileDelete($file)
			ElseIf $iDeleteOrigFile = $OPTION_ASK Then
				If Not $silentmode And Prompt(32 + 4, 'FILE_DELETE', $file) Then FileDelete($file)
			EndIf
			If $OpenOutDir And Not $silentmode Then Run("explorer.exe /e, " & $outdir)
	EndSwitch

	; Write error log if in batchmode
	If $exitcode <> 0 And $silentmode And $extract Then
		$handle = FileOpen($logdir & "errorlog.txt", 8 + 1)
		FileWrite($handle, ($filename = ""? $fname: $filenamefull) & " (" & StringUpper($status) & ")" & @CRLF & @TAB & $sFileType & @CRLF)
		FileClose($handle)
	EndIf

	; Delete empty output directory if failed
	If $createdir And $status <> $STATUS_SUCCESS And DirGetSize($outdir) = 0 Then DirRemove($outdir, 0)

	If $exitcode == 1 Or $exitcode == 3 Or $exitcode == 4 Or $exitcode == 12 Then
		If $FB_ask And $extract And Not $silentmode And Prompt(32+4, 'FEEDBACK_PROMPT', $file, 0) Then
			; Attach input file's first bytes for debug purpose
			Cout("--------------------------------------------------File dump--------------------------------------------------" & _
				 @CRLF & _HexDump($file, 1024))
			Cout("------------------------------------------------File metadata------------------------------------------------" & _
				 @CRLF & _ArrayToString(_GetExtProperty($file), @CRLF))
			; Prompt to send feedback
			GUI_Feedback($status, $file)
		EndIf
	EndIf

	If $batchEnabled = 1 And $status <> $STATUS_SILENT Then ; Don't start batch if gui is closed
		; Start next extraction
		BatchQueuePop()
	ElseIf $KeepOpen And $cmdline[0] = 0 And $status <> $STATUS_SILENT Then
		Run(@ScriptFullPath)
	EndIf

	; Check for updates
	If $status <> $STATUS_SILENT Then
		If Not $silentmode Then CheckUpdate($UPDATEMSG_FOUND_ONLY, True)
		SendStats($status, $sFileType)
	EndIf

	Exit $exitcode
EndFunc

; Function to prompt user for choice of extraction method
Func MethodSelect($aData, $arcdisp)
	; Auto choose first extraction method in silent mode
	If $silentmode Then
		Cout("Extractor selected automatically - run again in normal mode if not extracted correctly")
		Return 1
	EndIf

	_DeleteTrayMessageBox()
	Local Const $base_height = 130, $base_radio = 100
	Local $size = UBound($aData) - 1, $select[$size]

	; Create GUI and set header information
	Opt("GUIOnEventMode", 0)
	Local $hGUI = GUICreate($title, 330, $base_height + ($size * 20))
	_GuiSetColor()
	$header = GUICtrlCreateLabel(t('METHOD_HEADER', $aData[0]), 5, 5, 320, 20)
	GUICtrlSetFont(-1, -1, 1200)
	GUICtrlCreateLabel(t('METHOD_TEXT_LABEL', $aData[0]), 5, 25, 320, 65, $SS_LEFT)

	; Create radio selection options
	GUICtrlCreateGroup(t('METHOD_RADIO_LABEL'), 5, $base_radio, 215, 25 + ($size * 20))
	For $i = 0 To $size - 1
		$select[$i] = GUICtrlCreateRadio($aData[$i + 1], 10, $base_radio + 20 + ($i * 20), 205, 20)
	Next
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; Create buttons
	Local $idOk = GUICtrlCreateButton(t('OK_BUT'), 235, $base_radio - 10 + ($size * 10), 80, 20)
	Local $idCancel = GUICtrlCreateButton(t('CANCEL_BUT'), 235, $base_radio - 10 + ($size * 10) + 30, 80, 20)

	; Set properties
	GUICtrlSetState($select[0], $GUI_CHECKED)
	GUICtrlSetState($idOk, $GUI_DEFBUTTON)
	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			; Set extract command
			Case $idOk
				For $i = 0 To $size - 1
					If GUICtrlRead($select[$i]) == $GUI_CHECKED Then
						GUIDelete($hGUI)
						Opt("GUIOnEventMode", 1)
						_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & $arcdisp)
						Return $i + 1
					EndIf
				Next
				; Exit if Cancel clicked or window closed
			Case $GUI_EVENT_CLOSE, $idCancel
				If $createdir Then DirRemove($outdir, 0)
				terminate($STATUS_SILENT, '', '')
		EndSwitch
	WEnd
EndFunc

; Create GUI to select game and return selection
Func GameSelect($sEntries, $sStandard)
	Local $sSelection = 0
	$GameSelectGUI = GUICreate($title, 274, 460, -1, -1, BitOR($WS_SIZEBOX, $WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU))
	_GuiSetColor()
	$GameSelectLabel = GUICtrlCreateLabel(t('METHOD_GAME_LABEL', CreateArray($filenamefull, $sStandard)), 10, 8, 252, 210, $SS_CENTER)
	$GameSelectList = GUICtrlCreateList("", 24, 225, 225, 188, BitOR($WS_VSCROLL, $WS_HSCROLL, $LBS_NOINTEGRALHEIGHT))
	GUICtrlSetData(-1, $sStandard & '|' & $sEntries)
	$ok = GUICtrlCreateButton(t('OK_BUT'), 40, 427, 81, 25)
	$cancel = GUICtrlCreateButton(t('CANCEL_BUT'), 152, 427, 81, 25)
	_GUICtrlListBox_UpdateHScroll($GameSelectList)
	_GUICtrlListBox_SetCurSel($GameSelectList, 0)
	GUISetState(@SW_SHOW)
	Opt("GUIOnEventMode", 0)

	While 1
		$nMsg = GUIGetMsg($GameSelectGUI)
		Switch $nMsg
			Case $ok
				$sSelection = GUICtrlRead($GameSelectList)
				If $sSelection == $sStandard Then $sSelection = 0
				ExitLoop
			Case $GUI_EVENT_CLOSE, $cancel
				ExitLoop
		EndSwitch
	WEnd
	GUIDelete($GameSelectGUI)
	Opt("GUIOnEventMode", 1)
	Return $sSelection
EndFunc

; Warn user before executing files for extraction
Func Warn_Execute($command)
	If $warnexecute Then
		Cout("Displaying warn_execute message")
		If MsgBox($iTopmost + 49, $title, t('WARN_EXECUTE', $command)) <> 1 Then
			If $createdir Then DirRemove($outdir, 0)
			terminate($STATUS_SILENT, '', '')
		EndIf
	EndIf
	Return $command
EndFunc

; Create array on the fly
; Code based on _CreateArray UDF, which has been deprecated
Func CreateArray($i0, $i1 = 0, $i2 = 0, $i3 = 0, $i4 = 0, $i5 = 0, $i6 = 0, $i7 = 0, $i8 = 0, $i9 = 0)
	Local $arr[10] = [$i0, $i1, $i2, $i3, $i4, $i5, $i6, $i7, $i8, $i9]
	ReDim $arr[@NumParams]
	Return $arr
EndFunc

; Display a custom prompt and return user choice
Func Prompt($show_flag, $Message, $vars = 0, $terminate = 1)
	If $silentmode Then
		Cout("Assuming yes to message " & $Message)
		Return 1
	EndIf
	Local $return = MsgBox($iTopmost + $show_flag, $title, t($Message, $vars))
	If $return == 1 Or $return == 6 Then
		Return 1
	Else
		If Not $terminate Then Return 0
		If $createdir Then DirRemove($outdir, 0)
		terminate($STATUS_SILENT, '', '')
	EndIf
EndFunc

; Show Tray Message
; Based on work by Valuater (http://www.autoitscript.com/forum/topic/85977-system-tray-message-box-udf/)
Func _CreateTrayMessageBox($TBText)
	_DeleteTrayMessageBox()

	If $NoBox = 1 Then Return

	; Hide if in fullscreen
	If $bHideStatusBoxIfFullscreen Then
		$aReturn = WinGetPos("[ACTIVE]")
		If $aReturn[2] = @DesktopWidth And $aReturn[3] = @DesktopHeight Then Return
	EndIf

	Local Const $TBwidth = 225, $TBheight = 100, $left = 15, $top = 12, $width = 195, $iBetween = 5, $iMaxCharCount = 28, $bDark = True

	; Determine taskbar size
	Local $pos = WinGetPos("[CLASS:Shell_TrayWnd]")
	If @error Then Local $pos[4] = [0, 0, @DesktopWidth, 30]
	Local $iSpace = ($pos[0] = $pos[1])? $pos[3] + $iBetween: $pos[1] - $TBheight - $iBetween

	If $iSpace < 0 Or $iSpace > @DesktopHeight Then $iSpace = @DesktopHeight - $TBheight - $iBetween

	; Create GUI
	Global $TBgui = GUICreate($name, $TBwidth, $TBheight, $trayX > -1 ? $trayX : @DesktopWidth - ($TBwidth + $iBetween), $trayY > -1 ? $trayY : $iSpace, _
			$WS_POPUP, BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
	GUISetBkColor($bDark? 0x2D2D2D: 0xEEEEEE)
	_GuiRoundCorners($TBgui, 0, 0, 5, 5)

	; Labels
	Local $fname = $filename == ""? "": ($iUnicodeMode? $sUnicodeName: $filename) & "." & $fileext
	If StringLen($fname) > $iMaxCharCount Then $fname = StringLeft($fname, $iMaxCharCount) & " [...]"
	Local $idTrayFileName = GUICtrlCreateLabel($fname, $left, $top, $width, 16)
	Local $idTrayStatus = GUICtrlCreateLabel($TBText, $left, GetPos($TBgui, $idTrayFileName, 6, False), $width, 30)
	Global $idTrayStatusExt = GUICtrlCreateLabel("", $left, GetPos($TBgui, $idTrayStatus, 8, False), $width, 20, $SS_CENTER)

	GUICtrlSetFont($idTrayFileName, 9, 500, 0, "Arial")
	GUICtrlSetFont($idTrayStatus, 9, 500, 0, "Arial")
	GUICtrlSetFont($idTrayStatusExt, 9, 500, 0, "Arial")

	If $bDark Then
		GUICtrlSetColor($idTrayFileName, 0xFFFFFF)
		GUICtrlSetColor($idTrayStatus, 0xFFFFFF)
		GUICtrlSetColor($idTrayStatusExt, 0xFFFFFF)
	EndIf

;~     DllCall ( "user32.dll", "int", "AnimateWindow", "hwnd", $TBgui, "int", 250, "long", 0x00080000 )
	GUISetState(@SW_SHOWNOACTIVATE)

	; Workaround to keep corners round while fading in
	For $i = 0 To 225 Step 10
		WinSetTrans($TBgui, "", $i)
		Sleep(1)
	Next
EndFunc

; Close Tray Message
; Based on work by Valuater (http://www.autoitscript.com/forum/topic/85977-system-tray-message-box-udf/)
Func _DeleteTrayMessageBox()
	If Not $TBgui Then Return
	;DllCall ( "user32.dll", "int", "AnimateWindow", "hwnd", $TBgui, "int", 300, "long", 0x00090000 )

	; Workaround to keep corners round while fading out
	For $i = 225 To 0 Step -10
		WinSetTrans($TBgui, "", $i)
		Sleep(1)
	Next

	GUIDelete($TBgui)
	$TBgui = 0
EndFunc

; Create command line for current file
Func GetCmd($silent = True)
	If Not $file Then Return
	Local $return = Quote($file)

	If $extract Then
		If $outdir == "/sub" Then
			$return &= " " & $outdir
		ElseIf $outdir <> "" Then
			$return &= ' "' & $outdir & '"'
		EndIf
	Else
		$return &= " /scan"
	EndIf

	If $silentmode Or $silent Then $return &= " /silent"
	Return $return
EndFunc   ;==>GetCmd

; Add file to batch queue
Func AddToBatch()
	Local $cmdline = GetCmd()
	$handle = FileOpen($batchQueue, 32 + 8 + 1)
	FileSetPos($handle, 0, 0)
	Local $return = FileRead($handle)
	Local $ret = StringRegExpReplace($cmdline, '(".*?\.part)(\d+\.rar".*)', "$1", 1)
	If (@extended > 0 And StringInStr($return, $ret)) Or _ ; Only add one file if multiple part rar
	   (StringInStr($return, $cmdline) And Not Prompt(32 + 4, 'BATCH_DUPLICATE', $file, 0)) Then
		Cout("Not adding duplicate file " & $cmdline)
		FileClose($handle)
		Return
	EndIf
	FileWriteLine($handle, $cmdline)
	FileClose($handle)
	Cout("File added to batch queue: " & $cmdline)
	GetBatchQueue()
EndFunc

; Read batch queue from file
Func GetBatchQueue()
	_FileReadToArray($batchQueue, $queueArray)

	If Not @error And $queueArray[0] > 0 Then
;~ 		_ArrayDisplay($queueArray)
		If $guimain Then GUICtrlSetData($BatchBut, t('BATCH_BUT') & " (" & $queueArray[0] & ")")
		EnableBatchMode()
		Return 1
	EndIf

	Return 0
EndFunc   ;==>GetBatchQueue

; Write batch queue array to file
Func SaveBatchQueue()
	Cout("Saving batch queue")
;~ 	_ArrayDisplay($queueArray)
	$handle = FileOpen($batchQueue, 8 + 2)
	FileWrite($handle, _ArrayToString($queueArray, @CRLF, 1))
	FileClose($handle)
EndFunc   ;==>SaveBatchQueue

; Returns first element of batch queue
Func BatchQueuePop()
;~ 	_ArrayDisplay($queueArray)
	If Not IsArray($queueArray) Or UBound($queueArray) = 0 Then GetBatchQueue()

	If Not IsArray($queueArray) Or UBound($queueArray) = 0 Or $queueArray[0] = 0 Then ; Queue empty
		Cout("Batch queue empty")
		EnableBatchMode(False)
		If FileExists($fileScanLogFile) Then ShellExecute($fileScanLogFile)
		Local $return = _FileRead($logdir & "errorlog.txt", True)
		If $return <> "" Then MsgBox($iTopmost + 48, $name, t('BATCH_FINISH', $return))
		If $KeepOpen Then Run(@ScriptFullPath)
	Else ; Get next command and execute it
		Local $element = $queueArray[1]
		_ArrayDelete($queueArray, 1)
		$queueArray[0] -= 1
		Cout("Next batch element: " & $element)
		SaveBatchQueue()
		Run(@ScriptFullPath & " " & $element)
	EndIf
EndFunc   ;==>BatchQueuePop

; Enable batch mode
Func EnableBatchMode($enable = True)
	If $enable Then
		; Delete old filescan log file
		If FileExists($fileScanLogFile) And $extract Then
			If Not FileDelete($fileScanLogFile) Then
				Sleep(2000)
				FileDelete($fileScanLogFile)
			EndIf
		EndIf

		If $guimain Then
			GUICtrlSetOnEvent($ok, "GUI_Batch_OK")
			GUICtrlSetState($showitem, $GUI_ENABLE)
			GUICtrlSetState($clearitem, $GUI_ENABLE)
		EndIf
	Else
		; Delete empty batch queue file
		If FileExists($batchQueue) Then
			If Not FileDelete($batchQueue) Then
				Sleep(2000)
				FileDelete($batchQueue)
			EndIf
		EndIf

		If $guimain Then
			GUICtrlSetOnEvent($ok, "GUI_OK")
			GUICtrlSetData($BatchBut, t('BATCH_BUT'))
			GUICtrlSetState($showitem, $GUI_DISABLE)
			GUICtrlSetState($clearitem, $GUI_DISABLE)
		EndIf
	EndIf

;~ 	If $batchEnabled == $enable Then Return
	$batchEnabled = $enable
	SavePref("batchenabled", Number($batchEnabled))
EndFunc   ;==>EnableBatchMode

; Detect language of user's operating system
; Based on work by guinness (http://www.autoitscript.com/forum/topic/131832-getoslanguage-retrieve-the-language-of-the-os/)
Func _GetOSLanguage()
	Local $aString[36] = [35, "0409 0809 0c09 1009 1409 1809 1c09 2009 2409 2809 2c09 3009 3409", "0804 0c04 1004 0406", "0406", _
			"0413 0813", "0425", "040b", "040c 080c 0c0c 100c 140c 180c", "0407 0807 0c07 1007 1407", "040e", "0410 0810", _
			"0411", "0414 0814", "0415", "0816", "0418", "0419", "081a 0c1a", _
			"040a 080a 0c0a 100a 140a 180a 1c0a 200a 240a 280a 2c0a 300a 340a 380a 3c0a 400a 440a 480a 4c0a 500a", "041d 081d", _
			"0401 0801 0c01 1001 1401 1801 1c01 2001 2401 2801 2801 3001 3401 3801 3c01 4001", "042b", "0402", "041a", "0405", "0408", _
			"0412", "0429", "0416", "041b", "0404", "041e", "041f", "0422", "0403", "042a"]

	Local $aLanguage[36] = [35, "English", "Chinese (Simplified)", "Danish", "Dutch", "Estonian", "Finnish", "French", "German", "Hungarian", "Italian", _
			"Japanese", "Norwegian", "Polish", "Portuguese", "Romanian", "Russian", "Serbian", "Spanish", "Swedish", "Arabic", "Armenian", _
			"Bulgarian", "Croatian", "Czech", "Greek", "Korean", "Farsi", "Portuguese (Brazilian)", "Slovak", "Taiwanese", "Thai", "Turkish", _
			"Ukrainian", "Catalan", "Vietnamese"]

	For $i = 1 To $aString[0]
		If StringInStr($aString[$i], @OSLang) Then
			Cout("Selecting language based on OS language: " & $aLanguage[$i])
			Return $aLanguage[$i]
		EndIf
	Next
	Return $aLanguage[1]
EndFunc

; Determine whether JRE is installed or not and terminate if not found
Func IsJavaInstalled()
	Local $return = FetchStdout($cmd & "java", @ScriptDir, @SW_HIDE)
	If StringInStr($return, "java [-options]") Then
		Cout("Java is installed")
		Return True
	EndIf
	terminate($STATUS_MISSINGEXE, $file, "Java Runtime Environment")
EndFunc

; Determine versions of installed .NET frameworks
; Modified version of (https://www.autoitscript.com/forum/topic/164455-check-net-framework-4-or-45-is-installed/#comment-1199620)
Func HasNetFramework($iVersion)
	Local $sKey, $sBaseKeyName, $sBVersion, $sBBVersion, $z = 0, $i = 0
    $sKey = "HKLM" & $reg64 & "\SOFTWARE\Microsoft\NET Framework Setup\NDP"

    Do
        $z += 1
        $sBaseKeyName = RegEnumKey($sKey, $z)
        If @error Or StringLeft($sBaseKeyName,1) <> "v" Then ContinueLoop

        $sBVersion = RegRead($sKey & "\" & $sBaseKeyName, "Version")
		If Not @error Then
			If Number($sBVersion) >= $iVersion Then Return True
		Else
			$i = 0
			Do
				$i += 1
				$sKeyName = RegEnumKey($sKey & "\" & $sBaseKeyName, $i)
				If @error Then ExitLoop

				$sBBVersion = RegRead($sKey & "\" & $sBaseKeyName & "\" & $sKeyName, "Version")
			Until $sKeyName = '' Or $sBBVersion <> ''

			If $sBBVersion <> '' And Number($sBBVersion) >= $iVersion Then Return True
		EndIf
    Until $sBaseKeyName = ''

    Return False
EndFunc

; Determine whether Windows version >= Windows 7 or not; used for cascading context menu support
Func _IsWin7()
	Global $win7 = @OSVersion = "WIN_7" Or @OSVersion = "WIN_8" Or @OSVersion = "WIN_81" Or @OSVersion = "WIN_10" Or @OSVersion = "WIN_2016" Or @OSVersion = "WIN_2012R2" Or @OSVersion = "WIN_2012"
	Return $win7
EndFunc

; Determine whether Windows version is XP or not
Func _IsWinXP($sTrue = "", $sFalse = "")
	$return = StringInStr(@OSVersion, "WIN_XP")
	Return @NumParams == 0? $return: ($return? $sTrue: $sFalse)
EndFunc

; Determine if a key exists in registry
; Script by guinness (http://www.autoitscript.com/forum/topic/131425-registry-key-exists/page__view__findpost__p__915063)
Func RegExists($sKeyName, $sValueName)
	RegRead($sKeyName, $sValueName)
	Return Number(@error = 0)
EndFunc   ;==>RegExists

; Return a specific line of a multi line string
; http://www.autoitscript.com/forum/topic/103821-how-to-read-specific-line-from-a-string/page__view__findpost__p__735189
Func _StringGetLine($sString, $iLine, $bCountBlank = False)
	Local $sChar = "+"
	If $bCountBlank = True Then $sChar = "*"
	If Not IsInt($iLine) Then Return SetError(1, 0, "")
	If $iLine < 0 Then Return StringTrimLeft($sString, StringInStr($sString, @CRLF, 0, -2 + $iLine))
	Return StringRegExpReplace($sString, "((." & $sChar & "\n){" & $iLine - 1 & "})(." & $sChar & "\n)((." & $sChar & "\n?)+)", "\2")
EndFunc   ;==>_StringGetLine

; Return file metadata
; (Source: http://www.autoitscript.com/forum/topic/40684-querying-a-files-metadata/)
;===============================================================================
; Function Name:    GetExtProperty($sPath,$iProp)
; Description:      Returns an extended property of a given file.
; Parameter(s):     $sPath - The path to the file you are attempting to retrieve an extended property from.
;                   $iProp - The numerical value for the property you want returned. If $iProp is is set
;                             to -1 then all properties will be returned in a 1 dimensional array in their corresponding order.
;                           The properties are as follows:
;                           Name = 0
;                           Size = 1
;                           Type = 2
;                           DateModified = 3
;                           DateCreated = 4
;                           DateAccessed = 5
;                           Attributes = 6
;                           Status = 7
;                           Owner = 8
;                           Author = 9
;                           Title = 10
;                           Subject = 11
;                           Category = 12
;                           Pages = 13
;                           Comments = 14
;                           Copyright = 15
;                           Artist = 16
;                           AlbumTitle = 17
;                           Year = 18
;                           TrackNumber = 19
;                           Genre = 20
;                           Duration = 21
;                           BitRate = 22
;                           Protected = 23
;                           CameraModel = 24
;                           DatePictureTaken = 25
;                           Dimensions = 26
;                           Width = 27
;                           Height = 28
;                           Company = 30
;                           Description = 31
;                           FileVersion = 32
;                           ProductName = 33
;                           ProductVersion = 34
; Requirement(s):   File specified in $sPath must exist.
; Return Value(s):  On Success - The extended file property, or if $iProp = -1 then an array with all properties
;                   On Failure - 0, @Error - 1 (If file does not exist)
; Author(s):        Simucal (Simucal@gmail.com)
; Note(s):
;
;===============================================================================
Func _GetExtProperty($sPath, $iProp = -1)
    Local $sFile, $sDir, $oShellApp, $oDir, $oFile
    If Not FileExists($sPath) Then Return SetError(1, 0, 0)

    $sFile = StringTrimLeft($sPath, StringInStr($sPath, "\", 0, -1))
    $sDir = StringTrimRight($sPath, (StringLen($sPath) - StringInStr($sPath, "\", 0, -1)))
    $oShellApp = ObjCreate("shell.application")
    $oDir = $oShellApp.NameSpace ($sDir)
    $oFile = $oDir.Parsename ($sFile)
    If $iProp = -1 Then
        Local $aProperty[35]
        For $i = 0 To 34
            $aProperty[$i] = $oDir.GetDetailsOf ($oFile, $i)
        Next
		Return $aProperty
	Else
        Local $sProperty = $oDir.GetDetailsOf ($oFile, $iProp)
        If $sProperty = "" Then Return 0
		Return $sProperty
    EndIf
EndFunc

; Dump complete debug content to log file
Func CreateLog($status)
	Local $name = $logdir & @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & "-" & @MIN & "-" & @SEC & "_"
	If $status <> $STATUS_SUCCESS Then $name &= StringUpper($status)
	If $file <> "" Then $name &= "_" & ($iUnicodeMode? $sUnicodeName: $filename) & "." & $fileext
	$name &= ".log"
	$handle = FileOpen($name, 32 + 8 + 2)
	FileWrite($handle, $debug)
	FileClose($handle)
	Return $name
EndFunc

; Determine whether the archive is password protected or not and try passwords from list if necessary
Func _FindArchivePassword($sIsProtectedCmd, $sTestCmd, $sIsProtectedText = "encrypted", $sIsProtectedText2 = 0, $iLine = -3, $sTestText = "All OK")
	; Is archive encrypted?
	Local $return = FetchStdout(_MakeCommand($sIsProtectedCmd, True), $outdir, @SW_HIDE, $iLine)
	If Not StringInStr($return, $sIsProtectedText) And ($sIsProtectedText2 == 0 Or Not StringInStr($return, $sIsProtectedText2)) Then Return 0

	Cout("Archive is password protected")
	GUICtrlSetData($idTrayStatusExt, t('SEARCHING_PASSWORD'))
	$aPasswords = FileReadToArray($sPasswordFile)
	If @error Then
		Cout("Error reading password file " & $sPasswordFile)
		$aPasswords = FileReadToArray(@ScriptDir & "\passwords.txt")
		If @error Then Return 0
	EndIf

	; Try passwords from list
	Local $size = @extended, $sPassword = 0
	If $size > 0 Then Cout("Trying " & $size & " passwords from password list")
	$sTestCmd = _MakeCommand($sTestCmd, True)
	For $i = 0 To $size - 1
		If StringInStr(FetchStdout(StringReplace($sTestCmd, "%PASSWORD%", $aPasswords[$i], 1), $outdir, @SW_HIDE, 0, False), $sTestText) Then
			Cout("Password found")
			$sPassword = $aPasswords[$i]
			ExitLoop
		EndIf
	Next

	GUICtrlSetData($idTrayStatusExt, "")
	Return $sPassword
EndFunc

; Executes a program and log output using tee
Func _Run($f, $sWorkingdir = $outdir, $show_flag = @SW_MINIMIZE, $useCmd = True, $useTee = True, $patternSearch = True, $initialShow = True)
	Global $run = 0, $runtitle = 0
	Local $return = "", $pos = 0, $size = 1, $lastSize = 0
	Local Const $LogFile = $logdir & "teelog.txt"

	$f = _MakeCommand($f, $useCmd) & ($useTee? ' 2>&1 | ' & $tee & ' "' & $LogFile & '"': "")

	Cout("Executing: " & $f & " with options: patternSearch = " & $patternSearch & ", workingdir = " & $sWorkingdir)

	; Create log
	If $useTee Then
		HasPlugin($tee)
		If Not FileExists($logdir) Then DirCreate($logdir)

		$run = Run($f, $sWorkingdir, $initialShow? @SW_MINIMIZE: $show_flag)
		If @error Then
			$success = $RESULT_FAILED
			Return SetError(1)
		EndIf

		Local $TimerStart = TimerInit()
		Cout("Pid: " & $run)
		Do
			Sleep(1)
			If TimerDiff($TimerStart) > 5000 Then ExitLoop
		Until ProcessExists($run)

		$runtitle = _WinGetByPID($run)
		If $initialShow Then WinSetState($runtitle, "", $show_flag)
		Cout("Runtitle: " & $runtitle)

		; Wait until logfile exists
		$TimerStart = TimerInit()

		Do
			Sleep(10)
			If TimerDiff($TimerStart) > 5000 Then ExitLoop
		Until FileExists($LogFile)
		$handle = FileOpen($LogFile)
		$state = ""

		; Show progress (percentage) in status box
		While ProcessExists($run)
			$return = FileRead($handle)
;~ 			Cout($return)
			If $return <> $state Then
				$state = $return
				; Automatically show cmd window when user input needed
				If StringInStr($return, "already exist", 0) Or StringInStr($return, "overwrite", 0) Or StringInStr($return, " replace", 0) Or StringInStr($return, $STATUS_PASSWORD, 0) Then
					Cout("User input needed")
					WinSetState($runtitle, "", @SW_SHOW)
					GUICtrlSetFont($idTrayStatusExt, 8.5, 900)
					GUICtrlSetData($idTrayStatusExt, t('INPUT_NEEDED'))
					WinActivate($runtitle)
					$lastSize = Round((_DirGetSize($outdir, 0) - $initdirsize) / 1024 / 1024, 3)
					ContinueLoop
				EndIf
				; Percentage indicator
				If $patternSearch And _PatternSearch($return) Then $size = -1
			EndIf
			; Size of extracted file(s) as fallback
			If $size > -1 And $patternSearch > -1 Then
				$size = Round((_DirGetSize($outdir) - $initdirsize) / 1024 / 1024, 3)
;~ 				Cout("Size: " & $size & @TAB & $lastSize)
				If $size > 0 And $size <> $lastSize Then
					Cout("Size: " & $size & @TAB & $lastSize)
					GUICtrlSetData($idTrayStatusExt, $size & " MB")
				EndIf
				$lastSize = $size
				Sleep(50)
			EndIf
			Sleep(100)
		WEnd
		; Write tee log to UniExtract log file
		FileSetPos($handle, 0, $FILE_BEGIN)
		$return = FileRead($handle)
		If Not StringIsSpace($return) Then Cout("Teelog:" & @CRLF & $return)
		FileClose($handle)
		FileDelete($LogFile)

		; Check for success or failure indicator in log
		If StringInStr($return, "Wrong password?", 0) Or StringInStr($return, "The specified password is incorrect.", 0) _
		   Or StringInStr($return, "Archive encrypted.", 0) Then
			$success = $RESULT_FAILED
			SetError(1, 1)
		ElseIf StringInStr($return, "Break signaled") Then
			Cout("Cancelled by user")
			$success = $RESULT_CANCELED
		ElseIf StringInStr($return, "Everything is Ok") _
			   Or StringInStr($return, "0 failed") Or StringInStr($return, "All files OK") _
			   Or StringInStr($return, "All OK") Or StringInStr($return, "done.") _
			   Or StringInStr($return, "Done ...") Or StringInStr($return, ": done") _
			   Or StringInStr($return, "Result:	Successful, errorcode 0") _
			   Or StringInStr($return, "Extract files [ ") Or StringInStr($return, "Done; file is OK") Then
			Cout("Success evaluation passed")
			$success = $RESULT_SUCCESS
		ElseIf StringInStr($return, "err code(", 1) Or StringInStr($return, "stacktrace", 1) _
			   Or StringInStr($return, "Write error: ", 1) Or (StringInStr($return, "Cannot create", 1) _
			   And StringInStr($return, "No files to extract", 1)) Or StringInStr($return, "Archives with Errors: 1") _
			   Or StringInStr($return, "ERROR: Wrong tag in package", 1) Or StringInStr($return, "unzip:  cannot find", 1) _
			   Or StringInStr($return, "Open ERROR: Can not open the file as") Then
			$success = $RESULT_FAILED
			SetError(1)
		ElseIf StringInStr($return, "already exists.") Or StringInStr($return, "Overwrite") Then
			Cout("At least one output file already existed")
			; Folder size will most likely stay the same if files are overwritten,
			; so let's disable the check to avoid 'failed' message
			$success = $RESULT_SUCCESS
		EndIf

	; Do not create log
	Else
		Cout("Runtime logging disabled")
		$run = Run($f, $sWorkingdir, $show_flag)
		If @error Then
			$success = $RESULT_FAILED
			Return SetError(1)
		EndIf

		Do
			Sleep(10)
		Until ProcessExists($run)

		$runtitle = _WinGetByPID($run)
		WinSetState($runtitle, "", @SW_HIDE)
		$TimerStart = TimerInit()

		; Size of extracted file(s)
		While ProcessExists($run)
			$size = Round((DirGetSize($outdir) - $initdirsize) / 1024 / 1024, 3)
			If $size > 0 And $patternSearch > -1 Then
				If $TBgui Then GUICtrlSetData($idTrayStatusExt, $size & " MB")
			Else
				If $TimerStart And TimerDiff($TimerStart) > 60000 Then
					WinSetState($runtitle, "", @SW_SHOW)
					WinActivate($runtitle)
					Sleep(5000)
					$TimerStart = 0
				EndIf
			EndIf
			Sleep(100)
		WEnd
	EndIf

	; Reset run var so no wrong process is closed on tray exit
	$run = 0
	Return $return
EndFunc

; Search console output for progress indicator patterns and update status box
Func _PatternSearch($sString)
	If Not $TBgui Then Return False
;~ 	Cout($sString & _StringRepeat(@CRLF, 5))

;~ 	Cout("x %")
	If StringInStr($sString, "%", 0, -1) Then
		$aReturn = StringRegExp($sString, "(\d{1,3})[\d\.,]* ?%", 3)
;~ 		_ArrayDisplay($aReturn)
		If UBound($aReturn) > 0 Then Return GUICtrlSetData($idTrayStatusExt, _ArrayPop($aReturn) & "%")
	EndIf

;~ 	Cout("[x on y]")
	If StringInStr($sString, " on ", 0, -1) Then
		$aReturn = StringRegExp($sString, "\[(\d+) on (\d+)\]", 3)
		If UBound($aReturn) > 1 Then
			$Num = _ArrayPop($aReturn)
			Return GUICtrlSetData($idTrayStatusExt, t('TERM_FILE') & " " & _ArrayPop($aReturn) & "/" & $Num)
		EndIf
	EndIf

;~ 	Cout("x of y")
	If StringInStr($sString, " of ", 0, -1) Then
		$aReturn = StringRegExp($sString, "(\d+) of (\d+)", 3)
		If UBound($aReturn) > 1 Then
			$Num = _ArrayPop($aReturn)
			Return GUICtrlSetData($idTrayStatusExt, t('TERM_FILE') & " " & _ArrayPop($aReturn) & "/" & $Num)
		EndIf
	EndIf

;~ 	Cout("x/y")
	If StringInStr($sString, "/", 0, -1) Then
		$aReturn = StringRegExp($sString, "(\d+)/(\d+)", 3)
		If UBound($aReturn) > 1 Then
			$Num = _ArrayPop($aReturn)
			Return GUICtrlSetData($idTrayStatusExt, t('TERM_FILE') & " " & _ArrayPop($aReturn) & "/" & $Num)
		EndIf
	EndIf

;~ 	Cout("# x")
	$pos = StringInStr($sString, "#", 0, -1)
	If $pos Then
		$Num = Number(StringMid($sString, $pos + 1), 1)
		If $Num > 0 Then Return GUICtrlSetData($idTrayStatusExt, t('TERM_FILE') & " #" & $Num)
	EndIf
EndFunc

; Run a program and return stdout/stderr stream
Func FetchStdout($f, $sWorkingdir, $show_flag, $iLine = 0, $bOutput = True, $useCmd = True)
	Global $run = 0, $return = ""

	$f = _MakeCommand($f, $useCmd)
	If $bOutput Then Cout("Executing: " & $f)
	$run = Run($f, $sWorkingdir, $show_flag, $STDERR_MERGED)
	$runtitle = _WinGetByPID($run)
	$TimerStart = TimerInit()

	Do
		Sleep(1)
		If TimerDiff($TimerStart) > $Timeout Then ExitLoop
		$return &= StdoutRead($run)
	Until @error

	If $bOutput Then Cout($return)
	If ProcessExists($run) Then ProcessClose($run)
	$run = 0

	If $iLine <> 0 Then Return _StringGetLine($return, $iLine)
	Return $return
EndFunc

; Build final command line from parameters
Func _MakeCommand($f, $useCmd = False)
	If StringInStr($f, $cmd) Then Return $f

	If Not StringInStr($f, $bindir) Then
		$pos = StringInStr($f, " ")
		If $pos > 1 And $useCmd And FileExists($bindir & StringLeft($f, $pos)) Then
			$f = '""' & $bindir & _StringInsert($f, '"', $pos - 1)
		Else
			$f = $bindir & $f
		EndIf
	EndIf
	Return ($useCmd? $cmd: "") & $f
EndFunc

; DirGetSize wrapper with additional logic
Func _DirGetSize($f, $return = -1)
	; Calculating the size of a whole drive would take way too much time,
	; so let's only calculate size if less than 4 GB space used on drive
	If (StringLen($f) < 4 And DriveSpaceTotal($f) - DriveSpaceFree($f) > 4000) Then Return $return
	Return DirGetSize($f)
EndFunc

; Stop running helper process
Func KillHelper()
	If Not $run Then Return
	Cout("Killing helper process " & $run)
	StdioClose($run)

	If Not @error And Not StringIsSpace($runtitle) Then
		Cout("Runtitle: " & $runtitle)
		; Send SIGINT to console to terminate child processes
		WinActivate($runtitle)
		If WinActive($runtitle) Then Send("^c")
		; Close console
		WinClose($runtitle)
	EndIf

	; Force termination if other close commands failed
	If ProcessExists($run) Then ProcessClose($run)
EndFunc

; Restart Universal Extractor without elevated privileges
Func RestartWithoutAdminRights($sParameters = "")
	Run($cmd & 'runas /trustlevel:0x20000 "' & @ScriptFullPath & $sParameters & '"')
	terminate("silent")
EndFunc

; Write data to stdout stream if enabled in options
Func Cout($Data)
	Local $Output = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & ":" & @MSEC & @TAB & $Data & @CRLF; & @CRLF
	If $Opt_ConsoleOutput == 1 Then ConsoleWrite($Output)
	$debug &= $Output
	Return $Data
EndFunc

; Open URL and evaluate success
Func OpenURL($sURL, $hParent = 0)
	ShellExecute($sURL)
	If @error Then InputBox($title, t('OPEN_URL_FAILED'), $sURL, "", -1, -1, Default, Default, 0, $hParent)
EndFunc

; Send usage statistics if enabled in options
Func SendStats($a, $sResult = 1)
	If Not $bSendStats Then Return

	Cout($sStatsURL & $a & "&r=" & $sResult & "&id=" & $ID)
	InetRead($sStatsURL & $a & "&r=" & $sResult & "&id=" & $ID, 1)
EndFunc

; Check for new version
; $silent is used for automatic update check, supressing any error and 'no update found' messages
Func CheckUpdate($silent = $UPDATEMSG_PROMPT, $bCheckInterval = False, $iMode = $UPDATE_ALL)
	If @NumParams > 1 And $bCheckInterval And _DateDiff("D", $lastupdate, _NowCalc()) < $updateinterval Then Return
	If @NumParams < 1 Then
		$silent = $UPDATEMSG_PROMPT
		$iMode = $UPDATE_ALL
	EndIf
	$ret2 = $silentmode
	If $silent == $UPDATEMSG_SILENT Then $silentmode = 1

	Local $return = 0, $found = False
	Cout("Checking for update")

	; Save date of last check for update
	$lastupdate = @YEAR & "/" & @MON & "/" & @MDAY
	SavePref('lastupdate', $lastupdate)

	; Get index
	$aReturn = _UpdateGetIndex()
	If Not IsArray($aReturn) Then Return Cout("Failed to get update file listing")

	; UniExtract main executable - calling the updater is always necessary, because an executable file cannot overwrite itself while running
	If $iMode <> $UPDATE_HELPER Then
		If ($aReturn[0])[1] <> FileGetSize(@Compiled? @ScriptFullPath: StringReplace(@ScriptFullPath, "au3", "exe")) Then
			Cout("Update available")
			$found = True
			; $Because FFMPEG uses the same update message, we cannot use the %name constant here
			If Prompt(48 + 4, 'UPDATE_PROMPT', CreateArray($name, $version, $return, $aReturn[0] > 2? $aReturn[2]: ""), 0) Then
				If Not ShellExecute(CanAccess(@ScriptDir)? $sUpdaterNoAdmin: $sUpdaterNoAdmin, "/main") Then MsgBox($iTopmost + 16, $title, t('UPDATE_FAILED'))
				Exit
			Else
				; If the user does not want to install the main update, let's not bother him with more 'update found' messages
;~ 				$iMode = $UPDATE_MAIN
			EndIf
		EndIf
	EndIf

	; Other files - we can overwrite the files without a seperate updater
	If $iMode <> $UPDATE_MAIN Then
		If CheckUpdateHelpers($aReturn) Then
			$found = True
			If Prompt(48 + 4, 'UPDATE_PROMPT', t('UPDATE_TERM_PROGRAM_FILES'), 0) Then
				If Not CanAccess($bindir) Then Exit ShellExecute($sUpdater, "/helper")
				If Not _UpdateHelpers($aReturn) And Not $ret2 Then MsgBox($iTopmost + 16, $title, t('UPDATE_FAILED'))
			EndIf
		EndIf
		If _UpdateFFmpeg() Then $found = True
	EndIf

	If $found = False Then
		SendStats("CheckUpdate", 0)
		If $silent == $UPDATEMSG_PROMPT Then MsgBox($iTopmost + 64, $name, t('UPDATE_CURRENT'))
	EndIf
	Cout("Check for updates finished")

	If $silent == $UPDATEMSG_SILENT Then $silentmode = $ret2
	If IsAdmin() Then RestartWithoutAdminRights()
EndFunc

; Compare program files with server index to find if any file has an updated version available
Func CheckUpdateHelpers($aFiles)
	Local $i = 1, $iSize = UBound($aFiles)

	While $i < $iSize
		$a = $aFiles[$i]
		$i += 1
		$sPath = @ScriptDir & "\" & $a[0]
		If $sPath == @ScriptFullPath Then ContinueLoop

;~ 		Cout($sPath)
		$size = _UpdateGetSize($sPath)
		If $size == $a[1] Then ContinueLoop
		Cout($a[0] & ": " & $size & " - " & $a[1])

		; If it's a file and the size differs, update necessary
		If StringRight($a[0], 1) <> "/" Then Return True

		; Directory
		If Not FileExists($sPath) Then Return True

		$aReturn = _UpdateGetIndex($a[0])
		If Not IsArray($aReturn) Then ContinueLoop

		_ArrayAdd($aFiles, $aReturn)
		$iSize = UBound($aFiles)
	WEnd

	Return False
EndFunc

; Download updated program files and display status
Func _UpdateHelpers($aFiles)
	Local $sText = t('TERM_DOWNLOADING') & "... "

	Local $hGUI = GUICreate($title, 434, 130, -1, -1, $WS_POPUPWINDOW, -1, $guimain)
	Local $idLabel = GUICtrlCreateLabel($sText, 16, 16, 408, 17)
	GUICtrlCreateLabel(t('TERM_OVERALL_PROGRESS') & ":", 16, 72, 80, 17)
	Local $idProgressCurrent = GUICtrlCreateProgress(16, 32, 408, 25)
	Local $idProgressTotal = GUICtrlCreateProgress(16, 88, 406, 25)
	_GuiSetColor()
	GUISetState(@SW_SHOW)

	Local $i = 0, $iSize = UBound($aFiles), $iProgress = 0, $success = True, $iBytesReceived

	While $i < $iSize
		; Update progress
		$ret = (($i + 1) / $iSize) * 100
		$iProgress = $ret > $iProgress? $ret: $iProgress + 0.2
		GUICtrlSetData($idProgressTotal, $iProgress)

		$a = $aFiles[$i]
		$i += 1
		$sPath = @ScriptDir & "\" & $a[0]
		If $sPath == @ScriptFullPath Then ContinueLoop

;~ 		Cout($sPath)
		$size = _UpdateGetSize($sPath)
		If $size == $a[1] Then ContinueLoop
		Cout($a[0] & ": " & $size & " - " & $a[1])

		GUICtrlSetData($idProgressCurrent, 0)
		If StringRight($a[0], 1) = "/" Then ; Directory
			If Not FileExists($sPath) Then DirCreate($sPath)
			GUICtrlSetData($idLabel, t('UPDATE_STATUS_SEARCHING'))
			$aReturn = _UpdateGetIndex($a[0])
			If Not IsArray($aReturn) Then
				$success = False
				ContinueLoop
			EndIf
			_ArrayAdd($aFiles, $aReturn)
			$iSize = UBound($aFiles)
		Else
			GUICtrlSetData($idLabel, $sText & $a[0] & " (" & $a[1] & " bytes" & ")")

			$iBytesReceived = 0
			Local $hDownload = InetGet($sUpdateURL & $a[0], @ScriptDir & "\" & $a[0], 1, 1)

			; Update progress bar
			While Not InetGetInfo($hDownload, 2)
				Sleep(50)
				If InetGetInfo($hDownload, 4) <> 0 Then
					Cout("Download failed")
					$success = False
					ContinueLoop 2
				EndIf
				$iBytesReceived = InetGetInfo($hDownload, 0)
				GUICtrlSetData($idProgressCurrent, Int($iBytesReceived / $a[1] * 100))
			WEnd
			GUICtrlSetData($idProgressCurrent, 100)
		EndIf
	WEnd

	SendStats("UpdateHelpers", 1)
	GUIDelete($hGUI)
	Return $success
EndFunc

; Search for FFmpeg updates
Func _UpdateFFmpeg()
	If Not HasPlugin($ffmpeg, True) Then Return False

	; Determine FFmpeg version
	Local $return = FetchStdout($ffmpeg, @ScriptDir, @SW_HIDE, 0, False)
	Local $version = _StringBetween($return, "ffmpeg version ", " Copyright")
	If @error Then Return False

	$return = _INetGetSource(_IsWinXP()? $sLegacyUpdateURL & "?get=ffversion&xp": $sUpdateURL & "ffmpeg")
	Cout("FFmpeg: " & $version[0] & " <--> " & $return)

	; Download new
	If $return > $version[0] Then
		Cout("Found an update for FFmpeg")
		If Prompt(48 + 4, 'UPDATE_PROMPT', CreateArray("FFmpeg", $version[0], $return, ""), 0) Then Return GetFFmpeg()
	EndIf

	Return False
EndFunc

; Helper function for updater, downloads index file for subdirectories
Func _UpdateGetIndex($sURL = "")
	$sURL = $sUpdateURL & $sURL & "index"
;~ 	Cout("Sending request: " & $sURL)

	$return = _INetGetSource($sURL)
	If @error Then Return _UpdateCheckFailed(True)

	$aReturn = StringSplit($return, @LF, 2)
;~ 	_ArrayDisplay($aReturn)

	For $i = 0 To UBound($aReturn) - 1
		$aReturn[$i] = StringSplit($aReturn[$i], ",", 2)
		If @error Then Return _UpdateCheckFailed(True)
	Next

	Return $aReturn
EndFunc

; Return size of a file or directory, ignoring plugins, which are not on the server
Func _UpdateGetSize($sPath)
	If Not StringInStr(FileGetAttrib($sPath), "D") Then Return FileGetSize($sPath)
	$sPath = StringReplace($sPath, "/", "\")
;~ 	Cout("GetSize: " & $sPath)
	Local $iSize = DirGetSize($sPath)

	; Don't include plugins in calculations
	If $sPath = @ScriptDir & "\docs\" Then
		$iSize -= DirGetSize($sPath & "FFmpeg") ; FFmpeg does not exist on server
	ElseIf $sPath = $bindir & "x86\" Or $sPath = $bindir & "x64\" Then
		$iSize -= FileGetSize($sPath & "ffmpeg.exe")
	ElseIf $sPath = $bindir Then
		$ret = DirGetSize($bindir & "crass-0.4.14.0")
		If $ret > 0 Then $iSize -= $ret

		Local $aReturn[] = ["x86\ffmpeg.exe", "x64\ffmpeg.exe", "arc_conv.exe", "Extractor.exe", "iscab.exe", "ISTools.dll", "RPGDecrypter.exe", _
							"umodel.exe", "SDL2.dll", "dcp_unpacker.exe", "mpq.wcx", "mpq.wcx64", "ci-extractor.exe", "gea.dll", "gentee.dll", _
							"dgcac.exe", "bootimg.exe", "I5comp.exe", "ZD50149.DLL", "ZD51145.DLL", "sim_unpacker.exe"]

		For $i In $aReturn
			$iSize -= FileGetSize($bindir & $i)
		Next
	EndIf
	Return $iSize
EndFunc

; Display update failed message
Func _UpdateCheckFailed()
	If Not $silentmode Then MsgBox($iTopmost + 48, $title, t('UPDATECHECK_FAILED'))
	Return False
EndFunc

; Perform special actions after update, e.g. delete files
Func _AfterUpdate()
	; Remove unused files
	FileDelete($bindir & "languages\ChineseBig5_v0038.lng")
	FileDelete($bindir & "languages\exeinfope_Neutral_v0038.lng")
	FileDelete($bindir & "languages\exeinfope_Neutral_v0041.lng")
	FileDelete($bindir & "languages\exeinfope_turkish.lng")
	FileDelete($bindir & "languages\exeinfopeCHS.lng")
	FileDelete($bindir & "faad.exe")
	FileDelete($bindir & "x86\flac.exe")
	FileDelete($bindir & "x64\flac.exe")
	FileDelete($langdir & "Chinese.ini")
	FileDelete($bindir & "MediaInfo64.dll")
	FileDelete($bindir & "extract.exe")
	FileDelete($bindir & "dmgextractor.jar")
	DirRemove($bindir & "unrpa", 1)
	DirRemove($bindir & "file\contrib\file\5.03\file-5.03", 1)

	; Move files
	FileMove(@ScriptDir & "\UniExtractUpdater.exe.new", @ScriptDir & "\UniExtractUpdater.exe", 1)
	FileMove($bindir & "x86\sqlite3.dll", @ScriptDir)
	FileMove($bindir & "x64\sqlite3.dll", @ScriptDir & "\sqlite3_x64.dll")

	SendStats("UpdateMain", 1)

	; Update helpers
	CheckUpdate(True, False, $UPDATE_HELPER)

	RestartWithoutAdminRights()
EndFunc

; Start updater to download FFmpeg
Func GetFFmpeg()
	; As FFmpeg can be downloaded from the first start assistant, we use the updater to handle elevation and download.
	; Otherwise, it would be neccessary to exit the assistant and restart UniExtract with higher permissions, which would just look bad and scare the user :(
	$ret = ShellExecuteWait(CanAccess($bindir)? $sUpdaterNoAdmin: $sUpdater, "/ffmpeg")
	If @error Or Not HasPlugin($ffmpeg, True) Then Return SetError(1, 0, 0)

	Cout("FFmpeg successfully downloaded")
	If $FS_GUI Then GUI_FirstStart_Next()
	Return 1
EndFunc


; ------------------------ Begin GUI Control Functions ------------------------

; Build and display GUI if necessary
Func CreateGUI()
	Cout("Creating main GUI")
	GUIRegisterMsg($WM_DROPFILES, "WM_DROPFILES_UNICODE_FUNC")
	GUIRegisterMsg($WM_GETMINMAXINFO, "GUI_WM_GETMINMAXINFO_Main")

	Switch $language
		Case "Arabic", "Farsi", "Hebrew"
			$exStyle = $WS_EX_LAYOUTRTL
		Case Else
			$exStyle = -1
	EndSwitch

	; Create GUI
	If $StoreGUIPosition Then
		Global $guimain = GUICreate($title, 310, 160, $posx, $posy, BitOR($WS_SIZEBOX, $WS_MINIMIZEBOX), BitOR($WS_EX_ACCEPTFILES, $iTopmost, $exStyle < 0? 0: $exStyle))
	Else
		Global $guimain = GUICreate($title, 310, 160, -1, -1, BitOR($WS_SIZEBOX, $WS_MINIMIZEBOX), BitOR($WS_EX_ACCEPTFILES, $iTopmost, $exStyle < 0? 0: $exStyle))
	EndIf

	_GuiSetColor()
	Local $dropzone = GUICtrlCreateLabel("", 0, 0, 300, 135)

	; Menu controls
	Local $filemenu = GUICtrlCreateMenu(t('MENU_FILE_LABEL'))
	Local $openitem = GUICtrlCreateMenuItem(t('MENU_FILE_OPEN_LABEL'), $filemenu)
	GUICtrlCreateMenuItem("", $filemenu)
	Global $keepopenitem = GUICtrlCreateMenuItem(t('MENU_FILE_KEEP_OPEN_LABEL'), $filemenu)
	Global $topmostitem = GUICtrlCreateMenuItem(t('PREFS_TOPMOST_LABEL'), $filemenu)
	GUICtrlCreateMenuItem("", $filemenu)
	Global $showitem = GUICtrlCreateMenuItem(t('MENU_FILE_SHOW_LABEL'), $filemenu)
	Global $clearitem = GUICtrlCreateMenuItem(t('MENU_FILE_CLEAR_LABEL'), $filemenu)
	GUICtrlCreateMenuItem("", $filemenu)
	Local $logopenitem = GUICtrlCreateMenuItem(t('MENU_FILE_LOG_OPEN_LABEL'), $filemenu)
	Global $logitem = GUICtrlCreateMenuItem("DUMMY", $filemenu)
	GUICtrlCreateMenuItem("", $filemenu)
	Local $quititem = GUICtrlCreateMenuItem(t('MENU_FILE_QUIT_LABEL'), $filemenu)
	Local $editmenu = GUICtrlCreateMenu(t('MENU_EDIT_LABEL'))
	Global $silentitem = GUICtrlCreateMenuItem(t('MENU_EDIT_SILENT_MODE_LABEL'), $editmenu)
	GUICtrlCreateMenuItem("", $editmenu)
	Local $passworditem = GUICtrlCreateMenuItem(t('MENU_EDIT_PASSWORD_LABEL'), $editmenu)
	GUICtrlCreateMenuItem("", $editmenu)
	Local $contextitem = GUICtrlCreateMenuItem(t('MENU_EDIT_CONTEXT_LABEL'), $editmenu)
	Local $prefsitem = GUICtrlCreateMenuItem(t('MENU_EDIT_PREFS_LABEL'), $editmenu)
	Local $helpmenu = GUICtrlCreateMenu(t('MENU_HELP_LABEL'))
	Local $updateitem = GUICtrlCreateMenuItem(t('MENU_HELP_UPDATE_LABEL'), $helpmenu)
	GUICtrlCreateMenuItem("", $helpmenu)
	Local $feedbackitem = GUICtrlCreateMenuItem(t('MENU_HELP_FEEDBACK_LABEL'), $helpmenu)
	Local $pluginsitem = GUICtrlCreateMenuItem(t('MENU_HELP_PLUGINS_LABEL'), $helpmenu)
	Local $firststartitem = GUICtrlCreateMenuItem(t('FIRSTSTART_TITLE'), $helpmenu)
	GUICtrlCreateMenuItem("", $helpmenu)
	Local $webitem = GUICtrlCreateMenuItem(t('MENU_HELP_WEB_LABEL', $name), $helpmenu)
	Local $web2item = GUICtrlCreateMenuItem(t('MENU_HELP_WEB_LABEL', $name & " 2"), $helpmenu)
	Local $gititem = GUICtrlCreateMenuItem(t('MENU_HELP_GITHUB_LABEL'), $helpmenu)
	GUICtrlCreateMenuItem("", $helpmenu)
	Local $statsitem = GUICtrlCreateMenuItem(t('MENU_HELP_STATS_LABEL'), $helpmenu)
	Local $programdiritem = GUICtrlCreateMenuItem(t('MENU_HELP_PROGDIR_LABEL'), $helpmenu)
	Local $configfileitem = GUICtrlCreateMenuItem(t('MENU_HELP_CONFIGFILE_LABEL'), $helpmenu)
	GUICtrlCreateMenuItem("", $helpmenu)
	Local $aboutitem = GUICtrlCreateMenuItem(t('MENU_HELP_ABOUT_LABEL'), $helpmenu)
	GUI_UpdateLogItem()

	; File controls
	Local $filelabel = GUICtrlCreateLabel(t('MAIN_FILE_LABEL'), 5, 4, $exStyle == $WS_EX_LAYOUTRTL? 50: -1, 15)
	Global $GUI_Main_Extract = GUICtrlCreateRadio(t('TERM_EXTRACT'), GetPos($guimain, $filelabel, 5), 3, Default, 15)
	Global $GUI_Main_Scan = GUICtrlCreateRadio(t('TERM_SCAN'), GetPos($guimain, $GUI_Main_Extract, 10), 3, 100, 15)
	GUICtrlSetState($extract? $GUI_Main_Extract: $GUI_Main_Scan, $GUI_CHECKED)

	Global $filecont = $history? GUICtrlCreateCombo("", 5, 20, 260, 20): GUICtrlCreateInput("", 5, 20, 260, 20)
	Local $filebut = GUICtrlCreateButton("...", 270, 20, 25, 20)

	; Directory controls
	Local $GUI_Main_Destination_Label = GUICtrlCreateLabel(t('MAIN_DEST_DIR_LABEL'), 5, 45, $exStyle == $WS_EX_LAYOUTRTL? 50: -1, 15)
	Global $dircont = $history? GUICtrlCreateCombo("", 5, 60, 260, 20): GUICtrlCreateInput("", 5, 60, 260, 20)
	Local $dirbut = GUICtrlCreateButton("...", 270, 60, 25, 20)
	Global $GUI_Main_Lock = GUICtrlCreateCheckbox(t('MAIN_DIRECTORY_LOCK'), GetPos($guimain, $GUI_Main_Destination_Label, 5), 44, Default, 15)
	GUICtrlSetTip($GUI_Main_Lock, t('MAIN_DIRECTORY_LOCK_TOOLTIP'))

	; Buttons
	Global $ok = GUICtrlCreateButton(t('OK_BUT'), 10, 90, 80, 20)
	Local $cancel = GUICtrlCreateButton(t('CANCEL_BUT'), 110, 90, 80, 20)
	Global $BatchBut = GUICtrlCreateButton(t('BATCH_BUT'), 210, 90, 80, 20)

	; Set properties
	GUICtrlSetBkColor($dropzone, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetState($dropzone, $GUI_DISABLE)
	GUICtrlSetState($dropzone, $GUI_DROPACCEPTED)
	GUICtrlSetState($filecont, $GUI_FOCUS)
	GUICtrlSetState($ok, $GUI_DEFBUTTON)
	GUICtrlSetState($GUI_Main_Lock, $KeepOutdir? $GUI_CHECKED: $GUI_UNCHECKED)
	GUICtrlSetState($keepopenitem, $KeepOpen? $GUI_CHECKED: $GUI_UNCHECKED)
	GUICtrlSetState($topmostitem, $iTopmost? $GUI_CHECKED: $GUI_UNCHECKED)
	GUICtrlSetState($silentitem, $silentmode: $GUI_CHECKED: $GUI_UNCHECKED)

	If $batchEnabled = 0 Then
		GUICtrlSetState($showitem, $GUI_DISABLE)
		GUICtrlSetState($clearitem, $GUI_DISABLE)
	EndIf

	If $file <> "" Then
		FilenameParse($file)
		If $history Then
			$filelist = '|' & $file & '|' & ReadHist($HISTORY_FILE)
			GUICtrlSetData($filecont, $filelist, $file)
			$dirlist = '|' & $initoutdir & '|' & ReadHist($HISTORY_DIR)
			GUICtrlSetData($dircont, $dirlist, $initoutdir)
		Else
			GUICtrlSetData($filecont, $file)
			GUICtrlSetData($dircont, $initoutdir)
		EndIf
		GUICtrlSetState($dircont, $GUI_FOCUS)
	ElseIf $history Then
		GUICtrlSetData($filecont, ReadHist($HISTORY_FILE))
		GUICtrlSetData($dircont, ReadHist($HISTORY_DIR))
	EndIf

	; Set events
	GUISetOnEvent($GUI_EVENT_DROPPED, "GUI_Drop")
	GUICtrlSetOnEvent($filebut, "GUI_File")
	GUICtrlSetOnEvent($dirbut, "GUI_Directory")
	GUICtrlSetOnEvent($openitem, "GUI_File")
	GUICtrlSetOnEvent($keepopenitem, "GUI_KeepOpen")
	GUICtrlSetOnEvent($topmostitem, "GUI_Topmost")
	GUICtrlSetOnEvent($showitem, "GUI_Batch_Show")
	GUICtrlSetOnEvent($clearitem, "GUI_Batch_Clear")
	GUICtrlSetOnEvent($logopenitem, "GUI_OpenLogDir")
	GUICtrlSetOnEvent($logitem, "GUI_DeleteLogs")
	GUICtrlSetOnEvent($GUI_Main_Lock, "GUI_KeepOutdir")
	GUICtrlSetOnEvent($GUI_Main_Extract, "GUI_ScanOnly")
	GUICtrlSetOnEvent($GUI_Main_Scan, "GUI_ScanOnly")
	GUICtrlSetOnEvent($silentitem, "GUI_Silent")
	GUICtrlSetOnEvent($passworditem, "GUI_Password")
	GUICtrlSetOnEvent($contextitem, "GUI_ContextMenu")
	GUICtrlSetOnEvent($prefsitem, "GUI_Prefs")
	GUICtrlSetOnEvent($pluginsitem, "GUI_Plugins")
	GUICtrlSetOnEvent($feedbackitem, "GUI_Feedback")
	GUICtrlSetOnEvent($firststartitem, "GUI_FirstStart")
	GUICtrlSetOnEvent($updateitem, "CheckUpdate")
	GUICtrlSetOnEvent($webitem, "GUI_Website")
	GUICtrlSetOnEvent($web2item, "GUI_Website2")
	GUICtrlSetOnEvent($gititem, "GUI_Website_Github")
	GUICtrlSetOnEvent($statsitem, "GUI_Stats")
	GUICtrlSetOnEvent($programdiritem, "GUI_ProgDir")
	GUICtrlSetOnEvent($configfileitem, "GUI_ConfigFile")
	GUICtrlSetOnEvent($aboutitem, "GUI_About")
	GUICtrlSetOnEvent($ok, "GUI_Ok")
	GUICtrlSetOnEvent($cancel, "GUI_Exit")
	GUICtrlSetOnEvent($BatchBut, "GUI_Batch")
	GUICtrlSetOnEvent($quititem, "GUI_Exit")
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_Exit")

	GetBatchQueue()

	; Display GUI and wait for action
	GUISetState(@SW_SHOW)
EndFunc

; Return control width (for dynamic positioning)
Func GetPos($hGUI, $hControl, $iOffset = 0, $bX = True)
	$aReturn = ControlGetPos($hGUI, '', $hControl)
	If @error Then Return SetError(1, '', $iOffset)

	If $bX Then
		If $exStyle == $WS_EX_LAYOUTRTL Then $iOffset *= 0.4
		Return $aReturn[0] + $aReturn[2] + $iOffset
	EndIf

	Return $aReturn[1] + $aReturn[3] + $iOffset
EndFunc

; Return number of times a character appears in a string
Func CharCount($string, $char)
	Local $return = StringSplit($string, $char, 1)
	Return $return[0]
EndFunc   ;==>CharCount

; Get title of a window by PID as returned by Run()
; Script by SmOke_N (http://www.autoitscript.com/forum/topic/136271-solved-wingethandle-from-wingetprocess/#entry952135)
Func _WinGetByPID($iPID, $iString = 1) ; 0 Will Return 1 Base Array & 1 Will Return The First Window.
	Local $aError[1] = [0], $aWinList, $sReturn
	If IsString($iPID) Then $iPID = ProcessExists($iPID)
	$aWinList = WinList()
	For $A = 1 To $aWinList[0][0]
		If WinGetProcess($aWinList[$A][1]) = $iPID Then ;And BitAND(WinGetState($aWinList[$A][1]), 2) Then
			If $iString Then Return $aWinList[$A][1]
			$sReturn &= $aWinList[$A][1] & Chr(1)
		EndIf
	Next
	If $sReturn Then Return StringSplit(StringTrimRight($sReturn, 1), Chr(1))
	Return SetError(1, 0, $iString? 0: $aError)
EndFunc

; Round corners of status box
; Code by ? (http://www.autoitscript.com/forum/topic/100790-guiroundcorners-help/page__p__719767__hl__round%20corner__fromsearch__1#entry719767)
Func _GuiRoundCorners($h_win, $i_x1, $i_y1, $i_x3, $i_y3)
	Dim $pos, $ret, $ret2
	$pos = WinGetPos($h_win)
	$ret = DllCall("gdi32.dll", "long", "CreateRoundRectRgn", "long", $i_x1, "long", $i_y1, "long", $pos[2], "long", $pos[3], "long", $i_x3, "long", $i_y3)
	If $ret[0] Then
		$ret2 = DllCall("user32.dll", "long", "SetWindowRgn", "hwnd", $h_win, "long", $ret[0], "int", 1)
		If $ret2[0] Then
			Return 1
		Else
			Return 0
		EndIf
	Else
		Return 0
	EndIf
EndFunc   ;==>_GuiRoundCorners

; Set GUI color to white when using Windows 10
Func _GuiSetColor()
	If @OSVersion <> "WIN_10" Then Return
	GUISetBkColor(0xFFFFFF)
	GUICtrlSetDefBkColor(0xFFFFFF)
EndFunc

; Prompt user for file
Func GUI_File()
	$files = StringSplit(FileOpenDialog(t('OPEN_FILE'), "", t('SELECT_FILE') & " (*.*)|" & t('TERM_INSTALLER') & " (*.exe)|" & t('TERM_COMPRESSED') & " (*.rar;*.zip;*.7z)", 4 + 1, "", $guimain), "|", 2)
	If Not $files[0] = "" Then
		;_ArrayDisplay($files)
		$return = UBound($files)
		If $return == 1 Then
			Global $gaDropFiles = $files
		Else
			Global $gaDropFiles[$return]
			For $i = 0 To UBound($files) - 2
				$gaDropFiles[$i] = $files[0] & "\" & $files[$i + 1]
			Next
		EndIf
		;_ArrayDisplay($gaDropFiles)
		GUI_Drop()

		GUICtrlSetState($ok, $GUI_FOCUS)
	EndIf
EndFunc

; Prompt user for directory
Func GUI_Directory()
	Local $dir = GUICtrlRead($dircont)
	If Not FileExists($dir) Then
		$return = GUICtrlRead($filecont)
		$dir = FileExists($return)? StringLeft($return, StringInStr($return, '\', 0, -1) - 1): ""
	EndIf

	$outdir = FileSelectFolder(t('EXTRACT_TO'), "", 3, $dir, $guimain)
	If @error Then Return

	If $history Then
		$dirlist = '|' & $outdir & '|' & ReadHist($HISTORY_DIR)
		GUICtrlSetData($dircont, $dirlist, $outdir)
	Else
		GUICtrlSetData($dircont, $outdir)
	EndIf
EndFunc

; Option to keep the destination directory
Func GUI_KeepOutdir()
	$KeepOutdir = Int(BitAND(GUICtrlRead($GUI_Main_Lock), $GUI_CHECKED) = $GUI_CHECKED)
	SavePref('keepoutputdir', $KeepOutdir)
EndFunc

; Option to scan file without extracting
Func GUI_ScanOnly()
	If BitAND(GUICtrlRead($GUI_Main_Extract), $GUI_CHECKED) = $GUI_CHECKED Then
		GUICtrlSetState($GUI_Main_Extract, $GUI_CHECKED)
		$extract = 1
	Else
		GUICtrlSetState($GUI_Main_Scan, $GUI_CHECKED)
		$extract = 0
	EndIf

	SavePref('extract', $extract)
EndFunc   ;==>GUI_ScanOnly

; Option to scan file without extracting
Func GUI_Silent()
	If BitAND(GUICtrlRead($silentitem), $GUI_CHECKED) = $GUI_CHECKED Then
		GUICtrlSetState($silentitem, $GUI_UNCHECKED)
		$silentmode = 0
	Else
		GUICtrlSetState($silentitem, $GUI_CHECKED)
		$silentmode = 1
	EndIf

	SavePref('silentmode', $silentmode)
EndFunc   ;==>GUI_Silent

; Option to keep Universal Extractor open
Func GUI_KeepOpen()
	If BitAND(GUICtrlRead($keepopenitem), $GUI_CHECKED) = $GUI_CHECKED Then
		GUICtrlSetState($keepopenitem, $GUI_UNCHECKED)
		$KeepOpen = 0
	Else
		GUICtrlSetState($keepopenitem, $GUI_CHECKED)
		$KeepOpen = 1
	EndIf

	SavePref('keepopen', $KeepOpen)
EndFunc

; Option to keep Universal Extractor on top
Func GUI_Topmost()
	If BitAND(GUICtrlRead($topmostitem), $GUI_CHECKED) = $GUI_CHECKED Then
		GUICtrlSetState($topmostitem, $GUI_UNCHECKED)
		$iTopmost = 0
	Else
		GUICtrlSetState($topmostitem, $GUI_CHECKED)
		$iTopmost = $WS_EX_TOPMOST
	EndIf

	GUIDelete($guimain)
	CreateGUI()

	SavePref('topmost', Number($iTopmost > 0))
EndFunc

; Build and display preferences GUI
Func GUI_Prefs()
	Cout("Creating preferences GUI")

	; Load language list
	Local $aReturn = _FileListToArray($langdir, '*.ini', 1)
	If @error Then Local $aReturn[1]
	$aReturn[0] = 'English.ini'
	_ArraySort($aReturn)

	Local $langlist = ""
	For $i = 0 To UBound($aReturn) - 1
		$langlist &= StringTrimRight($aReturn[$i], 4) & '|'
	Next

	; Create GUI
	Global $guiprefs = GUICreate(t('PREFS_TITLE_LABEL'), 250, 470, -1, -1, -1, $exStyle, $guimain)
	_GuiSetColor()
	GUICtrlCreateGroup(t('PREFS_UNIEXTRACT_OPTS_LABEL'), 5, 5, 240, 122)

	; History option
	Global $historyopt = GUICtrlCreateCheckbox(t('PREFS_HISTORY_LABEL'), 10, 20, -1, 20)

	; Language controls
	Local $langlabel = GUICtrlCreateLabel(t('PREFS_LANG_LABEL'), 10, 45, -1, 15)
	Local $pos = GetPos($guiprefs, $langlabel, -8)
	Global $langselect = GUICtrlCreateCombo("", $pos, 42, 245 - $pos - 8, -1, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))

	; Timeout and update interval controls
	Local $TimeoutLabel = GUICtrlCreateLabel(t('PREFS_TIMEOUT_LABEL'), 10, 72, $exStyle == $WS_EX_LAYOUTRTL? 40: -1, 15)
	Local $UpdateIntervalLabel = GUICtrlCreateLabel(t('PREFS_UPDATEINTERVAL_LABEL'), 10, 102, $exStyle == $WS_EX_LAYOUTRTL? 50: -1, 15)
	$pos = _Max(GetPos($guiprefs, $TimeoutLabel, 5), GetPos($guiprefs, $UpdateIntervalLabel, 5))
	Global $TimeoutCont = GUICtrlCreateInput($Timeout / 1000, $pos, 70, 35, 20, $ES_NUMBER)
	Global $IntervalCont = GUICtrlCreateInput($updateinterval, $pos, 100, 35, 20, $ES_NUMBER)
	GUICtrlCreateLabel(t('PREFS_SECONDS_LABEL'), GetPos($guiprefs, $TimeoutCont, 5), 72, -1, 15)
	GUICtrlCreateLabel(t('PREFS_DAYS_LABEL'), GetPos($guiprefs, $IntervalCont, 5), 102, -1, 15)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; Format-specific preferences
	GUICtrlCreateGroup(t('PREFS_FORMAT_OPTS_LABEL'), 5, 130, 240, 260)
	Global $warnexecuteopt = GUICtrlCreateCheckbox(t('PREFS_WARN_EXECUTE_LABEL'), 10, 145, -1, 20)
	Global $freeSpaceCheckOpt = GUICtrlCreateCheckbox(t('PREFS_CHECK_FREE_SPACE_LABEL'), 10, 165, -1, 20)
	Global $unicodecheckopt = GUICtrlCreateCheckbox(t('PREFS_CHECK_UNICODE_LABEL'), 10, 185, -1, 20)
	Global $appendextopt = GUICtrlCreateCheckbox(t('PREFS_APPEND_EXT_LABEL'), 10, 205, -1, 20)
	Global $NoBoxOpt = GUICtrlCreateCheckbox(t('PREFS_HIDE_STATUS_LABEL'), 10, 225, -1, 20)
	Global $GameModeOpt = GUICtrlCreateCheckbox(t('PREFS_HIDE_STATUS_FULLSCREEN_LABEL'), 10, 245, -1, 20)
	Global $OpenOutDirOpt = GUICtrlCreateCheckbox(t('PREFS_OPEN_FOLDER_LABEL'), 10, 265, -1, 20)
	Global $FeedbackPromptOpt = GUICtrlCreateCheckbox(t('PREFS_FEEDBACK_PROMPT_LABEL'), 10, 285, -1, 20)
	Global $StoreGUIPositionOpt = GUICtrlCreateCheckbox(t('PREFS_WINDOW_POSITION_LABEL'), 10, 305, -1, 20)
	Global $UsageStatsOpt = GUICtrlCreateCheckbox(t('PREFS_SEND_STATS_LABEL'), 10, 325, -1, 20)
	Global $LogOpt = GUICtrlCreateCheckbox(t('PREFS_LOG_LABEL'), 10, 345, -1, 20)
	Global $VideoTrackOpt = GUICtrlCreateCheckbox(t('PREFS_VIDEOTRACK_LABEL'), 10, 365, -1, 20)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; Source file options
	GUICtrlCreateGroup(t('PREFS_SOURCE_FILES_LABEL'), 5, 395, 240, 40)
	$DeleteOrigFileOpt[$OPTION_KEEP] = GUICtrlCreateRadio(t('PREFS_SOURCE_FILES_OPT_KEEP'), 10, 410)
	$DeleteOrigFileOpt[$OPTION_DELETE] = GUICtrlCreateRadio(t('PREFS_SOURCE_FILES_OPT_DELETE'), GetPos($guiprefs, $DeleteOrigFileOpt[$OPTION_KEEP], 20), 410)
	$DeleteOrigFileOpt[$OPTION_ASK] = GUICtrlCreateRadio(t('PREFS_SOURCE_FILES_OPT_ASK'), GetPos($guiprefs, $DeleteOrigFileOpt[$OPTION_DELETE], 20), 410)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; Buttons
	Local $prefsok = GUICtrlCreateButton(t('OK_BUT'), 55, 443, 60, 20)
	Local $prefscancel = GUICtrlCreateButton(t('CANCEL_BUT'), 135, 443, 60, 20)

	; Tooltips
	GUICtrlSetTip($warnexecuteopt, t('PREFS_WARN_EXECUTE_TOOLTIP'))
	GUICtrlSetTip($freeSpaceCheckOpt, t('PREFS_CHECK_FREE_SPACE_TOOLTIP'))
	GUICtrlSetTip($unicodecheckopt, t('PREFS_CHECK_UNICODE_TOOLTIP'))
	GUICtrlSetTip($appendextopt, t('PREFS_APPEND_EXT_TOOLTIP'))
	GUICtrlSetTip($GameModeOpt, t('PREFS_HIDE_STATUS_FULLSCREEN_TOOLTIP'))
	GUICtrlSetTip($FeedbackPromptOpt, t('PREFS_FEEDBACK_PROMPT_TOOLTIP'))
	GUICtrlSetTip($UsageStatsOpt, t('PREFS_SEND_STATS_TOOLTIP'))
	GUICtrlSetTip($VideoTrackOpt, t('PREFS_VIDEOTRACK_TOOLTIP'))
	GUICtrlSetTip($DeleteOrigFileOpt[$OPTION_ASK], t('PREFS_SOURCE_FILES_OPT_KEEP_TOOLTIP'))

	; Set properties
	GUICtrlSetState($prefsok, $GUI_DEFBUTTON)
	If $history Then GUICtrlSetState($historyopt, $GUI_CHECKED)
	If $warnexecute Then GUICtrlSetState($warnexecuteopt, $GUI_CHECKED)
	If $freeSpaceCheck Then GUICtrlSetState($freeSpaceCheckOpt, $GUI_CHECKED)
	If $checkUnicode Then GUICtrlSetState($unicodecheckopt, $GUI_CHECKED)
	If $appendext Then GUICtrlSetState($appendextopt, $GUI_CHECKED)
	If $NoBox Then GUICtrlSetState($NoBoxOpt, $GUI_CHECKED)
	If $bHideStatusBoxIfFullscreen Then GUICtrlSetState($GameModeOpt, $GUI_CHECKED)
	If $OpenOutDir Then GUICtrlSetState($OpenOutDirOpt, $GUI_CHECKED)
	If $FB_ask Then GUICtrlSetState($FeedbackPromptOpt, $GUI_CHECKED)
	If $StoreGUIPosition Then GUICtrlSetState($StoreGUIPositionOpt, $GUI_CHECKED)
	If $bSendStats Then GUICtrlSetState($UsageStatsOpt, $GUI_CHECKED)
	If $Log Then GUICtrlSetState($LogOpt, $GUI_CHECKED)
	If $bExtractVideo Then GUICtrlSetState($VideoTrackOpt, $GUI_CHECKED)
	GUICtrlSetState($DeleteOrigFileOpt[$iDeleteOrigFile], $GUI_CHECKED)
	GUICtrlSetData($langselect, $langlist, $language)

	; Set events
	GUICtrlSetOnEvent($prefsok, "GUI_Prefs_Ok")
	GUICtrlSetOnEvent($prefscancel, "GUI_Prefs_Exit")
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_Prefs_Exit")

	; Display GUI and wait for action
	GUISetState(@SW_SHOW)
EndFunc

; Exit preferences GUI if Cancel clicked or window closed
Func GUI_Prefs_Exit()
	GUIDelete($guiprefs)
	$guiprefs = False
	Cout("Closing preferences GUI")
EndFunc   ;==>GUI_Prefs_Exit

; Exit preferences GUI if OK clicked
Func GUI_Prefs_OK()
	; Universal preferences
	$redrawgui = False

	If GUICtrlRead($historyopt) == $GUI_CHECKED Then
		If $history == 0 Then
			$history = 1
			$redrawgui = True
		EndIf
	Else
		If $history == 1 Then
			$history = 0
			IniDelete($prefs, $HISTORY_FILE)
			IniDelete($prefs, $HISTORY_DIR)
			$redrawgui = True
		EndIf
	EndIf
	If $language <> GUICtrlRead($langselect) Then
		$language = GUICtrlRead($langselect)
		$redrawgui = True
	EndIf
	If $Timeout / 1000 <> GUICtrlRead($TimeoutCont) And Int(GUICtrlRead($TimeoutCont)) > 9 Then $Timeout = Int(GUICtrlRead($TimeoutCont)) * 1000

	If $updateinterval <> GUICtrlRead($IntervalCont) Then $updateinterval = GUICtrlRead($IntervalCont)

	; Format-specific preferences
	If GUICtrlRead($NoBoxOpt) == $GUI_CHECKED Then
		$NoBox = 1
		TrayItemSetState($Tray_Statusbox, $TRAY_CHECKED)
	Else
		$NoBox = 0
		TrayItemSetState($Tray_Statusbox, $TRAY_UNCHECKED)
	EndIf

	$warnexecute = Number(GUICtrlRead($warnexecuteopt) == $GUI_CHECKED)
	$checkUnicode = Number(GUICtrlRead($unicodecheckopt) == $GUI_CHECKED)
	$freeSpaceCheck = Number(GUICtrlRead($freeSpaceCheckOpt) == $GUI_CHECKED)
	$appendext = Number(GUICtrlRead($appendextopt) == $GUI_CHECKED)
	$bHideStatusBoxIfFullscreen = Number(GUICtrlRead($GameModeOpt) == $GUI_CHECKED)
	$OpenOutDir = Number(GUICtrlRead($OpenOutDirOpt) == $GUI_CHECKED)
	$FB_ask = Number(GUICtrlRead($FeedbackPromptOpt) == $GUI_CHECKED)
	$Log = Number(GUICtrlRead($LogOpt) == $GUI_CHECKED)
	$bExtractVideo = Number(GUICtrlRead($VideoTrackOpt) == $GUI_CHECKED)
	$StoreGUIPosition = Number(GUICtrlRead($StoreGUIPositionOpt) == $GUI_CHECKED)
	$bSendStats = Number(GUICtrlRead($UsageStatsOpt) == $GUI_CHECKED)

	For $i = 0 To 2
		If GUICtrlRead($DeleteOrigFileOpt[$i]) == $GUI_CHECKED Then $iDeleteOrigFile = $i
	Next

	WritePrefs()

	GUIDelete($guiprefs)
	$guiprefs = False

	If Not $redrawgui Then Return
	GUIDelete($guimain)
	CreateGUI()
EndFunc

; Handle click on OK
Func GUI_OK()
	If Not GUI_OK_Set(True) Then Return
	GUIGetPosition()
	GUIDelete($guimain)
	$guimain = False
	Cout("Closing main GUI")
EndFunc   ;==>GUI_OK

; Set file to extract and target directory
Func GUI_OK_Set($bMsg = False)
	$file = EnvParse(GUICtrlRead($filecont))
	If FileExists($file) Then
		If EnvParse(GUICtrlRead($dircont)) == "" Then
			$outdir = '/sub'
		Else
			$outdir = EnvParse(GUICtrlRead($dircont))
		EndIf
		Return 1
	ElseIf $bMsg Then
		If $file <> '' Then $file &= ' ' & t('DOES_NOT_EXIST')
		MsgBox($iTopmost + 48, $title, t('INVALID_FILE_SELECTED', $file))
	EndIf
	Return 0
EndFunc

; Add file to batch queue
Func GUI_Batch()
	If GUI_OK_Set(True) Then
		AddToBatch()
		GUICtrlSetData($filecont, "")
		If Not $KeepOutdir Then GUICtrlSetData($dircont, "")
	Else ; Start batch process if items in queue and input fields empty
		If GetBatchQueue() Then GUI_Batch_OK()
	EndIf
EndFunc   ;==>GUI_Batch

; Execute batch queue
Func GUI_Batch_OK()
	Cout("Closing main GUI - batch mode")
	Local $file = GUICtrlRead($filecont)
	If $file <> "" And Not StringIsSpace($file) Then GUI_Batch()
	GUIGetPosition()
	GUIDelete($guimain)

	If FileExists($fileScanLogFile) Then FileDelete($fileScanLogFile)

	terminate($STATUS_BATCH, '', '')
EndFunc

; Display batch queue and allow changes
Func GUI_Batch_Show()
	Local Const $iListLeft = 8, $iListTop = 8
	Local $iLastIndex = -1, $tt = False
	Cout("Opening batch queue edit GUI")
	$GUI_Batch = GUICreate($name, 418, 267, 476, 262, -1, -1, $guimain)
	_GuiSetColor()
	$GUI_Batch_List = GUICtrlCreateList("", $iListLeft, $iListTop, 401, 201)
	GUICtrlSetData(-1, _ArrayToString($queueArray, "|", 1))
	$GUI_Batch_OK = GUICtrlCreateButton(t('OK_BUT'), 40, 225, 75, 25)
	$GUI_Batch_Cancel = GUICtrlCreateButton(t('CANCEL_BUT'), 171, 225, 75, 25)
	$GUI_Batch_Delete = GUICtrlCreateButton(t('DELETE_BUT'), 304, 224, 73, 25)
	GUISetState(@SW_SHOW)
	Opt("GUIOnEventMode", 0)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $GUI_Batch_Cancel
				GetBatchQueue()
				ExitLoop
			Case $GUI_Batch_OK
				Cout("Batch queue was modified")
				$return = UBound($queueArray)
				If $return = 1 Then
					EnableBatchMode(False)
					ExitLoop
				EndIf
				$queueArray[0] = $return
;~ 				_ArrayDisplay($queueArray)
				SaveBatchQueue()
				; Only called to update main GUI batch button
				GetBatchQueue()
				ExitLoop
			Case $GUI_Batch_Delete
				Local $pos = _GUICtrlListBox_GetCurSel($GUI_Batch_List)
				If $pos > -1 Then
					Local $return = _GUICtrlListBox_GetText($GUI_Batch_List, $pos)
;~ 					Cout($return)
					$pos = _ArraySearch($queueArray, $return)
;~ 					Cout($pos)
;~ 					_ArrayDisplay($queueArray)
					If @error Then ContinueLoop
					If _ArrayDelete($queueArray, $pos) Then GUICtrlSetData($GUI_Batch_List, "|" & _ArrayToString($queueArray, "|", 1))
				EndIf
			Case Else
				; Display tooltips if file name too long
				; Code by Malkey (https://www.autoitscript.com/forum/topic/146743-listbox-tooltip-for-long-items/?do=findComment&comment=1039835)
				$aCI = GUIGetCursorInfo($GUI_Batch)
				If $aCI[4] = $GUI_Batch_List Then
					$iIndex = _GUICtrlListBox_ItemFromPoint($GUI_Batch_List, $aCI[0] - $iListLeft, $aCI[1] - $iListTop)
					If $iLastIndex == $iIndex Then ContinueLoop
					$iLastIndex = $iIndex
					$sText = _GUICtrlListBox_GetText($GUI_Batch_List, $iIndex)
					If StringLen($sText) > 72 Then
						ToolTip($sText)
						$tt = True
					Else
						ToolTip("")
						$tt = False
						$iLastIndex = -1
					EndIf
				EndIf
				If $tt And $aCI[4] <> $GUI_Batch_List Then
					$tt = False
					$iLastIndex = -1
					ToolTip("")
				EndIf
		EndSwitch
	WEnd

	GUIDelete($GUI_Batch)
	Opt("GUIOnEventMode", 1)
EndFunc

; Clear batch queue
Func GUI_Batch_Clear()
	Cout("Batch queue cleared, batch mode disabled")
	EnableBatchMode(False)
EndFunc   ;==>GUI_Batch_Clear

; Process dropped files
Func GUI_Drop()
	;_ArrayDisplay($gaDropFiles)
	Cout("Drag and drop action detected")
	For $i = 0 To UBound($gaDropFiles) - 1
		If FileExists($gaDropFiles[$i]) Then
			; Folder is passed
			If StringInStr(FileGetAttrib($gaDropFiles[$i]), "D") Then
				Cout("Drag and drop - folder passed")
				$return = ReturnFiles($gaDropFiles[$i])
				Local $files = StringSplit($return, "|")
				;_ArrayDisplay($files)
				For $j = 1 To $files[0]
					If Not StringInStr(FileGetAttrib($gaDropFiles[$i] & "\" & $files[$j]), "D") Then
						Global $file = $gaDropFiles[$i] & "\" & $files[$j]
						GUI_Drop_Parse()
						GUI_Batch()
					EndIf
				Next

			; File is passed
			Else
				Global $file = $gaDropFiles[$i]
				GUI_Drop_Parse()

				If UBound($gaDropFiles) == 1 Then
					Return
				Else
					GUI_Batch()
				EndIf
			EndIf
		EndIf
	Next
	Cout("Drag and drop - a total of " & UBound($gaDropFiles) & " files added to batch queue")
EndFunc   ;==>GUI_Drop

; Process dropped files
Func GUI_Drop_Parse()
	If $history Then
		$filelist = '|' & $file & '|' & ReadHist($HISTORY_FILE)
		GUICtrlSetData($filecont, $filelist, $file)
	Else
		GUICtrlSetData($filecont, $file)
	EndIf

	If GUICtrlRead($dircont) == "" Or Not $KeepOutdir Then
		FilenameParse($file)
		If $history Then
			$dirlist = '|' & $initoutdir & '|' & ReadHist($HISTORY_DIR)
			GUICtrlSetData($dircont, $dirlist, $initoutdir)
		Else
			GUICtrlSetData($dircont, $initoutdir)
		EndIf
	EndIf
EndFunc

; Drag and drop handler for multiple file support
; http://www.autoitscript.com/forum/topic/28062-drop-multiple-files-on-any-control/page__view__findpost__p__635231
Func WM_DROPFILES_UNICODE_FUNC($hWnd, $msgID, $wParam, $lParam)
	Local $nSize, $pFileName
	Local $nAmt = DllCall("shell32.dll", "int", "DragQueryFileW", "hwnd", $wParam, "int", 0xFFFFFFFF, "ptr", 0, "int", 255)
	For $i = 0 To $nAmt[0] - 1
		$nSize = DllCall("shell32.dll", "int", "DragQueryFileW", "hwnd", $wParam, "int", $i, "ptr", 0, "int", 0)
		$nSize = $nSize[0] + 1
		$pFileName = DllStructCreate("wchar[" & $nSize & "]")
		DllCall("shell32.dll", "int", "DragQueryFileW", "hwnd", $wParam, "int", $i, "int", DllStructGetPtr($pFileName), "int", $nSize)
		ReDim $gaDropFiles[$i + 1]
		$gaDropFiles[$i] = DllStructGetData($pFileName, 1)
		$pFileName = 0
	Next
;~ 	_ArrayDisplay($gaDropFiles)
EndFunc   ;==>WM_DROPFILES_UNICODE_FUNC

; Create Feedback GUI
Func GUI_Feedback($Type = "", $file = "")
	Cout("Creating feedback GUI")
	Opt("GUIOnEventMode", 0)
	Global $FB_GUI = GUICreate(t('FEEDBACK_TITLE_LABEL'), 251, 370, -1, -1, BitOR($WS_SIZEBOX, $WS_SYSMENU), -1, $guimain)
	_GuiSetColor()
	GUICtrlCreateLabel(t('FEEDBACK_TYPE_LABEL'), 8, 8, -1, 15)
	If $Type == "" Then
		$FB_TypeCont = GUICtrlCreateCombo(t('FEEDBACK_TYPE_STANDARD'), 8, 24, 105, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
	Else
		$FB_TypeCont = GUICtrlCreateCombo($Type, 8, 24, 105, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
		GUICtrlSetData(-1, t('FEEDBACK_TYPE_STANDARD'))
	EndIf
	GUICtrlSetData(-1, t('FEEDBACK_TYPE_OPTIONS'))
	GUICtrlCreateLabel(t('FEEDBACK_SYSINFO_LABEL'), 120, 8, -1, 15)
	$FB_SysCont = GUICtrlCreateInput(@OSVersion & " " & @OSArch & (@OSServicePack = ""? "": " " & @OSServicePack) & ", Lang: " & @OSLang & ", UE: " & $language, 120, 24, 121, 21)
	GUICtrlCreateLabel(t('FEEDBACK_FILE_LABEL'), 8, 56, -1, 15)
	GUICtrlSetTip(-1, t('FEEDBACK_FILE_TOOLTIP'), "", 0, 1)
	$FB_FileCont = GUICtrlCreateInput($file, 8, 72, 233, 21)
	GUICtrlSetTip(-1, t('FEEDBACK_FILE_TOOLTIP'), "", 0, 1)
	GUICtrlCreateLabel(t('FEEDBACK_OUTPUT_LABEL'), 8, 104, -1, 15)
	GUICtrlSetTip(-1, t('FEEDBACK_OUTPUT_TOOLTIP'), "", 0, 1)
	$FB_OutputCont = GUICtrlCreateEdit("", 8, 120, 233, 49, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_WANTRETURN, $WS_VSCROLL))
	GUICtrlSetData(-1, $debug)
	GUICtrlSetTip(-1, t('FEEDBACK_OUTPUT_TOOLTIP'), "", 0, 1)
	GUICtrlCreateLabel(t('FEEDBACK_MESSAGE_LABEL'), 8, 176, -1, 15)
	GUICtrlSetTip(-1, t('FEEDBACK_MESSAGE_TOOLTIP'), "", 0, 1)
	$FB_MessageCont = GUICtrlCreateEdit("", 8, 192, 233, 73, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_WANTRETURN, $WS_VSCROLL))
	GUICtrlSetTip(-1, t('FEEDBACK_MESSAGE_TOOLTIP'), "", 0, 1)
	GUICtrlCreateLabel(t('FEEDBACK_EMAIL_LABEL'), 8, 272, -1, 15)
	GUICtrlSetTip(-1, t('FEEDBACK_EMAIL_TOOLTIP'), "", 0, 1)
	$FB_MailCont = GUICtrlCreateInput("", 8, 288, 233, 21)
	GUICtrlSetTip(-1, t('FEEDBACK_EMAIL_TOOLTIP'), "", 0, 1)
	$FB_Send = GUICtrlCreateButton(t('SEND_BUT'), 55, 318, 60, 20)
	$FB_Cancel = GUICtrlCreateButton(t('CANCEL_BUT'), 135, 318, 60, 20)
	$hSelectAll = GUICtrlCreateDummy()

	Local $accelKeys[1][2] = [["^a", $hSelectAll]]
	GUISetAccelerators($accelKeys)
	GUICtrlSetState($FB_MessageCont, $GUI_FOCUS)

	; Set minimum window size
	GUIRegisterMsg($WM_GETMINMAXINFO, "GUI_WM_GETMINMAXINFO_Feedback")
	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $FB_Send
				GUI_Feedback_Send(_GUICtrlComboBox_GetCurSel($FB_TypeCont), GUICtrlRead($FB_SysCont), GUICtrlRead($FB_FileCont), GUICtrlRead($FB_OutputCont), GUICtrlRead($FB_MessageCont), GUICtrlRead($FB_MailCont))
				ExitLoop
			Case $GUI_EVENT_CLOSE, $FB_Cancel
				ExitLoop
			Case $hSelectAll
				GUI_Edit_SelectAll()
		EndSwitch
	WEnd

	GUIDelete($FB_GUI)
	Opt("GUIOnEventMode", 1)
EndFunc

; Exit feedback GUI if OK clicked
Func GUI_Feedback_Send($FB_Type, $FB_Sys, $FB_File, $FB_Output, $FB_Message, $FB_Mail)
	If $FB_File = "" And $FB_Output = "" And $FB_Message = "" Then Return MsgBox($iTopmost + 16, $name, t('FEEDBACK_EMPTY'))

	; Opt-in privacy agreement
	If Not Prompt(64+4, 'FEEDBACK_PRIVACY', 0, 0) Then Return

	GUIDelete($FB_GUI)
	If $guimain Then GUISetState(@SW_HIDE, $guimain)
	_CreateTrayMessageBox(t('SENDING_FEEDBACK'))

	Local $FB_Types = StringSplit(t('FEEDBACK_TYPE_OPTIONS', '', 'english'), "|", $STR_NOCOUNT)
	_ArrayInsert($FB_Types, 0, t('FEEDBACK_TYPE_STANDARD', '', 'english'))
	$FB_Type = $FB_Types[$FB_Type]

	$FB_Text = $name & " Feedback: " & $FB_Type & @CRLF & _
			"------------------------------------------------------------------------------------------------" _
			 & @CRLF & "System Information: " & $title & ", " & $FB_Sys & @CRLF & @CRLF & "Sample File: " & $FB_File & @CRLF _
			 & "Type: " & StringReplace($filetype, @CRLF, "|") & @CRLF & @CRLF & " Output:" & @CRLF & $FB_Output & @CRLF & @CRLF & _
			"------------------------------------------------------------------------------------------------" _
			 & @CRLF & $FB_Message & @CRLF & @CRLF & "Sent by: " & @CRLF & $ID

	If StringInStr($FB_Mail, "@") Then $FB_Text &= @CRLF & $FB_Mail

	Const $boundary = "--UniExtractLog"

	$http = ObjCreate("winhttp.winhttprequest.5.1")
	$http.Open("POST", $sSupportURL, False)
	$http.SetRequestHeader("Content-Type", "multipart/form-data; boundary=" & StringTrimLeft($boundary, 2))

	Local $Data = $boundary & @CRLF & 'Content-Disposition: form-data; name="file"; filename="UE_Feedback"' & @CRLF & 'Content-Type: text/plain' & @CRLF & @CRLF & $FB_Text & @CRLF & $boundary & @CRLF & 'Content-Disposition: form-data; name="id"' & @CRLF & @CRLF & $ID & @CRLF & $boundary & '--'

	$http.Send($Data)
	$http.WaitForResponse()
	Local $sResponse = $http.ResponseText()

	_DeleteTrayMessageBox()

	If $sResponse = "1" Then
		Cout("Feedback successfully sent")
		MsgBox($iTopmost, $title, t('FEEDBACK_SUCCESS'))
	Else
		Cout("Error sending feedback")
		MsgBox($iTopmost+16, $title, t('FEEDBACK_ERROR'))
	EndIf

	GUISetState(@SW_SHOW, $guimain)
EndFunc

; Due to a bug in the Windows API, ctrl+a does not work for edit controls
; Workaround by Zedna (http://www.autoitscript.com/forum/topic/97473-hotkey-ctrla-for-select-all-in-the-selected-edit-box/#entry937287)
Func GUI_Edit_SelectAll()
	$hWnd = _WinAPI_GetFocus()
	$class = _WinAPI_GetClassName($hWnd)
	If $class = 'Edit' Then _GUICtrlEdit_SetSel($hWnd, 0, -1)
EndFunc

; Saves current position of main GUI
Func GUIGetPosition()
	If Not $StoreGUIPosition Then Return
	$pos = WinGetPos($guimain)
	SavePref('posx', $pos[0])
	SavePref('posy', $pos[1])
EndFunc   ;==>GUIGetPosition

; Set minimal size of main GUI
Func GUI_WM_GETMINMAXINFO_Main($hwnd, $Msg, $wParam, $lParam)
    $tagMaxinfo = DllStructCreate("int;int;int;int;int;int;int;int;int;int", $lParam)
    DllStructSetData($tagMaxinfo, 7, 320) ; min X
    DllStructSetData($tagMaxinfo, 8, 170) ; min Y
    ;DllStructSetData($tagMaxinfo, 9, 1200); max X
    ;DllStructSetData($tagMaxinfo, 10, 160) ; max Y
EndFunc

; Set minimal size of feedback GUI
Func GUI_WM_GETMINMAXINFO_Feedback($hWnd, $Msg, $wParam, $lParam)
	$tagMaxinfo = DllStructCreate("int;int;int;int;int;int;int;int;int;int", $lParam)
	DllStructSetData($tagMaxinfo, 7, 270) ; min width
	DllStructSetData($tagMaxinfo, 8, 380) ; min height
	;DllStructSetData($tagMaxinfo,  9, ) ; max width
	;DllStructSetData($tagMaxinfo, 10, )  ; max height
EndFunc

; Tooltip does not work for disabled controls, so here's a workaround
Func GUI_Create_Tooltip($gui, $hWnd, $Data)
	Local $pos = ControlGetPos($gui, "", $hWnd)
	If @error Then
		Cout("Error creating tooltip: failed to determine size of control")
		Return SetError(1, 0, -1)
	EndIf
	Local $label = GUICtrlCreateLabel("", $pos[0], $pos[1], $pos[2], $pos[3])
	GUICtrlSetTip($label, $Data)
	; Set initial control on top
	; Based on http://www.autoitscript.com/forum/topic/146182-solved-change-z-ordering-of-controls/#entry1034567
	If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	_WinAPI_SetWindowPos($hWnd, $HWND_BOTTOM, 0, 0, 0, 0, $SWP_NOMOVE + $SWP_NOSIZE + $SWP_NOCOPYBITS)
	Return $label
EndFunc   ;==>GUI_Create_Tooltip

; Create GUI to change context menu
Func GUI_ContextMenu()
	Cout("Creating context menu GUI")
	Local $iSize = UBound($CM_Shells) - 1
	Global $CM_Checkbox[$iSize + 1], $CM_Picture = False

	Global $CM_GUI = GUICreate(t('PREFS_TITLE_LABEL'), 450, 630, -1, -1, -1, $exStyle, $guimain)
	_GuiSetColor()

	GUICtrlCreateGroup(t('CONTEXT_ENTRIES_LABEL'), 8, 4, 434, 495)
	Global $CM_Checkbox_enabled = GUICtrlCreateCheckbox(t('CONTEXT_ENABLED_LABEL'), 24, 22, -1, 17)
	Global $CM_Checkbox_allusers = GUICtrlCreateCheckbox(t('CONTEXT_ALL_USERS_LABEL'), GetPos($CM_GUI, $CM_Checkbox_enabled, 25), 22, -1, 17)
	Global $CM_Simple_Radio = GUICtrlCreateRadio(t('CONTEXT_SIMPLE_RADIO'), 96, 50, 145, 17)
	Global $CM_Cascading_Radio = GUICtrlCreateRadio(t('CONTEXT_CASCADING_RADIO'), 296, 50, 137, 17)
	Global $CM_Picture = GUICtrlCreatePic("", 55, 78, 0, 0, -1, $WS_EX_LAYERED)

	Local $pos = 0, $iY = 428
	For $i = 0 To $iSize Step 2
		$CM_Checkbox[$i] = GUICtrlCreateCheckbox(t($CM_Shells[$i][2]), 25, $iY)
		If $pos == 0 Then $pos = GetPos($CM_GUI, $CM_Checkbox[0], 125)
		If $i == $iSize Then ExitLoop
		$CM_Checkbox[$i+1] = GUICtrlCreateCheckbox(t($CM_Shells[$i+1][2]), $pos, $iY)
		$iY += 20
	Next
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	GUICtrlCreateGroup(t('CONTEXT_FILE_ASSOC_LABEL'), 8, 505, 434, 80)
	Global $CM_Checkbox_add = GUICtrlCreateCheckbox(t('CONTEXT_ENABLED_LABEL'), 24, 525, -1, 17)
	Global $CM_Checkbox_allusers2 = GUICtrlCreateCheckbox(t('CONTEXT_ALL_USERS_LABEL'), GetPos($CM_GUI, $CM_Checkbox_enabled, 25), 525, -1, 17)
	Global $CM_add_input = GUICtrlCreateInput("", 24, 550, 401, 21)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	$CM_OK = GUICtrlCreateButton(t('OK_BUT'), 112, 595, 89, 25)
	$CM_Cancel = GUICtrlCreateButton(t('CANCEL_BUT'), 232, 595, 89, 25)

	GUICtrlSetState($CM_Checkbox_allusers, $GUI_DISABLE)
	GUICtrlSetState($CM_Checkbox_allusers2, $GUI_DISABLE)
	GUICtrlSetState($CM_Simple_Radio, $GUI_CHECKED)

	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_Close")
	GUICtrlSetOnEvent($CM_Cancel, "GUI_Close")
	GUICtrlSetOnEvent($CM_OK, "GUI_ContextMenu_OK")
	GUICtrlSetOnEvent($CM_Checkbox_enabled, "GUI_ContextMenu_activate")
	GUICtrlSetOnEvent($CM_Checkbox_add, "GUI_ContextMenu_activate")

	; Check which commands are activated
	For $i = 0 To $iSize
		If RegExists($regall & $CM_Shells[$i][0], "") Then
			Global $reguser = $regall
			GUICtrlSetState($CM_Checkbox_allusers, $GUI_CHECKED)
			GUICtrlSetState($CM_Checkbox[$i], $GUI_CHECKED)
			GUICtrlSetState($CM_Checkbox_enabled, $GUI_CHECKED)
		EndIf
		If RegExists($regcurrent & $CM_Shells[$i][0], "") Then
			Global $reguser = $regcurrent
			GUICtrlSetState($CM_Checkbox_allusers, $GUI_UNCHECKED)
			GUICtrlSetState($CM_Checkbox[$i], $GUI_CHECKED)
			GUICtrlSetState($CM_Checkbox_enabled, $GUI_CHECKED)
		EndIf
	Next

	; Disable Cascading context menu for non Win 7+ users as it is not supported
	If _IsWin7() Then
		; Check if Cascading context menu entries are enabled
		For $i = 0 To $iSize
			If RegExists($regall & "\Uniextract\Shell\" & $CM_Shells[$i][0], "") Then
				Global $reguser = $regall
				GUICtrlSetState($CM_Checkbox[$i], $GUI_CHECKED)
				GUICtrlSetState($CM_Cascading_Radio, $GUI_CHECKED)
				GUICtrlSetState($CM_Checkbox_enabled, $GUI_CHECKED)
			EndIf
			If RegExists($regcurrent & "\Uniextract\Shell\" & $CM_Shells[$i][0], "") Then
				Global $reguser = $regcurrent
				GUICtrlSetState($CM_Checkbox[$i], $GUI_CHECKED)
				GUICtrlSetState($CM_Cascading_Radio, $GUI_CHECKED)
				GUICtrlSetState($CM_Checkbox_enabled, $GUI_CHECKED)
			EndIf
		Next
		; Register function to change image
		GUICtrlSetOnEvent($CM_Simple_Radio, "GUI_ContextMenu_ChangePic")
		GUICtrlSetOnEvent($CM_Cascading_Radio, "GUI_ContextMenu_ChangePic")
	Else
		GUICtrlSetState($CM_Cascading_Radio, $GUI_DISABLE)
		GUI_Create_Tooltip($CM_GUI, $CM_Cascading_Radio, t('CONTEXT_CASCADING_RADIO_TOOLTIP'))
	EndIf

	; Create tooltips for disabled admin only options
	If Not IsAdmin() Then
		GUI_Create_Tooltip($CM_GUI, $CM_Checkbox_allusers, t('CONTEXT_ADMIN_REQUIRED'))
		GUI_Create_Tooltip($CM_GUI, $CM_Checkbox_allusers2, t('CONTEXT_ADMIN_REQUIRED'))
	EndIf

	; Check for additional file associations
	If $addassocenabled Then GUICtrlSetState($CM_Checkbox_add, $GUI_CHECKED)
	If $addassocallusers Then GUICtrlSetState($CM_Checkbox_allusers2, $GUI_CHECKED)

	GUICtrlSetData($CM_add_input, $addassoc)

	; Activate controls if context menu entries are enabled
	GUI_ContextMenu_activate()
	GUI_ContextMenu_ChangePic()

	GUISetState(@SW_SHOW)
EndFunc   ;==>GUI_ContextMenu

; Change picture according to selected context menu type
Func GUI_ContextMenu_ChangePic()
	If GUICtrlRead($CM_Cascading_Radio) = $GUI_CHECKED Then
		GUICtrlSetImage($CM_Picture, ".\support\Icons\cascading.jpg")
	Else
		GUICtrlSetImage($CM_Picture, ".\support\Icons\simple.jpg")
	EndIf
EndFunc   ;==>GUI_ContextMenu_ChangePic

; Close GUI and create context menu entries
Func GUI_ContextMenu_OK()
	Local $iSize = UBound($CM_Shells) - 1
	Sleep(100)
	GUISetState(@SW_HIDE)

	; Remove old associations
	GUI_ContextMenu_remove()

	;If NOT RegExists($reguser) Then RegWrite($reguser)

	Cout("Registering context menu entries")
	If GUICtrlRead($CM_Checkbox_enabled) == $GUI_CHECKED Then

		; Select registry key
		Global $reguser = GUICtrlRead($CM_Checkbox_allusers) == $GUI_CHECKED? $regall: $regcurrent

		; simple
		If GUICtrlRead($CM_Simple_Radio) == $GUI_CHECKED Then
			For $i = 0 To $iSize
				$command = '"' & @ScriptFullPath & '" "%1"' & $CM_Shells[$i][1]
				If GUICtrlRead($CM_Checkbox[$i]) == $GUI_CHECKED Then
					RegWrite($reguser & $CM_Shells[$i][0], "", "REG_SZ", t($CM_Shells[$i][2]))
					RegWrite($reguser & $CM_Shells[$i][0] & "\command", "", "REG_SZ", $command)
					; Add icon to context menu, seems to work only on win 7
					If $win7 Then RegWrite($reguser & $CM_Shells[$i][0], "Icon", "REG_SZ", @ScriptFullPath & ",0")
				EndIf
			Next

		; cascading
		ElseIf $win7 And GUICtrlRead($CM_Cascading_Radio) == $GUI_CHECKED Then
			RegWrite($reguser & "uniextract", "MUIVerb", "REG_SZ", "Universal Extractor")
			RegWrite($reguser & "uniextract", "Icon", "REG_SZ", @ScriptFullPath & ",0")
			RegWrite($reguser & "uniextract", "SubCommands", "REG_SZ", "")

			For $i = 0 To $iSize
				$command = '"' & @ScriptFullPath & '" "%1"' & $CM_Shells[$i][1]
				If GUICtrlRead($CM_Checkbox[$i]) == $GUI_CHECKED Then
					RegWrite($reguser & "Uniextract\Shell\" & $CM_Shells[$i][0], "", "REG_SZ", t($CM_Shells[$i][2]))
					RegWrite($reguser & "Uniextract\Shell\" & $CM_Shells[$i][0] & "\command", "", "REG_SZ", $command)

					; Icon
					RegWrite($reguser & "Uniextract\Shell\" & $CM_Shells[$i][0], "Icon", "REG_SZ", @ScriptFullPath & ",0")
				EndIf
			Next
		EndIf
	EndIf

	; File associations
	If GUICtrlRead($CM_add_input) == "" Then GUICtrlSetState($CM_Checkbox_add, $GUI_UNCHECKED)
	If GUICtrlRead($CM_Checkbox_add) == $GUI_CHECKED Then
		$return = MsgBox($iTopmost + 48 + 4, $name, t('CONTEXT_DANGEROUS'))
		If $return == 6 And ($addassocenabled == 0 Or ($addassocenabled = 1 And $addassoc <> GUICtrlRead($CM_add_input))) Then GUI_ContextMenu_fileassoc(1)
	Else
		If $addassocenabled Then GUI_ContextMenu_fileassoc(0)
	EndIf
	GUIDelete($CM_GUI)
EndFunc   ;==>GUI_ContextMenu_OK

; (De)activate controls if enabled but not checked
Func GUI_ContextMenu_activate()
	; Set state according to main enable checkbox
	Local $state = GUICtrlRead($CM_Checkbox_enabled) = $GUI_CHECKED? $GUI_ENABLE: $GUI_DISABLE

	If IsAdmin() Then GUICtrlSetState($CM_Checkbox_allusers, $state)
	GUICtrlSetState($CM_Simple_Radio, $state)

	For $i = 0 To UBound($CM_Shells) - 1
		GUICtrlSetState($CM_Checkbox[$i], $state)
	Next

	If $win7 Then GUICtrlSetState($CM_Cascading_Radio, $state)
	If GUICtrlRead($CM_Checkbox_add) = $GUI_CHECKED Then
		If IsAdmin() Then GUICtrlSetState($CM_Checkbox_allusers2, $GUI_ENABLE)
		GUICtrlSetState($CM_add_input, $GUI_ENABLE)
	Else
		GUICtrlSetState($CM_Checkbox_allusers2, $GUI_DISABLE)
		GUICtrlSetState($CM_add_input, $GUI_DISABLE)
	EndIf
EndFunc   ;==>GUI_ContextMenu_activate

; Create/remove file associations
Func GUI_ContextMenu_fileassoc($enable)
	; Delete old file associations

	If $addassocallusers Then
		$sRegistryKey = "HKLM" & $reg64 & "\SOFTWARE\Classes\"
	Else
		$sRegistryKey = "HKCU" & $reg64 & "\SOFTWARE\Classes\"
	EndIf

	Local $files = StringSplit($addassoc, ",")
	;_ArrayDisplay($files)
	For $i = 1 To $files[0]
		_ShellFile_Uninstall(StringStripWS($files[$i], 1), $sRegistryKey)
	Next
	$files = 0

	$addassocenabled = $enable
	SavePref("addassocenabled", $addassocenabled)

	; Return if associations are disabled
	If Not $enable Then Return

	; Select registry key
	If GUICtrlRead($CM_Checkbox_allusers2) == $GUI_CHECKED Then
		$sRegistryKey = "HKLM" & $reg64 & "\SOFTWARE\Classes\"
		$addassocallusers = 1
	Else
		$sRegistryKey = "HKCU" & $reg64 & "\SOFTWARE\Classes\"
		$addassocallusers = 0
	EndIf

	; Create new associations
	$addassoc = GUICtrlRead($CM_add_input)
	$files = StringSplit($addassoc, ",")
	For $i = 1 To $files[0]
		_ShellFile_Install($name & " " & StringStripWS($files[$i], 1), StringStripWS($files[$i], 1), $name, $sRegistryKey)
	Next
	$files = 0

	; Save associated filetypes
	SavePref('addassoc', $addassoc)
	SavePref('addassocallusers', $addassocallusers)
EndFunc   ;==>GUI_ContextMenu_fileassoc

; Creates file association for a specified file
; Based on _ShellFile.au3 by guinness (http://www.autoitscript.com/forum/topic/129955-shellfile-create-an-entry-in-the-
; shell-contextmenu-when-selecting-an-assigned-filetype-includes-the-program-icon-as-well/#entry903513)
Func _ShellFile_Install($sText, $sFileType, $sName, $sRegistryKey)
	Cout("Creating File Association: ." & $sFileType)
	If StringLeft($sFileType, 1) = "." Then $sFileType = StringTrimLeft($sFileType, 1)

	RegWrite($sRegistryKey & "." & $sFileType, "", "REG_SZ", $sName)
	RegWrite($sRegistryKey & $sName & "\DefaultIcon\", "", "REG_SZ", @ScriptFullPath & ",0")
	RegWrite($sRegistryKey & $sName & "\shell\open", "", "REG_SZ", $sText)
	RegWrite($sRegistryKey & $sName & "\shell\open", "Icon", "REG_EXPAND_SZ", @ScriptFullPath & ",0")
	RegWrite($sRegistryKey & $sName & "\shell\open\command\", "", "REG_SZ", '"' & @ScriptFullPath & '" "%1"')
	RegWrite($sRegistryKey & $sName, "", "REG_SZ", $sText)
	RegWrite($sRegistryKey & $sName, "Icon", "REG_EXPAND_SZ", @ScriptFullPath & ",0")
	RegWrite($sRegistryKey & $sName & "\command", "", "REG_SZ", '"' & @ScriptFullPath & '" "%1"')

	Return SetError(@error, 0, @error)
EndFunc   ;==>_ShellFile_Install

; Removes file association for a specified file
; Based on _ShellFile.au3 by guinness (http://www.autoitscript.com/forum/topic/129955-shellfile-create-an-entry-in-the-
; shell-contextmenu-when-selecting-an-assigned-filetype-includes-the-program-icon-as-well/#entry903513)
Func _ShellFile_Uninstall($sFileType, $sRegistryKey)
	Cout("Removing File Association: ." & $sFileType)
	If StringLeft($sFileType, 1) = "." Then $sFileType = StringTrimLeft($sFileType, 1)

	Local $sName = RegRead($sRegistryKey & "." & $sFileType, "")
	If @error Then Return SetError(@error, 0, 0)

	RegDelete($sRegistryKey & "." & $sFileType)
	Return RegDelete($sRegistryKey & $sName)
EndFunc   ;==>_ShellFile_Uninstall

; Remove Universal Extractor entries from registry
Func GUI_ContextMenu_remove()
	Cout("Deregistering context menu entries")
	; Context menu
	For $i = 0 To UBound($CM_Shells) - 1
		If RegExists($regall & $CM_Shells[$i][0], "") Then RegDelete($regall & $CM_Shells[$i][0])
		If RegExists($regcurrent & $CM_Shells[$i][0], "") Then RegDelete($regcurrent & $CM_Shells[$i][0])
	Next

	; Win 7 specific
	If $win7 Then
		If RegExists($regall & "uniextract", "MUIVerb") Then RegDelete($regall & "uniextract")
		If RegExists($regcurrent & "uniextract", "MUIVerb") Then RegDelete($regcurrent & "uniextract")
	EndIf

	; File associations
	If $addassocenabled Then GUI_ContextMenu_fileassoc(0)
EndFunc   ;==>GUI_ContextMenu_remove

; Perform special actions if Universal Extractor is started the first time
Func GUI_FirstStart()
	Cout("Creating first start assistant")
	GUISetState(@SW_HIDE, $guimain)
	; Create GUI
	Global $FS_GUI = GUICreate($title, 504, 387)
	_GuiSetColor()
	GUICtrlCreatePic(".\support\Icons\uniextract_inno.bmp", 8, 312, 65, 65)
	GUICtrlCreateLabel($name, 8, 8, 488, 60, $SS_CENTER)
	GUICtrlSetFont(-1, 24, 800, 0, "MS Sans Serif")
	GUICtrlCreateLabel(t('FIRSTSTART_TITLE'), 8, 50, 488, 60, $SS_CENTER)
	GUICtrlSetFont(-1, 14, 800, 0, "MS Sans Serif")
	Global $FS_Section = GUICtrlCreateLabel("", 16, 85, 382, 28)
	GUICtrlSetFont(-1, 14, 800, 4, "MS Sans Serif")
	Global $FS_Text = GUICtrlCreateLabel("", 16, 120, 468, 125)
	Global $FS_Next = GUICtrlCreateButton(t('NEXT_BUT'), 296, 344, 89, 25)
	Local $FS_Cancel = GUICtrlCreateButton(t('CANCEL_BUT'), 400, 344, 89, 25)
	Global $FS_Prev = GUICtrlCreateButton(t('PREV_BUT'), 192, 344, 89, 25)
	GUICtrlSetState(-1, $GUI_HIDE)
	Global $FS_Button = GUICtrlCreateButton("", 187, 260, 129, 41)
	Global $FS_Progress = GUICtrlCreateLabel("", 80, 350, 21, 17)

	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_FirstStart_Exit")
	GUICtrlSetOnEvent($FS_Cancel, "GUI_FirstStart_Exit")
	GUICtrlSetOnEvent($FS_Next, "GUI_FirstStart_Next")
	GUICtrlSetOnEvent($FS_Prev, "GUI_FirstStart_Prev")

	Global $page = 1
	Global $FS_Sections = StringSplit(t('FIRSTSTART_PAGES'), "|")
	If @error Then
		MsgBox(16, $name, "No language file was found. Please redownload " & $name & ".")
		SavePref("ID", "")
		GUI_Website_Github()
		Exit 99
	EndIf
	Global $FS_Texts[UBound($FS_Sections)] = [ _
				"", t('FIRSTSTART_PAGE1'), t('FIRSTSTART_PAGE2'), t('FIRSTSTART_PAGE3'), t('FIRSTSTART_PAGE4'), _
				t('FIRSTSTART_PAGE5', $bindir), t('FIRSTSTART_PAGE6') _
			]
;~  	_ArrayDisplay($FS_Texts)
	GUISetState(@SW_SHOW)
	GUI_FirstStart_ShowPage()
EndFunc   ;==>GUI_FirstStart

; Next clicked
Func GUI_FirstStart_Prev()
	If $page = 2 Then
		GUICtrlSetState($FS_Prev, $GUI_HIDE)
	ElseIf $page = $FS_Sections[0] Then
		GUICtrlSetData($FS_Next, t('NEXT_BUT'))
		GUICtrlSetOnEvent($FS_Next, "GUI_FirstStart_Next")
	EndIf
	$page -= 1
	GUI_FirstStart_ShowPage()
EndFunc

; Back clicked
Func GUI_FirstStart_Next()
	If $page = 1 Then
		GUICtrlSetState($FS_Prev, $GUI_SHOW)
	ElseIf $page = $FS_Sections[0] - 1 Then
		GUICtrlSetData($FS_Next, t('FINISH_BUT'))
		GUICtrlSetOnEvent($FS_Next, "GUI_FirstStart_Exit")
	EndIf
	$page += 1
	GUI_FirstStart_ShowPage()
EndFunc

; Load a page of the first start GUI
Func GUI_FirstStart_ShowPage()
	GUICtrlSetData($FS_Progress, $page & "/" & $FS_Sections[0])
	GUICtrlSetData($FS_Section, $FS_Sections[$page])
	GUICtrlSetData($FS_Text, $FS_Texts[$page])
	GUICtrlSetState($FS_Button, $GUI_ENABLE)
	Cout("First start assistant - step " & $page)
	Switch $page
		Case 2
			GUICtrlSetState($FS_Button, $GUI_SHOW)
			GUICtrlSetData($FS_Button, t('PREFS_TITLE_LABEL'))
			GUICtrlSetOnEvent($FS_Button, "GUI_Prefs")
		Case 3
			;GUICtrlSetState($FS_Button, $GUI_SHOW)
			GUICtrlSetData($FS_Button, t('CONTEXT_ENTRIES_LABEL'))
			GUICtrlSetOnEvent($FS_Button, "GUI_ContextMenu")
		Case 4
			GUICtrlSetState($FS_Button, $GUI_SHOW)
			If HasPlugin($ffmpeg, True) Then
				GUICtrlSetData($FS_Button, t('TERM_INSTALLED'))
				GUICtrlSetState($FS_Button, $GUI_DISABLE)
			Else
				GUICtrlSetData($FS_Button, t('TERM_DOWNLOAD'))
				GUICtrlSetOnEvent($FS_Button, "GetFFmpeg")
			EndIf
		Case 6
			GUICtrlSetState($FS_Button, $GUI_HIDE)
			GUICtrlSetPos($FS_Text, 16, 120, 468, 170)
		Case Else
			GUICtrlSetState($FS_Button, $GUI_HIDE)
	EndSwitch
EndFunc

; Close First Start GUI
Func GUI_FirstStart_Exit()
	GUIDelete($FS_GUI)
	$FS_GUI = False
	GUISetState(@SW_SHOW, $guimain)
	Cout("First start configuration finished")
EndFunc   ;==>GUI_FirstStart_Exit

; CreatePlugin GUI
Func GUI_Plugins()
	; Define plugins
	; executable|name|description|filetypes|url|filemask|extractionfilter|outdir|password
	Local $aPluginInfo[13][8] = [ _
		[$arc_conv, 'arc_conv', t('PLUGIN_ARC_CONV'), 'nsa, rgss2a, rgssad, wolf, xp3, ypf', 'arc_conv_r*.7z', 'arc_conv.exe', '', 'I Agree'], _
		[$thinstall, 'h4sh3m Virtual Apps Dependency Extractor', t('PLUGIN_THINSTALL'), 'exe (Thinstall)', 'Extractor.rar', '', '', 'h4sh3m'], _
		[$iscab, 'iscab', t('PLUGIN_ISCAB'), 'cab', 'iscab.exe;ISTools.dll', '', '', 0], _
		[$rgss3, 'RPGMaker Decrypter', t('PLUGIN_RPGMAKER'), 'rgss3a', 'RPGDecrypter.rar', '', '', 0], _
		[$unreal, 'Unreal Engine Resource Viewer', t('PLUGIN_UNREAL'), 'pak, u, uax, upk', 'umodel_win32.zip', 'umodel.exe|SDL2.dll', '', 0], _
		[$dcp, 'WinterMute Engine Unpacker', t('PLUGIN_WINTERMUTE'), 'dcp', $dcp, '', '', 0], _
		[$crage, 'Crass/Crage', t('PLUGIN_CRAGE'), 'exe (Livemaker)', 'Crass*.7z', '', '', 0], _
		[$mpq, 'MPQ Plugin', t('PLUGIN_MPQ'), 'mpq', 'wcx_mpq.zip', 'mpq.wcx|mpq.wcx64', '', 0], _
		[$ci, 'CreateInstall Extractor', t('PLUGIN_CI', CreateArray("ci-extractor.exe", "gea.dll", "gentee.dll")), 'exe (CreateInstall)', 'ci-extractor.exe;gea.dll;gentee.dll', '', '', 0], _
		[$dgca, 'DGCA', t('PLUGIN_DGCA'), 'dgca', 'dgca_v*.zip', 'dgcac.exe', '', 0], _
		[$bootimg, 'bootimg', t('PLUGIN_BOOTIMG'), 'boot.img', 'unpack_repack_kernel_redmi1s.zip', 'bootimg.exe', '', 0], _
		[$is5cab, 'is5comp', t('PLUGIN_IS5COMP'), 'cab (InstallShield)', 'i5comp21.rar', 'I5comp.exe|ZD50149.DLL|ZD51145.DLL', '', 0], _
		[$sim, 'Smart Install Maker unpacker', t('PLUGIN_SIM'), 'exe (Smart Install Maker)', 'sim_unpacker.7z', $sim, '', 0] _
	]

	Local Const $sSupportedFileTypes = t('PLUGIN_SUPPORTED_FILETYPES')
	Local $current = -1, $sWorkingdir = @WorkingDir, $aReturn[0]
	FileChangeDir(@UserProfileDir)

	$GUI_Plugins = GUICreate($name, 410, 167, -1, -1, -1, -1, $guimain)
	_GuiSetColor()
	$GUI_Plugins_List = GUICtrlCreateList("", 8, 8, 209, 149)
	GUICtrlSetData(-1, _ArrayToString($aPluginInfo, "|", -1, -1, "|", 1, 1))
	$GUI_Plugins_SelectClose = GUICtrlCreateButton(t('FINISH_BUT'), 320, 132, 83, 25)
	$GUI_Plugins_Download = GUICtrlCreateButton(t('TERM_DOWNLOAD'), 224, 132, 83, 25)
	GUICtrlSetState(-1, $GUI_DISABLE)
	$GUI_Plugins_Description = GUICtrlCreateEdit("", 224, 8, 177, 85, BitOR($ES_AUTOVSCROLL, $ES_WANTRETURN, $ES_READONLY, $ES_NOHIDESEL,$ES_MULTILINE))
	$GUI_Plugins_FileTypes = GUICtrlCreateEdit("", 224, 96, 177, 33, BitOR($ES_AUTOVSCROLL, $ES_WANTRETURN, $ES_READONLY, $ES_NOHIDESEL,$ES_MULTILINE))
	GUISetState(@SW_SHOW)

	Opt("GUIOnEventMode", 0)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $GUI_Plugins_List
				GUICtrlSetData($GUI_Plugins_FileTypes, $sSupportedFileTypes)
				Local $current = _GUICtrlListBox_GetCurSel($GUI_Plugins_List)
				If $current > -1 Then
					Local $return = _GUICtrlListBox_GetText($GUI_Plugins_List, $current)
					$current = _ArraySearch($aPluginInfo, $return)
					If @error Then ContinueLoop

					GUICtrlSetState($GUI_Plugins_Download, $GUI_DISABLE)
					GUICtrlSetData($GUI_Plugins_Description, $aPluginInfo[$current][2])
					GUICtrlSetData($GUI_Plugins_FileTypes, $sSupportedFileTypes & " " & $aPluginInfo[$current][3])

					If HasPlugin($aPluginInfo[$current][0], True) Then ; Installed
						GUICtrlSetData($GUI_Plugins_Download, t('TERM_INSTALLED'))
						GUICtrlSetData($GUI_Plugins_SelectClose, t('FINISH_BUT'))
					Else ; Not installed
						GUICtrlSetData($GUI_Plugins_Download, t('TERM_DOWNLOAD'))
						GUICtrlSetData($GUI_Plugins_SelectClose, t('SELECT_FILE'))
						If $aPluginInfo[$current][1] <> 'iscab' Then GUICtrlSetState($GUI_Plugins_Download, $GUI_ENABLE)
					EndIf
				EndIf
			Case $GUI_Plugins_SelectClose
				If $current = -1 Or HasPlugin($aPluginInfo[$current][0], True) Then ExitLoop

				Cout("Adding plugin " & $aPluginInfo[$current][1])
				$return = FileOpenDialog(t('OPEN_FILE'), @WorkingDir, $aPluginInfo[$current][1] & " (" & $aPluginInfo[$current][4] & ")", 4+1, "", $GUI_Plugins)
				If @error Then ContinueLoop
				GUICtrlSetState($GUI_Plugins_SelectClose, $GUI_DISABLE)
				Cout("Plugin file selected: " & $return)
				If $aPluginInfo[$current][6] = "" Then $aPluginInfo[$current][6] = $bindir

				; Check permissions
				If Not CanAccess($aPluginInfo[$current][6]) Then
					If IsAdmin() Then MsgBox($iTopmost + 64, $title, t('ACCESS_DENIED'))
					MsgBox($iTopmost + 64, $title, t('ELEVATION_REQUIRED'))
					ShellExecute($sUpdater, "/pluginst")
					terminate($STATUS_SILENT, '', '')
				EndIf

				; Determine filetype
				$ret = StringRight($return, 3)
				If $ret = ".7z" Or $ret = "rar" Or $ret = "zip" Then ; Unpack archive
					Local $command = $cmd & $7z & ($aPluginInfo[$current][5] == ''? ' x': ' e') & ($aPluginInfo[$current][7] == 0? '': ' -p"' & $aPluginInfo[$current][7] & '"')
					If $aPluginInfo[$current][5] <> "" Then ; Build include command for each file needed
						$aReturn = StringSplit($aPluginInfo[$current][5], "|", 2)
						For $sFile In $aReturn
							$command &= " -ir!" & $sFile
						Next
					EndIf
					$command &= ' -o"' & $aPluginInfo[$current][6] & '" "' & $return & '"'
					Cout("Plugin extraction command: " & $command)
					_Run($command, $aPluginInfo[$current][6], @SW_MINIMIZE)
				Else ; Copy files
					Local $aVars = StringSplit($aPluginInfo[$current][4], ";", 2), $success = True
					$aReturn = StringSplit($return, "|", 2)

					; Check if all files have been selected
					For $sFile In $aVars
						If _ArraySearch($aReturn, $sFile, 0, 0, 0, 1) < 0 Then
							MsgBox($iTopmost + 16, $title, t('PLUGIN_IMPORT_MISSINGFILES', CreateArray($aPluginInfo[$current][1], StringReplace($aPluginInfo[$current][4], ";", @CRLF))))
							$success = False
							ExitLoop
						EndIf
					Next

					; Copy files to \bin\
					If $success Then
						Local $size = UBound($aReturn)
						If $size = 1 Then ; Move single file directly
							Cout("Copying plugin file " & $aReturn[0] & " to " & $aPluginInfo[$current][6])
							FileCopy($aReturn[0], $aPluginInfo[$current][6], 1)
						Else ; Multiple files are returned as path|file1|filen
							For $i = 1 To $size - 1
								$aReturn[$i] = $aReturn[0] & "\" & $aReturn[$i]
								Cout("Copying plugin file " & $aReturn[$i] & " to " & $aPluginInfo[$current][6])
								FileCopy($aReturn[$i], $aPluginInfo[$current][6], 1)
							Next
						EndIf
					EndIf
				EndIf

				; Refresh GUI
				GUICtrlSetState($GUI_Plugins_SelectClose, $GUI_ENABLE)
				Local $aReturn = ["{UP}", "{DOWN}"]
				If $current = _GUICtrlListBox_GetTopIndex($GUI_Plugins_List) Then _ArrayReverse($aReturn)
				For $i = 0 To 1
					ControlSend($GUI_Plugins, "", $GUI_Plugins_List, $aReturn[$i])
				Next
			Case $GUI_Plugins_Download
				If $current = -1 Then ContinueLoop
				GUICtrlSetState($GUI_Plugins_Download, $GUI_DISABLE)
				Cout("Download clicked for plugin " & $aPluginInfo[$current][1])
				OpenURL($sGetLinkURL & $aPluginInfo[$current][1])
				GUICtrlSetState($GUI_Plugins_Download, $GUI_ENABLE)
		EndSwitch
	WEnd

	FileChangeDir($sWorkingdir)	; Reset working dir in case it was changed by FileOpenDialog
	GUIDelete($GUI_Plugins)
	Opt("GUIOnEventMode", 1)
EndFunc

; Open log directory
Func GUI_OpenLogDir()
	ShellExecute($logdir)
EndFunc

; Option to delete all log files
Func GUI_DeleteLogs()
	Cout("Deleting log files")
	Local $handle, $return, $i

	$handle = FileFindFirstFile($logdir & "*.log")
	If $handle == -1 Then Return

	While 1
		$return = FileFindNextFile($handle)
		If @error Then ExitLoop
		FileDelete($logdir & $return)
		$i += 1
	WEnd

	FileClose($handle)
	GUI_UpdateLogItem()

	Cout("Deleted a total of " & $i & " files")
EndFunc

; Update log directory size in menu entry after deleting log files
Func GUI_UpdateLogItem()
	If Not FileExists($logdir) Then DirCreate($logdir)
	Local $size = Round(DirGetSize($logdir) / 1024 / 1024, 2) & " MB"
	GUICtrlSetData($logitem, t('MENU_FILE_LOG_LABEL', $size))
EndFunc

; Display usage statistics
Func GUI_Stats()
	Local $aReturn = IniReadSection($prefs, "Statistics")
	If @error Or $aReturn[0][0] < 10 Then Return MsgBox($iTopmost + 48, $name, t('STATS_NO_DATA'))

	Local $ret = t('MENU_HELP_STATS_LABEL')
	$GUI_Stats = GUICreate($ret, 730, 434, 315, 209)
	$GUI_Stats_Status_Pie = GUICtrlCreatePic("", 8, 72, 209, 209)
	$GUI_Stats_Types_Pie = GUICtrlCreatePic("", 368, 72, 353, 353)
	$GUI_Stats_Types_Legend = GUICtrlCreatePic("", 8, 312, 337, 113)
	$GUI_Stats_Status_Legend = GUICtrlCreatePic("", 232, 72, 113, 209)
	GUICtrlCreateLabel($ret, 8, 8, 715, 33, $SS_CENTER)
	GUICtrlSetFont(-1, 18, 500, 0, "Arial")
	GUICtrlCreateLabel(t('STATS_HEADER_STATUS'), 8, 48, 212, 24, $SS_CENTER)
	GUICtrlSetFont(-1, 12, 300, 0, "Arial")
	GUICtrlCreateLabel(t('STATS_HEADER_TYPE'), 368, 48, 354, 24, $SS_CENTER)
	GUICtrlSetFont(-1, 12, 300, 0, "Arial")
	GUISetBkColor(0xFFFFFF)
	GUICtrlSetDefBkColor(0xFFFFFF)
	GUISetState(@SW_SHOW)

	Local $GUI_Stats_Types[0], $GUI_Stats_Status = [[0, t('STATS_STATUS_SUCCESS'), $COLOR_GREEN], [0, t('STATS_STATUS_FAILED'), $COLOR_RED], [0, t('STATS_STATUS_FILEINFO'), $COLOR_PURPLE], [0, t('STATS_STATUS_UNKNOWN'), $COLOR_GRAY]]

	For $i = 1 To $aReturn[0][0]
		Switch $aReturn[$i][0]
			Case $STATUS_FILEINFO
				$GUI_Stats_Status[2][0] += $aReturn[$i][1]
			Case $STATUS_NOTSUPPORTED, $STATUS_UNKNOWNEXE, $STATUS_UNKNOWNEXT
				$GUI_Stats_Status[3][0] += $aReturn[$i][1]
			Case $STATUS_FAILED, $STATUS_INVALIDDIR, $STATUS_INVALIDFILE, $STATUS_MISSINGDEF, $STATUS_MISSINGEXE, $STATUS_TIMEOUT
				$GUI_Stats_Status[1][0] += $aReturn[$i][1]
			Case $STATUS_SUCCESS, $STATUS_NOTPACKED, $STATUS_PASSWORD
				$GUI_Stats_Status[0][0] += $aReturn[$i][1]
			Case $STATUS_BATCH, $STATUS_SILENT, $STATUS_SYNTAX
				; Skip
			Case Else
				$iSize = UBound($GUI_Stats_Types)
				ReDim $GUI_Stats_Types[$iSize + 1][2]
				$GUI_Stats_Types[$iSize][0] = Number($aReturn[$i][1])
				$GUI_Stats_Types[$iSize][1] = $aReturn[$i][0]
		EndSwitch
	Next

	_ArraySort($GUI_Stats_Status, 1)
	_ArraySort($GUI_Stats_Types, 1)
	If UBound($GUI_Stats_Types) > 9 Then ReDim $GUI_Stats_Types[9][2]

	; Prepare values for the pie chart and setup GDI+ for both picture controls
	$GUI_Stats_Types_Handles = _Pie_PrepareValues($GUI_Stats_Types, $GUI_Stats_Types_Pie)
	$GUI_Stats_Types_Handles_Legend = _Pie_CreateContext($GUI_Stats_Types_Legend, 0)
	$GUI_Stats_Status_Handles = _Pie_PrepareValues($GUI_Stats_Status, $GUI_Stats_Status_Pie)
	$GUI_Stats_Status_Handles_Legend = _Pie_CreateContext($GUI_Stats_Status_Legend, 0)

	; Draw the initial pie chart and legend
	_Pie_Draw($GUI_Stats_Types_Handles, $GUI_Stats_Types, 1, 0)
	_Pie_Draw_Legend($GUI_Stats_Types_Handles_Legend, $GUI_Stats_Types)
	_Pie_Draw($GUI_Stats_Status_Handles, $GUI_Stats_Status, 1, 0)
	_Pie_Draw_Legend($GUI_Stats_Status_Handles_Legend, $GUI_Stats_Status, 20, 1, 8)

	Opt("GUIOnEventMode", 0)

	Local $GUI_Stats_Types_Rotation = 0, $GUI_Stats_Types_Aspect = 1, $GUI_Stats_Status_Rotation = 0, $GUI_Stats_Status_Aspect = 1
	While GUIGetMsg() <> $GUI_EVENT_CLOSE
		Sleep(10)

		_MouseOverRotation($GUI_Stats_Types_Rotation, $GUI_Stats_Types_Aspect, $GUI_Stats_Types_Handles, $GUI_Stats_Types, $GUI_Stats, $GUI_Stats_Types_Handles[4], 0, 0.5)
		_MouseOverRotation($GUI_Stats_Status_Rotation, $GUI_Stats_Status_Aspect, $GUI_Stats_Status_Handles, $GUI_Stats_Status, $GUI_Stats, $GUI_Stats_Status_Handles[4], 0, 0.7)
	WEnd

	Opt("GUIOnEventMode", 1)

	; Cleanup
	_Pie_Shutdown($GUI_Stats_Types_Handles, $GUI_Stats_Types, False)
	_Pie_Shutdown($GUI_Stats_Status_Handles, $GUI_Stats_Status, False)
	_Pie_Shutdown($GUI_Stats_Types_Handles_Legend, False)
	_Pie_Shutdown($GUI_Stats_Status_Handles_Legend)
	GUIDelete($GUI_Stats)
EndFunc

; Open password list file
Func GUI_Password()
	If Not FileExists($sPasswordFile) Then
		$handle = FileOpen($sPasswordFile, 1)
		FileClose($handle)
	EndIf
	ShellExecute($sPasswordFile)
EndFunc

; Open program directory
Func GUI_ProgDir()
	ShellExecute(@ScriptDir)
EndFunc

; Open configuration file
Func GUI_ConfigFile()
	ShellExecute($prefs)
EndFunc

; Create about GUI
Func GUI_About()
	Local Const $width = 437, $height = 285
	Cout("Creating about GUI")
	$About = GUICreate($title & " " & $codename, $width, $height, -1, -1, -1, $exStyle, $guimain)
	_GuiSetColor()
	GUICtrlCreateLabel($name, 16, 16, $width - 32, 52, $SS_CENTER)
	GUICtrlSetFont(-1, 25, 400, 0, "Arial")
	GUICtrlCreateLabel(t('ABOUT_VERSION', $version), 16, 72, $width - 32, 17, $SS_CENTER)
	GUICtrlCreateLabel(t('ABOUT_INFO_LABEL', CreateArray("Jared Breland <jbreland@legroom.net>", "uniextract@bioruebe.com", "TrIDLib (C) 2008 - 2011 Marco Pontello" & @CRLF & "<http://mark0.net/code-tridlib-e.html>", "GNU GPLv2")), 16, 104, $width - 32, -1, $SS_CENTER)
	GUICtrlCreateLabel($ID, 5, $height - 15, 175, 15)
	GUICtrlSetFont(-1, 8, 800, 0, "Arial")
	GUICtrlCreatePic(".\support\Icons\Bioruebe.jpg", $width - 89 - 10, $height - 55, 89, 50)
	$About_OK = GUICtrlCreateButton(t('OK_BUT'), $width / 2 - 45, $height - 50, 90, 25)
	GUISetState(@SW_SHOW)

	GUICtrlSetOnEvent($About_OK, "GUI_Close")
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_Close")
EndFunc

; Delete GUI if OK clicked or window closed
Func GUI_Close()
	GUIDelete()
EndFunc

; Launch Universal Extractor website if help menu item clicked
Func GUI_Website()
	Cout("Opening website")
	OpenURL($website)
EndFunc

; Launch Universal Extractor 2 website if help menu item clicked
Func GUI_Website2()
	Cout("Opening version 2 website")
	OpenURL($website2)
EndFunc

; Launch Universal Extractor 2 Github website if help menu item clicked
Func GUI_Website_Github()
	Cout("Opening Github website")
	OpenURL($websiteGithub)
EndFunc

; Exit if Cancel clicked or window closed
Func GUI_Exit()
	GUIGetPosition()
	terminate($STATUS_SILENT, '', '')
EndFunc

; Shows/hides cmd window when clicked on tray icon
Func Tray_ShowHide()
	If Not ProcessExists($run) Then Return
	If BitAND(WinGetState($runtitle), 2) Then
		WinSetState($runtitle, "", @SW_HIDE)
	Else
		WinSetState($runtitle, "", @SW_SHOW)
		WinActivate($runtitle)
	EndIf
EndFunc   ;==>Tray_ShowHide

; Change show statusbox option via tray
Func Tray_Statusbox()
	If TrayItemGetState($Tray_Statusbox) == 65 Then
		$NoBox = 0
		If $TBgui Then GUISetState(@SW_SHOW, $TBgui)
		TrayItemSetState($Tray_Statusbox, $TRAY_UNCHECKED)
	Else
		If $TBgui Then GUISetState(@SW_HIDE, $TBgui)
		$NoBox = 1
		TrayItemSetState($Tray_Statusbox, $TRAY_CHECKED)
	EndIf


	SavePref('nostatusbox', $NoBox)
EndFunc   ;==>Tray_Statusbox

; Exit and close helper binaries if necessary
Func Tray_Exit()
	Cout("Tray exit, helper PID: " & $run)
	KillHelper()

	If $guimain Then
		GUIGetPosition()
	Else
		CreateLog("trayexit")
	EndIf

	terminate($STATUS_SILENT, '', '')
EndFunc

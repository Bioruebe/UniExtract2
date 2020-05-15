#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=.\Support\Icons\uniextract_exe.ico
#AutoIt3Wrapper_Outfile=.\UniExtract.exe
#AutoIt3Wrapper_Res_Description=Universal Extractor
#AutoIt3Wrapper_Res_Fileversion=2.0.0
#AutoIt3Wrapper_Res_LegalCopyright=GNU General Public License v2
#AutoIt3Wrapper_Res_Field=Author|Jared Breland, Bioruebe
#AutoIt3Wrapper_Res_Field=Homepage|https://bioruebe.com/dev/uniextract/
#AutoIt3Wrapper_Res_Field=Timestamp|%date%
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
#include <FileConstants.au3>
#include <GDIPlus.au3>
#include <GUIConstantsEx.au3>
#include <GuiComboBox.au3>
#include <GuiComboBoxEx.au3>
#include <GUIEdit.au3>
#include <GuiListBox.au3>
#include <INet.au3>
#include <InetConstants.au3>
#include <Math.au3>
#include <Misc.au3>
#include <ProgressConstants.au3>
#include <SQLite.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <WinAPI.au3>
#include <WinAPIFiles.au3>
#include <WinAPIShellEx.au3>
#include <WinAPIShPath.au3>
#include <WindowsConstants.au3>
#include "HexDump.au3"
#include "Pie.au3"

Const $name = "Universal Extractor"
Const $sVersion = "2.0.0 RC 3"
Const $codename = '"in memoriam"'
Const $title = $name & " " & $sVersion
Const $website = "https://www.legroom.net/software/uniextract"
Const $website2 = "https://bioruebe.com/dev/uniextract"
Const $websiteGithub = "https://github.com/Bioruebe/UniExtract2"
Const $sDefaultUpdateURL = "https://update.bioruebe.com/uniextract/data/"
Const $sNightlyUpdateURL = "https://update.bioruebe.com/uniextract/nightly/"
Const $sLegacyUpdateURL = "https://update.bioruebe.com/uniextract/update2.php"
Const $sGetLinkURL = "https://update.bioruebe.com/uniextract/geturl.php?q="
Const $sSupportURL = "https://support.bioruebe.com/uniextract/upload.php"
Const $sStatsURL = "https://stat.bioruebe.com/uniextract/stats.php?a="
Const $sPrivacyPolicyURL = "https://bioruebe.com/dev/uniextract/privacypolicy"
Const $bindir = @ScriptDir & "\bin\"
Const $langdir = @ScriptDir & "\lang\"
Const $defdir = @ScriptDir & "\def\"
Const $docsdir = @ScriptDir & "\docs\"
Const $licensedir = $docsdir & "third-party\"
Const $sUpdater = @ScriptDir & '\UniExtractUpdater.exe'
Const $sUpdaterNoAdmin = @ScriptDir & '\UniExtractUpdater_NoAdmin.exe'
Const $sEnglishLangFile = @ScriptDir & '\English.ini'
Const $sLogoFile = @ScriptDir & "\support\Icons\uniextract.png"
Const $sUniExtract = @Compiled? @ScriptFullPath: StringReplace(@ScriptFullPath, "au3", "exe")
Const $sRegExAscii = "(?i)(?m)^[\w\Q @!§$%&/\()=?,.-:+~'²³{[]}*#ß°^âëöäüîêôûïáéíóúàèìòù\E]+$"
;~ Const $cmd = @ComSpec & ' /d /k ' ; Keep command prompt open for debugging
Const $cmd = (FileExists(@ComSpec)? @ComSpec: @WindowsDir & '\system32\cmd.exe') & ' /d /c '
Enum $OPTION_KEEP, $OPTION_DELETE, $OPTION_ASK, $OPTION_MOVE
Enum $PROMPT_ASK, $PROMPT_ALWAYS, $PROMPT_NEVER
Enum $RESULT_UNKNOWN, $RESULT_SUCCESS, $RESULT_FAILED, $RESULT_CANCELED, $RESULT_NOFREESPACE
Enum $UNICODE_NONE, $UNICODE_MOVE, $UNICODE_COPY
Enum $UPDATE_ALL, $UPDATE_HELPER, $UPDATE_MAIN
Enum $UPDATEMSG_PROMPT, $UPDATEMSG_SILENT, $UPDATEMSG_FOUND_ONLY
Const $FONT_ARIAL = "Arial", $COLOR_LINK = 0x000080
Const $PACKER_UPX = "UPX", $PACKER_ASPACK = "Aspack"
Const $HISTORY_FILE = "File History", $HISTORY_DIR = "Directory History"
Const $STATUS_SYNTAX = "syntax", $STATUS_FILEINFO = "fileinfo", $STATUS_UNKNOWNEXE = "unknownexe", $STATUS_UNKNOWNEXT = "unknownext", _
	  $STATUS_INVALIDFILE = "invalidfile", $STATUS_INVALIDDIR = "invaliddir", $STATUS_NOTPACKED = "notpacked", $STATUS_BATCH = "batch", _
	  $STATUS_NOTSUPPORTED = "notsupported", $STATUS_MISSINGEXE = "missingexe", $STATUS_TIMEOUT = "timeout", $STATUS_PASSWORD = "password", _
	  $STATUS_MISSINGDEF = "missingdef", $STATUS_MOVEFAILED = "movefailed", $STATUS_NOFREESPACE = "nofreespace", $STATUS_MISSINGPART = "missingpart", _
	  $STATUS_FAILED = "failed", $STATUS_SUCCESS = "success", $STATUS_SILENT = "silent"
Const $TYPE_7Z = "7z", $TYPE_ACE = "ace", $TYPE_ACTUAL = "actual", $TYPE_AI = "ai", $TYPE_ALZ = "alz", $TYPE_ARC_CONV = "arc_conv", _
	  $TYPE_AUDIO = "audio", $TYPE_BCM = "bcm", $TYPE_BOOTIMG = "bootimg", $TYPE_CAB = "cab", $TYPE_CHM = "chm", $TYPE_CI = "ci", _
	  $TYPE_CIC = "cic", $TYPE_CTAR = "ctar", $TYPE_DGCA = "dgca", $TYPE_DAA = "daa", $TYPE_DCP = "dcp", $TYPE_EI = "ei", $TYPE_ENIGMA = "enigma", _
	  $TYPE_FEAD = "fead", $TYPE_FREEARC = "freearc", $TYPE_FSB = "fsb", $TYPE_GARBRO = "garbro", $TYPE_GHOST = "ghost", $TYPE_HLP = "hlp", _
	  $TYPE_HOTFIX = "hotfix", $TYPE_INNO = "inno", $TYPE_ISCAB = "iscab", $TYPE_ISCRIPT = "installscript", $TYPE_ISEXE = "isexe", $TYPE_ISZ = "isz", _
	  $TYPE_KGB = "kgb", $TYPE_LZ = "lz", $TYPE_LZO = "lzo", $TYPE_LZX = "lzx", $TYPE_MHT = "mht", $TYPE_MOLE = "mole", $TYPE_MSCF = "mscf", _
	  $TYPE_MSI = "msi", $TYPE_MSM = "msm", $TYPE_MSP = "msp", $TYPE_NBH = "nbh", $TYPE_NSIS = "NSIS", $TYPE_PDF = "PDF", $TYPE_PEA = "pea", _
	  $TYPE_QBMS = "qbms", $TYPE_RAI = "rai", $TYPE_RAR = "rar", $TYPE_RGSS = "rgss", $TYPE_ROBO = "robo", $TYPE_RPA = "rpa", $TYPE_SFARK = "sfark", _
	  $TYPE_SGB = "sgb", $TYPE_SIM = "sim", $TYPE_SIS = "sis", $TYPE_SQLITE = "sqlite", $TYPE_SUPERDAT = "superdat", $TYPE_SWF = "swf", _
	  $TYPE_SWFEXE = "swfexe", $TYPE_TAR = "tar", $TYPE_THINSTALL = "thinstall", $TYPE_TTARCH = "ttarch", $TYPE_UHA = "uha", $TYPE_UIF = "uif", _
	  $TYPE_UNITY = "unity", $TYPE_UNREAL = "unreal", $TYPE_VIDEO = "video", $TYPE_VIDEO_CONVERT = "videoconv", $TYPE_VISIONAIRE3 = "visionaire3", _
	  $TYPE_VSSFX = "vssfx", $TYPE_VSSFX_PATH = "vssfxpath", $TYPE_WISE = "wise", $TYPE_WIX = "wix", $TYPE_WOLF = "wolf", $TYPE_ZIP = "zip", _
	  $TYPE_ZOO = "zoo", $TYPE_ZPAQ = "zpaq"
Const $aExtractionTypes = [$TYPE_7Z, $TYPE_ACE, $TYPE_ACTUAL, $TYPE_AI, $TYPE_ALZ, $TYPE_ARC_CONV, $TYPE_AUDIO, $TYPE_BCM, $TYPE_BOOTIMG, _
	  $TYPE_CAB, $TYPE_CHM, $TYPE_CI, $TYPE_CIC, $TYPE_CTAR, $TYPE_DGCA, $TYPE_DAA, $TYPE_DCP, $TYPE_EI, $TYPE_ENIGMA, $TYPE_FEAD, _
	  $TYPE_FREEARC, $TYPE_FSB, $TYPE_GARBRO, $TYPE_GHOST, $TYPE_HLP, $TYPE_HOTFIX, $TYPE_INNO, $TYPE_ISCAB, $TYPE_ISCRIPT, $TYPE_ISEXE, _
	  $TYPE_ISZ, $TYPE_KGB, $TYPE_LZ, $TYPE_LZO, $TYPE_LZX, $TYPE_MHT, $TYPE_MOLE, $TYPE_MSCF, $TYPE_MSI, $TYPE_MSM, $TYPE_MSP, $TYPE_NBH, _
	  $TYPE_NSIS, $TYPE_PDF, $TYPE_PEA, $TYPE_QBMS, $TYPE_RAI, $TYPE_RAR, $TYPE_RGSS, $TYPE_ROBO, $TYPE_RPA, $TYPE_SFARK, $TYPE_SGB, _
	  $TYPE_SIM, $TYPE_SIS, $TYPE_SQLITE, $TYPE_SUPERDAT, $TYPE_SWF, $TYPE_SWFEXE, $TYPE_TAR, $TYPE_THINSTALL, $TYPE_TTARCH, $TYPE_UHA, _
	  $TYPE_UIF, $TYPE_UNITY, $TYPE_UNREAL, $TYPE_VIDEO, $TYPE_VIDEO_CONVERT, $TYPE_VISIONAIRE3, $TYPE_VSSFX, $TYPE_VSSFX_PATH, $TYPE_WISE, _
	  $TYPE_WIX, $TYPE_WOLF, $TYPE_ZIP, $TYPE_ZOO, $TYPE_ZPAQ]


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
Global $bOptNoStatusBox = 0
Global $bHideStatusBoxIfFullscreen = 1
Global $bOptOpenOutDir = 0
Global $iDeleteOrigFile = $OPTION_KEEP
Global $Timeout = 60000 ; milliseconds
Global $updateinterval = 1 ; days
Global $lastupdate = "2010/12/05" ; last official version
Global $addassocenabled = 0
Global $addassocallusers = 0
Global $addassoc = ""
Global $ID = ""
Global $FB_ask = 0
Global $Log = 0
Global $CheckGame = 1
Global $bSendStats = 1
Global $bOptNightlyUpdates = 0
Global $iCleanup = $OPTION_MOVE
Global $KeepOutdir = 0
Global $KeepOpen = 0
Global $silentmode = 0
Global $extract = 1
Global $checkUnicode = 1
Global $bExtractVideo = 1
Global $bOptRememberGuiSizePosition = 0
Global $iTopmost = 0
Global $iOptGuiPosX = -1, $iOptGuiPosY = -1, $iOptGuiWidth = -1, $iOptGuiHeight = -1
Global $trayX = -1, $trayY = -1

; Global variables
Dim $file, $filename, $filenamefull, $filedir, $fileext, $sFileSize, $initoutdir, $outdir, $initdirsize
Dim $prompt, $return, $Output, $hMutex, $prefs = "", $sUpdateURL = $sDefaultUpdateURL, $eCustomPromptSetting = $PROMPT_ASK
Dim $Type, $win7, $silent, $iUnicodeMode = $UNICODE_NONE, $reg64 = "", $iOsArch = 32
Dim $sFullLog = "", $success = $RESULT_UNKNOWN, $isofile = 0, $sArcTypeOverride = 0, $sMethodSelectOverride = 0
Dim $innofailed, $arjfailed, $7zfailed, $zipfailed, $iefailed, $isofailed, $tridfailed, $gamefailed, $observerfailed
Dim $unpackfailed, $exefailed, $ttarchfailed
Dim $oldpath, $oldoutdir, $sUnicodeName, $createdir
Dim $guimain = False, $TBgui = 0, $exStyle = -1, $FS_GUI = False, $idTrayStatusExt, $BatchBut, $hProgress, $idProgress, $sComError = 0
Dim $isexe = False, $Message, $run = 0, $runtitle, $DeleteOrigFileOpt[3]
Dim $gaDropFiles[1], $aFiletype[0][2], $queueArray[0], $aTridDefinitions[0][0], $aFileDefinitions[0][0], $aGUIs[0]

; Check if OS is 64 bit version
If @OSArch == "X64" Or @OSArch == "IA64" Then
	Global $reg64 = 64
	Global $iOsArch = 64
	Global Const $archdir = $bindir & "x64\"
Else
	Global Const $archdir = $bindir & "x86\"
EndIf

; Extractors
Const $7z = Quote($archdir & '7z.exe', True)
Const $7zsplit = "7ZSplit.exe"
Const $ace = "acefile.exe"
Const $alz = "unalz.exe"
Const $arj = "arj.exe"
Const $aspack = "AspackDie.exe"
Const $bcm = Quote($archdir & "bcm.exe", True)
Const $cic = "cicdec.exe"
Const $daa = "daa2iso.exe"
Const $enigma = "EnigmaVBUnpacker.exe"
Const $exeinfope = Quote($bindir & "exeinfope.exe")
Const $filetool = Quote($bindir & "file\bin\file.exe", True)
Const $freearc = "unarc.exe"
Const $fsb = "fsbext.exe"
Const $garbro = Quote($bindir & "GARbro\GARbro.Console.exe", True)
Const $gcf = $archdir & "GCFScape.exe"
Const $hlp = "helpdeco.exe"
Const $inno = "innounp.exe"
Const $is6cab = "i6comp.exe"
Const $isxunp = "IsXunpack.exe"
Const $isz = "unisz.exe"
Const $kgb = "kgb\kgb2_console.exe"
Const $lit = "clit.exe"
Const $lz = "lzip.exe"
Const $lzo = "lzop.exe"
Const $lzx = "unlzx.exe"
Const $mole = "demoleition.exe"
Const $msi_msix = "MsiX.exe"
Const $msi_jsmsix = "jsMSIx.exe"
Const $msi_lessmsi = Quote($bindir & 'lessmsi\lessmsi.exe', True)
Const $nbh = "NBHextract.exe"
Const $pea = Quote($bindir & "pea.exe")
Const $pdfdetach = "pdfdetach.exe"
Const $pdftohtml = "pdftohtml.exe"
Const $pdftopng = "pdftopng.exe"
Const $pdftotext = "pdftotext.exe"
Const $peid = Quote($bindir & "peid.exe")
Const $quickbms = Quote($bindir & "quickbms.exe", True)
Const $rai = "RAIU.EXE"
Const $rar = Quote($archdir & "UnRAR.exe", True)
Const $rgss = "RgssDecrypter.exe"
Const $rpa = "unrpa.exe"
Const $sfark = "sfarkxtc.exe"
Const $sqlite = "sqlite3.exe"
Const $swf = "swfextract.exe"
Const $trid = "trid.exe"
Const $ttarch = "ttarchext.exe"
Const $uharc = "UNUHARC06.EXE"
Const $uharc04 = "UHARC04.EXE"
Const $uharc02 = "UHARC02.EXE"
Const $uif = "uif2iso.exe"
;~ Const $unity = ""
Const $unshield = "unshield.exe"
Const $upx = "upx.exe"
Const $visionaire3 = "VIS3Ext.exe"
Const $wise_ewise = "e_wise_w.exe"
Const $wise_wun = "wun.exe"
Const $wix = Quote($bindir & "dark\dark.exe", True)
Const $zip = "unzip.exe"
Const $zpaq = Quote($archdir & "zpaq.exe", True)
Const $zoo = "unzoo.exe"

; Exractor plugins
Const $bms = "BMS.bms"
Const $dbx = "dbxplug.wcx"
Const $gaup = "gaup_pro.wcx"
Const $ie = "InstExpl.wcx"
Const $iso = "Iso.wcx"
Const $msi_plug = "msi.wcx"
Const $observer = "TotalObserver.wcx"
Const $sis = "PDunSIS.wcx"

; Other
Const $tee = Quote($bindir & "mtee.exe")
Const $mediainfo = $bindir & "MediaInfo.dll"
Const $xor = "xor.exe"

; UniExtract plugins
Const $arc_conv = "arc_conv.exe"
Const $bootimg = "bootimg.exe"
Const $ci = "ci-extractor.exe"
Const $dcp = "dcp_unpacker.exe"
Const $dgca = "dgcac.exe"
Const $extsis = "extsis.exe"
Const $ffmpeg = Quote($archdir & "ffmpeg.exe", True)
Const $iscab = "iscab.exe"
Const $is5cab = "i5comp.exe"
Const $sim = "sim_unpacker.exe"
Const $thinstall = Quote($bindir & "Extractor.exe")
Const $unreal = "umodel.exe"
Const $wolf = "WolfDec.exe"

; Define registry keys
Global Const $reg = "HKCU" & $reg64 & "\Software\UniExtract"
Global Const $regcurrent = "HKCU" & $reg64 & "\Software\Classes\*\shell\"
Global Const $regall = "HKCR" & $reg64 & "\*\shell\"
Global $reguser = $regcurrent

; Define context menu commands
; On top to make remove via command line parameter possible
; Shell	| Commandline Parameter | Translation | MultiSelectModel
Global $CM_Shells[5][4] = [ _
	['uniextract_files', '', 'EXTRACT_FILES', "Single"], _
	['uniextract_here', ' .', 'EXTRACT_HERE', "Player"], _
	['uniextract_sub', ' /sub', 'EXTRACT_SUB', "Player"], _
	['uniextract_last', ' /last', 'EXTRACT_LAST', "Player"], _
	['uniextract_scan', ' /scan', 'SCAN_FILE', "Player"] _
]

; Make sure a language file exists
If Not FileExists($sEnglishLangFile) And Not FileExists($langdir) Then
	If MsgBox(48+1, $title, "No language file found." & @CRLF & @CRLF & "Click OK to download all missing files or Cancel to exit") == 1 Then
		CheckUpdate($UPDATEMSG_SILENT, False, $UPDATE_HELPER)
		Run($sUniExtract, @ScriptDir)
	EndIf
	Exit 99
EndIf

ReadPrefs()

Cout("Starting " & $name & " " & $sVersion)

ParseCommandLine()

; Create tray menu items
$Tray_Statusbox = TrayCreateItem(t('PREFS_HIDE_STATUS_LABEL'))
If $bOptNoStatusBox Then TrayItemSetState(-1, $TRAY_CHECKED)
TrayCreateItem("")
$Tray_Exit = TrayCreateItem(t('MENU_FILE_QUIT_LABEL'))

TrayItemSetOnEvent($Tray_Statusbox, "Tray_Statusbox")
TrayItemSetOnEvent($Tray_Exit, "Tray_Exit")
TraySetToolTip($name)
TraySetClick(8)

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

; If no file passed, display GUI to select file and set options
If $prompt Then
	CreateGUI()
	CheckUpdate($UPDATEMSG_FOUND_ONLY, True, $UPDATE_ALL, False)

	While 1
		If Not $guimain Then ExitLoop
		Sleep(100)
	WEnd
EndIf

; Prevent multiple instances to avoid errors
; Only necessary when extraction starts
; Do not do this in StartExtraction, the function can be called twice
$hMutex = _Singleton($name & " " & $sVersion, 1)
If $hMutex = 0 And $extract Then
	AddToBatch()
	terminate($STATUS_SILENT)
EndIf

StartExtraction()

; -------------------------- Begin Custom Functions ---------------------------

; Start extraction process
Func StartExtraction()
	Cout("------------------------------------------------------------")
	$iUnicodeMode = False

	If _IsDirectory($file) Then
		GUI_Batch_AddDirectory($file)
		terminate($STATUS_BATCH)
	EndIf

	FilenameParse($file)

	; Collect file information, for log/feedback only
	Local $iSize = Round(FileGetSize($file) / 1048576, 2)
	Global $sFileSize = $iSize < 1? Round(FileGetSize($file) / 1024, 2) & " KB": $iSize & " MB"
	Cout("File size: " & $sFileSize)

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
	$gamefailed = False
	$ttarchfailed = False
	$unpackfailed = False
	ReDim $aFiletype[0][2]

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

	; ExeInfo PE supports non-executables as well
	If Not $exefailed Then advexescan()

	; Display file information and terminate if scan only mode
	If Not $extract Then
		MediaFileScan($file)
		terminate($STATUS_FILEINFO, $filenamefull, $fileext)
	EndIf

	; Else perform additional extraction methods
	CheckIso()
	CheckGame()
	CheckTotalObserver()

	; Use file extension if signature not recognized
	CheckExt()

	check7z()

	; Cannot determine filetype, all checks failed - abort
	_DeleteTrayMessageBox()
	terminate($STATUS_UNKNOWNEXT, $file, $fileext & "; " & StringLeft($aFiletype[0][1], 45))
EndFunc

; Extract if exe file detected
Func IsExe()
	If $exefailed Then Return
	$isexe = True
	Cout("File seems to be executable")

	; Just for fun
	If $file = @ScriptFullPath Or $file = $sUpdater Or $file = $sUpdaterNoAdmin Then
		_FiletypeAdd($name, $name)
		terminate($STATUS_NOTPACKED, $file, $name, $name)
	EndIf

	; Check executable using Exeinfo PE
	advexescan()

	; Check executable using PEiD
	exescan($file, 'ext', $extract) ; Userdb is much faster, so do that first
	exescan($file, 'hard', $extract)

	; Make sure TrID doesn't call IsExe again
	$exefailed = True

	If Not $extract Then Return

	; Perform additional tests if necessary
	checkInno()
	checkIE()

	CheckGame()

	; Scan using TrID
	filescan($file)

	check7z()

	; Exit with unknown file type
	terminate($STATUS_UNKNOWNEXE, $file, StringLeft($aFiletype[0][1], 50))
EndFunc

; Parse filename
Func FilenameParse($f)
	$file = _PathFull($f)
	$filedir = StringLeft($f, StringInStr($f, '\', 0, -1) - 1)
	$filename = StringTrimLeft($f, StringInStr($f, '\', 0, -1))
	Local $iPos = StringInStr($filename, '.', 0, -1)
	If $iPos Then
		$fileext = StringTrimLeft($filename, $iPos)
		$filename = StringTrimRight($filename, StringLen($fileext) + 1)
		$initoutdir = $filedir & '\' & $filename
		If StringInStr($filename, ".") And  FileExists($initoutdir) And Not _IsDirectory($initoutdir) Then $initoutdir = $filedir & '\' & StringReplace($filename, ".", "_")
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
EndFunc

; Translate text
Func t($t, $aVars = 0, $lang = $language, $sDefault = 0)
	$return = IniRead($lang = 'English'? $sEnglishLangFile: $langdir & '\' & $lang & '.ini', 'UniExtract', $t, '')
	If $return == '' Then
		Cout("Translation not found for term " & $t)
		$return = IniRead($sEnglishLangFile, 'UniExtract', $t, '')
		If $return = '' Then
			Cout("Warning: term " & $t & " is not defined")
			Return $sDefault == 0? $t: $sDefault
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
	Local $iArgs = $cmdline[0]

	If $iArgs = 0 Then
		$prompt = 1
		Return
	EndIf

	$extract = True

	Cout("Command line parameters: " & _ArrayToString($cmdline, " ", 1))

	If _ArraySearch($cmdline, "/silent") > -1 Then $silentmode = True
	If _ArraySearch($cmdline, "/nolog") > -1 Then $Log = False
	If _ArraySearch($cmdline, "/nostats") > -1 Then $bSendStats = False

	If $cmdline[1] = "/prefs" Then
		GUI_Prefs()
		While $guiprefs
			Sleep(250)
		WEnd
		terminate($STATUS_SILENT)

	ElseIf $cmdline[1] = "/help" Or $cmdline[1] = "/?" Or $cmdline[1] = "-h" Or $cmdline[1] = "/h" Or $cmdline[1] = "-?" Or $cmdline[1] = "--help" Then
		terminate($STATUS_SYNTAX, "", $iArgs > 1)

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
		$file = $cmdline[1]
		If Not FileExists($file) Then terminate($STATUS_INVALIDFILE, $file)

		If $iArgs > 1 Then
			; Scan only
			If $cmdline[2] = "/scan" Then
				$extract = False
				$Log = False
			Else ; Outdir specified
				$outdir = $cmdline[2]
				; When executed from context menu, opening the outdir is not wanted
				$bOptOpenOutDir = 0
			EndIf

			; /type=arctype
			If $iArgs > 2 And StringLeft($cmdline[3], 5) = "/type" Then
				Local $aReturn = _FileListToArray($defdir, "*.ini", 1)
				Local $iPos = _ArraySearch($aReturn, "registry.ini")
				If $iPos > -1 Then _ArrayDelete($aReturn, $iPos)
				_ArrayTrim($aReturn, 4, 1, 1)
				_ArrayDelete($aReturn, 0)
				_ArrayConcatenate($aReturn, $aExtractionTypes)
				$aReturn = _ArrayUnique($aReturn, 0, 0, 0, 0)
				_ArraySort($aReturn)
;~ 				_ArrayDisplay($aFiles)

				$sArcTypeOverride = StringTrimLeft($cmdline[3], 6)
				If StringLen($sArcTypeOverride) > 0 Then
					If _ArrayBinarySearch($aReturn, $sArcTypeOverride) < 0 Then
						Local $tmp = StringRight($sArcTypeOverride, 1)
						If StringIsInt($tmp) Then $sMethodSelectOverride = ""
						While StringLen($sArcTypeOverride) > 0 And StringIsInt($tmp)
							$sMethodSelectOverride = $tmp & $sMethodSelectOverride
							$sArcTypeOverride = StringTrimRight($sArcTypeOverride, 1)

							$tmp = StringRight($sArcTypeOverride, 1)
						WEnd
					EndIf
				Else
					$sArcTypeOverride = GUI_MethodSelectList($aReturn, StringReplace(t('SCAN_FILE'), "&", ""), 'METHOD_EXTRACTOR_SELECT_LABEL')
					If $sArcTypeOverride < 0 Then terminate($STATUS_SILENT)
				EndIf
				Cout("Arctype override: " & $sArcTypeOverride)
				Cout("Method select override: " & $sMethodSelectOverride)
			EndIf
		Else
			$prompt = 1
		EndIf

		If _ArraySearch($cmdline, "/batch") > -1 Then
			AddToBatch()
			terminate($STATUS_SILENT)
		EndIf
	EndIf

	If _ArraySearch($cmdline, "/close") > -1 Then terminate($STATUS_SILENT)
EndFunc

; Read complete preferences
Func ReadPrefs()
	If IsAdmin() Then Cout("Warning: running as admin")

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

	LoadPref("language", $language, False)
	LoadPref("batchqueue", $batchQueue, False)
	If $batchQueue Then $batchQueue = _PathFull($batchQueue, $settingsdir)
	LoadPref("filescanlogfile", $fileScanLogFile, False)
	If Not @error Then $fileScanLogFile = _PathFull($fileScanLogFile, $settingsdir)
	LoadPref("batchenabled", $batchEnabled, 0)
	LoadPref("history", $history)
	LoadPref("appendext", $appendext)
	LoadPref("warnexecute", $warnexecute)
	LoadPref("nostatusbox", $bOptNoStatusBox)
	If Not $bOptNoStatusBox Then LoadPref("hidestatusboxiffullscreen", $bHideStatusBoxIfFullscreen)
	LoadPref("openfolderafterextr", $bOptOpenOutDir)
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
	LoadPref("storeguiposition", $bOptRememberGuiSizePosition)
	LoadPref("cleanup", $iCleanup)

	If $bOptRememberGuiSizePosition Then
		LoadPref("posx", $iOptGuiPosX, True, $iOptGuiPosX)
		LoadPref("posy", $iOptGuiPosY, True, $iOptGuiPosY)
		LoadPref("GuiWidth", $iOptGuiWidth, True, $iOptGuiWidth)
		LoadPref("GuiHeight", $iOptGuiHeight, True, $iOptGuiHeight)
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
	LoadPref("nightlyupdates", $bOptNightlyUpdates)
	If $bOptNightlyUpdates == 1 Then $sUpdateURL = $sNightlyUpdateURL

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
	SavePref('nostatusbox', $bOptNoStatusBox)
	SavePref("hidestatusboxiffullscreen", $bHideStatusBoxIfFullscreen)
	SavePref('openfolderafterextr', $bOptOpenOutDir)
	SavePref('deletesourcefile', $iDeleteOrigFile)
	SavePref('freespacecheck', $freeSpaceCheck)
	SavePref('unicodecheck', $checkUnicode)
	SavePref('feedbackprompt', $FB_ask)
	SavePref('log', $Log)
	SavePref('checkgame', $CheckGame)
	SavePref('sendstats', $bSendStats)
	SavePref("extractvideotrack", $bExtractVideo)
	SavePref('storeguiposition', $bOptRememberGuiSizePosition)
	SavePref('updateinterval', $updateinterval)
	SavePref('nightlyupdates', $bOptNightlyUpdates)
	SavePref('cleanup', $iCleanup)
	SavePref("topmost", Number($iTopmost > 0))
EndFunc

; Save single preference
Func SavePref($sName, $value)
	IniWrite($prefs, "UniExtract Preferences", $sName, $value)
	Cout("Saving: " & $sName & " = " & $value)
EndFunc

; Load single preference
Func LoadPref($sName, ByRef $value, $bInt = True, $iMin = -1)
	Local $return = IniRead($prefs, "UniExtract Preferences", $sName, "#Error#")
	If @error Or $return = "#Error#" Then
		Cout("Failed to read option " & $sName)
		SavePref($sName, $value)
		Return SetError(1, "", -1)
	EndIf

	If $bInt Then
		$value = _Max(Int($return), $iMin)
	Else
		$value = $return
	EndIf

	Cout("Option: " & $sName & " = " & $value)
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

; Return available languages as string, seperated by |
Func GetLanguageList()
	Local $aReturn = _FileListToArray($langdir, '*.ini', 1)
	If @error Then Local $aReturn[1]
	$aReturn[0] = 'English.ini'
	_ArraySort($aReturn)

	Local $return = ""
	For $i = 0 To UBound($aReturn) - 1
		$return &= StringTrimRight($aReturn[$i], 4) & '|'
	Next

	Return $return
EndFunc

; Return original file name, used to rename unicode files correctly
Func GetFileName()
	Return $iUnicodeMode? $sUnicodeName: $filename
EndFunc

; Scan file using TrID
Func filescan($f, $analyze = 1)
	If $tridfailed Then Return

	_CreateTrayMessageBox(t('SCANNING_FILE', "TrID"))
	Cout("Starting file scan using TrID")

	If $extract Then
		Local $return = ""
		$hDll = DllOpen($bindir & "TrIDLib.dll")
		Local $aReturn = DllCall($hDll, "int", "TrID_LoadDefsPack", "str", $bindir)
		If Not @error Then Cout($aReturn[0] & " definitions loaded")
		DllCall($hDll, "int", "TrID_SubmitFileA", "str", $f)
		DllCall($hDll, "int", "TrID_Analyze")

		$aReturn = DllCall($hDll, "int", "TrID_GetInfo", "int", 1, "int", 0, "str", $return)
		If $aReturn[0] = 0 Then
			Cout("Unknown filetype!")
			_DeleteTrayMessageBox()
			Return advfilescan($f)
		EndIf

		For $i = 1 To $aReturn[0]
			$aReturn = DllCall($hDll, "int", "TrID_GetInfo", "int", 2, "int", $i, "str", $return)
			_FiletypeAdd("TrID", $aReturn[3])
			If $analyze And $i < 4 Then tridcompare($aReturn[3])
		Next

		RenameWithTridExtension($hDll)

	Else ; Run TrID and fetch output to include additional information about the file type
		$return = StringSplit(FetchStdout($trid & ' "' & $f & '"' & ($appendext ? " -ce" : "") & ($analyze ? "" : " -v"), $filedir, @SW_HIDE, 0, True, False), @CRLF)
		Local $sFileType = ""
		For $i = 1 To UBound($return) - 1
			If StringInStr($return[$i], "%") Or (Not $analyze And (StringInStr($return[$i], "Related URL") Or StringInStr($return[$i], "Remarks"))) Then _
				$sFileType &= $return[$i] & @CRLF
		Next

		If $sFileType <> "" Then
			_FiletypeAdd("TrID", $sFileType)
			If $analyze Then tridcompare($sFileType)
		EndIf
	EndIf

	; Scan file using unix file tool
	advfilescan($f)

	$tridfailed = True
EndFunc

; Change file extension to the one TrID suggests if enabled in options
Func RenameWithTridExtension($hDll)
	If Not $appendext Then Return False

	Local $return = ""
	Local $aReturn = DllCall($hDll, "int", "TrID_GetInfo", "int", 3, "int", 1, "str", $return)
	$aReturn[3] = StringLower($aReturn[3])
	If $aReturn[3] == "" Then Return

	Local $ret = $filedir & "\" & $filename & "." & $aReturn[3]
	If $ret = $file Then Return False

	Cout("Changing file extension from ." & $fileext & " to ." & $aReturn[3])
	If Not _FileMove($file, $ret) Then Return False

	FilenameParse($file)
	Return True
EndFunc

; Additional file scan using unix file tool
Func advfilescan($f)
	Local $sFileType = ""

	_CreateTrayMessageBox(t('SCANNING_FILE', "Unix File Tool"))

	Cout("Start file scan using unix file tool")
	$sFileType = StringReplace(StringReplace(FetchStdout($filetool & ' "' & $f & '"', $filedir, @SW_HIDE), $f & "; ", ""), @CRLF, "")
	If $sFileType And $sFileType <> "data" Then _FiletypeAdd("Unix File Tool", $sFileType)

	_DeleteTrayMessageBox()

	If Not $extract Then
		; Text files are often misdetected, renaming them is not a good idea
		If $appendext And (StringInStr($sFileType, "text", 0) Or StringInStr($sFileType, "ASCII", 0)) Then $appendext = False
		Return
	EndIf

	filecompare($sFileType)
EndFunc

; Compare unix file tool's return to supported file types
Func filecompare($sFileType)
	Select
		Case StringInStr($sFileType, "7 zip archive data") Or StringInStr($sFileType, "7-zip archive data")
			extract($TYPE_7Z, '7-Zip ' & t('TERM_ARCHIVE'))
		Case StringInStr($sFileType, "RAR archive data")
			extract($TYPE_RAR, 'RAR ' & t('TERM_ARCHIVE'))
		Case StringInStr($sFileType, "lzip compressed data")
			extract($TYPE_LZ, "LZIP " & t('TERM_COMPRESSED') & " " & t('TERM_ARCHIVE'))
		Case StringInStr($sFileType, "Zip archive data") And Not StringInStr($sFileType, "7")
			extract($TYPE_ZIP, 'ZIP ' & t('TERM_ARCHIVE'))
		Case StringInStr($sFileType, "UHarc archive data", 0)
			extract($TYPE_UHA, 'UHARC ' & t('TERM_ARCHIVE'))
		Case StringInStr($sFileType, "Symbian installation file", 0)
			extract($TYPE_SIS, 'SymbianOS ' & t('TERM_INSTALLER'))
		Case StringInStr($sFileType, "Zoo archive data", 0)
			extract($TYPE_ZOO, 'ZOO ' & t('TERM_ARCHIVE'))
		Case StringInStr($sFileType, "MS Outlook Express DBX file", 0)
			extract($TYPE_QBMS, 'Outlook Express ' & t('TERM_ARCHIVE'), $dbx)
		Case StringInStr($sFileType, "bzip2 compressed data", 0)
			extract($TYPE_7Z, 'bzip2 ' & t('TERM_COMPRESSED'), "bz2")
		Case StringInStr($sFileType, "ASCII cpio archive", 0)
			extract($TYPE_7Z, 'CPIO ' & t('TERM_ARCHIVE'))
		Case StringInStr($sFileType, "gzip compressed", 0)
			extract($TYPE_7Z, 'gzip ' & t('TERM_COMPRESSED'), "gz")
		Case StringInStr($sFileType, "LZX compressed archive", 0)
			extract($TYPE_LZX, 'LZX ' & t('TERM_COMPRESSED'))
		Case StringInStr($sFileType, "ar archive", 0)
			extract($TYPE_7Z, 'AR ' & t('TERM_ARCHIVE'))
		Case StringInStr($sFileType, "ARJ archive", 0)
			extract($TYPE_7Z, 'ARJ ' & t('TERM_ARCHIVE'))
		Case StringInStr($sFileType, "POSIX tar archive", 0)
			extract($TYPE_TAR, 'Tar ' & t('TERM_ARCHIVE'))
		Case StringInStr($sFileType, "LHa", 0) And StringInStr($sFileType, "archive data", 0)
			extract($TYPE_7Z, 'LZH ' & t('TERM_COMPRESSED'))
		Case StringInStr($sFileType, "Macromedia Flash data", 0)
			extract($TYPE_SWF, 'Shockwave Flash ' & t('TERM_CONTAINER'))
		Case StringInStr($sFileType, "PowerISO Direct-Access-Archive", 0)
			extract($TYPE_DAA, 'DAA/GBI ' & t('TERM_IMAGE'))
		Case StringInStr($sFileType, "sfArk compressed Soundfont")
			extract($TYPE_SFARK, 'sfArk ' & t('TERM_COMPRESSED'))
		Case StringInStr($sFileType, "SQLite", 0)
			extract($TYPE_SQLITE, 'SQLite ' & t('TERM_FILE'))
		Case StringInStr($sFileType, "XZ compressed data")
			extract($TYPE_7Z, 'XZ ' & t('TERM_COMPRESSED'), "xz")
		Case StringInStr($sFileType, "MS Windows HtmlHelp Data")
			extract($TYPE_CHM, 'Compiled HTML ' & t('TERM_HELP'))
		Case StringInStr($sFileType, "MIME entity text")
			extract($TYPE_MHT, 'MHTML ' & t('TERM_ARCHIVE'))
		Case StringInStr($sFileType, "MoPaQ", 0)
			CheckTotalObserver('MPQ ' & t('TERM_ARCHIVE'))
		Case (StringInStr($sFileType, "RIFF", 0) And Not StringInStr($sFileType, "WAVE audio", 0)) Or _
			 StringInStr($sFileType, "MPEG v", 0) Or StringInStr($sFileType, "MPEG sequence") Or _
			 StringInStr($sFileType, "Microsoft ASF") Or StringInStr($sFileType, "GIF image") Or _
			 StringInStr($sFileType, "PNG image") Or StringInStr($sFileType, "MNG video")
			extract($TYPE_VIDEO, t('TERM_VIDEO') & ' ' & t('TERM_FILE'))
		Case StringInStr($sFileType, "AAC,")
			extract($TYPE_AUDIO, 'AAC ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))
		Case StringInStr($sFileType, "FLAC audio")
			extract($TYPE_AUDIO, 'FLAC ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))
		Case StringInStr($sFileType, "Ogg data, Vorbis audio")
			extract($TYPE_AUDIO, 'OGG Vorbis ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))
		Case StringInStr($sFileType, "Audio file", 0) Or StringInStr($sFileType, "Dolby Digital stream", 0)
			extract($TYPE_AUDIO, t('TERM_AUDIO') & ' ' & t('TERM_FILE'))
		Case StringInStr($sFileType, "ISO", 0) And StringInStr($sFileType, "filesystem", 0)
			CheckIso()
		Case Else
			UserDefCompare($aFileDefinitions, $sFileType, "File")
	EndSelect

	; Not extractable filetypes
	If StringInStr($sFileType, "CDF V2 document") Then Return

	If (StringInStr($sFileType, "text") And (StringInStr($sFileType, "CRLF") Or _
	  StringInStr($sFileType, "long lines") Or StringInStr($sFileType, "ASCII")) Or _
	  StringInStr($sFileType, "batch file") Or StringInStr($sFileType, "XML") Or _
	  StringInStr($sFileType, "HTML") Or StringInStr($sFileType, "source") Or _
	  StringInStr($sFileType, "Rich ")) Or _
	  StringInStr($sFileType, "image") Or StringInStr($sFileType, "icon resource") Or _
	 (StringInStr($sFileType, "bitmap") And Not StringInStr($sFileType, "MGR bitmap")) Or _
	  StringInStr($sFileType, "WAVE audio") Or StringInStr($sFileType, "boot sector;") Or _
	  StringInStr($sFileType, "shortcut") Or StringInStr($sFileType, "empty") Or _
	  StringInStr($sFileType, "directory") Or StringInStr($sFileType, "BitTorrent file") Or _
	  StringInStr($sFileType, "Standard MIDI data") Or StringInStr($sFileType, "MSVC program database") Then _
		terminate($STATUS_NOTPACKED, $file, $fileext, $sFileType)

	If StringInStr($sFileType, "MS-DOS executable") Then terminate($STATUS_NOTSUPPORTED, $file, $sFileType, $sFileType)
EndFunc

; Compare TrID's return to supported file types
Func tridcompare($sFileType)
	Cout("--> " & $sFileType)
	Select
		Case StringInStr($sFileType, "7-Zip compressed archive")
			extract($TYPE_7Z, '7-Zip ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "ACE compressed archive") Or StringInStr($sFileType, "ACE Self-Extracting Archive")
			extract($TYPE_ACE, 'ACE ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Android boot image")
			extract($TYPE_BOOTIMG, ' Android boot ' & t('TERM_IMAGE'))

		Case StringInStr($sFileType, "ALZip compressed archive")
			CheckAlz()

		Case StringInStr($sFileType, "LZIP compressed archive")
			extract($TYPE_LZ, "LZIP " & t('TERM_COMPRESSED'))

		Case StringInStr($sFileType, "FreeArc compressed archive")
			extract($TYPE_FREEARC, 'FreeArc ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Android Package")
			extract($TYPE_7Z, "Android " & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "ARJ compressed archive")
			extract($TYPE_7Z, 'ARJ ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "asar Electron Archive")
			HasPlugin($archdir & "\Formats\Asar." & $iOsArch & ".dll")
			extract($TYPE_7Z, 'ASAR ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "BCM compressed")
			extract($TYPE_BCM, 'BCM ' & t('TERM_COMPRESSED'))

		Case StringInStr($sFileType, "BZA compressed") Or StringInStr($sFileType, "GZA compressed")
			extract($TYPE_7Z, 'BGA ' & t('TERM_COMPRESSED'))

		Case StringInStr($sFileType, "bzip2 compressed archive")
			extract($TYPE_7Z, 'bzip2 ' & t('TERM_COMPRESSED'), "bz2")

		Case StringInStr($sFileType, "Microsoft Cabinet Archive") Or StringInStr($sFileType, "IncrediMail letter/ecard")
			extract($TYPE_CAB, 'Microsoft CAB ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Magic ISO Universal Image Format")
			extract($TYPE_UIF, 'UIF ' & t('TERM_IMAGE'))

		Case StringInStr($sFileType, "CDImage") Or StringInStr($sFileType, "null bytes")
			CheckIso()
			check7z()

		Case StringInStr($sFileType, "CPIO Archive")
			extract($TYPE_7Z, 'CPIO ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "PowerISO Direct-Access-Archive") Or StringInStr($sFileType, "gBurner Image")
			extract($TYPE_DAA, 'DAA/GBI ' & t('TERM_IMAGE'))

		Case StringInStr($sFileType, "Debian Linux Package")
			extract($TYPE_7Z, 'Debian ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "DGCA Digital G Codec Archiver")
			extract($TYPE_DGCA, 'DGCA ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Disk Image (Macintosh)")
			extract($TYPE_7Z, 'Macintosh ' & t('TERM_DISK') & ' ' & t('TERM_IMAGE'))

		Case StringInStr($sFileType, "FMOD Sample Bank Format")
			extract($TYPE_FSB, 'FMOD ' & t('TERM_CONTAINER'))

		Case StringInStr($sFileType, "Gentee Installer executable") Or StringInStr($sFileType, "Installer VISE executable") Or _
			 StringInStr($sFileType, "Setup Factory")
			checkIE()

		Case StringInStr($sFileType, "GZipped")
			extract($TYPE_7Z, 'gzip ' & t('TERM_COMPRESSED'), "gz")

		Case StringInStr($sFileType, "Windows Help File")
			extract($TYPE_HLP, 'Windows ' & t('TERM_HELP'), "", False, True)
			extract($TYPE_CHM, 'Compiled HTML ' & t('TERM_HELP'))

		Case StringInStr($sFileType, "VirtualBox Disk Image") Or StringInStr($sFileType, "Virtual HD image") Or _
			 StringInStr($sFileType, "VMware 4 Virtual Disk")
			extract($TYPE_7Z, t('TERM_DISK') & " " & t('TERM_IMAGE'))

		Case StringInStr($sFileType, "Generic PC disk image") Or StringInStr($sFileType, "WinImage compressed disk image")
			CheckIso()
			check7z()

		Case StringInStr($sFileType, "Reflexive Arcade Installer")
			extract($TYPE_RAI, "Reflexive Arcade " & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "Inno Setup installer")
			checkInno()

		Case StringInStr($sFileType, "InstallShield Z archive")
			If Not (StringLower($fileext) = "z") Then CreateRenamedCopy("z")
			CheckTotalObserver('InstallShield Z ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "InstallShield compressed archive")
			extract($TYPE_ISCAB, 'InstallShield CAB ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "ISo Zipped format")
			extract($TYPE_ISZ, 'Zipped ISO ' & t('TERM_IMAGE'))

		Case StringInStr($sFileType, "KGB archive")
			extract($TYPE_KGB, 'KGB ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "LHARC/LZARK compressed archive")
			extract($TYPE_7Z, 'LZH ' & t('TERM_COMPRESSED'))

		Case StringInStr($sFileType, "lzop compressed")
			extract($TYPE_LZO, 'LZO ' & t('TERM_COMPRESSED'))

		Case StringInStr($sFileType, "LZX Amiga compressed archive")
			extract($TYPE_LZX, 'LZX ' & t('TERM_COMPRESSED'))

		Case StringInStr($sFileType, "MIME HTML archive format") Or StringInStr($sFileType, "E-Mail message")
			extract($TYPE_MHT, 'MHTML ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Microsoft Windows Installer merge module")
			extract($TYPE_MSM, 'Windows Installer (MSM) ' & t('TERM_MERGE_MODULE'))

		Case StringInStr($sFileType, "Microsoft Windows Installer") Or StringInStr($sFileType, "Generic OLE2 / Multistream Compound File")
			extract($TYPE_MSI, 'Windows Installer (MSI) ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "Windows Installer Patch")
			extract($TYPE_MSP, 'Windows Installer (MSP) ' & t('TERM_PATCH'))

		Case StringInStr($sFileType, "MPQ Archive - Blizzard game data")
			CheckTotalObserver('MPQ ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "HTC NBH ROM Image")
			extract($TYPE_NBH, 'NBH ' & t('TERM_IMAGE'))

		Case StringInStr($sFileType, "Outlook Express E-mail folder")
			extract($TYPE_QBMS, 'Outlook Express ' & t('TERM_ARCHIVE'), $dbx)

		Case StringInStr($sFileType, "Portable Document Format")
			extract($TYPE_PDF, 'PDF ' & t('TERM_FILE'))

		Case StringInStr($sFileType, "PEA compressed archive")
			extract($TYPE_PEA, 'Pea ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "RAR compressed archive")
			extract($TYPE_RAR, 'RAR ' & t('TERM_ARCHIVE'))

		; Game Archives

		Case StringInStr($sFileType, "BGI (Buriko General Interpreter) engine")
			CheckGarbro()

		Case StringInStr($sFileType, "Broken Age package")
			CheckGame(False)

		Case StringInStr($sFileType, "Bruns Engine encrypted") Or StringInStr($sFileType, "Ultramarine 3 encrypted audio file")
			CheckGarbro()

		Case StringInStr($sFileType, "ClsFileLink") Or StringInStr($sFileType, "ERISA archive file")
			CheckGarbro()
			extract($TYPE_ARC_CONV, "ERISA archive file" & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "KiriKiri Adventure Game System Package")
			CheckGarbro()
			extract($TYPE_ARC_CONV, 'KiriKiri Adventure Game System ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "Livemaker Engine main game executable")
			CheckGarbro()

		Case StringInStr($sFileType, "NScripter archive, version 1")
			CheckGarbro()

		Case StringInStr($sFileType, "Ren'Py data file")
			extract($TYPE_RPA, "Ren'Py " & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "RPG Maker") And Not StringInStr($sFileType, "MV encrypted")
			extract($TYPE_RGSS, "RPG Maker " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Smile Game Builder")
			extract($TYPE_SGB, "Smile Game Builder " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Telltale Games ressource archive")
			extract($TYPE_TTARCH, "Telltale " & t('TERM_GAME') & t('TERM_ARCHIVE'))

;~ 		Case StringInStr($sFileType, "Unity Engine Asset file")
;~ 			extract($TYPE_UNITY, 'Unity Engine Asset ' & t('TERM_FILE'))

		Case StringInStr($sFileType, "Unreal Package")
			extract($TYPE_UNREAL, 'Unreal Engine ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "Valve package") Or StringInStr($sFileType, "WAD3 game data") Or _
			 StringInStr($sFileType, "Valve Source map") Or StringInStr($sFileType, "Valve Source BSP")
			CheckTotalObserver('Valve ' & StringUpper($fileext) & " " & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "Visionaire Studio V3 archive")
			extract($TYPE_VISIONAIRE3, "Visionaire Studio V3 " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Wintermute Engine data")
			extract($TYPE_DCP, 'Wintermute Engine ' & t('TERM_GAME') & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "Wolf RPG Editor")
			CheckGarbro()
			extract($TYPE_WOLF, "Wolf RPG Editor " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "YU-RIS Script Engine")
			CheckGarbro()
			extract($TYPE_ARC_CONV, "YU-RIS Script Engine " & t('TERM_GAME') & t('TERM_ARCHIVE'))


		Case StringInStr($sFileType, "RPM Package")
			extract($TYPE_7Z, 'RPM ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "sfArk compressed SoundFont")
			extract($TYPE_SFARK, 'sfArk ' & t('TERM_COMPRESSED'))

		Case StringInStr($sFileType, "EPOC Installation package")
			extract($TYPE_SIS, 'SymbianOS ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "Macromedia Flash Player")
			extract($TYPE_SWF, 'Shockwave Flash ' & t('TERM_CONTAINER'))

		Case StringInStr($sFileType, "TAR - Tape ARchive")
			extract($TYPE_TAR, 'Tar ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "UHARC compressed archive")
			extract($TYPE_UHA, 'UHARC ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Base64 Encoded file")
			extract("uu", 'Base64 ' & t('TERM_ENCODED'))

		Case StringInStr($sFileType, "Quoted-Printable Encoded file")
			extract("uu", 'Quoted-Printable ' & t('TERM_ENCODED'))

		Case StringInStr($sFileType, "UUencoded file") Or StringInStr($sFileType, "XXencoded file")
			extract("uu", 'UUencoded ' & t('TERM_ENCODED'))

		Case StringInStr($sFileType, "yEnc Encoded file")
			extract("uu", 'yEnc ' & t('TERM_ENCODED'))

		Case StringInStr($sFileType, "Windows Imaging Format")
			extract($TYPE_7Z, 'WIM ' & t('TERM_IMAGE'))

		Case StringInStr($sFileType, "Wise Installer Executable")
			extract($TYPE_WISE, 'Wise Installer ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "UNIX Compressed")
			extract($TYPE_7Z, 'LZW ' & t('TERM_COMPRESSED'), "Z")

		Case StringInStr($sFileType, "xz compressed container")
			extract($TYPE_7Z, 'XZ ' & t('TERM_COMPRESSED'), "xz")

		Case StringInStr($sFileType, "ZIP compressed archive") Or StringInStr($sFileType, "Winzip Win32 self-extracting archive")
			extract($TYPE_ZIP, 'ZIP ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "ZOO compressed archive")
			extract($TYPE_ZOO, 'ZOO ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "ZPAQ compressed archive")
			extract($TYPE_ZPAQ, 'ZPAQ ' & t('TERM_ARCHIVE'))

		; Forced to bottom of list due to false positives
		Case StringInStr($sFileType, "LZMA compressed archive") Or StringInStr($sFileType, "Windows Thumbnail Database")
			check7z()

		Case StringInStr($sFileType, "Enigma Virtual Box virtualized executable")
			extract($TYPE_ENIGMA, 'Enigma Virtual Box ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "InstallShield setup")
			extract($TYPE_ISEXE, 'InstallShield ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "audio") Or StringInStr($sFileType, "FLAC lossless")
			extract($TYPE_AUDIO, t('TERM_AUDIO') & ' ' & t('TERM_FILE'))

		Case StringInStr($sFileType, "Smacker movie/video") Or StringInStr($sFileType, "Bink video")
			extract($TYPE_VIDEO_CONVERT, t('TERM_VIDEO') & ' ' & t('TERM_FILE'))

		Case StringInStr($sFileType, "Video") Or StringInStr($sFileType, "QuickTime Movie") Or _
			 StringInStr($sFileType, "Matroska") Or StringInStr($sFileType, "Material Exchange Format") Or _
			 StringInStr($sFileType, "Windows Media (generic)") Or StringInStr($sFileType, "GIF animated") Or _
			 StringInStr($sFileType, "MPEG-2 Transport Stream")
			extract($TYPE_VIDEO, t('TERM_VIDEO') & ' ' & t('TERM_FILE'))

		; Not packed
		Case StringInStr($sFileType, "null bytes") Or StringInStr($sFileType, "phpMyAdmin SQL dump") Or _
			 StringInStr($sFileType, "ELF Executable and Linkable format") Or StringInStr($sFileType, "Generic XML") Or _
			 StringInStr($sFileType, "Microsoft Program DataBase") Or StringInStr($sFileType, "Windows Minidump") Or _
			 StringInStr($sFileType, "Windows Shortcut") Or StringInStr($sFileType, "JPEG bitmap") Or StringInStr($sFileType, "Windows Registry Data") Or _
			 StringInStr($sFileType, "X509 Certificate")
			terminate($STATUS_NOTPACKED, $file, $fileext, $sFileType)

		; Not supported
		Case StringInStr($sFileType, "Long Range ZIP") Or StringInStr($sFileType, "Kremlin Encrypted File") Or _
			 StringInStr($sFileType, "Foxit Reader Add-on")
			terminate($STATUS_NOTSUPPORTED, $file, $fileext, $sFileType)

		; Check for .exe file, only when fileext not .exe
		Case Not $isexe And (StringInStr($sFileType, 'Executable') Or StringInStr($sFileType, '(.EXE)', 1))
			IsExe()

		Case Else
			UserDefCompare($aTridDefinitions, $sFileType, "Trid")
	EndSelect
EndFunc

; Compare file type with definitions stored in def/registry.ini
Func UserDefCompare(ByRef $aDefinitions, $sFileType, $sSection)
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
			If (StringInStr($sFileType, $aDefinitions[$i][1])) Then extract($aDefinitions[$i][0])
		Next
	Next
EndFunc

; Scan .exe file using PEiD
Func exescan($f, $scantype, $analyze = 1)
	Local $sFileType = "", $bHasRegKey = True
	Local Const $key = "HKCU\Software\PEiD"

	Cout("Start file scan using PEiD (" & $scantype & ")")
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
	While ($sFileType = "") Or ($sFileType = "Scanning...")
		Sleep(100)
		$sFileType = ControlGetText("PEiD v", "", "Edit2")
		$TimerDiff = TimerDiff($TimerStart)
		If $TimerDiff > $Timeout Then ExitLoop
	WEnd
	WinClose("PEiD v")

	_FiletypeAdd("PEiD (" & $scantype & ")", $sFileType)
	Cout($sFileType)

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
	If Not $analyze Then Return $sFileType

	; Match known patterns
	Select
		; ExeInfo cannot detect big files, so PEiD is used as a fallback here
		Case StringInStr($sFileType, "Enigma Virtual Box")
			extract($TYPE_ENIGMA, 'Enigma Virtual Box ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "ARJ SFX", 0)
			extract($TYPE_7Z, t('TERM_SFX') & ' ARJ ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Gentee Installer", 0)
			checkIE()

		Case StringInStr($sFileType, "Inno Setup", 0)
			checkInno()

		Case StringInStr($sFileType, "Installer VISE", 0)
			extract("ie", 'Installer VISE ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "KGB SFX", 0)
			extract($TYPE_KGB, t('TERM_SFX') & ' KGB ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "Microsoft Visual C++ 7.0", 0) And StringInStr($sFileType, "Custom", 0) And Not StringInStr($sFileType, "Hotfix", 0)
			extract($TYPE_VSSFX, 'Visual C++ ' & t('TERM_SFX') & ' ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "Microsoft Visual C++ 6.0", 0) And StringInStr($sFileType, "Custom", 0)
			extract($TYPE_VSSFX_PATH, 'Visual C++ ' & t('TERM_SFX') & '' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "Nullsoft PiMP SFX", 0)
			checkNSIS()

		Case StringInStr($sFileType, "PEtite", 1)
			If Not checkArj() Then extract($TYPE_ACE, t('TERM_SFX') & ' ACE ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "RAR SFX", 0)
			extract($TYPE_RAR, t('TERM_SFX') & ' RAR ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Reflexive Arcade Installer", 0)
			extract($TYPE_INNO, 'Reflexive Arcade ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "RoboForm Installer", 0)
			extract($TYPE_ROBO, 'RoboForm ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "Setup Factory 6.x", 0)
			extract("ie", 'Setup Factory ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "SPx Method", 0) Or StringInStr($sFileType, "CAB SFX", 0)
			extract($TYPE_CAB, t('TERM_SFX') & ' Microsoft CAB ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "SuperDAT", 0)
			extract($TYPE_SUPERDAT, 'McAfee SuperDAT ' & t('TERM_UPDATER'))

		Case StringInStr($sFileType, "Wise", 0) Or StringInStr($sFileType, "PEncrypt 4.0", 0)
			extract($TYPE_WISE, 'Wise Installer ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "ZIP SFX", 0)
			extract($TYPE_ZIP, t('TERM_SFX') & ' ZIP ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "upx", 0)
			unpack($PACKER_UPX)

		Case StringInStr($sFileType, "aspack", 0)
			unpack($PACKER_ASPACK)

		Case StringInStr($sFileType, "Unable to open file", 0)
			$isexe = False

	EndSelect
EndFunc

; Scan file using Exeinfo PE
Func advexescan($bUseCmd = $extract)
	Local $sFileType = ""

	Cout("Start file scan using Exeinfo PE")
	_CreateTrayMessageBox(t('SCANNING_EXE', "Exeinfo PE"))

	; Analyze file
	If $bUseCmd Then ; Use log command line for best speed
		Local Const $LogFile = $logdir & "exeinfo.log"
		RunWait($exeinfope & ' "' & $file & '*" /sx /log:"' & $LogFile & '"', $bindir, @SW_HIDE)
		$sFileType = _FileRead($LogFile, True)
		If StringInStr($sFileType, "File corrupted or Buffer Error") Or StringIsSpace($sFileType) Then Return advexescan(False)
	Else ; In scan only mode run and read GUI fields to get additional information on how to extract
		$aReturn = OpenExeInfo()
		$TimerStart = TimerInit()

		While $sFileType = "" Or StringInStr($sFileType, "File too big") Or StringInStr($sFileType, "Antivirus may slow") Or _
			  StringInStr($sFileType, "File corrupted or Buffer Error")
			Sleep(200)
			$sFileType = ControlGetText($aReturn[0], "", "TEdit6")
			$TimerDiff = TimerDiff($TimerStart)
			If $TimerDiff > $Timeout Then ExitLoop
		WEnd

		$sFileType &= @CRLF & @CRLF & ControlGetText($aReturn[0], "", "TEdit5")

		CloseExeInfo($aReturn)
	EndIf

	_DeleteTrayMessageBox()

	If StringInStr($sFileType, $filenamefull) Then $sFileType = StringTrimLeft(StringStripWS(StringReplace($sFileType, $filenamefull, ""), 1), 2)

	; Return if file is too big
	If StringInStr($sFileType, "Skipped") Then Return

	; Do not display 'unknown file type' scan result in scan only mode
	If Not $extract And StringInStr($sFileType, "file is not EXE or DLL") Then Return

	_FiletypeAdd("Exeinfo PE", $sFileType)

	; Return filetype without matching if specified
	If Not $extract Then Return $sFileType

	; Match known patterns
	Select
		Case StringInStr($sFileType, "Inno Setup")
			checkInno()

		Case StringInStr($sFileType, "WinAce / SFX Factory")
			extract($TYPE_ACE, t('TERM_SFX') & ' ACE ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Actual Installer")
			extract($TYPE_ACTUAL, 'Actual Installer ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "Advanced Installer")
			extract($TYPE_AI, 'Advanced Installer ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "FreeArc")
			extract($TYPE_FREEARC, 'FreeArc ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "CreateInstall")
			extract($TYPE_CI, 'CreateInstall ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "Excelsior Installer")
			extract($TYPE_EI, 'Excelsior Installer ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "Ghost Installer Studio")
			extract($TYPE_GHOST, 'Ghost Installer Studio ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "Gentee Installer") Or StringInStr($sFileType, "Installer VISE") Or _
			 StringInStr($sFileType, "Setup Factory 6.x")
			checkIE()

		Case StringInStr($sFileType, "install4j")
			BmsExtract("install4j")

		; Needs to be before InstallShield
		Case StringInStr($sFileType, "InstallAware")
			extract($TYPE_7Z, 'InstallAware ' & t('TERM_INSTALLER') & ' ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "Install Creator/Pro")
			extract($TYPE_CIC, 'Clickteam Install Creator ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "InstallScript Setup Launcher")
			extract($TYPE_ISCRIPT, 'InstallScript ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "InstallShield")
			extract($TYPE_ISEXE, 'InstallShield ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "KGB SFX")
			extract($TYPE_KGB, t('TERM_SFX') & ' KGB ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "Microsoft Visual C++ 7.0") And StringInStr($sFileType, "Custom") And Not StringInStr($sFileType, "Hotfix")
			extract($TYPE_VSSFX, 'Visual C++ ' & t('TERM_SFX') & ' ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "Microsoft Visual C++ 6.0") And StringInStr($sFileType, "Custom")
			extract($TYPE_VSSFX_PATH, 'Visual C++ ' & t('TERM_SFX') & '' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "www.molebox.com")
			extract($TYPE_MOLE, 'Mole Box ' & t('TERM_CONTAINER'))

		Case StringInStr($sFileType, "Netopsystems AG INSTALLER FEAD")
			extract($TYPE_FEAD, 'Netopsystems FEAD ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "Nullsoft")
			checkNSIS()

		Case StringInStr($sFileType, "RAR SFX")
			extract($TYPE_RAR, t('TERM_SFX') & ' RAR ' & t('TERM_ARCHIVE'));

		Case StringInStr($sFileType, "Reflexive Arcade Installer")
			extract($TYPE_INNO, 'Reflexive Arcade ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "RoboForm Installer")
			extract($TYPE_ROBO, 'RoboForm ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "WiX Installer")
			extract($TYPE_WIX, 'WiX ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "Microsoft SFX CAB") And StringInStr($sFileType, "rename file *.exe as *.cab")
			CreateRenamedCopy("cab")
			check7z()

		Case StringInStr($sFileType, "Smart Install Maker")
			extract($TYPE_SIM, 'Smart Install Maker ' & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "SPx Method") Or StringInStr($sFileType, "Microsoft SFX CAB")
			extract($TYPE_CAB, t('TERM_SFX') & ' Microsoft CAB ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Overlay :  SWF flash object ver", 0)
			extract($TYPE_SWFEXE, 'Shockwave Flash ' & t('TERM_CONTAINER'))

		Case StringInStr($sFileType, "VMware ThinApp") Or StringInStr($sFileType, "Thinstall") Or StringInStr($sFileType, "ThinyApp Packager", 0)
			extract($TYPE_THINSTALL, "ThinApp/Thinstall" & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Wise") Or StringInStr($sFileType, "PEncrypt 4.0")
			extract($TYPE_WISE, 'Wise Installer ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, "ZIP SFX") Or (StringInStr($sFileType, "WinZip") And StringInStr($sFileType, "Sfx ver"))
			extract($TYPE_ZIP, t('TERM_SFX') & ' ZIP ' & t('TERM_ARCHIVE'))

		Case StringInStr($sFileType, "Enigma Virtual Box")
			extract($TYPE_ENIGMA, 'Enigma Virtual Box ' & t('TERM_PACKAGE'))

		Case StringInStr($sFileType, ".dmg  Mac OS")
			extract($TYPE_7Z, 'DMG ' & t('TERM_IMAGE'))

		Case StringInStr($sFileType, "MSCF Cab file detected") Or StringInStr($sFileType, "VirtualBox Installer")
			extract($TYPE_MSCF, "MSCF " & t('TERM_INSTALLER'))

		Case StringInStr($sFileType, "aspack")
			unpack($PACKER_ASPACK)

		; Not supported
		Case StringInStr($sFileType, "Astrum InstallWizard") Or StringInStr($sFileType, "clickteam") Or _
			 StringInStr($sFileType, "BitRock InstallBuilder") Or StringInStr($sFileType, "NE <- Windows 16bit")
			terminate($STATUS_NOTSUPPORTED, $file, $sFileType, $sFileType)

		; Terminate if file cannot be unpacked
		Case (StringInStr($sFileType, "Not packed") And Not StringInStr($sFileType, "Microsoft Visual C++")) Or _
			  StringInStr($sFileType, "ELF executable") Or StringInStr($sFileType, "Microsoft Visual C# / Basic.NET") Or _
			  StringInStr($sFileType, "Autoit") Or StringInStr($sFileType, "LE <- Linear Executable") Or _
			  StringInStr($sFileType, "NOT EXE - Empty file") Or StringInStr($sFileType, "Native - System driver") Or _
			  StringInStr($sFileType, "Denuvo protector") Or StringInStr($sFileType, "Kaspersky AV Pack") Or _
			  StringInStr($sFileType, "TASM / MASM / FASM - assembler")
			terminate($STATUS_NOTPACKED, $file, $sFileType, $sFileType)

		; Needs to be at the end, otherwise files might not be recognized
		Case StringInStr($sFileType, "upx")
			unpack($PACKER_UPX)
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
	WinWait($aReturn[0], "", $Timeout)
	WinSetState($aReturn[0], "", @SW_HIDE)

	Return $aReturn
EndFunc

; Use ExeInfo PE's rip feature
Func RipExeInfo($f, $tempoutdir, $command)
	DirCreate($tempoutdir)
	$tmp = $tempoutdir & $filenamefull
	Cout("Moving file to " & $tmp)
	_FileMove($file, $tmp)

	$aReturn = OpenExeInfo($tmp)

	WinWait($aReturn[0], "", $Timeout)
	MouseMove(0, 0, 0)
	ControlClick($aReturn[0], "", "[CLASS:TBitBtn; INSTANCE:16]")
	ControlSend($aReturn[0], "", "[CLASS:TBitBtn; INSTANCE:16]", $command & "{ENTER}")

	Local $hWnd = WinWait("[CLASS:TSViewer]", "", $Timeout)
	Local $hControl = ControlGetHandle($hWnd, "", "TListBox1")

	$TimerStart = TimerInit()
	$return = -1

	While $return < 0
		Sleep(200)
		$return = _GUICtrlListBox_FindString($hControl, "--- End of file ---", True)
		If $return < 0 Then $return = _GUICtrlListBox_FindString($hControl, "-- End of file --", True)
		If TimerDiff($TimerStart) > $Timeout Then ExitLoop
	WEnd

	Local $success = _GUICtrlListBox_FindString($hControl, "--- Not found , sorry ---", True) == -1

	CloseExeInfo($aReturn)

	Cout("Moving file back")
	_FileMove($tmp, $filedir & "\")

	Return $success
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
	Local $sFileType = ""
	Cout("Start filescan using MediaInfo")
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
	$aReturn = StringSplit($return[0], @CRLF, 2)
	If UBound($aReturn) < 10 Then Return _DeleteTrayMessageBox()

	; Format returned string to align in message box
	For $i in $aReturn
		Local $aSplit = StringSplit($i, " : ", 2+1)

		If @error Then
			If Not StringIsSpace($i) Then $sFileType &= @CRLF & "[" & $i & "]" & @CRLF
			ContinueLoop
		EndIf

		$sType = StringStripWS($aSplit[0], 4+2+1)
		If $sType == "Complete name" Then ContinueLoop

		$sFileType &= StringFormat("%-24s%s\r\n", $sType, StringStripWS($aSplit[1], 4+2+1))
	Next

	_FiletypeAdd("MediaInfo", $sFileType)
	_DeleteTrayMessageBox()
EndFunc

; Determine if 7-zip can extract the file
Func check7z()
	If $7zfailed Then Return
	Cout("Testing 7zip")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' 7-Zip')
	Local $return = FetchStdout($7z & ' l "' & $file & '"', $filedir, @SW_HIDE)

	If StringInStr($return, "Listing archive:") And Not (StringInStr($return, "Can not open the file as ") Or _
	StringInStr($return, "There are data after the end of archive")) Then
		; Failsafe in case TrID misidentifies MS SFX
		If StringInStr($return, "_sfx_manifest_") Then
			_DeleteTrayMessageBox()
			extract($TYPE_HOTFIX, 'Microsoft ' & t('TERM_HOTFIX'))
		EndIf
		_DeleteTrayMessageBox()
		If $fileext = "exe" Then
			extract($TYPE_7Z, '7-Zip ' & t('TERM_INSTALLER') & ' ' & t('TERM_PACKAGE'))
		ElseIf $fileext = "iso" Then
			extract($TYPE_7Z, 'Iso ' & t('TERM_IMAGE'))
		ElseIf $fileext = "xz" Then
			extract($TYPE_7Z, 'XZ ' & t('TERM_COMPRESSED'), "xz")
		ElseIf $fileext = "z" Then
			extract($TYPE_7Z, 'LZW ' & t('TERM_COMPRESSED'), "Z")
		Else
			extract($TYPE_7Z, '7-Zip ' & t('TERM_ARCHIVE'))
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

	_DeleteTrayMessageBox()
	If StringInStr($return, "Listing archive:") And Not (StringInStr($return, "corrupted file") _
	Or StringInStr($return, "file open error")) Then extract($TYPE_ALZ, -1)

	Return False
EndFunc

; Determine if file is self-extracting ARJ archive
Func checkArj()
	If $arjfailed Then Return False
	Cout("Testing ARJ")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' SFX ARJ ' & t('TERM_ARCHIVE'))
	$return = FetchStdout($arj & ' l "' & $file & '"', $filedir, @SW_HIDE)

	If StringInStr($return, "Archive created:", 0) Then
		_DeleteTrayMessageBox()
		extract($TYPE_7Z, t('TERM_SFX') & ' ARJ ' & t('TERM_ARCHIVE'))
	EndIf

	_DeleteTrayMessageBox()
	$arjfailed = True
	Return False
EndFunc

; Determine if folder contains .bin files
Func CheckBin()
	Cout("Searching additional .bin files to be extracted")

	Local $hSearch = FileFindFirstFile($filedir & "\data*.bin")
	If $hSearch == -1 Then Return

	Local $tmp = $file
	If Prompt(64 + 1, "NSIS_BINFILES", CreateArray($file, $filenamefull)) Then
		While 1
			$file = $filedir & "\" & FileFindNextFile($hSearch)
			If @error Then ExitLoop
			FilenameParse($file)
			extract($TYPE_7Z, ".bin " & t('TERM_ARCHIVE'), "", True, True)
		WEnd
	EndIf
	FileClose($hSearch)
	FilenameParse($tmp) ; Make sure the log/stats are correct
	terminate($STATUS_SUCCESS, $filenamefull, "NSIS", "NSIS")
EndFunc

; Determine if file is supported game archive
Func CheckGame($bUseGaup = True)
	If (Not $CheckGame) Or $gamefailed Then Return

	Cout("Testing Game archive")

	_CreateTrayMessageBox(t('TERM_TESTING') & ' ' & t('TERM_GAME') & t('TERM_PACKAGE'))

	CheckGarbro()

	If $bUseGaup Then
		; Check GAUP first
		$return = FetchStdout($quickbms & ' -l "' & $bindir & $gaup & '" "' & $file & '"', $filedir, @SW_HIDE, -1)

		If @error Or StringInStr($return, "Target directory:", 0) Or StringInStr($return, "0 files found", 0) Or StringInStr($return, "Error", 0) _
		Or StringInStr($return, "exception occured", 0) Or StringInStr($return, "not supported", 0) Or $return == "" Then

		Else
			_DeleteTrayMessageBox()
			extract($TYPE_QBMS, t('TERM_GAME') & t('TERM_PACKAGE'), $gaup, False, True)
		EndIf
	EndIf

	$gamefailed = True

	If $silentmode And Number($sMethodSelectOverride) < 1  Then
		Cout("INFO: File may be extractable via BMS script, but user input is needed. Disable silent mode to try this method.")
		_DeleteTrayMessageBox()
		Return False
	EndIf

	; Check if game specific bms script is available
	Local $hDB = OpenDB("BMS.db")
	Local $aReturn[0], $iRows, $iColumns

	_SQLite_GetTable($hDB, "SELECT n.Name FROM Names n, Scripts s, Extensions e WHERE s.SID = e.EID AND s.SID = n.NID AND e.Extension= '" _
						  & StringLower($fileext) & "' ORDER BY n.Name", $aReturn, $iRows, $iColumns)
 	_ArrayDelete($aReturn, 1)

	If $aReturn[0] > 1 Then
		_ArrayDelete($aReturn, 0)
		_ArraySort($aReturn)
		Local $iChoice = GUI_MethodSelectList($aReturn, t('METHOD_GAME_NOGAME'))
		If $iChoice > -1 Then BmsExtract($iChoice, $hDB)
	EndIf

	_SQLite_Close()
	_SQLite_Shutdown()

	_DeleteTrayMessageBox()
	Return False
EndFunc

; Determine if file can be extracted with GARbro
Func CheckGarbro()
	HasNetFramework(4.6)
	Cout("Testing GARbro")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' GARbro ' & t('TERM_ARCHIVE'))
	Local $return = FetchStdout($garbro & ' l "' & $file & '"', $filedir, @SW_HIDE)
	If Not @error And Not StringInStr($return, "Error: Input file has an unknown format") And Not StringInStr($return, "Error: Archive is empty") Then
		$return = StringStripWS(StringStripCR(FetchStdout($garbro & ' i "' & $file & '"', $filedir, @SW_HIDE, -1)), 8)
		extract($TYPE_GARBRO, $return & ' ' & t('TERM_GAME') & t('TERM_FILE'))
	EndIf

	_DeleteTrayMessageBox()
	Return False
EndFunc

; Determine if InstallExplorer can extract the file
Func checkIE()
	If $iefailed Then Return False

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

	extract($TYPE_QBMS, 'InstallExplorer ' & t('TERM_INSTALLER'), $ie)
EndFunc

; Determine if file is Inno Setup installer
Func checkInno()
	If $innofailed Then Return False

	Cout("Testing Inno Setup")
	_CreateTrayMessageBox(t('TERM_TESTING') & " Inno Setup " & t('TERM_INSTALLER'))

	$return = FetchStdout($inno & ' "' & $file & '"', $filedir, @SW_HIDE)

	_DeleteTrayMessageBox()

	If (StringInStr($return, "Version detected:", 0) And Not (StringInStr($return, "error", 0))) _
	Or (StringInStr($return, "Signature detected:", 0) And Not StringInStr($return, "not a supported version", 0)) Then _
		extract($TYPE_INNO, "Inno Setup " & t('TERM_INSTALLER'))

	$innofailed = True
	checkIE()
	Return False
EndFunc

; Determine if file can be extracted by TotalObserver
Func CheckTotalObserver($arcdisp = 0)
	If $observerfailed Then Return False
	Cout("Testing TotalObserver")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' TotalObserver ' & t('TERM_ARCHIVE'))

	Local $return = FetchStdout($quickbms & ' -l "' & $bindir & $observer & '" "' & $file & '"', $filedir, @SW_HIDE)

	_DeleteTrayMessageBox()
	If StringInStr($return, "not supported by this WCX plugin") Or StringInStr($return, "0 files found") Or _
	   StringInStr($return, "exception occured") Or StringInStr($return, "EXCEPTION HANDLER") Then
	   $observerfailed = True
	   Return False
	EndIf

	extract($TYPE_QBMS, $arcdisp, $observer)
EndFunc

; Determine if file is CD/DVD image
Func CheckIso()
	If $isofailed Then Return False
	Cout("Testing image file")
	_CreateTrayMessageBox(t('TERM_TESTING') & " " & t('TERM_IMAGE') & " " & t('TERM_FILE'))
	If $fileext = "cue" And FileExists(StringTrimRight($file, 3) & "bin") Then $file = StringTrimRight($file, 3) & "bin"

	Local $return = FetchStdout($quickbms & ' -l "' & $bindir & $iso & '" "' & $file & '"', $filedir, @SW_HIDE)
	_DeleteTrayMessageBox()
	If StringInStr($return, "Target directory:") Or StringInStr($return, "0 files found") Or StringInStr($return, "Error") _
	Or StringInStr($return, "exception occured") Or StringInStr($return, "not supported") Or $return == "" Then
		$isofailed = True
		Return False
	EndIf
	Return extract($TYPE_QBMS, t('TERM_IMAGE') & " " & t('TERM_FILE'), $iso, $isofile, $isofile)
EndFunc

; Try listing msi contents with lessmsi
Func CheckLessmsi()
	If Not HasNetFramework(4, False) Then Return False

	Cout("Testing lessmsi")
	$return = FetchStdout($msi_lessmsi & ' l "' & $file & '"', $outdir)

	Return StringInStr($return, "Listing msi file") And Not StringInStr($return, "Error: ")
EndFunc

; Determine if file is NSIS installer
Func checkNSIS()
	Cout("Testing NSIS")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' NSIS ' & t('TERM_INSTALLER'))

	$return = FetchStdout($7z & ' l "' & $file & '"', $filedir, @SW_HIDE)
	If StringInStr($return, "Listing archive:") And Not StringInStr($return, "Can not open the file as") Then _
		extract($TYPE_NSIS, 'NSIS ' & t('TERM_INSTALLER'))

	_DeleteTrayMessageBox()
	checkIE()
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
EndFunc

; Perform special actions for some file types
Func InitialCheckExt()
	If Not $extract Then Return

	Switch $fileext
		; Compound compressed files that require multiple actions
		Case "ipk", "tbz2", "tgz", "tz", "tlz", "txz"
			extract($TYPE_CTAR, 'Compressed Tar ' & t('TERM_ARCHIVE'))
		; Disk images - file type identification is not always reliable
		Case "bin", "cue", "cdi", "mdf"
			CheckIso()
		Case "dmg"
			extract($TYPE_7Z, 'DMG ' & t('TERM_IMAGE'))
		Case "iso", "nrg"
			check7z()
			CheckIso()
	EndSwitch
EndFunc

; Check for unicode characters in path
Func MoveInputFileIfNecessary()
	$bIsUnc = _WinAPI_PathIsUNC($file)

	Local $new = 0
	If $checkUnicode And (Not StringRegExp($file, $sRegExAscii, 0) Or StringLeft($filename, 2) == "--") Then
		Cout("File name seems to be unicode")

		If StringRegExp($filedir, $sRegExAscii, 0) Then
			; Directory is ASCII, only rename file
			$new = _TempFile($filedir, "Unicode_", $fileext)
		Else
			Cout("Path seems to be unicode")
			If Not StringRegExp(@TempDir, $sRegExAscii, 0) Then Return Cout("Temp directory contains unicode characters: " & @TempDir)
			$new = StringRegExp($filename, $sRegExAscii, 0)? @TempDir & "\" & $filenamefull: _TempFile(@TempDir, "Unicode_", $fileext)
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
		If Not _FileMove($file, $new) Then Return terminate($STATUS_MOVEFAILED, $new)
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

	If Not StringRegExp($outdir, $sRegExAscii, 0) Then
		Cout("Output directory seems to be unicode")
		$outdir = $initoutdir
	EndIf
EndFunc

; Extract from known archive format
Func extract($arctype, $arcdisp = 0, $additionalParameters = "", $returnSuccess = False, $returnFail = False)
	Local $dirmtime = -1
	$success = $RESULT_UNKNOWN

	Cout("Starting " & $arctype & " extraction")
	If $arcdisp == 0 Then $arcdisp = "." & $fileext & " " & t('TERM_FILE')
	If $arcdisp <> -1 Then _CreateTrayMessageBox(t('EXTRACTING') & @CRLF & $arcdisp)

	; Create subdirectory
	If StringRight($outdir, 1) = '\' Then $outdir = StringTrimRight($outdir, 1)
	If FileExists($outdir) Then
		If Not _IsDirectory($outdir) Then terminate($STATUS_INVALIDDIR, $outdir, "")
		$dirmtime = FileGetTime($outdir, 0, 1)
		If @error Then $dirmtime = -1
		If Not CanAccess($outdir) Then terminate($STATUS_INVALIDDIR, $outdir)
	Else
		If Not DirCreate($outdir) Then terminate($STATUS_INVALIDDIR, $outdir)
		$createdir = True
	EndIf

	HasFreeSpace()

	$initdirsize = _DirGetSize($outdir)
	$tempoutdir = TempDir($outdir, 7)
	Local $sFileType = _FiletypeGet(False)

	; Extract archive based on filetype
	Switch $arctype
		Case $TYPE_7Z
			Local $sPassword = _FindArchivePassword($7z & ' l -p -slt "' & $file & '"', $7z & ' t -p"%PASSWORD%" "' & $file & '"', "Encrypted = +", "Wrong password?", 0, "Everything is Ok")
			_Run($7z & ' x ' & ($sPassword == 0? '"': '-p"' & $sPassword & '" "') & $file & '"', $outdir, @SW_HIDE, True, True, True, True)
			If @error = 3 Then terminate($STATUS_MISSINGPART)
			If @extended Then terminate($STATUS_PASSWORD, $file, $arctype, $arcdisp)

			If FileExists($outdir & '\.text') Then
				; Generic .exe extraction should not be considered successful
				$success = $RESULT_FAILED
			ElseIf StringInStr($sFileType, 'RPM Linux Package', 0) Then
				; Extract inner CPIO for RPMs
				If FileExists($outdir & '\' & $filename & '.cpio') Then
					_Run($7z & ' x "' & $outdir & '\' & $filename & '.cpio"', $outdir)
					FileDelete($outdir & '\' & $filename & '.cpio')
				EndIf
			ElseIf StringInStr($sFileType, 'Debian Linux Package', 0) Then
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

		Case $TYPE_ACE
			; TODO: _FindArchivePassword
			_Run($ace & ' -x -v -d "' & $outdir & '" "' & $file & '"', $outdir, @SW_HIDE, True, True, True, True)

		Case $TYPE_ACTUAL
			; Actual installers contain two blobs of zip data.
			; First, extract the meta data, which is needed later.
			DirCreate($tempoutdir)
			_Run($zip & ' "' & $file & '"', $tempoutdir, @SW_MINIMIZE)
			Local $aFiles = IniReadSection($tempoutdir & "aisetup.ini", "Files")
			Cleanup($tempoutdir & "*")

			; Now get the actual files
			; Different zip implementations parse archives differently.
			; We can just abuse this fact instead of writing a proper extractor for these installers.
			If Not extract($TYPE_7Z, -1, "", True, True) Then
				Cout("Failed to extract files")
				$success = $RESULT_FAILED
			ElseIf Not IsArray($aFiles) Then
				Cout("Failed to read file names")
			Else
				; Extracted files do not have original names, we need to parse the installer configuration file
				; and rename the files accordingly.
				For $i = 1 To $aFiles[0][0]
					; Remove invalid characters
					Local $sDestination = StringReplace($aFiles[$i][1], "<", "[")
					$sDestination = StringReplace($sDestination, ">", "]")
					Local $iPos = StringInStr($sDestination, "?")
					If $iPos > -1 Then $sDestination = StringLeft($sDestination, $iPos - 1)

					Local $sSource = $outdir & "\" & $aFiles[$i][0]
					$sDestination = $outdir & "\" & $sDestination

					Cout("Renaming " & $sSource & " to " & $sDestination)
					_FileMove($sSource, $sDestination, $FC_CREATEPATH)
				Next
			EndIf

		Case $TYPE_AI
			Warn_Execute($file & ' /extract:"' & $outdir & '"')
			; ShellExecute is needed here to display UAC prompt, fails with Run()
			ShellExecute($file, ' /extract:"' & $outdir & '"', $outdir)
			ProcessWait($filenamefull, $Timeout)
			ProcessWaitClose($filenamefull, $Timeout)

		Case $TYPE_ARC_CONV
			If Not HasPlugin($arc_conv, $returnFail) Then Return

			$run = Run(Cout(_MakeCommand($arc_conv & ' "' & $file & '"', True)), $outdir, @SW_HIDE)
			If @error Then Return

			Local $hWnd = WinWait("arc_conv", "", $Timeout)
			If $hWnd == 0 Then terminate($STATUS_TIMEOUT, $file, $arctype, $arcdisp)
			Local $current = "", $last = ""
			; Hide not possible as window text has to be read
			WinSetState("arc_conv", "", @SW_MINIMIZE)
			While WinExists("arc_conv")
				$current = WinGetText("arc_conv")
				If $current <> $last Then
					_SetTrayMessageBoxText(StringInStr($current, "/")? $current: (t('TERM_FILE') & " #" & $current))
					$last = $current
					Sleep(10)
				EndIf
			WEnd
			$run = 0
			MoveFiles($file & "~", $outdir, True, "", True, True)

		Case $TYPE_AUDIO
			HasFFMPEG()
			_Run($cmd & $ffmpeg & ' -i "' & $file & '" "' & GetFileName() & '.wav"', $outdir, @SW_HIDE)

		Case $TYPE_BCM
			_Run($bcm & ' -d "' & $file & '" "' & $outdir & '\' & GetFileName() & '"', $filedir, @SW_HIDE, True, True, False)

		Case $TYPE_BOOTIMG
			HasPlugin($bootimg)
			$ret = $outdir & "\" & $bootimg
			$ret2 = $outdir & '\boot.img'
			FileCopy($bindir & $bootimg, $outdir)
			_FileMove($file, $ret2)
			_Run($cmd & '"' & $ret & ' --unpack-bootimg"', $outdir, @SW_MINIMIZE, False, False)
			_FileMove($ret2, $file)
			FileDelete($ret)

		Case $TYPE_CAB
			If StringInStr($sFileType, 'Type 1', 0) Then
				RunWait(Warn_Execute(Quote($file & '" /q /x:"' & $outdir), $outdir)
			Else
				check7z()
			EndIf

		Case $TYPE_CHM
			_Run($7z & ' x "' & $file & '"', $outdir)
			Local $aCleanup[] = ['#*', '$*']
			Cleanup($aCleanup)
			$hSearch = FileFindFirstFile($outdir & '\*')
			If $hSearch <> -1 Then
				$dir = FileFindNextFile($hSearch)
				Do
					$char = StringLeft($dir, 1)
					If $char = '#' Or $char = '$' Then Cleanup($outdir & "\" & $dir)
					$dir = FileFindNextFile($hSearch)
				Until @error
			EndIf
			FileClose($hSearch)

		Case $TYPE_CI
			HasPlugin($ci)
			$return = @TempDir & "\ci.txt"
			$ret = FileOpen($return, $FO_CREATEPATH + $FO_OVERWRITE)
			FileWrite($ret, "1" & @LF & $file & @LF & $outdir & @LF & "3" & @LF & "1")
			FileClose($ret)
			$run = Run(_MakeCommand($ci & ' ' & $return, False), $outdir, @SW_SHOW)
			WinWait("CreateInstall Setup Extractor", "Click Finish to close the program", $Timeout)
			ControlClick("CreateInstall Setup Extractor", "", "Button1")
			ProcessClose($run)
			FileDelete($return)
			terminate($STATUS_SILENT)

		Case $TYPE_CIC
			HasNetFramework(4.5)
			_Run($cic & ' -db "' & $file & '" "' & $outdir & '"', $filedir, @SW_HIDE)
			Cleanup("Block 0x*.bin")

		Case $TYPE_CTAR
			$oldfiles = ReturnFiles($outdir)

			; Decompress archive with 7-zip
			_Run($7z & ' x "' & $file & '"', $outdir)

			; Check for new files
			$hSearch = FileFindFirstFile($outdir & "\*")
			If Not @error Then
				While 1
					$fname = FileFindNextFile($hSearch)
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
			FileClose($hSearch)

		Case $TYPE_DGCA
			HasPlugin($dgca)
			Local $sPassword = _FindArchivePassword($dgca & ' e "' & $file & '"', $dgca & ' l -p%PASSWORD% "' & $file & '"', "Archive encrypted.", 0, -2, "-------------------------")
			_Run($dgca & ' e ' & ($sPassword == 0? '"': '-p' & $sPassword & ' "') & $file & '" "' & $outdir & '"', $outdir, @SW_HIDE, True, True, False, False)
			If @extended Then terminate($STATUS_PASSWORD, $file, $arctype, $arcdisp)

		Case $TYPE_DAA
			; Prompt user to continue
			_DeleteTrayMessageBox()
			Prompt(32 + 4, 'CONVERT_CDROM', 'DAA/GBI', True)

			; Begin conversion to .iso format
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'DAA ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 1)')
			$isofile = $filedir & '\' & $filename & '.iso'
			_Run($daa & ' "' & $file & '" "' & $isofile & '"', $filedir)

		Case $TYPE_DCP
			HasPlugin($dcp)
			_Run($dcp & ' "' & $file & '"', $outdir)

		Case $TYPE_EI
			Warn_Execute($file & ' /batch /no-reg /no-postinstall /dest "' & $outdir & '"')
			ShellExecuteWait($file, '/batch /no-reg /no-postinstall /dest "' & $outdir & '"', $outdir)

		Case $TYPE_ENIGMA ; Test
			_RunInTempOutdir($tempoutdir, $enigma & ' /nogui "' & $file & '"', $tempoutdir, @SW_HIDE, True, False, False)

			_FileMove($outdir & "\" & $filename & "_unpacked.exe", $outdir & "\" & GetFileName() & "_" & t('TERM_UNPACKED') & ".exe")
			; TODO: move other folders

			; Read log file
			Local $sPath = $outdir & "\!unpacker.log"
			$return = Cout(FileRead($sPath))
			FileDelete($sPath)

			; Success evaluation
			If StringInStr($return, 'Expected section name ".enigma2"') Then
				$success = $RESULT_FAILED
			ElseIf StringInStr($return, '[+] Finished!') Then
				$success = $RESULT_SUCCESS
			EndIf

		Case $TYPE_FEAD
			Local $tmp = ' /s -nos_ne -nos_o"' & $tempoutdir & '\"'
			Warn_Execute($file & $tmp)
			ShellExecuteWait($file, $tmp, $filedir)
			FileSetAttrib($tempoutdir & '*', '-R', 1)
			MoveFiles($tempoutdir, $outdir, False, "", True, True)
			DirRemove($tempoutdir)

		Case $TYPE_FREEARC
			_Run($freearc & ' x -dp"' & $outdir & '" "' & $file & '"', $filedir, @SW_HIDE, True, True, False, False)

		Case $TYPE_FSB
			_Run($fsb & ' -o -1 -A -d "' & $outdir & '" "' & $file & '"', $filedir, @SW_MINIMIZE, True, True, False)

			; Ogg files are raw dumps and not playable
			Cleanup("*.ogg")

		Case $TYPE_GARBRO
			_Run($garbro & ' x -ocu -if png -o "' & $outdir & '" "' & $file & '"', $outdir, @SW_MINIMIZE)

		Case $TYPE_GHOST
			$ret = $outdir & "\" & $filename & ".exe"
			Cout("Moving file to " & $ret)
			_FileMove($file, $ret)

			$aReturn = OpenExeInfo($ret)

			WinWait($aReturn[0], "", $Timeout)
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
			_FileMove($ret, $filedir & "\")

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

		Case $TYPE_HLP
			_Run($hlp & ' "' & $file & '"', $outdir)
			If _DirGetSize($outdir, $initdirsize + 1) > $initdirsize Then
				DirCreate($tempoutdir)
				_Run($hlp & ' /r /n "' & $file & '"', $tempoutdir)
				_FileMove($tempoutdir & $filename & '.rtf', $outdir & '\' & $filename & '_' & t('TERM_RECONSTRUCTED') & '.rtf')
				DirRemove($tempoutdir, 1)
			EndIf

		; Failsafe in case TrID misidentifies MS SFX hotfixes
		Case $TYPE_HOTFIX ; Test
			Cout("Executing: " & Warn_Execute(Quote($file & '" /q /x:"' & $outdir)))
			ShellExecuteWait($file, '/q /x:' & Quote($outdir), $outdir)

		Case $TYPE_INNO
			_Run($inno & ' -x -m -a "' & $file & '"', $outdir)

			; Inno setup files can contain multiple versions of files, they are named ',1', ',2',... after extraction
			; rename the first file(s), so extracted programs do not fail with not found exceptions
			; This is a convenience function, so the user does not have to rename them manually
			$return = $outdir & "\{app}\"
			$aReturn = _FileListToArrayRec($return, "*,1.*", 1, 1)
			If Not @error Then
				For $i = 1 To $aReturn[0]
					$ret = StringReplace($aReturn[$i], ",1", "", -1)
					Cout("Renaming " & $return & $aReturn[$i] & " to " & $return & $ret)
					_FileMove($return & $aReturn[$i], $return & $ret)
				Next
			EndIf

			; (Re)move ',2' files and install_script.iss
			Local $aCleanup = _FileListToArrayRec($return, "*,2.*;*,3.*", 1, 1, 0, 2)
			_ArrayDelete($aCleanup, 0)
			Cleanup($aCleanup)

			; Change output directory structure
			Local $aCleanup[] = ["embedded", "{tmp}", "{commonappdata}", "{cf}", "{cf32}", "{group}", "{{userappdata}}", "{{userdocs}}"]
			Cleanup($aCleanup)
			MoveFiles($outdir & "\{app}", $outdir, True, '', True)
			; TODO: {syswow64}, {sys} - move files to outdir as dlls might be needed by the program?

			Local $aCleanup[] = ["install_script.iss", "setup.iss"]
			Cleanup($aCleanup)

		Case $TYPE_ISCAB
			; Unshield only works with UNIX-style paths
			Local $sPath = StringReplace($file, "\", "/")
			Local $sReturn = _Run($unshield & ' -D 2 -d "' & $outdir & '" x "' & $sPath & '"', $outdir)
			If StringInStr($sReturn, "Try unshield_file_save_old()") Then $sReturn = _Run($unshield & ' -O -D 2 -d "' & $outdir & '" x "' & $sPath & '"', $outdir)
			If StringInStr($sReturn, "Failed to extract file") Then

			Local $aReturn = ['InstallShield Cabinet ' & t('TERM_ARCHIVE'), t('METHOD_EXTRACTION_RADIO', 'is6comp'), t('METHOD_EXTRACTION_RADIO', 'is5comp'), t('METHOD_EXTRACTION_RADIO', 'iscab')]
			$iChoice = GUI_MethodSelect($aReturn, $arcdisp)

			Switch $iChoice
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
			EndIf

		Case $TYPE_ISCRIPT
			If Not extract($TYPE_QBMS, $arcdisp, $observer, False, True) Then
				$success = $RESULT_UNKNOWN
				Warn_Execute($file & ' /extract_all:"' & $outdir & '"')
				ShellExecuteWait($file, ' /extract_all:"' & $outdir & '"', $outdir, "open", @SW_MINIMIZE)
			EndIf

		Case $TYPE_ISEXE
			CheckTotalObserver($arcdisp)

			Local $aReturn = ["InstallShield " & t('TERM_INSTALLER'), t('METHOD_EXTRACTION_RADIO', 'isxunpack'), t('METHOD_SWITCH_RADIO', 'InstallShield /b'), t('METHOD_NOTIS_RADIO')]
			$iChoice = GUI_MethodSelect($aReturn, $arcdisp)

			Switch $iChoice
				; Extract using isxunpack
				Case 1
					_FileMove($file, $outdir)
					Run(_MakeCommand($isxunp & ' "' & $outdir & '\' & $filenamefull & '"', True), $outdir)
					WinWait(@ComSpec)
					WinActivate(@ComSpec)
					Send("{ENTER}")
					ProcessWaitClose($isxunp)
					_FileMove($outdir & '\' & $filenamefull, $filedir)

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
										If $fname <> $msiname Then FileCopy($isdir & "\" & $fname, $tempoutdir)
										$fname = FileFindNextFile($ishandle)
									Until @error
									FileClose($ishandle)
								WEnd
								FileClose($msihandle)
							EndIf

							; Move files to outdir
							_DeleteTrayMessageBox()
							Prompt(64, 'INIT_COMPLETE')
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
						Prompt(64, 'INIT_COMPLETE')
					EndIf

				; Not InstallShield
				Case 3
					Return False
			EndSwitch

		Case $TYPE_ISZ
			_DeleteTrayMessageBox()
			Prompt(32 + 4, 'CONVERT_CDROM', 'ISZ', True)
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'ISZ ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 1)')

			_RunInTempOutdir($tempoutdir, $isz & ' "' & $file & '"', $tempoutdir, True, True)
			$isofile = $outdir & "\" & $filename & '.iso'

		Case $TYPE_KGB
			_Run($kgb & ' "' & $file & '"', $outdir, @SW_MINIMIZE, True, False, False)

		Case $TYPE_LZ
			_RunInTempOutdir($tempoutdir, $lz & ' -d -k -v -v "' & $file & '"', $tempoutdir, @SW_SHOW, True, True, False)

		Case $TYPE_LZO
			_Run($lzo & ' -d -p"' & $outdir & '" "' & $file & '"', $filedir)

		Case $TYPE_LZX
			_Run($lzx & ' -x "' & $file & '"', $outdir)

		Case $TYPE_MHT
			If HasPlugin($archdir & "\Formats\eDecoder." & $iOsArch & ".dll", True) Then
				extract($TYPE_7Z, $arcdisp)
			Else
				extract($TYPE_QBMS, $arcdisp, $observer)
			EndIf

;~ 			Local $aReturn = ['MHTML ' & t('TERM_ARCHIVE'), t('METHOD_EXTRACTION_RADIO', '7zip'), t('METHOD_EXTRACTION_RADIO', 'TotalObserver')]
;~ 			$iChoice = GUI_MethodSelect($aReturn, $arcdisp)

		Case $TYPE_MOLE
			_RunInTempOutdir($tempoutdir, $mole & ' /nogui "' & $file & '"', $outdir, @SW_HIDE, True, False, False)

			; Move files
			Local $sPath = $outdir & "\" & $filename & "_unpacked.exe"
			If FileExists($sPath) Then _FileMove($sPath, $outdir & "\" & GetFileName() & "_" & t('TERM_UNPACKED') & ".exe")

			$sPath = $outdir & "\_extracted"
			If FileExists($sPath) Then MoveFiles($sPath, $outdir, False, "", True, True)

			; Read log file
			$sPath = $outdir & "\!unpacker.log"
			Local $sLog = Cout(FileRead($sPath))
			FileDelete($sPath)

			; Success evaluation
			If StringInStr($sLog, '[x] Not a Molebox or unknown version') Then
				$success = $RESULT_FAILED
			ElseIf StringInStr($sLog, '[i] Finished! Have a nice day!') Then
				$success = $RESULT_SUCCESS
			EndIf

		Case $TYPE_MSCF
			$oldfiles = ReturnFiles($outdir)
			extract($TYPE_7Z, -1, "", False, True)

			; If 7z fails, remove useless files and extract cab files from installer
			MoveFiles($outdir, $tempoutdir, False, $oldfiles, True, False)
			DirRemove($tempoutdir, True)
			Sleep(1000)

			If RipExeInfo($file, $tempoutdir, "{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{RIGHT}{DOWN}{DOWN}{DOWN}") Then
				$cabfiles = FileSearch($tempoutdir & "\*.cab")
				For $i = 1 To $cabfiles[0]
					Cout("Extracting cab file " & $cabfiles[$i])
					_Run($7z & ' x "' & $cabfiles[$i] & '"', $tempoutdir, @SW_HIDE, True, True, True, False)
					If $success == $RESULT_SUCCESS Then FileRecycle($cabfiles[$i])
				Next

				MoveFiles($tempoutdir, $outdir, False, '', True, True)
				Local $aCleanup[] = ["resource.dat", "cp*.bin", "*.cab"]
				Cleanup($aCleanup)
				$success = $RESULT_UNKNOWN
			Else
				$success = $RESULT_FAILED
			EndIf

		Case $TYPE_MSI
			; Try Lessmsi first
			$ret = CheckLessmsi()
			If $ret Then
				_Run($msi_lessmsi & ' x "' & $file & '" "' & $outdir & '\"', $outdir, @SW_HIDE, True, True, True)
				MoveFiles($outdir & "\SourceDir", $outdir, False, '', True)
				If $success == $RESULT_UNKNOWN And DirGetSize($outdir) == $initdirsize Then $success = $RESULT_FAILED
			EndIf

			; If lessmsi fails or .NET framework is not available, the user can choose between legacy extractors
			If Not $ret Or $success == $RESULT_FAILED Then
				$success = $RESULT_UNKNOWN
				Local $aReturn = ['MSI ' & t('TERM_INSTALLER'), t('METHOD_EXTRACTION_RADIO', 'jsMSI Unpacker'), t('METHOD_EXTRACTION_RADIO', 'MsiX'), t('METHOD_EXTRACTION_RADIO', 'MSI TC Packer'), t('METHOD_ADMIN_RADIO', 'MSI')]
				$iChoice = GUI_MethodSelect($aReturn, $arcdisp)

				Switch $iChoice
					Case 1 ; jsMSI Unpacker
						_Run($msi_jsmsix & ' "' & $file & '"|"' & $outdir & '"', $filedir, @SW_HIDE, False, False)
						_FileRead($outdir & "\MSI Unpack.log", True)
						Cleanup("*.cab")

					Case 2 ; MsiX
						Local $appendargs = $appendext? '/ext': ''
						_Run($msi_msix & ' "' & $file & '" /out "' & $outdir & '" ' & $appendargs, $filedir)

					Case 3 ; MSI Total Commander plugin
						DirCreate($tempoutdir)
						_Run($quickbms & ' "' & $bindir & $msi_plug & '" "' & $file & '" "' & $tempoutdir & '"', $outdir, @SW_MINIMIZE, True, False)

						; Extract files from extracted CABs
						$cabfiles = FileSearch($tempoutdir & "*.cab")
						For $i = 1 To $cabfiles[0]
							_Run($cmd & $7z & ' x "' & $cabfiles[$i] & '"', $outdir)
							FileDelete($cabfiles[$i])
						Next

						; Append missing file extensions
						If $appendext Then AppendExtensions($tempoutdir)

						; Move files to output directory and remove tempdir
						MoveFiles($tempoutdir, $outdir, False, "", True)

					Case 4 ; Administrative install
						RunWait(Warn_Execute('msiexec.exe /a "' & $file & '" /qb TARGETDIR="' & $outdir & '"'), $filedir, @SW_SHOW)
				EndSwitch
			EndIf

		Case $TYPE_MSM ; Test
			; Due to the appendext argument, a definition file cannot be used here
			_Run($msi_msix & ' "' & $file & '" /out "' & $outdir & '" ' & $appendext? '/ext': '', $filedir)

		Case $TYPE_MSP ; Test
			Local $aReturn = ['MSP ' & t('TERM_PACKAGE'), t('METHOD_EXTRACTION_RADIO', 'MSI TC Packer'), t('METHOD_EXTRACTION_RADIO', 'MsiX'), t('METHOD_EXTRACTION_RADIO', '7-Zip')]
			$iChoice = GUI_MethodSelect($aReturn, $arcdisp)

			DirCreate($tempoutdir)
			Switch $iChoice
				Case 1 ; TC MSI
					_Run($quickbms & ' "' & $bindir & $msi_plug & '" "' & $file & '" "' & $tempoutdir & '"', $tempoutdir, @SW_MINIMIZE, True, False)
;~ 					extract($TYPE_QBMS, $arcdisp, $msi_plug, True)
				Case 2 ; MsiX
					_Run($msi_msix & ' "' & $file & '" /out "' & $tempoutdir & '"', $filedir)
				Case 3 ; 7-Zip
					_Run($7z & ' x "' & $file & '"', $tempoutdir)
			EndSwitch

			; Regardless of method, extract files from extracted CABs
			; TODO: This does not work with new filescan function.
			; 		Basically we need to do: open dll, query each file, change extension if enabled, extract if cab, close DLL
;~ 			$cabfiles = FileSearch($tempoutdir)
;~ 			For $i = 1 To $cabfiles[0]
;~ 				filescan($cabfiles[$i], 0)
;~ 				If StringInStr($sFileType, "Microsoft Cabinet Archive", 0) Then
;~ 					_Run($7z & ' x "' & $cabfiles[$i] & '"', $outdir)
;~ 					FileDelete($cabfiles[$i])
;~ 				EndIf
;~ 			Next

			; Append missing file extensions
			If $appendext Then AppendExtensions($tempoutdir)

			; Move files to output directory and remove tempdir
			MoveFiles($tempoutdir, $outdir, False, "", True)

		Case $TYPE_NBH ; Test
			RunWait(_MakeCommand($nbh, True) & ' "' & $file & '"', $outdir)

		Case $TYPE_NSIS
			; Rename duplicates and extract
			_Run($7z & ' x -aou' & ' "' & $file & '"', $outdir)

			If $success == $RESULT_FAILED Then checkIE()

			Local $aCleanup[] = ["[NSIS].nsi", "[LICENSE].*", "$PLUGINSDIR", "$TEMP", "uninstall.exe", "[LICENSE]"]
			Cleanup($aCleanup)

			; Determine if there are .bin files in filedir
			CheckBin()

		Case $TYPE_PDF
			_Run($pdfdetach & ' -saveall "' & $file & '"', $outdir, @SW_HIDE, True, True, False, False)
			_Run($pdftohtml & ' "' & $file & '" "' & $outdir & '\' & $filename & '-HTML"', $outdir, @SW_HIDE, True, True, False, False)
			_Run($pdftopng & ' "' & $file & '" "' & $outdir & '\' & $filename & '-' & t('TERM_PAGE') & '"', $outdir, @SW_HIDE, True, True, False, False)
			_Run($pdftotext & ' "' & $file & '" "' & $outdir & '\' & $filename & '.txt"', $outdir, @SW_HIDE, True, True, False, False)

		Case $TYPE_PEA
			DirCreate($tempoutdir)
			Local $pid = Run($pea & ' UNPEA "' & $file & '" "' & $tempoutdir & '" RESETDATE SETATTR EXTRACT2DIR INTERACTIVE', $filedir)
			While ProcessExists($pid)
				$return = ControlGetText(_WinGetByPID($pid), '', 'Button1')
				If StringLeft($return, 4) = 'Done' Then ProcessClose($pid)
				Sleep(10)
			WEnd
			MoveFiles($tempoutdir, $outdir, False, "", True)
			FileDelete($bindir & "rnd")

		Case $TYPE_QBMS
			_Run($quickbms & ' "' & $bindir & $additionalParameters & '" "' & $file & '" "' & $outdir & '"', $outdir, @SW_MINIMIZE, True, False)
			If FileExists($bindir & $bms) Then FileDelete($bindir & $bms)

			If $additionalParameters == $ie Then
				Local $aCleanup[] = ["[NSIS].nsi", "[LICENSE].*", "$PLUGINSDIR", "$TEMP", "uninstall.exe", "[LICENSE]"]
				Cleanup($aCleanup)
			EndIf

		Case $TYPE_RAI
			DirCreate($tempoutdir)
			Local $tmp = $tempoutdir & $filename & '_' & t('TERM_UNPACKED') & '.exe'
			_Run($rai & ' "' & $file & '" "' & $tmp & '"', $filedir)
			$file = $tmp
			extract($TYPE_INNO, $arcdisp, "", True)
			Cleanup($tmp)
			DirRemove($tempoutdir)

		Case $TYPE_RAR
			Local $sPassword = _FindArchivePassword($rar & ' lt -p- "' & $file & '"', $rar & ' t -p"%PASSWORD%" "' & $file & '"', "encrypted", 0, 0)
			_Run($rar & ' x ' & ($sPassword == 0? '"': '-p"' & $sPassword & '" "') & $file & '"', $outdir, @SW_SHOW)
			If @error = 3 Then terminate($STATUS_MISSINGPART)
			If @extended Then terminate($STATUS_PASSWORD, $file, $arctype, $arcdisp)

		Case $TYPE_RGSS
			HasNetFramework(2)
			_Run($rgss & ' -p -o="' & $outdir & '" "' & $file & '"', $outdir, @SW_HIDE)

		Case $TYPE_ROBO ; Test
			RunWait(Warn_Execute($file & ' /unpack="' & $outdir & '"'), $filedir)

		Case $TYPE_RPA
			_Run($rpa & ' -m -v --continue-on-error -p "' & $outdir & '" "' & $file & '"', @ScriptDir, True, True, True)

		Case $TYPE_SFARK
			_Run($sfark & ' "' & $file & '" "' & $outdir & '\' & $filename & '.sf2"', $filedir, @SW_SHOW)

		Case $TYPE_SGB ; Test
			$return = _Run($7z & ' x "' & $file & '"', $outdir, @SW_HIDE)
			; Effect.sgr cannot be extracted by 7zip, that should not be considered failed extraction as all other files are OK
			If StringInStr($return, "Headers Error : Effect.sgr") And StringInStr($return, "Sub items Errors: 1") Then $success = $RESULT_SUCCESS

		Case $TYPE_SIM
			HasPlugin($sim)
			_Run($sim & ' "' & $file & '" "' & $outdir & '"', $outdir, @SW_SHOW)

			; Extract cab files
			$ret = $outdir & "\data.cab"
			_Run($7z & ' x "' & $ret & '"', $outdir, @SW_SHOW, True, True, True)
			If $success == $RESULT_SUCCESS Then Cleanup($ret, $OPTION_DELETE)

			; Restore original file names and folder structure
			$ret = _FileRead($outdir & "\installer.config", False, $FO_BINARY)
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
						_FileMove($return, $ret, 9)
					Else
						DirCreate($ret)
					EndIf

					; Update next file name variable
					If $bNext Then $j += 1
				Next
			EndIf

			Local $aCleanup[] = ["runtime.cab", "installer.config"]
			Cleanup($aCleanup)

		Case $TYPE_SIS
			extract($TYPE_QBMS, -1, $sis, False, True)
			HasPlugin($extsis)
			DirCreate($tempoutdir)
			_Run($extsis & ' -x -xcsd "' & $file & '" -d "' & $tempoutdir & '"', $tempoutdir, @SW_MINIMIZE)
			MoveFiles($tempoutdir & StringLower($filename), $outdir, False, '', True, True)
			FileDelete($bindir & "extsis.ini")
			DirRemove($bindir & "Shell\", 1)
			DirRemove(@MyDocumentsDir & "\SISContents", 0)

		Case $TYPE_SQLITE
			$return = FetchStdout($sqlite & ' "' & $file & '" .dump"', $filedir, @SW_HIDE, 0)
			$hFile = FileOpen($outdir & '\' & $filename & '.sql', $FO_CREATEPATH + $FO_OVERWRITE)
			FileWrite($hFile, $return)
			FileClose($hFile)

		Case $TYPE_SUPERDAT
			Local $sPath = $outdir & '\SuperDAT.log'
			Local $sParameters = ' /LOGFILE "' & $sPath & '" /e "' & $outdir & '"'
			Warn_Execute($file & $sParameters)
			ShellExecuteWait($file, $sParameters, $outdir)
			_FileRead($sPath, True)

		Case $TYPE_SWF
			; Run swfextract to get list of contents
			$aReturn = StringSplit(FetchStdout($swf & ' "' & $file & '"', $filedir, @SW_HIDE), @CRLF)
			If @error Then
				$success = $RESULT_FAILED
			Else
;~ 				_ArrayDisplay($return)
				For $i = 2 To $aReturn[0] - 1
					$line = $aReturn[$i]
					; Extract files
					If StringInStr($line, "MP3 Soundstream") Then
						_Run($swf & ' -m "' & $file & '"', $outdir, @SW_HIDE, True, True, False, False)
						If FileExists($outdir & "\output.mp3") Then _FileMove($outdir & "\output.mp3", $outdir & "\MP3 Soundstream\soundstream.mp3", 8 + 1)
					ElseIf $line <> "" Then
						$swf_arr = StringSplit(StringRegExpReplace(StringStripWS($line, 8), '(?i)\[(-\w)\]\d+(.+):(.*?)\)', "$1,$2,"), ",")
;~ 						_ArrayDisplay($swf_arr)
						$j = 3
						Do
							;Cout("$j = " & $j & @TAB & $swf_arr[$j])
							$swf_obj = StringInStr($swf_arr[$j], "-")
							If $swf_obj Then
								For $k = StringMid($swf_arr[$j], 1, $swf_obj - 1) To StringMid($swf_arr[$j], $swf_obj + 1)
									_ArrayAdd($swf_arr, $k)
								Next
								$swf_arr[0] = UBound($swf_arr) - 1
;~ 								_ArrayDisplay($swf_arr)
							Else
								; Progress indicator
								_SetTrayMessageBoxText($swf_arr[2] & ": " & $j & "/" & $swf_arr[0] + 1)

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
;~								_ArrayDisplay($swf_arr)

								_FileMove($outdir & "\" & $fname, $outdir & "\" & $swf_arr[2] & "\", 8 + 1)
							EndIf
							$j += 1
						Until $j = $swf_arr[0] + 1
					EndIf
				Next
			EndIf

		Case $TYPE_SWFEXE
			If RipExeInfo($file, $tempoutdir, "{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}") Then
				MoveFiles($tempoutdir, $outdir, False, '', True, True)
			Else
				$success = $RESULT_FAILED
			EndIf

		Case $TYPE_TAR
			_Run($7z & ' x "' & $file & '"', $outdir)

			If $fileext <> "tar" Then
				_Run($7z & ' x "' & $outdir & '\' & $filename & '.tar"', $outdir)
				FileDelete($outdir & '\' & $filename & '.tar')
			EndIf

		Case $TYPE_THINSTALL ; Test
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

		Case $TYPE_TTARCH
			If $ttarchfailed Then Return 0

			; Get all supported games
			$aReturn = _StringBetween(FetchStdout(Quote($bindir & $ttarch), @ScriptDir, @SW_HIDE, 0, False), "Games", "Examples")
			If @error Then terminate($STATUS_FAILED, $file, $arctype, $arcdisp)
			$aGames = StringRegExp($aReturn[0], "\d+ +(.+)", 3)
;~ 			_ArrayDisplay($aGames)

			Local $tmp = $aGames
			_ArraySort($tmp)

			; Display game select GUI
			Local $iChoice = GUI_MethodSelectList($tmp, t('METHOD_GAME_NOGAME'))
			Cout("Selected game: " & $iChoice)
			If $iChoice Then
				$iChoice = _ArraySearch($aGames, $iChoice)
				If $iChoice > -1 Then _Run($ttarch & ' -m ' & $iChoice & ' "' & $file & '" "' & $outdir & '"', $outdir, @SW_HIDE)
			Else
				$ttarchfailed = True
				$returnFail = True
			EndIf

		Case $TYPE_UHA
			_Run($uharc & ' x -t"' & $outdir & '" "' & $file & '"', $outdir)
			If Not $success And _DirGetSize($outdir, $initdirsize + 1) <= $initdirsize Then
				_Run($uharc04 & ' x -t"' & $outdir & '" "' & $file & '"', $outdir)
				If Not $success And _DirGetSize($outdir, $initdirsize + 1) <= $initdirsize Then _
					_Run($uharc02 & ' x -t' & FileGetShortName($outdir) & ' ' & FileGetShortName($file), $outdir)
			EndIf

		Case $TYPE_UIF
			_DeleteTrayMessageBox()
			Prompt(32 + 4, 'CONVERT_CDROM', 'UIF', True)

			; Begin conversion, format selected by uif2iso
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'UIF ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 1)')
			_Run($uif & ' "' & $file & '" "' & $outdir & "\" & $filename & '"', $filedir, True, True, True)
			$hSearch = FileFindFirstFile($outdir & "\" & $filename & ".*")
			$isofile = $outdir & "\" & FileFindNextFile($hSearch)
			FileClose($hSearch)

;~ 		Case $TYPE_UNITY
;~ 			_Run($unity & ' extract "' & $file & '"', $filedir, @SW_MINIMIZE, True, True, True, False)

		Case $TYPE_UNREAL ; Test
			HasPlugin($unreal)
			; Currently extracts files from all packages in folder, not only the selected one!
			_Run($unreal & ' -export -all -sounds -3rdparty -path="' & $filedir & '" -out="' & $outdir & '" *', $outdir, @SW_MINIMIZE, True, True, False)

		Case $TYPE_VIDEO
			HasFFMPEG()

			; Collect information about number of tracks
			$command = $ffmpeg & ' -i "' & $file & '"'
			$return = FetchStdout($command, $outdir, @SW_HIDE)

			; Terminate if file could not be read by FFmpeg
			If StringInStr($return, "Invalid data found when processing input") Or Not StringInStr($return, "Stream") Then terminate($STATUS_FAILED, $file, $arctype, $arcdisp)

			$aStreams = StringSplit($return, "Stream", 1)
			$iStreams = $aStreams[0] - 2
			Cout($iStreams & " streams found in file")
;~ 			_ArrayDisplay($aStreams)

			; We don't want to extract a .wma file from a .wma file
			If $fileext == "wma" And $iStreams < 2 Then extract($TYPE_AUDIO, t('TERM_AUDIO') & ' ' & t('TERM_FILE'))

			; Otherwise, extract all tracks
			Local $iVideo = 0, $iAudio = 0
			For $i = 2 To $aStreams[0]
				$aStreams[$i] = StringRegExpReplace($aStreams[$i], "(?i)(?s).*?#(\d:\d)(.*?): (\w+): (\w+).*", "$3,$4,$1,$2")
				$aStreamType = StringSplit($aStreams[$i], ",")
;~ 				_ArrayDisplay($aStreamType)

				If $aStreamType[1] == "Video" Then
					; Split gif files into single images
					If $aStreamType[2] == "gif" Or $aStreamType[2] == "apng" Or $aStreamType[2] == "webp" Then
						_Run($command & ' "' & GetFileName() & '%05d.png"', $outdir, @SW_HIDE, True, False)
						$iVideo += 1
						ContinueLoop
					EndIf

					If Not $bExtractVideo Then ContinueLoop
					$iVideo += 1
					If $aStreamType[2] == "h264" Then
						_Run(_MakeFFmpegCommand($command & ' -vcodec copy -an -bsf:v h264_mp4toannexb -map ', $aStreamType, t('TERM_VIDEO'), $iVideo), $outdir, @SW_HIDE, True, False)
					Else
						; Special cases
						If StringInStr($aStreamType[2], "wmv") Then
							$aStreamType[2] = "wmv" ;wmv3
						ElseIf StringInStr($aStreamType[2], "mpeg") Then
							$aStreamType[2] = "mpeg" ;mpeg1video
						ElseIf StringInStr($aStreamType[2], "vp8") Then
							$aStreamType[2] = "webm"
						ElseIf StringInStr($aStreamType[2], "flv") Then
							$aStreamType[2] = "flv" ;flv1
						EndIf
						_Run(_MakeFFmpegCommand($command & ' -vcodec copy -an -map ', $aStreamType, t('TERM_VIDEO'), $iVideo), $outdir, @SW_HIDE, True, True, False)
					EndIf
				ElseIf $aStreamType[1] == "Audio" Then
					$iAudio += 1
					; Special cases:
					; The stream type can be different from the file extension, so we need to change it for some files
					If StringInStr($aStreamType[2], "wma") Then
						$aStreamType[2] = "wma" ;wmav2
					ElseIf StringInStr($aStreamType[2], "vorbis") Then
						$aStreamType[2] = "ogg"
					ElseIf StringInStr($aStreamType[2], "pcm") Then
						$aStreamType[2] = "wav"
					EndIf

					_Run(_MakeFFmpegCommand($command & ' -acodec copy -vn -map ', $aStreamType, t('TERM_AUDIO'), $iAudio), $outdir, @SW_HIDE, True, True, False)
				Else
					Cout("Unknown stream type: " & $aStreamType[1])
				EndIf
			Next
			If $iVideo + $iAudio < 1 Then terminate($STATUS_NOTPACKED, $file, $arctype, $arcdisp)

		Case $TYPE_VIDEO_CONVERT
			HasFFMPEG()

			_Run($ffmpeg & ' -i "' & $file & '" "' & GetFileName() & '.mp4"', $outdir, @SW_HIDE, True, True, False)

		Case $TYPE_VISIONAIRE3
			Local $tmp = $outdir & "\names.txt"
			Local $f = $file

			If Not FileExists($tmp) Then
				For $i = 0 To 2
					Local $sPath = _PathFull(_StringRepeat("..\", $i), $filedir & "\")
					Cout("Searching for main data file in " & $sPath)

					Local $aReturn = _FileListToArray($sPath, "*.vis", $FLTA_FILES)
					If @error Then ContinueLoop

					_ArrayDelete($aReturn, 0)
					Local $sChoice = $aReturn[0]
					If UBound($aReturn) > 1 Then
						Local $sChoice = GUI_MethodSelectList($aReturn, t('METHOD_NOT_IN_LIST'), 'METHOD_FILE_SELECT_LABEL')
						If $sChoice < 1 Then ContinueLoop
					EndIf

					$sPath &= $sChoice
					Cout("Generating names with main file " & $sPath)
					_Run($visionaire3 & ' "' & $sPath & '" /force /generateNames="' & $tmp & '"', $outdir, @SW_HIDE, True, True, False, False)
					ExitLoop
				Next
			EndIf

			If FileGetSize($tmp) > 0 Then
				Cout("Names generated successfully. Extracting...")
				_Run($visionaire3 & ' "' & $file & '" /force /names="' & $tmp & '"', $outdir, @SW_HIDE, True, True, False, False)
			Else
				Cout("Failed to extract names. Some files may not be usable.")
				_Run($visionaire3 & ' "' & $file & '" /force', $outdir, @SW_HIDE, True, True, False, False)
			EndIf

		Case $TYPE_VSSFX ; Test
			_FileMove($file, $outdir)
			RunWait(Warn_Execute($outdir & '\' & $filenamefull & ' /extract'), $outdir)
			_FileMove($outdir & '\' & $filenamefull, $filedir)

		Case $TYPE_VSSFX_PATH ; Test
			RunWait(Warn_Execute($file & ' /extract:"' & $outdir & '" /quiet'), $outdir)

		Case $TYPE_WISE
			_Run($wise_ewise & ' "' & $file & '" "' & $outdir & '"', $filedir)
			If $success == $RESULT_FAILED Then
				$success = $RESULT_UNKNOWN
				Local $aOptions = ['Wise ' & t('TERM_INSTALLER'), t('METHOD_UNPACKER_RADIO', 'Wise UNpacker'), t('METHOD_SWITCH_RADIO', 'Wise Installer /x'), t('METHOD_EXTRACTION_RADIO', 'Wise MSI'), t('METHOD_EXTRACTION_RADIO', 'Unzip')]
				$iChoice = GUI_MethodSelect($aOptions, $arcdisp)

				Switch $iChoice
					; Extract with WUN
					Case 1
						RunWait(_MakeCommand($wise_wun, True) & ' "' & $filename & '" "' & $tempoutdir & '"', $filedir)

						Local $aCleanup[] = [$tempoutdir & "INST0*", $tempoutdir & "WISE0*"]
						Cleanup($aCleanup)
						MoveFiles($tempoutdir, $outdir, False, "", True)

					; Extract using the /x switch
					Case 2
						Warn_Execute($file & ' /x ' & $outdir)
						ShellExecuteWait($file, ' /x ' & $outdir, $filedir)

					; Attempt to extract MSI
					Case 3
						; Some Wise installers contain a msi installer, which is unpacked to CommonFilesDir & "\Wise Installation Wizard"
						; when the main file is executed. Trying to find the correct file inside this directory is unreliable, so we simply
						; search the msi inside the exe file.
						If RipExeInfo($file, $tempoutdir, "{DOWN}{DOWN}{DOWN}") Then MoveFiles($tempoutdir, $outdir, False, '', True, True)

					; Extract using unzip, falling back to 7-Zip
					Case 4
						_Run($zip & ' -x "' & $file & '"', $outdir)
						If $success == $RESULT_FAILED Then _Run($7z & ' x "' & $file & '"', $outdir)
				EndSwitch
			Else
				RunWait($cmd & '00000000.BAT', $outdir, @SW_HIDE)
				FileDelete($outdir & '\00000000.BAT')
			EndIf

		Case $TYPE_WIX
			HasNetFramework(4)
			_Run($wix & ' -x "' & $outdir & '" "' & $file & '"', $outdir, @SW_MINIMIZE, True, True, False)

		Case $TYPE_WOLF
			extract($TYPE_ARC_CONV, -1, "", False, True)
			HasPlugin($wolf)
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & "Wolf RPG Editor " & t('TERM_GAME') & t('TERM_ARCHIVE'))
			_RunInTempOutdir($tempoutdir, $wolf & ' ' & Quote($file), $outdir, @SW_MINIMIZE, True, True, False)
			MoveFiles($outdir & "\" & $filename, $outdir, True, '', True, True)

		Case $TYPE_ZIP
			If Not extract($TYPE_7Z, -1, $additionalParameters, False, True) Then
				If $arcdisp > -1 Then _CreateTrayMessageBox(t('EXTRACTING') & @CRLF & $arcdisp)
				_Run($zip & ' -x "' & $file & '"', $outdir, @SW_MINIMIZE, True, False)
			EndIf

		Case $TYPE_ZOO
			_FileMove($file, $tempoutdir, 8)
			_Run($zoo & ' -x ' & $filenamefull, $tempoutdir, @SW_HIDE)
			_FileMove($tempoutdir & $filenamefull, $file)
			MoveFiles($tempoutdir, $outdir, False, "", True)

		Case $TYPE_ZPAQ
			; ZPaq uses a different executable for Windows XP, so a definition file cannot be used
			_Run($zpaq & ' x "' & $file & '" -to "' & $outdir & '"', $outdir, @SW_SHOW, True, True, False)

		Case Else
			pluginExtract($arctype, $tempoutdir)
			If @error Then Cout("Unknown arctype: " & $arctype & ". Feature not implemented!")
	EndSwitch

	Opt("WinTitleMatchMode", 1)
	If Not $returnFail Then _DeleteTrayMessageBox()


	If $isofile And FileExists($isofile) And Not ($returnSuccess Or $returnFail) Then
		; Extract converted iso file
		If FileGetSize($isofile) > 0 Then
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & StringUpper($arctype) & ' ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 2)')
			$file = $isofile
			If Not CheckIso() Then _Run($7z & ' x "' & $isofile & '"', $outdir)
			If $success <> $RESULT_FAILED Or Prompt(16 + 4, 'CONVERT_CDROM_STAGE2_FAILED') = 0 Then
				$initdirsize -= FileGetSize($isofile)
				FileDelete($isofile)
			EndIf
		Else
			Prompt(16, 'CONVERT_CDROM_STAGE1_FAILED')
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
			If $arctype == $TYPE_7Z And ($fileext = "exe" Or StringInStr($sFileType, "SFX")) Then
				; Check if sfx archive and extract sfx script using 7ZSplit if possible
				_CreateTrayMessageBox(t('SCANNING_FILE', "7z SFX Archives splitter"))
				Cout("Trying to extract sfx script")
				Run(_MakeCommand($7zsplit & ' "' & $file & '"'), $outdir, @SW_HIDE)
				WinWait("7z SFX Archives splitter")
				ControlClick("7z SFX Archives splitter", "", "Button8")
				ControlClick("7z SFX Archives splitter", "", "Button1")
				$TimerStart = TimerInit()
				Do
					Sleep(100)
					If WinExists("7z SFX Archives splitter warning") Then WinClose("7z SFX Archives splitter warning")
					$TimerDiff = TimerDiff($TimerStart)
					If $TimerDiff > $Timeout Then ExitLoop
				Until FileExists($filedir & "\" & $filename & ".txt") Or WinExists("7z SFX Archives splitter error")
				; Force close all messages
				ProcessClose("7ZSplit.exe")

				; Move sfx script to outdir
				If FileExists($filedir & "\" & $filename & ".txt") Then _FileMove($filedir & "\" & $filename & ".txt", $outdir & "\sfx_script_" & $filename & ".txt")
				_DeleteTrayMessageBox()
			EndIf

		Case $RESULT_NOFREESPACE
			terminate($STATUS_NOFREESPACE)
		Case $RESULT_FAILED

		Case $RESULT_CANCELED

		Case $RESULT_UNKNOWN
			; Otherwise, check directory size
			If ($initdirsize > -1 And _DirGetSize($outdir, $initdirsize + 1) <= $initdirsize) Or (FileGetTime($outdir, 0, 1) == $dirmtime) Then
				If $arctype = "ace" And $fileext = "exe" Then Return False
				$success = $RESULT_FAILED
			EndIf
	EndSwitch

	If $success = $RESULT_FAILED Then
		If Not $returnFail Then terminate($STATUS_FAILED, $file, $arctype, $arcdisp)
		$success = $RESULT_UNKNOWN
		Return 0
	EndIf

	If Not $returnSuccess Then terminate($STATUS_SUCCESS, $filenamefull, $arctype, $arcdisp)
	$success = $RESULT_UNKNOWN
	Return 1
EndFunc

Func pluginExtract($sPlugin, $tempoutdir)
	Cout("Starting custom " & $sPlugin & " extraction")

	Local $sPluginFile = $userDefDir & $sPlugin & ".ini"
	If Not FileExists($sPluginFile) Then
		$sPluginFile = $defdir & $sPlugin & ".ini"
		If Not FileExists($sPluginFile) Then terminate($STATUS_MISSINGDEF, $sPluginFile, $sPlugin)
	EndIf

	Local Const $sSection = "Plugin"
	Local $aIniSection = IniReadSection($sPluginFile, "Plugin")

	Local $sBinary = _ArrayGet($aIniSection, "executable", $sPlugin)
	HasPlugin($sBinary)

	; Dependencies
	Local $ret = _ArrayGet($aIniSection, "requireNetFramework", 0)
	If $ret > 0 Then HasNetFramework($ret)

	; Set status box
	If Not $bOptNoStatusBox Then
		Local $arcdisp = t('EXTRACTING') & @CRLF & _ArrayGet($aIniSection, "display", $sPlugin)
		$arcdisp = ReplacePlaceholders($arcdisp)
		_CreateTrayMessageBox($arcdisp)
	EndIf

	Local $sParameters = " " & _ArrayGet($aIniSection, "parameters", "")
	Local $sWorkingDir = _ArrayGet($aIniSection, "workingdir", "")
	Local $bRunInTempDir = _ArrayGet($aIniSection, "runInTempOutdir", 0, True) == 1
	Local $show_flag = _ArrayGet($aIniSection, "hide", 0, True) == 1? @SW_HIDE: @SW_MINIMIZE
	Local $bUseCmd = _ArrayGet($aIniSection, "useCmd", 1, True) == 1
	Local $bUseTee = _ArrayGet($aIniSection, "log", 1, True) == 1
	Local $bPatternSearch = _ArrayGet($aIniSection, "patternSearch", 0, True) == 1
	Local $bInitialShow = _ArrayGet($aIniSection, "initialShow", 1, True) == 1

	If Not $sWorkingDir Or $sWorkingDir = "" Then $sWorkingDir = $outdir

	$sParameters = ReplacePlaceholders($sParameters)
	$sWorkingDir = StringReplace($sWorkingDir, "%tempoutdir%", $tempoutdir)
	$sWorkingDir = ReplacePlaceholders($sWorkingDir)

	If $bRunInTempDir Then
		_RunInTempOutdir($tempoutdir, $sBinary & $sParameters, $sWorkingDir, $show_flag, $bUseCmd, $bUseTee, $bPatternSearch, $bInitialShow)
	Else
		_Run($sBinary & $sParameters, $sWorkingDir, $show_flag, $bUseCmd, $bUseTee, $bPatternSearch, $bInitialShow)
	EndIf

	Local $sCleanup = IniRead($sPluginFile, $sSection, "cleanup", 0)
	If Not $sCleanup Or StringLen($sCleanup) < 1 Then Return

	$aCleanup = StringSplit($sCleanup, "|", 2)
	Cleanup($aCleanup)
EndFunc

; Replace % placeholders with variable contents
Func ReplacePlaceholders($sString, $bQuote = True)
	If Not StringInStr($sString, "%") Then Return $sString

	$sString = StringReplace($sString, "%filename%", $filename)
	$sString = StringReplace($sString, "%fileext%", $fileext)
	$sString = StringReplace($sString, "%filedir%", $filedir)
	$sString = StringReplace($sString, "%file%", $bQuote? Quote($file): $file)
	$sString = StringReplace($sString, "%outdir%", $bQuote? Quote($outdir): $outdir)

	$aReturn = _StringBetween($sString, "%", "%")
	If @error Then Return $sString

	For $sPlaceholder In $aReturn
		If StringInStr($sPlaceholder, " ") Then ContinueLoop
		$sString = StringReplace($sString, "%" & $sPlaceholder & "%", t($sPlaceholder))
	Next

	Return $sString
EndFunc

; Load a BMS script from the database and start extraction
Func BmsExtract($sName, $hDB = 0)
	If Not $sName Then Return
	Cout('Extracting using BMS script "' & $sName & '"')

	If $hDB == 0 Then $hDB = OpenDB("BMS.db")

	If $hDB Then
		Local $aReturn[0], $iRows, $iColumns
		_SQLite_GetTable($hDB, Cout("SELECT s.Script FROM Scripts s, Names n WHERE s.SID = n.NID AND Name = '" & $sName & "'"), $aReturn, $iRows, $iColumns)
;~ 		_ArrayDisplay($aReturn)

		; Write script to file and execute it
		$bmsScript = FileOpen($bindir & $bms, $FO_OVERWRITE)
		FileWrite($bmsScript, $aReturn[2])
		FileClose($bmsScript)
		$return = FetchStdout($quickbms & ' -l "' & $bindir & $bms & '" "' & $file & '"', $filedir, @SW_HIDE, -1)

		If Not StringInStr($return, "0 files found") And Not StringInStr($return, "Error") And Not StringInStr($return, "invalid") _
		And Not StringInStr($return, "expected: ") And $return <> "" Then
			_SQLite_Close($hDB)
			_SQLite_Shutdown()
			extract($TYPE_QBMS, $sName & " " & t('TERM_PACKAGE'), $bms)
		EndIf
	EndIf

	terminate($STATUS_FAILED, $filenamefull, $sName, $sName)
EndFunc

; Start SQLite and open a database
Func OpenDB($sName)
	_SQLite_Startup()
	If @error Then Return Cout("[ERROR] SQLite startup failed with code " & @error)

	$hDB = _SQLite_Open($bindir & $sName, $SQLITE_OPEN_READONLY)
	If @error Then Return Cout("[ERROR] Failed to open database " & $sName)

	Return $hDB
EndFunc

; Retrieve array value be key
Func _ArrayGet(ByRef $aData, $sKey, $sDefault, $bNumber = False, $iKeyIndex = 0, $iValueIndex = 1)
	Local $ret = _ArraySearch($aData, $sKey, 0, 0, 0, 0, 1, $iKeyIndex)
	If $ret < 0 Then Return $sDefault

	Local $return = $aData[$ret][$iValueIndex]
	Return $bNumber? Number($return): $return
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
	If Not Prompt(32 + 4, 'UNPACK_PROMPT', CreateArray($packer, $ret)) Then Return

	; Unpack file
	Switch $packer
		Case $PACKER_UPX
			_Run($upx & ' -d -k "' & $file & '"', $filedir)
			$tempext = StringTrimRight($fileext, 1) & '~'
			If FileExists($filedir & "\" & $filename & "." & $tempext) Then
				_FileMove($file, $ret)
				_FileMove($filedir & "\" & $filename & "." & $tempext, $file)
			EndIf
		Case $PACKER_ASPACK
			RunWait($aspack & ' "' & $file & '" "' & $filedir & '\' & $filename & '_' & t('TERM_UNPACKED') & _
					$fileext & '" /NO_PROMPT', $filedir)
	EndSwitch

	; Success evaluation
	If FileExists($ret) Then
		; Prompt if unpacked file should be scanned
		If Prompt(32 + 4, 'UNPACK_AGAIN', CreateArray($file, $filename & '_' & t('TERM_UNPACKED') & "." & $fileext)) Then
			$file = $ret
			$outdir = $filedir & "\" & $filename & "_" & t('TERM_UNPACKED') & "\"
			StartExtraction()
		Else
			terminate($STATUS_SUCCESS, $filenamefull, $packer, $packer)
		EndIf
	Else
		Prompt(16, 'UNPACK_FAILED', $file)
	EndIf
EndFunc

; Perform outdir cleanup: move/delete given files according to $iCleanup setting
Func Cleanup($aFiles, $iMode = $iCleanup, $sDestination = 0)
	If Not $iMode Then Return
	If Not IsArray($aFiles) Then
		If Not FileExists($aFiles) And Not FileExists($outdir & "\" & $aFiles) Then
			Cout("Cleanup: Invalid path " & $aFiles)
			Return SetError(1, 0, 0)
		EndIf
		$return = $aFiles
		Dim $aFiles = [$return]
	EndIf

	If $iMode = $OPTION_MOVE And $sDestination == 0 Then $sDestination = $outdir & "\" & t('DIR_ADDITIONAL_FILES')
;~ 	Cout("Cleanup - " & _ArrayToString($aFiles))

	For $sFile In $aFiles
		If Not StringInStr($sFile, $outdir) Then $sFile = $outdir & "\" & $sFile
		If Not FileExists($sFile) Then ContinueLoop
		$bIsFolder = _IsDirectory($sFile)
		If $iMode = $OPTION_DELETE Then
			Cout("Cleanup: Deleting " & $sFile)
			If $bIsFolder Then
				DirRemove($sFile, 1)
			Else
				FileDelete($sFile)
			EndIf
		Else
			If Not FileExists($sDestination) Then DirCreate($sDestination)
			Cout("Cleanup: Moving " & $sFile & " to " & $sDestination)
			If $bIsFolder Then
				_DirMove($sFile, $sDestination, 1)
			Else
				_FileMove($sFile, $sDestination, 1)
			EndIf
		EndIf
	Next
EndFunc

; Check write permissions for specified file or folder
Func CanAccess($sPath)
	Local $bIsTemp = False

	Cout("Checking permissions for path " & $sPath)
	$bExists = FileExists($sPath)
	If _IsDirectory($sPath) Or ($bExists = False And StringRight($sPath, 1)) Then
		$bIsTemp = True
		$sPath = _TempFile($sPath)
	Else
		If Not $bExists Then Return SetError(1, 0, False)
	EndIf

	Local $hFile = FileOpen($sPath, $FO_APPEND)
	If $hFile == -1 Then
		Cout("Access denied")
		Return False
	EndIf
	FileClose($hFile)

	If $bIsTemp Then FileDelete($sPath)
	Return True
EndFunc

; Terminate if specified plugin was not found
Func HasPlugin($sPlugin, $returnFail = False)
	$sPlugin = StringReplace($sPlugin, '"', '')
	Cout("Searching for plugin " & $sPlugin)
	If FileExists($sPlugin) Or (_WinAPI_PathIsRelative($sPlugin) And (FileExists(_PathFull($sPlugin, $bindir)) Or FileExists(_PathFull($sPlugin, $archdir)))) Then Return True

	Cout("Plugin not found")
	If $returnFail Then Return False
	If $silentmode Then terminate($STATUS_MISSINGEXE, $file, $sPlugin, $sPlugin)

	Opt("GUIOnEventMode", 0)
	Local $hGUI = GUICreate($name, 416, 176, -1, -1, $GUI_SS_DEFAULT_GUI)
	_GuiSetColor()
	GUICtrlCreateLabel(t('PLUGIN_MISSING', CreateArray($filenamefull, t('SELECT_FILE'))), 72, 20, 330, 113)
	Local $idPluginManager = GUICtrlCreateButton(t('MENU_HELP_PLUGINS_LABEL'), 194, 142, 123, 25)
	Local $idCancel = GUICtrlCreateButton(t('EXIT_BUT'), 332, 142, 75, 25)
	_GUICtrlCreatePic($sLogoFile, 8, 20, 49, 49)
	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $idCancel
				GUIDelete($hGUI)
				SendStats($STATUS_MISSINGEXE, $sPlugin)
				terminate($STATUS_SILENT)
			Case $idPluginManager
				GUI_Plugins($hGUI, $sPlugin)
				Opt("GUIOnEventMode", 0)
				If HasPlugin($sPlugin, True) Then ExitLoop
		EndSwitch
	WEnd

	Opt("GUIOnEventMode", 1)
	GUIDelete($hGUI)
	Return True
EndFunc

; Search for translation file for given language and return result
Func HasTranslation($language)
	If $language = "English" Then Return True
	$return = FileExists($langdir & $language & ".ini")
	If Not $return Then Cout("Language file for " & $language & " does not exist")
	Return $return
EndFunc

; Check if enough free space is available
Func HasFreeSpace($sPath = $outdir, $fModifier = 2)
	While Not _IsDirectory($sPath)
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
		If $silentmode Then terminate($STATUS_FAILED, $filenamefull, $STATUS_NOFREESPACE, $Msg)

		$return = MsgBox($iTopmost + 48 + 2, $name, $Msg)
		If $return = 4 Then ; Retry
			Return HasFreeSpace($sPath, $fModifier)
		ElseIf $return = 3 Then ; Cancel
			If $createdir Then DirRemove($outdir, 0)
			terminate($STATUS_SILENT)
		EndIf
	EndIf
EndFunc

; Search for FFMPEG and prompt to download it if not found
Func HasFFMPEG()
	If HasPlugin($ffmpeg, True) Then Return
	If $silentmode Then terminate($STATUS_MISSINGEXE, $filenamefull, "FFmpeg", "FFmpeg")

	Opt("GUIOnEventMode", 0)
	Local $sTranslation = t('TERM_DOWNLOAD')

	Local $hGUI = GUICreate($name, 416, 201, 438, 245, $GUI_SS_DEFAULT_GUI)
	_GuiSetColor()
	GUICtrlCreateLabel(t('FFMPEG_NEEDED', CreateArray($filenamefull, $sTranslation)), 72, 20, 330, 107)
	Local $idDownload = GUICtrlCreateButton($sTranslation, 242, 166, 75, 25)
	Local $idCancel = GUICtrlCreateButton(t('CANCEL_BUT'), 332, 166, 75, 25)
	_GUICtrlCreatePic($sLogoFile, 8, 20, 49, 49)
	Local $idCheckbox = GUICtrlCreateCheckbox(t('FFMPEG_LICENSE_ACCEPT_LABEL'), 10, 138, 259, 17)
	Local $idViewLicense = GUICtrlCreateLabel(t('FFMPEG_LICENSE_VIEW_LABEL'), 266, 140, 138, 17, $SS_RIGHT)
	_GuiCtrlLinkFormat()
	Local $idSelectFile = GUICtrlCreateLabel(t('FFMPEG_SELECT_LABEL'), 10, 172, 133, 17)
	_GuiCtrlLinkFormat()
	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $idCancel
				GUIDelete($hGUI)
				terminate($STATUS_SILENT)
			Case $idDownload
				If GUICtrlRead($idCheckbox) = $GUI_CHECKED Then
					GUIDelete($hGUI)
					GetFFmpeg()
					If @error Then terminate($STATUS_SILENT)
					ExitLoop
				EndIf
				MsgBox($iTopmost + 48, $name, t('LICENSE_NOT_ACCEPTED'))
			Case $idViewLicense
				ShellExecute("https://ffmpeg.org/legal.html")
			Case $idSelectFile
				Local $tmp = @WorkingDir
				$sPath = FileOpenDialog(t('FFMPEG_SELECT_TITLE'), _GetFileOpenDialogInitDir(), "FFmpeg (ffmpeg.exe)", 1, "", $hGUI)
				If @error Or Not FileExists($sPath) Then ContinueLoop

				; Make sure the executable is really FFmpeg
				GUICtrlSetState($idDownload, $GUI_DISABLE)
				GUICtrlSetState($idSelectFile, $GUI_DISABLE)
				Local $ret = FetchStdout($sPath, @WorkingDir, @SW_HIDE, 0, True, True, False)
				FileChangeDir($tmp)
				GUICtrlSetState($idDownload, $GUI_ENABLE)
				GUICtrlSetState($idSelectFile, $GUI_ENABLE)

				; Shared version is not supported, because the DLLs cannot be found with a hardlinked executable
				If FileGetSize($sPath) < 1024 * 1024 Or Not StringInStr($ret, "ffmpeg version") Then
					MsgBox($iTopmost + 16, $title, t('FFMPEG_INVALID_FILE'))
					ContinueLoop
				EndIf
				GUIDelete($hGUI)
				Cout("FFmpeg selected: " & $sPath)

				Local $sDestination = StringReplace($ffmpeg, '"', '')
				; Create a hardlink because:
				;  -FFmpeg is available even after deleting the linked file
				;  -Compatibility with Windows XP (symlinks not available on older OS)
				;  -Yes, using the direct path to FFmpeg would be the cleanest solution, but the whole extraction logic
				;   is built around having everything in \bin. Saving a few MB is not worth the work necessary to change
				;   all functions to support extractors outside this directory.
				If FileCreateNTFSLink($sPath, $sDestination) Then ExitLoop

				; If creating a hardlink fails, simply copy the binary to \bin directory
				If FileCopy($sPath, $sDestination) Then ExitLoop

				MsgBox($iTopmost + 16, $title, t('FFMPEG_MOVE_FAILED'))
				terminate($STATUS_SILENT)
		EndSwitch
	WEnd

	GUIDelete($hGUI)

	Opt("GUIOnEventMode", 1)
EndFunc

; Determine versions of installed .NET frameworks
; Modified version of (https://www.autoitscript.com/forum/topic/164455-check-net-framework-4-or-45-is-installed/#comment-1199620)
Func HasNetFramework($iVersion, $bTerminate = True)
	Cout("Searching for .NET Framework " & $iVersion)
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

	If $bTerminate Then terminate($STATUS_MISSINGEXE, $file, ".Net Framework " & $iVersion)
    Return False
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
Func ReturnFiles($sDir)
	Local $hSearch, $files, $fname
	$hSearch = FileFindFirstFile($sDir & "\*")
	If @error Then Return SetError(1)

	While 1
		$fname = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		$files &= $fname & '|'
	WEnd

	$files = StringTrimRight($files, 1)
	FileClose($hSearch)
	Return $files
EndFunc

; Create a copy of the input file and change its extension
Func CreateRenamedCopy($sExtension)
	Cout("Creating copy with extension ." & $sExtension)
	Prompt(64 + 1, "FILE_COPY", $file, True)
	Local $tmp = _TempFile($filedir, $filename & "_", "." & $sExtension)
	If Not FileCopy($file, $tmp) Then Return False
	$file = $tmp
	FilenameParse($file)
	Global $iDeleteOrigFile = $OPTION_DELETE
EndFunc

; Append missing file extensions using TrID
Func AppendExtensions($dir)
	Cout("Fixing file extensions")

	Local $files = FileSearch($dir)
	If $files[1] = '' Then Return
	For $i = 1 To $files[0]
		_SetTrayMessageBoxText(t('RENAMING_FILES', CreateArray($i, $files[0])))
		If _IsDirectory($files[$i]) Then ContinueLoop
		$filename = StringTrimLeft($files[$i], StringInStr($files[$i], '\', 0, -1))
		If StringInStr($filename, '.') And Not (StringLeft($filename, 7) = 'Binary.' Or StringRight($filename, 4) = '.bin') Then ContinueLoop
		RunWait(Cout(_MakeCommand($trid & ' "' & $files[$i] & '" -ae')), $dir, @SW_HIDE)
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

	If UBound($s_Buf) == 2 And $s_Buf[1] = '' Then
		ReDim $s_Buf[1]
		$s_Buf[0] = 0
		SetError(1)
	EndIf

	Return $s_Buf
EndFunc

; Open file and return contents
Func _FileRead($f, $bDelete = False, $iFlag = 0)
	Cout("Reading file " & $f)
	Local $hFile = FileOpen($f, $iFlag)
	If $hFile = -1 Then Return SetError(1, 0, "")

	$return = FileRead($hFile)
	FileClose($hFile)
	If $iFlag <> $FO_BINARY Then Cout($return)

	If $bDelete Then FileDelete($f)
	Return $return
EndFunc

; Delete a file and retry if it fails
Func _FileDelete($sFile, $iSleep = 100)
	If Not FileExists($sFile) Then Return SetError(1, 0, False)
	If _IsDirectory($sFile) Then Return SetError(2, 0, False)

	If FileDelete($sFile) Then Return True

	Cout('Failed to delete file "' & $sFile & '", retrying')
	If $iSleep > 0 Then Sleep($iSleep)
	If _WinAPI_DeleteFile($sFile) Then Return True
	Cout("Failed again, error " & _WinAPI_GetLastError() & ": " & _WinAPI_GetLastErrorMessage())
EndFunc

; Handle program termination with appropriate error message
Func terminate($status, $fname = '', $arctype = '', $arcdisp = '')
	Local $LogFile = 0, $exitcode = 0, $sFileType = _FiletypeGet(False), $shortStatus = ($status = $STATUS_SUCCESS)? $arctype: $status

	; Rename unicode file
	If $iUnicodeMode Then
		Cout("Renaming unicode file")
		_CreateTrayMessageBox(t('MOVING_FILE') & @CRLF & $oldoutdir)
		If $iUnicodeMode = $UNICODE_MOVE Then
			_FileMove($file, $oldpath, 1)
		Else
			If Not FileRecycle($file) Then Cout("Failed to recycle file")
		EndIf
		Cout("Moving extracted files: " & _DirMove($outdir, $oldoutdir))
		$fname = $sUnicodeName
		$file = $oldpath
		$outdir = $oldoutdir
	EndIf

	_DeleteTrayMessageBox()
	Cout("Terminating - Status: " & $status)

	; When multiple files are selected and executed via command line, they are added to batch queue, but the working instance uses in-memory data.
	; So we need to look for changes in the batch queue file, so batch mode could be enabled if necessary.
	If Not $silentmode And GetBatchQueue() Then $silentmode = True

	; Create log file if enabled in options
	If $Log And Not ($status = $STATUS_SILENT Or $status = $STATUS_SYNTAX Or $status = $STATUS_FILEINFO Or $status = $STATUS_NOTPACKED Or $status = $STATUS_BATCH) Or ($status = $STATUS_FILEINFO And $silentmode) Then _
		$LogFile = CreateLog($shortStatus)

	; Save local statistics
	IniWrite($prefs, "Statistics", $status, Number(IniRead($prefs, "Statistics", $status, 0)) + 1)
	If $status = $STATUS_SUCCESS Then IniWrite($prefs, "Statistics", $arctype, Number(IniRead($prefs, "Statistics", $arctype, 0)) + 1)

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
			; If UniExtract was started with invalid parameters, a timeout is set to keep batch processing alive
			MsgBox($iTopmost + 32, $title, $syntax, $arctype? 15: 0)

		; Display file type information and exit
		Case $STATUS_FILEINFO
			If $silentmode Then ; Save info to result file if in silent mode
				Local $hFile = FileOpen($fileScanLogFile, $FO_CREATEPATH + $FO_APPEND)
				FileWrite($hFile, $file & @CRLF & @CRLF & $sFileType & @CRLF & "------------------------------------------------------------" & @CRLF)
				FileClose($hFile)
			Else
				If UBound($aFiletype) < 1 Then
					GUI_Error_UnknownExt()
					$exitcode = 4
				Else
					_GUI_FileScan()
				EndIf
			EndIf

		; Display error information and exit
		Case $STATUS_UNKNOWNEXE
			GUI_Error_UnknownExt()
			$exitcode = 3
		Case $STATUS_UNKNOWNEXT
			GUI_Error_UnknownExt()
			$exitcode = 4
		Case $STATUS_INVALIDFILE
			Prompt(16, 'INVALID_FILE', $fname)
			$exitcode = 5
		Case $STATUS_INVALIDDIR
			Prompt(16, 'INVALID_DIR', $fname)
			$exitcode = 5
		Case $STATUS_NOTPACKED
			Prompt(48, 'NOT_PACKED', CreateArray($file, $sFileType))
			$exitcode = 6
		Case $STATUS_NOTSUPPORTED
			Prompt(16, 'NOT_SUPPORTED', CreateArray($file, StringReplace(t('MENU_HELP_FEEDBACK_LABEL'), "&", ""), StringReplace(t('MENU_HELP_LABEL'), "&", "")))
			$exitcode = 7
		Case $STATUS_MISSINGEXE
			Prompt(48, 'MISSING_EXE', CreateArray($file, $arctype))
			$exitcode = 8
		Case $STATUS_TIMEOUT
			Prompt(48, 'EXTRACT_TIMEOUT', $file)
			$exitcode = 9
		Case $STATUS_PASSWORD
			Prompt(48, 'WRONG_PASSWORD', CreateArray($file, StringReplace(t('MENU_EDIT_LABEL'), "&", "")))
			$exitcode = 10
		Case $STATUS_MISSINGDEF
			Prompt(48, 'MISSING_DEFINITION', CreateArray($file, $fname))
			$exitcode = 11
		Case $STATUS_MOVEFAILED
			Prompt(48, 'MOVE_FAILED', CreateArray($file, $fname))
			$exitcode = 12
		Case $STATUS_NOFREESPACE
			Prompt(48, 'NO_FREE_SPACE_ERROR', $file)
			$exitcode = 13
		Case $STATUS_MISSINGPART
			Prompt(48, 'MISSING_PART', $file)
			$exitcode = 14

			; Display failed attempt information and exit
		Case $STATUS_FAILED
			If Not $silentmode And Prompt(256 + 16 + 4, 'EXTRACT_FAILED', CreateArray($file, $arcdisp)) Then ShellExecute($LogFile? $LogFile: CreateLog($status))
			$exitcode = 1

			; Exit successfully
		Case $STATUS_SUCCESS
			If $iDeleteOrigFile = $OPTION_DELETE Or ($iDeleteOrigFile = $OPTION_ASK And Not $silentmode And Prompt(32 + 4, 'FILE_DELETE', $file)) Then
				Cout("Deleting source file " & $file)
				FileDelete($file)
			EndIf
			If $bOptOpenOutDir And Not $silentmode Then ShellExecute($outdir)
	EndSwitch

	; Write error log if in batchmode
	If $exitcode <> 0 And $silentmode And $extract Then
		Local $hFile = FileOpen($logdir & "errorlog.txt", $FO_CREATEPATH + $FO_APPEND)
		Local $sMsg = GetDateTime() & " " & ($filenamefull = ""? $fname: $filenamefull) & " (" & StringUpper($status)& ")" & " - " & $arctype  & @CRLF
		FileWrite($hFile, $sMsg)
		FileClose($hFile)
	EndIf

	; Delete empty output directory if failed
	If $createdir And $status <> $STATUS_SUCCESS And DirGetSize($outdir) = 0 Then DirRemove($outdir, 1)

	If ($exitcode == 1 Or $exitcode == 3 Or $exitcode == 4 Or $exitcode == 7 Or $exitcode == 12) And $fileext <> "dll" Then GUI_Feedback_Prompt()

	If $batchEnabled = 1 And $status <> $STATUS_SILENT Then ; Don't start batch if gui is closed
		; Start next extraction
		BatchQueuePop()
	ElseIf $KeepOpen And $cmdline[0] = 0 And $status <> $STATUS_SILENT Then
		Run(@ScriptFullPath)
	EndIf

	; Check for updates
	If $status <> $STATUS_SILENT Then
		If Not $silentmode Then CheckUpdate($UPDATEMSG_FOUND_ONLY, True, $UPDATE_ALL, False)
		SendStats($status, $arctype)
	EndIf

	Exit $exitcode
EndFunc

; Create array on the fly
; Code based on _CreateArray UDF, which has been deprecated
Func CreateArray($i0, $i1 = 0, $i2 = 0, $i3 = 0, $i4 = 0, $i5 = 0, $i6 = 0, $i7 = 0, $i8 = 0, $i9 = 0)
	Local $arr[10] = [$i0, $i1, $i2, $i3, $i4, $i5, $i6, $i7, $i8, $i9]
	ReDim $arr[@NumParams]
	Return $arr
EndFunc

; Show tray message box
; Based on work by Valuater (http://www.autoitscript.com/forum/topic/85977-system-tray-message-box-udf/)
Func _CreateTrayMessageBox($TBText)
	_DeleteTrayMessageBox()

	If $bOptNoStatusBox = 1 Then Return

	; Hide if in fullscreen
	If $bHideStatusBoxIfFullscreen Then
		$aReturn = WinGetPos("[ACTIVE]")
		If $aReturn[2] = @DesktopWidth And $aReturn[3] = @DesktopHeight Then Return
	EndIf

	Local Const $TBwidth = 225, $TBheight = 100, $left = 15, $top = 12, $width = 200, $iBetween = 5, $iMaxCharCount = 28, $bDark = True

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
	Local $fname = $filename == ""? "": GetFileName() & "." & $fileext
	If StringLen($fname) > $iMaxCharCount Then $fname = StringLeft($fname, $iMaxCharCount) & " [...]"
	If StringLen($TBText) > $iMaxCharCount * 2 Then $TBText = StringLeft($TBText, $iMaxCharCount * 2) & " [...]"
	Local $idTrayFileName = GUICtrlCreateLabel($fname, $left, $top, $width, 16, $SS_LEFTNOWORDWRAP)
	Local $idTrayStatus = GUICtrlCreateLabel($TBText, $left, GetPos($TBgui, $idTrayFileName, 6, False), $width, 30)
	Global $idTrayStatusExt = GUICtrlCreateLabel("", $left, GetPos($TBgui, $idTrayStatus, 8, False), $width, 20, $SS_CENTER)

	GUICtrlSetFont($idTrayFileName, 9, 500, 0, $FONT_ARIAL)
	GUICtrlSetFont($idTrayStatus, 9, 500, 0, $FONT_ARIAL)
	GUICtrlSetFont($idTrayStatusExt, 9, 500, 0, $FONT_ARIAL)

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

; Set tray message extended status text
Func _SetTrayMessageBoxText($sText)
	If Not $TBgui Then Return 0
	Return GUICtrlSetData($idTrayStatusExt, $sText)
EndFunc

; Close tray message box
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

; Test if a file is a known multipart archive and already in batch queue
Func IsMultipartArchive($sBatchQueueContent)
	If Not $filenamefull Then FilenameParse($file)

	Return TestMultipart('(.*?\.part)(\d+\.rar)', $sBatchQueueContent) Or _
		   TestMultipart('(.*?\.7z.)(\d{3})', $sBatchQueueContent) Or _
		   TestMultipart('(.*?\.r)((\d{2})|ar)', $sBatchQueueContent)
EndFunc

; Test if a file matches a given regex and compare capture group with batch queue content
Func TestMultipart($sRegEx, $sBatchQueueContent)
;~ 	Cout("Testing " & $sRegEx)
	Local $ret = StringRegExpReplace($filenamefull, $sRegEx, "$1", 1)
	Return @extended > 0 And StringInStr($sBatchQueueContent, $ret)
EndFunc

; Create command line for current file
Func GetCmd($silent = True)
	If Not $file Then Return SetError(1)
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
EndFunc

; Add file to batch queue
Func AddToBatch()
	Local $cmdline = GetCmd()
	If @error Then Return Cout("Failed to add file to batch queue: invalid file parameter: " & $file)

	Local $hFile = FileOpen($batchQueue, $FO_UNICODE + $FO_CREATEPATH + $FO_APPEND)
	If @error Then Return Cout("Failed to open batch queue")
;~ 	FileSetPos($hFile, 0, 0)
	Local $sBatchQueueContent = FileRead($hFile)

	Local $bAddFile = True
	If StringInStr($sBatchQueueContent, $cmdline) Then
		$bAddFile = CustomPrompt('BATCH_DUPLICATE', $filenamefull)
	Else
		; Only add one file if multipart archive
		$bAddFile = Not IsMultipartArchive($sBatchQueueContent)
	EndIf

	If Not $bAddFile Then
		Cout("Not adding duplicate file " & $filenamefull)
		FileClose($hFile)
		Return
	EndIf

	FileWrite($hFile, $cmdline & @CRLF)
	FileClose($hFile)
	Cout("File added to batch queue: " & $cmdline)
	EnableBatchMode()
EndFunc

; Read batch queue from file
Func GetBatchQueue()
	Local $hFile = FileOpen($batchQueue, $FO_UNICODE)
	$queueArray = FileReadToArray($hFile)
	FileClose($hFile)

	Local $iSize = UBound($queueArray)
	If IsArray($queueArray) And $iSize > 0 Then
;~ 		_ArrayDisplay($queueArray)
		If $guimain Then GUICtrlSetData($BatchBut, t('BATCH_BUT') & " (" & $iSize & ")")
		EnableBatchMode()
		Return 1
	EndIf

	Return 0
EndFunc

; Write batch queue array to file
Func SaveBatchQueue()
	Cout("Saving batch queue")
	Local $hFile = FileOpen($batchQueue, $FO_UNICODE + $FO_CREATEPATH + $FO_OVERWRITE)
	FileWrite($hFile, _ArrayToString($queueArray, @CRLF))
	FileClose($hFile)
EndFunc

; Returns first element of batch queue
Func BatchQueuePop()
;~ 	_ArrayDisplay($queueArray)
	If Not IsArray($queueArray) Or UBound($queueArray) < 1 Then GetBatchQueue()

	If Not IsArray($queueArray) Or UBound($queueArray) < 1 Then
		Cout("Batch queue empty")
		EnableBatchMode(False)
		If FileExists($fileScanLogFile) Then ShellExecute($fileScanLogFile)
		Local $return = _FileRead($logdir & "errorlog.txt", True)
		If $return <> "" Then MsgBox($iTopmost + 48, $name, t('BATCH_FINISH', $return))
		If $KeepOpen Then Run(@ScriptFullPath)
	Else ; Get next command and execute it
		Local $element = $queueArray[0]
		_ArrayDelete($queueArray, 0)
		Cout("Next batch element: " & $element)
		SaveBatchQueue()
		Run(@ScriptFullPath & " " & $element)
	EndIf
EndFunc

; Enable batch mode
Func EnableBatchMode($bEnable = True)
	If $bEnable Then
		; Delete old filescan log file
		_FileDelete($fileScanLogFile)

		If $guimain Then
			GUICtrlSetOnEvent($GUI_Main_Ok, "GUI_Batch_OK")
			GUICtrlSetState($showitem, $GUI_ENABLE)
			GUICtrlSetState($clearitem, $GUI_ENABLE)
		EndIf
	Else
		; Delete empty batch queue file
		_FileDelete($batchQueue)

		If $guimain Then
			GUICtrlSetOnEvent($GUI_Main_Ok, "GUI_OK")
			GUICtrlSetData($BatchBut, t('BATCH_BUT'))
			GUICtrlSetState($showitem, $GUI_DISABLE)
			GUICtrlSetState($clearitem, $GUI_DISABLE)
		EndIf
	EndIf

	$batchEnabled = $bEnable
	SavePref("batchenabled", Number($batchEnabled))
EndFunc

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

; Determine whether a path is a directory or not
Func _IsDirectory($sPath)
	; Wildcards should not be considered directories
	If StringRight($sPath, 1) == "*" Then Return False

	Return StringInStr(FileGetAttrib($sPath), "D")
EndFunc

; Determine if a key exists in registry
; Script by guinness (http://www.autoitscript.com/forum/topic/131425-registry-key-exists/page__view__findpost__p__915063)
Func RegExists($sKeyName, $sValueName)
	RegRead($sKeyName, $sValueName)
	Return Number(@error = 0)
EndFunc

; Return a specific line of a multi line string
; https://www.autoitscript.com/forum/topic/103821-how-to-read-specific-line-from-a-string/
Func _StringGetLine($sString, $iLine, $bCountBlank = False)
	Local $sChar = "+"
	If $bCountBlank = True Then $sChar = "*"
	If Not IsInt($iLine) Then Return SetError(1, 0, "")
	If $iLine < 0 Then Return StringTrimLeft($sString, StringInStr($sString, @CRLF, 0, -1 + $iLine))
	Return StringRegExpReplace($sString, "((." & $sChar & "\n){" & $iLine - 1 & "})(." & $sChar & "\n)((." & $sChar & "\n?)+)", "\2")
EndFunc

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
    $oDir = $oShellApp.NameSpace($sDir)
    $oFile = $oDir.Parsename($sFile)
    If $iProp = -1 Then
        Local $aProperty[35]

        For $i = 0 To 34
            $aProperty[$i] = $oDir.GetDetailsOf($oFile, $i)
        Next

		; Remove empty cells
		For $i = 34 To 0 Step -1
            If $aProperty[$i] = "" Then _ArrayDelete($aProperty, $i)
        Next

		Return $aProperty
	Else
        Local $sProperty = $oDir.GetDetailsOf($oFile, $iProp)
        If $sProperty = "" Then Return 0
		Return $sProperty
    EndIf
EndFunc

; Compress with zlib
; Based on https://www.autoitscript.com/forum/topic/87284-zlib-udf
Func _Zlib_Compress($Data)
	If Not IsBinary($Data) Then $Data = StringToBinary($Data, 4)
	Local Const $iCompression = 9
	Local $hDll = DllOpen($bindir & "zlib1.dll")
	If @error Then Return SetError(1)

	$aInput = DllStructCreate("byte[" & BinaryLen($Data) + 1 & "]")
	DllStructSetData($aInput, 1, $Data)

	$aOutput = DllStructCreate("byte[" & Round(BinaryLen($Data) * 1.0001) + 12 & "]")

	Local $ret = DllCall($hDll, "int:cdecl", "compress2", "ptr", DllStructGetPtr($aOutput), "long*", DllStructGetSize($aOutput), "ptr", DllStructGetPtr($aInput), "long", DllStructGetSize($aInput), "int", $iCompression)
	If $ret[0] <> 0 Then Return $ret[0]

	$ret2 = DllStructGetData(DllStructCreate("byte[" & $ret[2] & "]", DllStructGetPtr($aOutput)), 1)

	DllClose($hDll)
	Return $ret2
EndFunc

; Error handler for COM objects, currently used for sending feedback
Func _ComErrorHandler($oError)
	Global $sComError = $oError.description & "(0x" & Hex($oError.number) & ") in " & $oError.source
EndFunc

; Dump complete debug content to log file
Func CreateLog($status)
	Local $sName = $logdir & @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & "-" & @MIN & "-" & @SEC & "_"
	If $status <> $STATUS_SUCCESS Then $sName &= StringUpper($status)
	If $file <> "" Then $sName &= "_" & GetFileName() & "." & $fileext
	$sName &= ".log"
	$hFile = FileOpen($sName, $FO_UNICODE + $FO_CREATEPATH + $FO_OVERWRITE)
	FileWrite($hFile, $sFullLog)
	FileClose($hFile)
	Return $sName
EndFunc

; Determine whether the archive is password protected or not and try passwords from list if necessary
Func _FindArchivePassword($sIsProtectedCmd, $sTestCmd, $sIsProtectedText = "encrypted", $sIsProtectedText2 = 0, $iLine = -3, $sTestText = "All OK")
	; Is archive encrypted?
	Local $return = FetchStdout(_MakeCommand($sIsProtectedCmd, True), $outdir, @SW_HIDE, $iLine)
	If Not StringInStr($return, $sIsProtectedText) And ($sIsProtectedText2 == 0 Or Not StringInStr($return, $sIsProtectedText2)) Then Return 0

	Cout("Archive is password protected")
	_SetTrayMessageBoxText(t('SEARCHING_PASSWORD'))
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
		_SetTrayMessageBoxText(t('TESTING_PASSWORD', CreateArray($i, $size)))
		If StringInStr(FetchStdout(StringReplace($sTestCmd, "%PASSWORD%", $aPasswords[$i], 1), $outdir, @SW_HIDE, 0, False), $sTestText) Then
			Cout("Password found")
			$sPassword = $aPasswords[$i]
			ExitLoop
		EndIf
	Next

	_SetTrayMessageBoxText("")
	Return $sPassword
EndFunc

; Execute a program and log output using tee
Func _Run($f, $sWorkingDir = $outdir, $show_flag = @SW_MINIMIZE, $bUseCmd = True, $bUseTee = True, $bPatternSearch = True, $bInitialShow = True)
	Global $run = 0, $runtitle = 0
	Local $return = "", $pos = 0, $size = 1, $lastSize = 0
	Local Const $LogFile = $logdir & "teelog.txt"

	$f = _MakeCommand($f, $bUseCmd) & ($bUseTee? ' 2>&1 | ' & $tee & ' "' & $LogFile & '"': '')

	Cout("Executing: " & $f)
	Cout("           with options: showFlag = " & $show_flag & ", initialShow = " & $bInitialShow & ", patternSearch = " & $bPatternSearch & ", workingdir = " & $sWorkingDir)

	; Create log
	If $bUseTee Then
		HasPlugin($tee)
		If Not FileExists($logdir) Then DirCreate($logdir)

		$run = Run($f, $sWorkingDir, $bInitialShow? @SW_MINIMIZE: $show_flag)
		If @error Then
			Cout("Failed to execute command")
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
		If $bInitialShow Then WinSetState($runtitle, "", $show_flag)
		Cout("Runtitle: " & $runtitle)

		; Wait until logfile exists
		$TimerStart = TimerInit()

		Do
			Sleep(10)
			If TimerDiff($TimerStart) > 5000 Then ExitLoop
		Until FileExists($LogFile)
		Local $hFile = FileOpen($LogFile)
		$state = ""

		; Show progress (percentage) in status box
		While ProcessExists($run)
			$return = FileRead($hFile)
;~ 			Cout($return)
			If $return <> $state Then
				$state = $return
				; Automatically show cmd window when user input needed
				If StringInStr($return, "already exist") Or StringInStr($return, "overwrite") Or StringInStr($return, " replace") _
				Or StringInStr($return, "password") Or StringInStr($return, "Not enough free space available") _
				Or StringInStr($return, "you must choose a new filename") Or StringInStr($return, "Insert disk with") Then
					Cout("User input needed")
					WinSetState($runtitle, "", @SW_SHOW)
					GUICtrlSetFont($idTrayStatusExt, 8.5, 900)
					_SetTrayMessageBoxText(t('INPUT_NEEDED'))
					WinActivate($runtitle)
					$lastSize = Round((_DirGetSize($outdir, 0) - $initdirsize) / 1024 / 1024, 3)
					ContinueLoop
				EndIf
				; Percentage indicator
				If $bPatternSearch And _PatternSearch($return) Then $size = -1
			EndIf
			; Size of extracted file(s) as fallback
			If $size > -1 And $bPatternSearch > -1 Then
				$size = Round((_DirGetSize($outdir) - $initdirsize) / 1024 / 1024, 3)
;~ 				Cout("Size: " & $size & @TAB & $lastSize)
				If $size > 0 And $size <> $lastSize Then
					Cout("Size: " & $size & @TAB & $lastSize)
					_SetTrayMessageBoxText($size & " MB")
				EndIf
				$lastSize = $size
				Sleep(50)
			EndIf
			Sleep(100)
		WEnd
		; Write tee log to UniExtract log file
		FileSetPos($hFile, 0, $FILE_BEGIN)
		$return = FileRead($hFile)
		If Not StringIsSpace($return) Then Cout("Teelog:" & @CRLF & $return)
		FileClose($hFile)
		FileDelete($LogFile)

		; Check for success or failure indicator in log
		If StringInStr($return, "Wrong password?", 0) Or StringInStr($return, "The specified password is incorrect.", 0) Or _
		   StringInStr($return, "Archive encrypted.", 0) Or StringInStr($return, "Corrupt file or wrong password", 0) Or _
		   StringInStr(_StringGetLine($return, -1), "Enter password") Then
			$success = $RESULT_FAILED
			SetError(1, 1)
		ElseIf StringInStr($return, "Break signaled") Or StringInStr($return, "Program aborted") Or StringInStr($return, "User break") Then
			Cout("Cancelled by user")
			$success = $RESULT_CANCELED
		ElseIf StringInStr($return, "There is not enough space on the disk") Then
			$success = $RESULT_NOFREESPACE
			SetError(2)
		ElseIf StringInStr($return, "You need to start extraction from a previous volume") Or _
			   StringInStr($return, "Unavailable start of archive") Or StringInStr($return, "Missing volume") Then
			$success = $RESULT_FAILED
			SetError(3)
		ElseIf StringInStr($return, "Everything is Ok") Or _
			   StringInStr($return, "0 failed") Or StringInStr($return, "All files OK") Or _
			   StringInStr($return, "All OK") Or StringInStr($return, "done.") Or _
			   StringInStr($return, "Done ...") Or StringInStr($return, ": done") Or _
			   StringInStr($return, "Result:	Successful, errorcode 0") Or StringInStr($return, "... Successful") Or _
			   StringInStr($return, "Extract files [ ") Or StringInStr($return, "Done; file is OK") Or _
			   StringInStr($return, "Successfully extracted to") Then
			Cout("Success evaluation passed")
			$success = $RESULT_SUCCESS
		ElseIf StringInStr($return, "err code(", 1) Or StringInStr($return, "stacktrace", 1) _
			   Or StringInStr($return, "Write error: ", 1) Or (StringInStr($return, "Cannot create", 1) _
			   And StringInStr($return, "No files to extract", 1)) Or StringInStr($return, "Archives with Errors: 1") _
			   Or StringInStr($return, "ERROR: Wrong tag in package", 1) Or StringInStr($return, "unzip:  cannot find", 1) _
			   Or StringInStr($return, "Open ERROR: Can not open the file as") Or StringInStr($return, "Error: System.Exception:") _
			   Or StringInStr($return, "unknown WISE-version -> contact author") Then
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
		$run = Run($f, $sWorkingDir, $show_flag)
		If @error Then
			Cout("Failed to execute command")
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
			If $size > 0 And $bPatternSearch > -1 Then
				If $TBgui Then _SetTrayMessageBoxText($size & " MB")
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

; Move file to tempoutdir and use _Run to execute a program
; $file is automatically replaced with the new temporary path
Func _RunInTempOutdir($tempoutdir, $f, $sWorkingDir = $outdir, $show_flag = @SW_MINIMIZE, $bUseCmd = True, $bUseTee = True, $bPatternSearch = True, $bInitialShow = True)
	Local $tmp = $tempoutdir & $filenamefull
	_FileMove($file, $tempoutdir, 8)
	$f = StringReplace($f, $file, $tmp)

	_Run($f, $sWorkingDir, $show_flag, $bUseCmd, $bUseTee, $bPatternSearch, $bInitialShow)

	_FileMove($tmp, $file)
	MoveFiles($tempoutdir, $outdir, True, "", True, True)
EndFunc

; Search console output for progress indicator patterns and update status box
Func _PatternSearch($sString)
	If Not $TBgui Then Return False

	Local $iNum
	Static $sTranslation = t('TERM_FILE') & " "
;~ 	Cout($sString & _StringRepeat(@CRLF, 5))

;~ 	Cout("x %")
	If StringInStr($sString, "%", 0, -1) Then
		$aReturn = StringRegExp($sString, "(\d{1,3})[\d\.,]* ?%", 3)
;~ 		_ArrayDisplay($aReturn)
		If UBound($aReturn) > 0 Then Return _SetTrayMessageBoxText(_ArrayPop($aReturn) & "%")
	EndIf

;~ 	Cout("[x on y]")
	If StringInStr($sString, " on ", 0, -1) Then
		$aReturn = StringRegExp($sString, "\[(\d+) on (\d+)\]", 3)
		If UBound($aReturn) > 1 Then
			$iNum = _ArrayPop($aReturn)
			Return _SetTrayMessageBoxText($sTranslation & _ArrayPop($aReturn) & "/" & $iNum)
		EndIf
	EndIf

;~ 	Cout("x of y")
	If StringInStr($sString, " of ", 0, -1) Then
		$aReturn = StringRegExp($sString, "(\d+) of (\d+)", 3)
		If UBound($aReturn) > 1 Then
			$iNum = _ArrayPop($aReturn)
			Return _SetTrayMessageBoxText($sTranslation & _ArrayPop($aReturn) & "/" & $iNum)
		EndIf
	EndIf

;~ 	Cout("x/y")
	If StringInStr($sString, "/", 0, -1) Then
		$aReturn = StringRegExp($sString, "(\d+)/(\d+)", 3)
		If UBound($aReturn) > 1 Then
			$iNum = _ArrayPop($aReturn)
			Return _SetTrayMessageBoxText($sTranslation & _ArrayPop($aReturn) & "/" & $iNum)
		EndIf
	EndIf

;~ 	Cout("# x")
	$pos = StringInStr($sString, "#", 0, -1)
	If $pos Then
		$iNum = Number(StringMid($sString, $pos + 1), 1)
		If $iNum > 0 Then Return _SetTrayMessageBoxText($sTranslation & "#" & $iNum)
	EndIf
EndFunc

; Run a program and return stdout/stderr stream
Func FetchStdout($f, $sWorkingDir, $show_flag = @SW_HIDE, $iLine = 0, $bOutput = True, $bUseCmd = True, $bMakeCommand = True)
	Global $run = 0, $return = ""

	If $bMakeCommand Then $f = _MakeCommand($f, $bUseCmd)
	If $bOutput Then Cout("Executing: " & $f)
	$run = Run($f, $sWorkingDir, $show_flag, $STDERR_MERGED)
	If @error Then Return SetError(1, 0, -1)

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
Func _MakeCommand($f, $bUseCmd = False)
;~ 	Cout("MakeCommand: " & $f)
	If StringInStr($f, $cmd) Then Return $f

	If Not StringInStr($f, $bindir) Then
		$pos = StringInStr($f, " ")
		If $pos > 1 And $bUseCmd And FileExists($bindir & StringLeft($f, $pos)) Then
			$f = '""' & $bindir & _StringInsert($f, '"', $pos - 1)
		Else
			$f = $bindir & $f
		EndIf
	EndIf
	Return ($bUseCmd? $cmd: "") & $f
EndFunc

; Build command line for FFMPEG extractions
Func _MakeFFmpegCommand($sPrefix, $aStreamType, $sType, $iIndex)
	Local $sName = GetFileName()

	While StringLeft($sName, 1) == "-"
		$sName = StringTrimLeft($sName, 1)
	WEnd

	Return $sPrefix & $aStreamType[3] & ' "' & $sName & "_" & $sType & StringFormat("_%02s", $iIndex) & $aStreamType[4] & "." & $aStreamType[2] & '"'
EndFunc

; DirGetSize wrapper with additional logic
Func _DirGetSize($f, $return = -1)
	; Calculating the size of a whole drive would take way too much time,
	; so let's only calculate size if less than 4 GB space used on drive
	If (StringLen($f) < 4 And DriveSpaceTotal($f) - DriveSpaceFree($f) > 4000) Then Return $return
	Return DirGetSize($f)
EndFunc

; DirMove wrapper with error handling and auto-retry
Func _DirMove($sDir, $sDestination, $iFlag = 0)
	Return MovePath($sDir, $sDestination, $iFlag, True)
EndFunc

; FileMove wrapper with error handling and auto-retry
Func _FileMove($sFile, $sDestination, $iFlag = 0)
	Return MovePath($sFile, $sDestination, $iFlag, False)
EndFunc

; Move file/folder specified by path with error handling and auto-retry
Func MovePath($sPath, $sDestination, $iFlag = 0, $bIsFolder = False)
	Cout("Moving " & $sPath & " to " & $sDestination)

	If Not FileExists($sPath) Then
		Cout("Error: input file does not exist")
		Return SetError(1, 0, False)
	EndIf

	If $bIsFolder Then
		If DirMove($sPath, $sDestination, $iFlag) Then Return True
	Else
		If FileMove($sPath, $sDestination, $iFlag) Then Return True
	EndIf

	Cout("Failed to move " & ($bIsFolder? "directory": "file") & ", retrying")
	Sleep($bIsFolder? 100: 50)
	If _WinAPI_MoveFileEx($sPath, $sDestination, BitOR($MOVE_FILE_COPY_ALLOWED, $iFlag)) Then Return True

	Local $iError = _WinAPI_GetLastError()
	Cout("Failed again, error " & $iError & ": " & _WinAPI_GetLastErrorMessage())
	Return SetError($iError, 0, False)
EndFunc

; Move all files and subdirectories from one directory to another
; $force is an integer that specifies whether or not to replace existing files
; $omit is a string that includes files to be excluded from move
Func MoveFiles($source, $dest, $force = False, $omit = '', $removeSourceDir = False, $bShowStatus = False)
	Local $hSearch, $fname, $iCount = 0, $iErrors = 0
	Static $sTranslation = t('TERM_FILE') & " "

	Cout("Moving files from " & $source & " to " & $dest)
	If $bShowStatus Then _CreateTrayMessageBox(t('MOVING_FILE') & @CRLF & $dest)
	DirCreate($dest)

	$hSearch = FileFindFirstFile($source & "\*")
	If @error Then Return SetError(1)
	While 1
		$fname = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		If StringInStr($omit, $fname) Then ContinueLoop
		$iCount += 1

		_SetTrayMessageBoxText($sTranslation & $iCount)
		Local $sPath = $source & '\' & $fname
		If _IsDirectory($sPath) Then
			If Not _DirMove($sPath, $dest, 1) Then
				$iErrors += 1
				Cout("Failed to move directory " & $fname)
			EndIf
		Else
			If Not _FileMove($sPath, $dest, $force) Then
				$iErrors += 1
				Cout("Failed to move file " & $fname)
			EndIf
		EndIf
	WEnd
	FileClose($hSearch)

	If $iErrors > 0 Then Cout($iErrors & " files/folders could not be moved")
	If $bShowStatus Then _DeleteTrayMessageBox()
	If $removeSourceDir Then Return DirRemove($source, ($omit = "" And $iErrors < 1? 1: 0))
EndFunc

; Return the path to the download directory
Func _GetFileOpenDialogInitDir()
	Local $sDir = _WinAPI_ShellGetKnownFolderPath($FOLDERID_Downloads)
	If @error Or $sDir == "" Or Not FileExists($sDir) Then $sDir = @WorkingDir
	Return $sDir
EndFunc

; Add new scan result to filetype array
Func _FiletypeAdd($sScanner, $sType)
	If StringRight($sType, 2) == @CRLF Then $sType = StringTrimRight($sType, 2)
	If Not $sType Or $sType == "" Then Return

	Local $iPos = _ArraySearch($aFiletype, $sScanner, 0, 0, 0, 0, 1, 0)
	If $iPos > -1 Then
		$aFiletype[$iPos][1] &= @CRLF & $sType
		Return
	EndIf

	Local $iSize = UBound($aFiletype)
	ReDim $aFiletype[$iSize + 1][2]
	$aFiletype[$iSize][0] = $sScanner
	$aFiletype[$iSize][1] = $sType
EndFunc

; Return formatted file scan results
Func _FiletypeGet($bHeader = True, $iWidth = 50)
	Local $return = "", $tmp, $iSize, $sPadding

	For $i = 0 To UBound($aFiletype) - 1
		If $return <> "" Then $return &= @CRLF & @CRLF
		If Not $bHeader Then
			$return &= $aFiletype[$i][1]
			ContinueLoop
		EndIf

		$tmp = $aFiletype[$i][0]
		If $iWidth > 0 Then
			$tmp = " " & $tmp & " "
			$iSize = Floor(($iWidth - StringLen($tmp)) / 2)
			$sPadding = _StringRepeat("-", $iSize)
			$tmp = $sPadding & $tmp & $sPadding
		EndIf
		$return &= $tmp & @CRLF & @CRLF & $aFiletype[$i][1]
	Next

	Return $return
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

; Restart Universal Extractor
Func Restart()
	Run(@ScriptFullPath)
	terminate($STATUS_SILENT)
EndFunc

; Restart Universal Extractor without elevated privileges
Func RestartWithoutAdminRights($sParameters = "")
	Run($cmd & 'runas /trustlevel:0x20000 "' & @ScriptFullPath & $sParameters & '"')
	terminate($STATUS_SILENT)
EndFunc

; Return current date and time
Func GetDateTime()
	Return @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
EndFunc

; Write data to stdout stream if enabled in options
Func Cout($sMsg)
	Local $sOutput = GetDateTime() & ":" & @MSEC & @TAB & $sMsg & @CRLF
	If Not @Compiled Then ConsoleWrite($sOutput)
	$sFullLog &= $sOutput
	Return $sMsg
EndFunc

; Open URL and evaluate success
Func OpenURL($sURL, $hParent = 0)
	ShellExecute($sURL)
	If @error Then InputBox($title, t('OPEN_URL_FAILED'), $sURL, "", -1, -1, Default, Default, 0, $hParent)
EndFunc

; Send usage statistics if enabled in options
Func SendStats($a, $sResult = 1)
	If Not $bSendStats Then Return

	InetRead(Cout($sStatsURL & $a & "&r=" & $sResult & "&id=" & $ID & "&v=" & $sVersion), 1)
EndFunc

; Check for new version
; $silent is used for automatic update check, supressing any error and 'no update found' messages
Func CheckUpdate($silent = $UPDATEMSG_PROMPT, $bCheckInterval = False, $iMode = $UPDATE_ALL, $bShowProgress = True)
	If @NumParams > 1 And $bCheckInterval And _DateDiff("D", $lastupdate, _NowCalc()) < $updateinterval Then Return
	If @NumParams < 1 Then
		$silent = $UPDATEMSG_PROMPT
		$iMode = $UPDATE_ALL
		$bShowProgress = True
	EndIf

	Cout("Checking for update")
	Local $return = 0, $found = False

	; Get index
	$aReturn = _UpdateGetIndex("", $silent == $UPDATEMSG_SILENT Or $silent == $UPDATEMSG_FOUND_ONLY)
	If Not IsArray($aReturn) Then Return Cout("Failed to get update file listing")

	; Save date of last check for update
	Global $lastupdate = @YEAR & "/" & @MON & "/" & @MDAY
	; In case of missing files, CheckUpdate can be run without any preferences being loaded
	If StringLen($prefs) > 0 Then SavePref('lastupdate', $lastupdate)

	; UniExtract main executable - calling the updater is always necessary, because an executable file cannot overwrite itself while running
	If $iMode <> $UPDATE_HELPER Then
		If ($aReturn[0])[1] <> FileGetSize($sUniExtract) Or StringTrimLeft(_Crypt_HashFile($sUniExtract, $CALG_MD5), 2) <> ($aReturn[0])[2] Then
			Cout("Update available")
			$found = True
			If GUI_UpdatePrompt() Then
				Local $sParameters = "/main"
				If $bOptNightlyUpdates == 1 Then $sParameters &= " /nightly"
				If CanAccess(@ScriptDir) Then
					If Not ShellExecute($sUpdaterNoAdmin, $sParameters) Then MsgBox($iTopmost + 16, $title, t('UPDATE_FAILED'))
				Else
					If Not ShellExecute($sUpdater, $sParameters) Then MsgBox($iTopmost + 16, $title, t('UPDATE_NOADMIN'))
				EndIf
				Exit
			Else
				; If the user does not want to install the main update, let's not bother him with more 'update found' messages
				$iMode = $UPDATE_MAIN
				SendStats("UpdateMain", 0)
			EndIf
		EndIf
	EndIf

	; Other files - we can overwrite the files without a seperate updater
	If $iMode <> $UPDATE_MAIN Then
		If CheckUpdateHelpers($aReturn, $bShowProgress) Then
			If $bShowProgress Then
				_ProgressSet(100)
				Sleep(200)
				_ProgressOff()
			EndIf
			$found = True
			If Prompt(48 + 4, 'UPDATE_PROMPT', t('UPDATE_TERM_PROGRAM_FILES')) Then
				If Not CanAccess($bindir) Then
					If Not ShellExecute($sUpdater, "/helper") Then MsgBox($iTopmost + 16, $title, t('UPDATE_NOADMIN'))
					Exit
				EndIf
				If Not _UpdateHelpers($aReturn) And Not $silentmode Then MsgBox($iTopmost + 16, $title, t('UPDATE_FAILED'))
			Else
				SendStats("UpdateHelpers", 1)
			EndIf
		EndIf
		If _UpdateFFmpeg($bShowProgress) Then $found = True
	EndIf

	If $found = False Then
		SendStats("CheckUpdate", 0)
		If $silent == $UPDATEMSG_PROMPT Then MsgBox($iTopmost + 64, $name, t('UPDATE_CURRENT'))
	EndIf
	Cout("Check for updates finished")

	If IsAdmin() Then RestartWithoutAdminRights()
EndFunc

; Compare program files with server index to find if any file has an updated version available
Func CheckUpdateHelpers($aFiles, $bShowProgress = True)
	If $bShowProgress Then _ProgressOn(t('UPDATE_STATUS_SEARCHING'), $guimain)
	Local $i = 1, $iSize = UBound($aFiles)

	While $i < $iSize
		$a = $aFiles[$i]
		$i += 1
		If $bShowProgress Then _ProgressSet(($i / _Max($iSize, 200)) * 100)
		$sPath = @ScriptDir & "\" & $a[0]
		If $sPath == @ScriptFullPath Then ContinueLoop

;~ 		Cout($sPath)
		If Not _UpdateFileCompare($sPath, $a) Then ContinueLoop

		; If it's a file and the size differs, update necessary
		If StringRight($a[0], 1) <> "/" Then Return True

		; Directory
		If Not FileExists($sPath) Then Return True

		$aReturn = _UpdateGetIndex($a[0])
		If Not IsArray($aReturn) Then ContinueLoop

		_ArrayAdd($aFiles, $aReturn)
		$iSize = UBound($aFiles)
	WEnd

	If $bShowProgress Then _ProgressOff()
	Return False
EndFunc

; Download updated program files and display status
Func _UpdateHelpers($aFiles)
	Local $sText = t('TERM_DOWNLOADING', 0, $language, "Downloading") & "... "

	Local $hGUI = GUICreate($title, 434, 130, -1, -1, $WS_POPUPWINDOW, -1, $guimain)
	Local $idLabel = GUICtrlCreateLabel($sText, 16, 16, 408, 17)
	GUICtrlCreateLabel(t('TERM_OVERALL_PROGRESS', 0, $language, "Overall progress") & ":", 16, 72, 80, 17)
	Local $idProgressCurrent = GUICtrlCreateProgress(16, 32, 408, 25)
	Local $idProgressTotal = GUICtrlCreateProgress(16, 88, 406, 25)
	_GuiSetColor()
	GUISetState(@SW_SHOW)

	Local $i = 0, $iSize = UBound($aFiles), $iProgress = 0, $success = True

	While $i < $iSize
		; Update progress
		Local $ret = (($i + 1) / $iSize) * 100
		$iProgress = $ret > $iProgress? $ret: $iProgress + 0.2
		GUICtrlSetData($idProgressTotal, $iProgress)

		$a = $aFiles[$i]
		$i += 1
		$sPath = @ScriptDir & "\" & $a[0]
		If $sPath == @ScriptFullPath Then ContinueLoop

		If Not _UpdateFileCompare($sPath, $a) Then ContinueLoop

		GUICtrlSetData($idProgressCurrent, 0)
		If StringRight($a[0], 1) = "/" Then ; Directory
			If Not FileExists($sPath) Then DirCreate($sPath)
			GUICtrlSetData($idLabel, t('UPDATE_STATUS_SEARCHING', 0, $language, "Searching for updates..."))
			$aReturn = _UpdateGetIndex($a[0])
			If Not IsArray($aReturn) Then
				$success = False
				ContinueLoop
			EndIf
			_ArrayAdd($aFiles, $aReturn)
			$iSize = UBound($aFiles)
		Else
			GUICtrlSetData($idLabel, $sText & $a[0] & " (" & $a[1] & " bytes" & ")")

			; Failsafe. Overwriting existing files fails under certain (unknown) conditions. Deleting the old file beforehand helps.
			_FileDelete($sPath)

			Local $iBytesReceived = 0
			Local $hDownload = InetGet($sUpdateURL & $a[0], $sPath, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)

			; Update progress bar
			While Not InetGetInfo($hDownload, $INET_DOWNLOADCOMPLETE)
				Sleep(50)
				Local $iError = InetGetInfo($hDownload, $INET_DOWNLOADERROR)
				If $iError <> 0 Then
					Cout("Download failed: error " & $iError)
					$success = False
					ContinueLoop 2
				EndIf
				$iBytesReceived = InetGetInfo($hDownload, $INET_DOWNLOADREAD)
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
Func _UpdateFFmpeg($bShowProgress = True)
	If Not HasPlugin($ffmpeg, True) Then Return False
	If $bShowProgress Then _ProgressOn(t('UPDATE_STATUS_SEARCHING'), $guimain)

	; Determine FFmpeg version
	Local $return = FetchStdout($ffmpeg, @ScriptDir, @SW_HIDE, 0, False)
	If $bShowProgress Then _ProgressSet(50)
	Local $aReturn = _StringBetween($return, "ffmpeg version ", " Copyright")
	Local $sVersion = @error? 0: $aReturn[0] ; In case ffmpeg exists but crashes, redownload it

	$return = _INetGetSource(_IsWinXP()? $sLegacyUpdateURL & "?get=ffversion&xp": $sUpdateURL & "ffmpeg")
	Cout("FFmpeg: " & $sVersion & " <--> " & $return)
	If $bShowProgress Then _ProgressSet(100)
	Sleep(300)
	If $bShowProgress Then _ProgressOff()

	; Download new
	If $return > $sVersion Then
		Cout("Found an update for FFmpeg")
		If Prompt(48 + 4, 'UPDATE_PROMPT', CreateArray("FFmpeg", $sVersion, $return)) Then Return GetFFmpeg()
	EndIf

	Return False
EndFunc

; Helper function for updater, downloads index file for subdirectories
Func _UpdateGetIndex($sURL = "", $bSilent = $silentmode)
	$sURL = $sUpdateURL & $sURL & "index"
;~ 	Cout("Sending request: " & $sURL)

	$return = _INetGetSource($sURL)
	If @error Then Return _UpdateCheckFailed($bSilent)

	$aReturn = StringSplit($return, @LF, 2)
;~ 	_ArrayDisplay($aReturn)

	For $i = 0 To UBound($aReturn) - 1
		$aReturn[$i] = StringSplit($aReturn[$i], ",", 2)
		If @error Then Return _UpdateCheckFailed($bSilent)
	Next

	Return $aReturn
EndFunc

; Return size of a file or directory, ignoring plugins, which are not on the server
Func _UpdateGetSize($sPath)
	If Not _IsDirectory($sPath) Then Return FileGetSize($sPath)
	$sPath = StringReplace($sPath, "/", "\")
;~ 	Cout("GetSize: " & $sPath)
	Local $iSize = DirGetSize($sPath)

	; Don't include plugins in calculations
	If $sPath = @ScriptDir & "\docs\" Then
		If FileExists($sPath & "FFmpeg") Then $iSize -= DirGetSize($sPath & "FFmpeg") ; FFmpeg does not exist on server
	ElseIf $sPath = $bindir & "x86\" Or $sPath = $bindir & "x64\" Then
		$iSize -= FileGetSize($sPath & "ffmpeg.exe")
	ElseIf $sPath = $bindir Then
		$ret = DirGetSize($bindir & "crass-0.4.14.0")
		If $ret > 0 Then $iSize -= $ret

		Local $aReturn[] = ["x86\ffmpeg.exe", "x64\ffmpeg.exe", "arc_conv.exe", "Extractor.exe", "iscab.exe", "ISTools.dll", _
							"umodel.exe", "SDL2.dll", "dcp_unpacker.exe", "ci-extractor.exe", "gea.dll", "gentee.dll", _
							"dgcac.exe", "bootimg.exe", "I5comp.exe", "ZD50149.DLL", "ZD51145.DLL", "sim_unpacker.exe"]

		For $i In $aReturn
			$iSize -= FileGetSize($bindir & $i)
		Next
	EndIf
	Return $iSize
EndFunc

; Compare file size and hash with the server value
Func _UpdateFileCompare($sPath, $a)
	Local $iSize = _UpdateGetSize($sPath)

	; Directory
	If StringRight($a[0], 1) = "/" Then
		If $iSize == $a[1] Then Return False

	ElseIf $iSize == $a[1] Then
		Local $sHash = _Crypt_HashFile($sPath, $CALG_MD5)
		If $sHash <> -1 Then $sHash = StringLower(StringTrimLeft($sHash, 2))
		If $sHash == $a[2] Then Return False
		Cout($a[0] & ": " & $sHash & " - " & $a[2])
		Return True
	EndIf

	Cout($a[0] & ": " & $iSize & " - " & $a[1])
	Return True
EndFunc

; Display update failed message
Func _UpdateCheckFailed($bSilent = $silentmode)
	If Not $bSilent Then MsgBox($iTopmost + 48, $title, t('UPDATECHECK_FAILED'))
	Return False
EndFunc

; Custom styled ProgressOn replacement
Func _ProgressOn($sText, $hParent)
	$hProgress = GUICreate($title, 270, 54, -1, -1, $WS_POPUPWINDOW, -1, $hParent)
	$idProgress = GUICtrlCreateProgress(4, 6, 259, 25)
	GUICtrlCreateLabel($sText, 6, 36, 261, 17, $SS_CENTER)
	_GuiSetColor()
	GUISetState(@SW_SHOW)
EndFunc

; ProgressSet replacement
Func _ProgressSet($iPercent)
	GUICtrlSetData($idProgress, $iPercent)
EndFunc

; ProgressOff replacement
Func _ProgressOff()
	GUIDelete($hProgress)
EndFunc

; Delete a file from both OSArch subdirectories
Func _DeleteFromArchDir($sFile)
	FileDelete($bindir & "x86\" & $sFile)
	FileDelete($bindir & "x64\" & $sFile)
EndFunc

; Perform special actions after update, e.g. delete files
Func _AfterUpdate()
	; Move files
	FileMove($bindir & "x86\sqlite3.dll", @ScriptDir)
	FileMove($bindir & "x64\sqlite3.dll", @ScriptDir & "\sqlite3_x64.dll")
	If FileExists($docsdir & "7zip_readme.txt") Then MoveFiles($docsdir, $licensedir, True)

	; Remove unused files
	FileDelete($bindir & "faad.exe")
	FileDelete($bindir & "MediaInfo64.dll")
	FileDelete($bindir & "extract.exe")
	FileDelete($bindir & "dmgextractor.jar")
	FileDelete($bindir & "RPGDecrypter.exe")
	FileDelete($bindir & "mpq.wcx")
	FileDelete($bindir & "mpq.wcx64")
	FileDelete($bindir & "Expander.exe")
	FileDelete($bindir & "stuffit5.engine-5.1.dll")
	FileDelete($bindir & "FLVExtractCL.exe")
	FileDelete($bindir & "zpaqxp.exe")
	FileDelete($bindir & "unrar.exe")
	FileDelete($bindir & "xace.exe")
	FileDelete($bindir & "disunity.bat")
	FileDelete($bindir & "disunity.jar")
	FileDelete($bindir & "extractMHT.exe")
	FileDelete($bindir & "MhtUnPack.wcx")
	FileDelete($bindir & "STIX_D.exe")
	FileDelete($bindir & "WDOSXLE.exe")
	FileDelete($bindir & "wtee.exe")
	FileDelete($bindir & "ns2dec.exe")
	FileDelete($bindir & "EXTRNT.EXE")
	FileDelete($bindir & "ethornell.exe")
	FileDelete($bindir & "libpng12.dll")
	FileDelete($bindir & "brunsdec.exe")

	FileDelete($defdir & "flv.ini")
	FileDelete($defdir & "ns2.ini")
	FileDelete($defdir & "bruns.ini")
	FileDelete($licensedir & "flac_authors.txt")
	FileDelete($licensedir & "flac_readme.txt")
	FileDelete($licensedir & "Expander_license.txt")
	FileDelete($licensedir & "flvextractcl_icons.txt")
	FileDelete($licensedir & "bcm_readme.txt")
	FileDelete($licensedir & "wixtoolset_source.nz")
	FileDelete($licensedir & "disunity_license.md")
	FileDelete($licensedir & "disunity_readme.md")
	FileDelete($licensedir & "xace_license.txt")
	FileDelete($licensedir & "GCFScape_license.txt")
	FileDelete($licensedir & "ns2dec_readme.txt")
	FileDelete($licensedir & "extract_license.txt")
	FileDelete($licensedir & "Arc-reader_licence.txt")
	FileDelete($licensedir & "Arc-reader_readme.txt")
	FileDelete($licensedir & "libpng_license.txt")

	FileDelete($langdir & "Chinese.ini")
	FileDelete($langdir & "changes.txt")
	FileDelete(@ScriptDir & "\todo.txt")
	FileDelete(@ScriptDir & "\useful_software.txt")
	FileDelete(@ScriptDir & "\support\Icons\Bioruebe.jpg")

	_DeleteFromArchDir("flac.exe")
	_DeleteFromArchDir("7z.dll.new")
	_DeleteFromArchDir("7z.exe.new")
	_DeleteFromArchDir("GCFScape.exe")
	_DeleteFromArchDir("hllib.dll")

	DirRemove($bindir & "unrpa", 1)
	DirRemove($bindir & "languages", 1)
	DirRemove($bindir & "plugins", 1)
	DirRemove($bindir & "crass-0.4.14.0", 1)
	DirRemove($bindir & "lib", 1)
	DirRemove($bindir & "file\contrib\file\5.03\file-5.03", 1)
	DirRemove($bindir & "file\contrib\file\5.03\file-5.03-src", 1)

	; Ini changes
	IniDelete($prefs, "UniExtract Preferences", "removetemp")
	IniDelete($prefs, "UniExtract Preferences", "consoleoutput")

	SendStats("UpdateMain", 1)

	; Update helpers
	CheckUpdate($UPDATEMSG_SILENT, False, $UPDATE_HELPER)

	If IsAdmin() Then RestartWithoutAdminRights()
	Restart()
EndFunc

; Start updater to download FFmpeg
Func GetFFmpeg()
	; Use the updater to handle elevation and download.
	$ret = ShellExecuteWait(CanAccess($bindir)? $sUpdaterNoAdmin: $sUpdater, "/ffmpeg")
	If @error Or Not HasPlugin($ffmpeg, True) Then Return SetError(1, 0, 0)

	Cout("FFmpeg successfully downloaded")
	Return 1
EndFunc


; ------------------------ Begin GUI Control Functions ------------------------

; Build and display GUI if necessary
Func CreateGUI()
	Global $iGuiMainWidth = 344, $iGuiMainHeight = 182
	Local Const $iLeft = 12, $iTop = 10, $iInputWidth = 290
	Local $iPosY = $iTop - 1

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
	Global $guimain = GUICreate($title, $iGuiMainWidth, $iGuiMainHeight + GUI_GetFontScalingModifier(True), -1, -1, BitOR($WS_SIZEBOX, $WS_MINIMIZEBOX), BitOR($WS_EX_ACCEPTFILES, $iTopmost, $exStyle < 0? 0: $exStyle))

	_GuiSetColor()
	Local $dropzone = GUICtrlCreateLabel("", 0, 0, $iGuiMainWidth, $iGuiMainHeight)

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
	Local $logdiropenitem = GUICtrlCreateMenuItem(t('MENU_FILE_LOG_FOLDER_OPEN_LABEL'), $filemenu)
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
	Local $filelabel = GUICtrlCreateLabel(t('MAIN_FILE_LABEL'), $iLeft, $iTop, $exStyle == $WS_EX_LAYOUTRTL? 50: -1, 15)
	Global $GUI_Main_Extract = GUICtrlCreateRadio(t('TERM_EXTRACT'), GetPos($guimain, $filelabel, 5), $iPosY, Default, 15)
	Global $GUI_Main_Scan = GUICtrlCreateRadio(t('TERM_SCAN'), GetPos($guimain, $GUI_Main_Extract, 10), $iPosY, 100, 15)
	GUICtrlSetState($extract? $GUI_Main_Extract: $GUI_Main_Scan, $GUI_CHECKED)

	$iPosY = GetPos($guimain, $filelabel, 1, False)
	Global $filecont = $history? GUICtrlCreateCombo("", $iLeft, $iPosY, $iInputWidth, 20): GUICtrlCreateInput("", $iLeft, $iPosY, $iInputWidth, 20)
	Local $filebut = GUICtrlCreateButton("...", GetPos($guimain, $filecont, 4), $iPosY, 25, 20)

	; Directory controls
	$iPosY = GetPos($guimain, $filecont, 10, False)
	Global $GUI_Main_Destination_Label = GUICtrlCreateLabel(t('MAIN_DEST_DIR_LABEL'), $iLeft, $iPosY, $exStyle == $WS_EX_LAYOUTRTL? 50: -1, 15)
	Global $GUI_Main_Lock = GUICtrlCreateCheckbox(t('MAIN_DIRECTORY_LOCK'), GetPos($guimain, $GUI_Main_Destination_Label, 5), $iPosY - 1, Default, 15)
	GUICtrlSetTip($GUI_Main_Lock, t('MAIN_DIRECTORY_LOCK_TOOLTIP'))

	$iPosY = GetPos($guimain, $GUI_Main_Destination_Label, 1, False)
	Global $dircont = $history? GUICtrlCreateCombo("", $iLeft, $iPosY, $iInputWidth, 20): GUICtrlCreateInput("", $iLeft, $iPosY, $iInputWidth, 20)
	Global $dirbut = GUICtrlCreateButton("...", GetPos($guimain, $dircont, 4), $iPosY, 25, 20)

	; Buttons
	$iPosY = GetPos($guimain, $dircont, 12, False)
	Global $GUI_Main_Ok = GUICtrlCreateButton(t('OK_BUT'), $iLeft + 20, $iPosY, 80, 22)
	Local $idCancel = GUICtrlCreateButton(t('CANCEL_BUT'), $iLeft + 118, $iPosY, 80, 22)
	Global $BatchBut = GUICtrlCreateButton(t('BATCH_BUT'), $iLeft + 212, $iPosY, 80, 22)

	; Set properties
	GUICtrlSetBkColor($dropzone, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetState($dropzone, $GUI_DISABLE)
	GUICtrlSetState($dropzone, $GUI_DROPACCEPTED)
	GUICtrlSetState($filecont, $GUI_FOCUS)
	GUICtrlSetState($GUI_Main_Ok, $GUI_DEFBUTTON)
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
	GUICtrlSetOnEvent($logopenitem, "GUI_OpenLastLog")
	GUICtrlSetOnEvent($logdiropenitem, "GUI_OpenLogDir")
	GUICtrlSetOnEvent($logitem, "GUI_DeleteLogs")
	GUICtrlSetOnEvent($GUI_Main_Lock, "GUI_KeepOutdir")
	GUICtrlSetOnEvent($GUI_Main_Extract, "GUI_ScanOnly")
	GUICtrlSetOnEvent($GUI_Main_Scan, "GUI_ScanOnly")
	GUICtrlSetOnEvent($filecont, "GUI_OnFileInputChanged")
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
	GUICtrlSetOnEvent($GUI_Main_Ok, "GUI_Ok")
	GUICtrlSetOnEvent($idCancel, "GUI_Exit")
	GUICtrlSetOnEvent($BatchBut, "GUI_Batch")
	GUICtrlSetOnEvent($quititem, "GUI_Exit")
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_Exit")

	GUI_ScanOnly(False)
	GetBatchQueue()

	; Set minimum GUI size for WM_GETMINMAXINFO and $bOptRememberGuiSizePosition
	; GuiCreate width/height refers to the client area while resizing sets the dimensions for the whole window
	; including window decorations (title bar, window borders). These elements can be of different sizes,
	; depending on the theme and version of Windows, so we have to get the real window size dynamically.
	Local $aPos = WinGetPos($guimain)
	$iGuiMainWidth = $aPos[2]
	$iGuiMainHeight = $aPos[3]

	GUISetState(@SW_SHOW)

	If $bOptRememberGuiSizePosition Then WinMove($guimain, "", $iOptGuiPosX, $iOptGuiPosY, $iOptGuiWidth, $iOptGuiHeight)
EndFunc

; Display a standard prompt and return user choice
Func Prompt($iShowFlag, $sMsg, $aVars = 0, $bTerminate = False)
	If $silentmode Then
		Cout("Assuming yes to message " & $sMsg)
		Return 1
	EndIf
	Local $return = MsgBox($iTopmost + $iShowFlag, $title, t($sMsg, $aVars))
	If $return == 1 Or $return == 6 Then
		Return 1
	Else
		If Not $bTerminate Then Return 0
		If $createdir Then DirRemove($outdir, 0)
		terminate($STATUS_SILENT)
	EndIf
EndFunc

; Display a custom prompt with always, never buttons
Func CustomPrompt($sMsg, $aVars)
	If $eCustomPromptSetting == $PROMPT_ALWAYS Then Return True
	If $eCustomPromptSetting == $PROMPT_NEVER Then Return False
	If $silentmode Then Return True

	Opt("GUIOnEventMode", 0)
	Local $return = False

	Local $hGUI = GUICreate($title, 417, 177, -1, -1, $GUI_SS_DEFAULT_GUI)
	GUICtrlCreateLabel(t($sMsg, $aVars), 72, 20, 332, 113)
	Local $idYes = GUICtrlCreateButton(t('YES_BUT'), 251, 142, 75, 25)
	Local $idNo = GUICtrlCreateButton(t('NO_BUT'), 332, 142, 75, 25)
	Local $idAlways = GUICtrlCreateButton(t('ALWAYS_BUT'), 71, 142, 75, 25)
	Local $idNever = GUICtrlCreateButton(t('NEVER_BUT'), 154, 142, 75, 25)
	_GUICtrlCreatePic($sLogoFile, 8, 20, 49, 49)
	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $idNo
				ExitLoop
			Case $idYes
				$return = True
				ExitLoop
			Case $idAlways
				$eCustomPromptSetting = $PROMPT_ALWAYS
				$return = True
				ExitLoop
			Case $idNever
				$eCustomPromptSetting = $PROMPT_NEVER
				ExitLoop
		EndSwitch
	WEnd

	GUIDelete($hGUI)
	Opt("GUIOnEventMode", 1)
	Return $return
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

; Return the checked state of a checkbox
Func _IsChecked($idControlID)
    Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc

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

; Drop-in replacement for GUICtrlCreatePic with PNG support
Func _GUICtrlCreatePic($sPath, $left, $top, $iWidth, $iHeight)
	Local Const $sPlaceholder = @ScriptDir & "\support\Icons\uniextract_inno.bmp"

	$idImage = GUICtrlCreatePic($sPlaceholder, $left, $top, $iWidth, $iHeight)
	_GDIPlus_LoadImage($sPath, $idImage, $iWidth, $iHeight)

	Return $idImage
EndFunc

; Load an image into a picture GUI control, used for PNG support
Func _GDIPlus_LoadImage($sPath, $idImage, $iWidth, $iHeight)
	If Not _GDIPlus_Startup() Then
		Cout("Failed to start GDI+")
		Return SetError(1)
	EndIf

	$hImage = _GDIPlus_ImageLoadFromFile($sPath)
	$hResized = _GDIPlus_ImageResize($hImage, $iWidth, $iHeight)
	$hBitmap = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hResized)

	If Not @error Then GUICtrlSendMsg($idImage, $STM_SETIMAGE, 0, $hBitmap)

	_GDIPlus_BitmapDispose($hResized)
	_GDIPlus_ImageDispose($hImage)
	_GDIPlus_Shutdown()
EndFunc

; Set GUI color to white when using Windows 10
Func _GuiSetColor()
	If @OSVersion <> "WIN_10" Then Return
	GUISetBkColor($COLOR_WHITE)
	GUICtrlSetDefBkColor($COLOR_WHITE)
EndFunc

; Format a label to look like a link
Func _GuiCtrlLinkFormat($idControlID = -1)
	GUICtrlSetFont($idControlID, 8, 800, 4, $FONT_ARIAL)
	GUICtrlSetColor($idControlID, $COLOR_LINK)
	GUICtrlSetCursor($idControlID, 0)
EndFunc

; Return the amount of pixels to increase GUI heigth if Windows font scaling is enabled
Func GUI_GetFontScalingModifier($bOutput = False)
	; Get the size of the title bar and calculate difference to the default value
	Local $ret = _WinAPI_GetSystemMetrics($SM_CYCAPTION)
	Local $iModifier = _Max(0, $ret - 23) * 2.5

	If $bOutput Then Cout("Font scaling: height = " & $ret & ", modifier = " & $iModifier)

	Return $iModifier
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
EndFunc

; Warn user before executing files for extraction
Func Warn_Execute($sCommand)
	If Not $warnexecute Then Return $sCommand

	Cout("Displaying execution warning message")
	If MsgBox($iTopmost + 49, $title, t('WARN_EXECUTE', $sCommand)) <> 1 Then
		If $createdir Then DirRemove($outdir, 0)
		terminate($STATUS_SILENT)
	EndIf

	Return $sCommand
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

		GUICtrlSetState($GUI_Main_Ok, $GUI_FOCUS)
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
	$KeepOutdir = Int(_IsChecked($GUI_Main_Lock))
	SavePref('keepoutputdir', $KeepOutdir)
EndFunc

; Option to scan file without extracting
Func GUI_ScanOnly($bSave = True)
	Global $extract = Number(_IsChecked($GUI_Main_Extract))
	Local $state = $GUI_ENABLE

	If $extract Then
		GUICtrlSetState($GUI_Main_Extract, $GUI_CHECKED)
	Else
		GUICtrlSetState($GUI_Main_Scan, $GUI_CHECKED)
		$state = $GUI_DISABLE
	EndIf

	; Enable/disable destination directory input
	GUICtrlSetState($dircont, $state)
	GUICtrlSetState($dirbut, $state)
	GUICtrlSetState($GUI_Main_Destination_Label, $state)
	GUICtrlSetState($GUI_Main_Lock, $state)

	If @NumParams < 1 Or $bSave Then SavePref('extract', $extract)
EndFunc

; Option to scan file without extracting
Func GUI_Silent()
	If _IsChecked($silentitem) Then
		GUICtrlSetState($silentitem, $GUI_UNCHECKED)
		$silentmode = 0
	Else
		GUICtrlSetState($silentitem, $GUI_CHECKED)
		$silentmode = 1
	EndIf

	SavePref('silentmode', $silentmode)
EndFunc

; Option to keep Universal Extractor open
Func GUI_KeepOpen()
	If _IsChecked($keepopenitem) Then
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
	If _IsChecked($topmostitem) Then
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
	Local $iPosX, $iWidth = 230
	Cout("Creating preferences GUI")

	; Create GUI
	Global $guiprefs = _GUICreate(t('PREFS_TITLE_LABEL'), 466, 330, -1, -1, -1, $exStyle, $guimain)
	_GuiSetColor()

	; General options
	GUICtrlCreateGroup(t('PREFS_UNIEXTRACT_OPTS_LABEL'), 8, 6, 260, 98)
	GUICtrlCreateLabel(t('PREFS_LANG_LABEL'), 14, 36, 72, 15)
	GUICtrlCreateLabel(t('PREFS_UPDATEINTERVAL_LABEL'), 14, 72, 128, 15)
	Global $langselect = GUICtrlCreateCombo("", 100, 32, 160, 25, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
	Global $IntervalCont = GUICtrlCreateCombo("", 140, 68, 120, 21, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
	Local $aUpdateInterval = [t('PREFS_UPDATE_DAILY'), t('PREFS_UPDATE_WEEKLY'), t('PREFS_UPDATE_MONTHLY'), t('PREFS_UPDATE_YEARLY'), t('PREFS_UPDATE_NEVER'), t('PREFS_UPDATE_CUSTOM', $updateinterval)]
	GUICtrlSetData($IntervalCont, _ArrayToString($aUpdateInterval), $aUpdateInterval[0])
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; Source file options
	$iPosX = 330
	GUICtrlCreateGroup(t('PREFS_SOURCE_FILES_LABEL'), 276, 6, 180, 98)
	$DeleteOrigFileOpt[$OPTION_KEEP] = GUICtrlCreateRadio(t('PREFS_SOURCE_FILES_OPT_KEEP'), $iPosX, 22, 113, 17)
	$DeleteOrigFileOpt[$OPTION_ASK] = GUICtrlCreateRadio(t('PREFS_SOURCE_FILES_OPT_ASK'), $iPosX, 40, 113, 17)
	$DeleteOrigFileOpt[$OPTION_DELETE] = GUICtrlCreateRadio(t('PREFS_SOURCE_FILES_OPT_DELETE'), $iPosX, 58, 113, 17)
	Global $idOptDeleteAdditionalFiles = GUICtrlCreateCheckbox(t('PREFS_DELETE_ADDITIONAL_FILES_LABEL'), 300, 76)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; Format-specific preferences
	$iPosX = 14
	GUICtrlCreateGroup(t('PREFS_FORMAT_OPTS_LABEL'), 8, 116, 448, 168)
	Global $historyopt = GUICtrlCreateCheckbox(t('PREFS_HISTORY_LABEL'), $iPosX, 136, $iWidth, 20)
	Global $idOptOpenOutDir = GUICtrlCreateCheckbox(t('PREFS_OPEN_FOLDER_LABEL'), $iPosX, 156, $iWidth, 20)
	Global $freeSpaceCheckOpt = GUICtrlCreateCheckbox(t('PREFS_CHECK_FREE_SPACE_LABEL'), $iPosX, 176, $iWidth, 20)
	Global $idOptRememberGuiSizePosition = GUICtrlCreateCheckbox(t('PREFS_WINDOW_POSITION_LABEL'), $iPosX, 196, $iWidth, 20)
	Global $idOptNoStatusBox = GUICtrlCreateCheckbox(t('PREFS_HIDE_STATUS_LABEL'), $iPosX, 216, $iWidth, 20)
	Global $idOptGameMode = GUICtrlCreateCheckbox(t('PREFS_HIDE_STATUS_FULLSCREEN_LABEL'), $iPosX, 236, $iWidth, 20)
	Global $idOptExtractVideo = GUICtrlCreateCheckbox(t('PREFS_VIDEOTRACK_LABEL'), $iPosX, 256, $iWidth, 20)

	$iPosX += 236
	$iWidth = 204
	Global $warnexecuteopt = GUICtrlCreateCheckbox(t('PREFS_WARN_EXECUTE_LABEL'), $iPosX, 136, $iWidth, 20)
	Global $unicodecheckopt = GUICtrlCreateCheckbox(t('PREFS_CHECK_UNICODE_LABEL'), $iPosX, 156, $iWidth, 20)
	Global $appendextopt = GUICtrlCreateCheckbox(t('PREFS_APPEND_EXT_LABEL'), $iPosX, 176, $iWidth, 20)
	Global $idOptCreateLog = GUICtrlCreateCheckbox(t('PREFS_LOG_LABEL'), $iPosX, 196, $iWidth, 20)
	Global $idOptFeedbackPrompt = GUICtrlCreateCheckbox(t('PREFS_FEEDBACK_PROMPT_LABEL'), $iPosX, 216, $iWidth, 20, $BS_AUTO3STATE)
	Global $idOptSendStats = GUICtrlCreateCheckbox(t('PREFS_SEND_STATS_LABEL'), $iPosX, 236, $iWidth, 20)
	Global $idOptBetaUpdates = GUICtrlCreateCheckbox(t('PREFS_BETA_UPDATES_LABEL'), $iPosX, 256, $iWidth, 20)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; Buttons
	Local $idOk = GUICtrlCreateButton(t('OK_BUT'), 132, 296, 80, 24)
	Local $idCancel = GUICtrlCreateButton(t('CANCEL_BUT'), 248, 296, 80, 24)

	; Tooltips
	GUICtrlSetTip($warnexecuteopt, t('PREFS_WARN_EXECUTE_TOOLTIP'))
	GUICtrlSetTip($freeSpaceCheckOpt, t('PREFS_CHECK_FREE_SPACE_TOOLTIP'))
	GUICtrlSetTip($unicodecheckopt, t('PREFS_CHECK_UNICODE_TOOLTIP'))
	GUICtrlSetTip($appendextopt, t('PREFS_APPEND_EXT_TOOLTIP'))
	GUICtrlSetTip($idOptGameMode, t('PREFS_HIDE_STATUS_FULLSCREEN_TOOLTIP'))
	GUICtrlSetTip($idOptFeedbackPrompt, t('PREFS_FEEDBACK_PROMPT_TOOLTIP'))
	GUICtrlSetTip($idOptSendStats, t('PREFS_SEND_STATS_TOOLTIP'))
	GUICtrlSetTip($idOptExtractVideo, t('PREFS_VIDEOTRACK_TOOLTIP'))
	GUICtrlSetTip($DeleteOrigFileOpt[$OPTION_ASK], t('PREFS_SOURCE_FILES_OPT_KEEP_TOOLTIP'))
	GUICtrlSetTip($idOptDeleteAdditionalFiles, t('PREFS_DELETE_ADDITIONAL_FILES_TOOLTIP', t('DIR_ADDITIONAL_FILES')))

	; Set properties
	GUICtrlSetState($idOk, $GUI_DEFBUTTON)
	If $history Then GUICtrlSetState($historyopt, $GUI_CHECKED)
	If $warnexecute Then GUICtrlSetState($warnexecuteopt, $GUI_CHECKED)
	If $freeSpaceCheck Then GUICtrlSetState($freeSpaceCheckOpt, $GUI_CHECKED)
	If $checkUnicode Then GUICtrlSetState($unicodecheckopt, $GUI_CHECKED)
	If $appendext Then GUICtrlSetState($appendextopt, $GUI_CHECKED)
	If $bOptNoStatusBox Then GUICtrlSetState($idOptNoStatusBox, $GUI_CHECKED)
	If $bHideStatusBoxIfFullscreen Then GUICtrlSetState($idOptGameMode, $GUI_CHECKED)
	If $bOptOpenOutDir Then GUICtrlSetState($idOptOpenOutDir, $GUI_CHECKED)
	If $FB_ask == 1 Then
		GUICtrlSetState($idOptFeedbackPrompt, $GUI_CHECKED)
	ElseIf $FB_ask == 2 Then
		GUICtrlSetState($idOptFeedbackPrompt, $GUI_INDETERMINATE)
	EndIf
	If $bOptRememberGuiSizePosition Then GUICtrlSetState($idOptRememberGuiSizePosition, $GUI_CHECKED)
	If $bSendStats Then GUICtrlSetState($idOptSendStats, $GUI_CHECKED)
	If $Log Then GUICtrlSetState($idOptCreateLog, $GUI_CHECKED)
	If $bExtractVideo Then GUICtrlSetState($idOptExtractVideo, $GUI_CHECKED)
	If $iCleanup == $OPTION_DELETE Then GUICtrlSetState($idOptDeleteAdditionalFiles, $GUI_CHECKED)
	If $bOptNightlyUpdates Then GUICtrlSetState($idOptBetaUpdates, $GUI_CHECKED)

	; Update interval
	; For convenience we use presets instead of numeral values, so we need to convert them here
	Local $iIndex = 5
	Switch $updateinterval
		Case 1:
			$iIndex = 0
		Case 7:
			$iIndex = 1
		Case 30:
			$iIndex = 2
		Case 365:
			$iIndex = 3
		Case 999999:
			$iIndex = 4
	EndSwitch
	_GUICtrlComboBoxEx_SetCurSel(GUICtrlGetHandle($IntervalCont), $iIndex)

	GUICtrlSetState($DeleteOrigFileOpt[$iDeleteOrigFile], $GUI_CHECKED)
	GUICtrlSetData($langselect, GetLanguageList(), $language)

	; Set events
	GUICtrlSetOnEvent($idOk, "GUI_Prefs_Ok")
	GUICtrlSetOnEvent($idCancel, "GUI_Prefs_Exit")
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_Prefs_Exit")

	; Display GUI and wait for action
	GUISetState(@SW_SHOW)
EndFunc

; Exit preferences GUI if Cancel clicked or window closed
Func GUI_Prefs_Exit()
	Cout("Closing preferences GUI")
	GUI_Close()
	$guiprefs = False
EndFunc

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

	Local $tmp = GUICtrlRead($langselect)
	If $language <> $tmp Then
		$language = $tmp
		$redrawgui = True
	EndIf

	$tmp = _GUICtrlComboBoxEx_GetCurSel(GUICtrlGetHandle($IntervalCont))
	Local $aReturn = [1, 7, 30, 365, 999999, $updateinterval]
	$updateinterval = $aReturn[$tmp]

	$bOptNoStatusBox = Number(GUICtrlRead($idOptNoStatusBox) == $GUI_CHECKED)
	TrayItemSetState($Tray_Statusbox, $bOptNoStatusBox? $TRAY_CHECKED: $TRAY_UNCHECKED)

	$warnexecute = Number(GUICtrlRead($warnexecuteopt) == $GUI_CHECKED)
	$checkUnicode = Number(GUICtrlRead($unicodecheckopt) == $GUI_CHECKED)
	$freeSpaceCheck = Number(GUICtrlRead($freeSpaceCheckOpt) == $GUI_CHECKED)
	$appendext = Number(GUICtrlRead($appendextopt) == $GUI_CHECKED)
	$bHideStatusBoxIfFullscreen = Number(GUICtrlRead($idOptGameMode) == $GUI_CHECKED)
	$bOptOpenOutDir = Number(GUICtrlRead($idOptOpenOutDir) == $GUI_CHECKED)
	$FB_ask = Number(GUICtrlRead($idOptFeedbackPrompt))
	If $FB_ask > 2 Then $FB_ask = 0
	$Log = Number(GUICtrlRead($idOptCreateLog) == $GUI_CHECKED)
	$bExtractVideo = Number(GUICtrlRead($idOptExtractVideo) == $GUI_CHECKED)
	$bOptRememberGuiSizePosition = Number(GUICtrlRead($idOptRememberGuiSizePosition) == $GUI_CHECKED)
	$bSendStats = Number(GUICtrlRead($idOptSendStats) == $GUI_CHECKED)
	$iCleanup = GUICtrlRead($idOptDeleteAdditionalFiles) == $GUI_CHECKED? $OPTION_DELETE: $OPTION_MOVE
	$tmp = Number(GUICtrlRead($idOptBetaUpdates) == $GUI_CHECKED)
	$bUpdate = Not ($bOptNightlyUpdates == $tmp)
	$bOptNightlyUpdates = $tmp
	$sUpdateURL = $bOptNightlyUpdates == 1? $sNightlyUpdateURL: $sDefaultUpdateURL

	For $i = 0 To 2
		If GUICtrlRead($DeleteOrigFileOpt[$i]) == $GUI_CHECKED Then $iDeleteOrigFile = $i
	Next

	WritePrefs()

	GUI_Prefs_Exit()

	If $bUpdate Then CheckUpdate()

	If Not $redrawgui Then Return
	GUIDelete($guimain)
	CreateGUI()
EndFunc

; Handle change event of file input
Func GUI_OnFileInputChanged()
	If StringLen(GUICtrlRead($dircont)) > 0 Then Return

	Global $file = GUICtrlRead($filecont)
	GUI_Drop_Parse()
EndFunc

; Handle click on OK
Func GUI_OK()
	If Not GUI_OK_Set(True) Then Return
	GUI_SavePosition()
	GUIDelete($guimain)
	$guimain = False
EndFunc

; Set file to extract and target directory
Func GUI_OK_Set($bShowError = False)
	$file = EnvParse(GUICtrlRead($filecont))
	If Not FileExists($file) Then
		If $bShowError Then MsgBox($iTopmost + 48, $title, t('INVALID_FILE_SELECTED', $file))
		Return 0
	EndIf

	Local $ret = EnvParse(GUICtrlRead($dircont))
	$outdir = $ret == ""? '/sub': $ret

	Return 1
EndFunc

; Add file to batch queue
Func GUI_Batch()
	If GUI_OK_Set() Then
		AddToBatch()
		GUICtrlSetData($filecont, "")
		If Not $KeepOutdir Then GUICtrlSetData($dircont, "")
	Else ; Start batch process if items in queue and input fields empty
		If GetBatchQueue() Then Return GUI_Batch_OK()
		MsgBox($iTopmost + 48, $title, t('INVALID_FILE_SELECTED', $file))
	EndIf
EndFunc

; Execute batch queue
Func GUI_Batch_OK()
	Cout("Closing main GUI - batch mode")
	Local $file = GUICtrlRead($filecont)
	If $file <> "" And Not StringIsSpace($file) Then GUI_Batch()
	GUI_SavePosition()
	GUIDelete($guimain)

	If FileExists($fileScanLogFile) Then FileDelete($fileScanLogFile)

	terminate($STATUS_BATCH)
EndFunc

; Add all files from a directory to batch queue
Func GUI_Batch_AddDirectory($sDir)
	Local Static $bRecurse = Number(IniRead($prefs, "UniExtract Preferences", "BatchRecurse", 1) )
	If $bRecurse > 1 Then $bRecurse = 1

	Local $aFiles = _FileListToArrayRec($sDir, "*", $FLTAR_FILES, $bRecurse, $FLTAR_NOSORT, $FLTAR_FULLPATH)
	If @error Then
		Cout("Add to batch queue: Failed to read contents of directory " & $sDir)
		Return False
	EndIf
;~ 	_ArrayDisplay($aFiles)

	For $j = 1 To $aFiles[0]
		If $guimain Then
			GUI_Drop_Parse($aFiles[$j])
			GUI_Batch()
		Else
			$file = $aFiles[$j]
			AddToBatch()
		EndIf
	Next
	$eCustomPromptSetting = $PROMPT_ASK

	Return $aFiles[0]
EndFunc

; Display batch queue and allow changes
Func GUI_Batch_Show()
	Local Const $iListLeft = 8, $iListTop = 8
	Local $iLastIndex = -1, $bTooltip = False
	Cout("Opening batch queue edit GUI")
	Local $hGUI = GUICreate($name, 418, 267, 476, 262, BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU, $WS_SIZEBOX), -1, $guimain)
	_GuiSetColor()
	Local $idList = GUICtrlCreateList("", $iListLeft, $iListTop, 401, 201)
	GUICtrlSetData(-1, _ArrayToString($queueArray, "|"))
	Local $idOk = GUICtrlCreateButton(t('OK_BUT'), 40, 225, 75, 25)
	Local $idCancel = GUICtrlCreateButton(t('CANCEL_BUT'), 171, 225, 75, 25)
	Local $idDelete	= GUICtrlCreateButton(t('DELETE_BUT'), 304, 224, 73, 25)
	GUISetState(@SW_SHOW)
	Opt("GUIOnEventMode", 0)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $idCancel
				GetBatchQueue()
				ExitLoop
			Case $idOk
;~ 				Cout("Batch queue was modified")
				If UBound($queueArray) < 1 Then
					EnableBatchMode(False)
					ExitLoop
				EndIf
;~ 				_ArrayDisplay($queueArray)
				SaveBatchQueue()
				; Only called to update main GUI batch button
				GetBatchQueue()
				ExitLoop
			Case $idDelete
				Local $iPos = _GUICtrlListBox_GetCurSel($idList)
				If $iPos < 0 Then ContinueLoop Cout("No item selected")

				If _ArrayDelete($queueArray, $iPos) > -1 Then GUICtrlSetData($idList, "|" & _ArrayToString($queueArray, "|"))
			Case Else
				; Display tooltips if file name too long
				; Code by Malkey (https://www.autoitscript.com/forum/topic/146743-listbox-tooltip-for-long-items/?do=findComment&comment=1039835)
				Local $aCursorInfo = GUIGetCursorInfo($hGUI)
				If $aCursorInfo[4] = $idList Then
					$iIndex = _GUICtrlListBox_ItemFromPoint($idList, $aCursorInfo[0] - $iListLeft, $aCursorInfo[1] - $iListTop)
					If $iLastIndex == $iIndex Then ContinueLoop
					$iLastIndex = $iIndex
					$sText = _GUICtrlListBox_GetText($idList, $iIndex)
					If StringLen($sText) > 72 Then
						ToolTip($sText)
						$bTooltip = True
					Else
						ToolTip("")
						$bTooltip = False
						$iLastIndex = -1
					EndIf
				EndIf
				If $bTooltip And $aCursorInfo[4] <> $idList Then
					$bTooltip = False
					$iLastIndex = -1
					ToolTip("")
				EndIf
		EndSwitch
	WEnd

	GUIDelete($hGUI)
	Opt("GUIOnEventMode", 1)
EndFunc

; Clear batch queue
Func GUI_Batch_Clear()
	Cout("Batch queue cleared, batch mode disabled")
	EnableBatchMode(False)
EndFunc   ;==>GUI_Batch_Clear

; Process dropped files
Func GUI_Drop()
;~ 	_ArrayDisplay($gaDropFiles)
	Cout("Drag and drop action detected")

	$iCount = 0
	For $sPath In $gaDropFiles
		If Not FileExists($sPath) Then ContinueLoop

		If _IsDirectory($sPath) Then
			Cout("Drag and drop: folder passed")
			$iCount += GUI_Batch_AddDirectory($sPath)
		Else
			$iCount += 1
			GUI_Drop_Parse($sPath)
			If UBound($gaDropFiles) == 1 Then Return

			GUI_Batch()
		EndIf
	Next
	$eCustomPromptSetting = $PROMPT_ASK

	GetBatchQueue()
	Cout("Drag and drop - a total of " & $iCount & " files were added to batch queue")
EndFunc

; Process dropped files
Func GUI_Drop_Parse($sFile = $file)
	Global $file = $sFile
	If $file == "" Then Return

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

; Create Feedback GUI
Func GUI_Feedback()
	; Attach input file information
	If $file Then
		If Not $isexe Then Cout("--------------------------------------------------File dump--------------------------------------------------" & _
								 @CRLF & _HexDump($file, 1024))
		Cout("------------------------------------------------File metadata------------------------------------------------" & _
			  @CRLF & _ArrayToString(_GetExtProperty($file), @CRLF))
		Global $FB_ask = 0
	EndIf

	Global $FB_GUI = GUICreate(t('FEEDBACK_TITLE_LABEL'), 402, 528 + GUI_GetFontScalingModifier(), -1, -1, BitOR($WS_SIZEBOX, $WS_SYSMENU), -1, $guimain)
	_GuiSetColor()

	GUICtrlCreateLabel(t('FEEDBACK_SYSINFO_LABEL'), 8, 8, 384, 17)
	$FB_SysCont = GUICtrlCreateInput(@OSVersion & " " & @OSArch & (@OSServicePack = ""? "": " " & @OSServicePack) & ", Lang: " & @OSLang & ", UE: " & $language, 8, 24, 385, 21)

	GUICtrlCreateLabel(t('FEEDBACK_OUTPUT_LABEL'), 8, 56, 384, 17)
	GUICtrlSetTip(-1, t('FEEDBACK_OUTPUT_TOOLTIP'), "", 0, 1)
	$FB_OutputCont = GUICtrlCreateEdit("", 8, 72, 385, 161, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_WANTRETURN, $WS_VSCROLL))
	GUICtrlSetData(-1, $sFullLog)
	GUICtrlSetTip(-1, t('FEEDBACK_OUTPUT_TOOLTIP'), "", 0, 1)

	GUICtrlCreateLabel(t('FEEDBACK_MESSAGE_LABEL'), 8, 248, 384, 17)
	GUICtrlSetTip(-1, t('FEEDBACK_MESSAGE_TOOLTIP'), "", 0, 1)
	$FB_MessageCont = GUICtrlCreateEdit("", 8, 264, 385, 169, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_WANTRETURN, $WS_VSCROLL))
	GUICtrlSetTip(-1, t('FEEDBACK_MESSAGE_TOOLTIP'), "", 0, 1)

	$FB_PrivacyPolicyCheckbox = GUICtrlCreateCheckbox(t('FEEDBACK_PRIVACY_ACCEPT_LABEL'), 8, 442, 217, 17)
	$FB_PrivacyPolicyOpen = GUICtrlCreateLabel(t('FEEDBACK_PRIVACY_VIEW_LABEL'), 246, 444, 147, 17, $SS_RIGHT)
	_GuiCtrlLinkFormat()

	$FB_Send = GUICtrlCreateButton(t('SEND_BUT'), 111, 470, 75, 25)
	Local $idCancel = GUICtrlCreateButton(t('CANCEL_BUT'), 215, 470, 75, 25)
	$hSelectAll = GUICtrlCreateDummy()

	Local $accelKeys[1][2] = [["^a", $hSelectAll]]
	GUISetAccelerators($accelKeys)
	GUICtrlSetState($FB_MessageCont, $GUI_FOCUS)

	; Set minimum window size
	GUIRegisterMsg($WM_GETMINMAXINFO, "GUI_WM_GETMINMAXINFO_Feedback")
	GUISetState(@SW_SHOW)

	; Warn if UniExtract is outdated
	GUICtrlSetState($FB_Send, $GUI_DISABLE)
	Local $aReturn = _UpdateGetIndex("", True)
	If IsArray($aReturn) And (($aReturn[0])[1] <> FileGetSize($sUniExtract) Or StringTrimLeft(_Crypt_HashFile($sUniExtract, $CALG_MD5), 2) <> ($aReturn[0])[2]) Then GUI_Feedback_Outdated()
	GUICtrlSetState($FB_Send, $GUI_ENABLE)
	Opt("GUIOnEventMode", 0)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $FB_Send
				If GUICtrlRead($FB_PrivacyPolicyCheckbox) = $GUI_CHECKED Then
					GUICtrlSetState($FB_Send, $GUI_DISABLE)
					If GUI_Feedback_Send(GUICtrlRead($FB_SysCont), $file, GUICtrlRead($FB_OutputCont), GUICtrlRead($FB_MessageCont)) Then ExitLoop
					GUICtrlSetState($FB_Send, $GUI_ENABLE)
				Else
					MsgBox($iTopmost + 48, $name, t('FEEDBACK_PRIVACY_NOT_ACCETPED'))
				EndIf
			Case $GUI_EVENT_CLOSE, $idCancel
				ExitLoop
			Case $FB_PrivacyPolicyOpen
				ShellExecute($sPrivacyPolicyURL)
			Case $hSelectAll
				GUI_Edit_SelectAll()
		EndSwitch
	WEnd

	GUIDelete($FB_GUI)
	Opt("GUIOnEventMode", 1)
EndFunc

; Display warning when using an outdated version of UniExtract
Func GUI_Feedback_Outdated()
	Opt("GUIOnEventMode", 0)

	Local $hGUI = GUICreate($name, 416, 156, -1, -1, $GUI_SS_DEFAULT_GUI, -1, $FB_GUI)
	_GuiSetColor()
	GUICtrlCreateLabel(t('FEEDBACK_OUTDATED'), 72, 20, 330, 113)
	Local $idInstall = GUICtrlCreateButton(t('UPDATE_ACTION_INSTALL'), 194, 120, 123, 25)
	Local $idContinue = GUICtrlCreateButton(t('CONTINUE_BUT'), 332, 120, 75, 25)
	_GUICtrlCreatePic($sLogoFile, 8, 20, 49, 49)
	GUISetState(@SW_SHOW)
	SendStats("FeedbackOutdated")

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $idContinue
				ExitLoop
			Case $idInstall
				GUIDelete($hGUI)
				Opt("GUIOnEventMode", 1)
				Return CheckUpdate($UPDATEMSG_FOUND_ONLY)
		EndSwitch
	WEnd

	GUIDelete($hGUI)
	Opt("GUIOnEventMode", 1)
EndFunc

; Exit feedback GUI if OK clicked
Func GUI_Feedback_Send($FB_Sys, $FB_File, $FB_Output, $FB_Message)
	If $FB_File = "" And $FB_Output = "" And $FB_Message = "" Then Return MsgBox($iTopmost + 16, $name, t('FEEDBACK_EMPTY'))

	GUISetState(@SW_HIDE, $FB_GUI)
	If $guimain Then GUISetState(@SW_HIDE, $guimain)
	_CreateTrayMessageBox(t('SENDING_FEEDBACK'))

	$FB_Text = $name & " Feedback v" & $sVersion & " (" & FileGetVersion($sUniExtract, "Timestamp") & ")" & @CRLF & _
			"----------------------------------------------------------------------------------------------------" _
			 & @CRLF & @CRLF & "System Information: " & $title & ", " & $FB_Sys & @CRLF & @CRLF & "Sample file: " & _
			 $FB_File & @CRLF & "File size: " & $sFileSize & @CRLF & "File type: " & _FiletypeGet(False) & @CRLF _
			 & @CRLF & "Message: " & $FB_Message & @CRLF & @CRLF & _
			"----------------------------------------------------------------------------------------------------" _
			 & @CRLF & @CRLF & "Output:" & @CRLF & $FB_Output & @CRLF & @CRLF & _
			"----------------------------------------------------------------------------------------------------" _
			 & @CRLF & "Sent by: " & @CRLF & $ID

	Const $boundary = "--UniExtractLog"
	Local $iSize = BinaryLen(StringToBinary($FB_Text))
	Const $bUseGzip = StringInStr($language, "Chinese") Or $language = "Japanese" Or $iSize > 1024 * 1024

	If $bUseGzip Then
		$FB_Text = 'Content-Type: gzip' & @CRLF & @CRLF & StringTrimLeft(String(_Zlib_Compress($FB_Text)), 2)
	Else
		$FB_Text = 'Content-Type: text/plain' & @CRLF & @CRLF & $FB_Text
	EndIf

	Local $sData = $boundary & @CRLF & 'Content-Disposition: form-data; name="file"; filename="UE_Feedback"' & @CRLF & $FB_Text & @CRLF & _
				   $boundary & @CRLF & 'Content-Disposition: form-data; name="id"' & @CRLF & @CRLF & $ID & @CRLF & $boundary & '--'
	Local $sResponse = 0

	$http = ObjCreate("winhttp.winhttprequest.5.1")
	If @error Then
		_DeleteTrayMessageBox()
		Return GUI_Feedback_Error("Failed to create winhttp object")
	Else
		Global $sComError = 0
		Local $oErrorHandler = ObjEvent("AutoIt.Error", "_ComErrorHandler")

		$http.Open("POST", $sSupportURL, False)
		$http.SetRequestHeader("Content-Type", "multipart/form-data; boundary=" & StringTrimLeft($boundary, 2))

		; Debug only: use MITM proxy to see raw HTTP data
;~ 		$http.SetProxy(2, "127.0.0.1:8080", "")
;~ 		$http.Option(4) = 0x3300

		Cout("Sending feedback (" & ($bUseGzip? "gzip": "plain") & " @ " & Round($iSize / 1024, 2) & "kb/" & Round(BinaryLen(StringToBinary($FB_Text)) / 1024, 2) & "kb)")
		$http.Send($sData)
		$http.WaitForResponse()
		$sResponse = $http.ResponseText()
	EndIf

	_DeleteTrayMessageBox()

	If $sResponse = "1" Then
		Cout("Feedback successfully sent")
		GUIDelete($FB_GUI)
		MsgBox($iTopmost, $title, t('FEEDBACK_SUCCESS'))
	Else
		Return GUI_Feedback_Error($sComError == 0? "Invalid response from server": $sComError)
	EndIf

	GUISetState(@SW_SHOW, $guimain)
	Return True
EndFunc

; Display error message if sending feedback failed
Func GUI_Feedback_Error($sError)
	Cout("Error sending feedback: " & $sError)
	MsgBox($iTopmost + 16, $title, t('FEEDBACK_ERROR', $sError))

	GUISetState(@SW_SHOW, $FB_GUI)
	GUISetState(@SW_SHOW, $guimain)

	Return False
EndFunc

; Ask for feedback
Func GUI_Feedback_Prompt()
	If Not ($FB_ask And $extract) Or $silentmode Then Return
	If $FB_ask == 2 Then Return GUI_Feedback()

	Opt("GUIOnEventMode", 0)
	Local $hGUI = GUICreate($name, 416, 176, -1, -1, $GUI_SS_DEFAULT_GUI)
	_GuiSetColor()
	GUICtrlCreateLabel(t('FEEDBACK_PROMPT'), 72, 20, 330, 111)
	Local $idYes = GUICtrlCreateButton(t('YES_BUT'), 242, 142, 75, 25)
	Local $idNo = GUICtrlCreateButton(t('NO_BUT'), 332, 142, 75, 25)
	_GUICtrlCreatePic($sLogoFile, 8, 20, 49, 49)
	Local $idRemember = GUICtrlCreateCheckbox(t('CHECKBOX_REMEMBER'), 12, 148, 217, 17)
	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $idYes
				If _IsChecked($idRemember) Then
					$FB_ask = 2
					SavePref("feedbackprompt", $FB_ask)
				EndIf
				GUIDelete($hGUI)
				GUI_Feedback()
				ExitLoop
			Case $idNo
				If _IsChecked($idRemember) Then
					$FB_ask = 0
					SavePref("feedbackprompt", $FB_ask)
				EndIf
				ExitLoop
		EndSwitch
	WEnd

	GUIDelete($hGUI)
	Opt("GUIOnEventMode", 1)
EndFunc

; Due to a bug in the Windows API, ctrl+a does not work for edit controls
; Workaround by Zedna (http://www.autoitscript.com/forum/topic/97473-hotkey-ctrla-for-select-all-in-the-selected-edit-box/#entry937287)
Func GUI_Edit_SelectAll()
	$hWnd = _WinAPI_GetFocus()
	$class = _WinAPI_GetClassName($hWnd)
	If $class = 'Edit' Then _GUICtrlEdit_SetSel($hWnd, 0, -1)
EndFunc

; Saves current position of main GUI
Func GUI_SavePosition()
	If Not $guimain Or Not $bOptRememberGuiSizePosition Then Return

	$aPos = WinGetPos($guimain)
	If @error Then Return

	SavePref('posx', $aPos[0])
	SavePref('posy', $aPos[1])
	SavePref('GuiWidth', $aPos[2])
	SavePref('GuiHeight', $aPos[3])
EndFunc

; Set minimal size of main GUI
Func GUI_WM_GETMINMAXINFO_Main($hwnd, $Msg, $wParam, $lParam)
    $tagMaxinfo = DllStructCreate("int;int;int;int;int;int;int;int;int;int", $lParam)
    DllStructSetData($tagMaxinfo, 7, $iGuiMainWidth) ; min X
    DllStructSetData($tagMaxinfo, 8, $iGuiMainHeight) ; min Y
    ;DllStructSetData($tagMaxinfo, 9, 1200); max X
    ;DllStructSetData($tagMaxinfo, 10, 160) ; max Y
EndFunc

; Set minimal size of feedback GUI
Func GUI_WM_GETMINMAXINFO_Feedback($hWnd, $Msg, $wParam, $lParam)
	$tagMaxinfo = DllStructCreate("int;int;int;int;int;int;int;int;int;int", $lParam)
	DllStructSetData($tagMaxinfo, 7, 400) ; min width
	DllStructSetData($tagMaxinfo, 8, 500) ; min height
EndFunc

; Tooltip does not work for disabled controls, so here's a workaround
Func GUI_Create_Tooltip($hGui, $hWnd, $sMsg)
	Local $pos = ControlGetPos($hGui, "", $hWnd)
	If @error Then
		Cout("Error creating tooltip: failed to determine size of control")
		Return SetError(1, 0, -1)
	EndIf

	Local $idLabel = GUICtrlCreateLabel("", $pos[0], $pos[1], $pos[2], $pos[3])
	GUICtrlSetTip($idLabel, $sMsg)

	; Set initial control on top
	; Based on http://www.autoitscript.com/forum/topic/146182-solved-change-z-ordering-of-controls/#entry1034567
	If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	_WinAPI_SetWindowPos($hWnd, $HWND_BOTTOM, 0, 0, 0, 0, $SWP_NOMOVE + $SWP_NOSIZE + $SWP_NOCOPYBITS)

	Return $idLabel
EndFunc

; Create GUI to change context menu
Func GUI_ContextMenu()
	Cout("Creating context menu GUI")
	Local $iSize = UBound($CM_Shells) - 1
	Global $CM_Checkbox[$iSize + 1]

	Global $CM_GUI = _GUICreate(t('PREFS_TITLE_LABEL'), 450, 630, -1, -1, -1, $exStyle, $guimain)
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
EndFunc

; Change picture according to selected context menu type
Func GUI_ContextMenu_ChangePic()
	Local $sPath = @ScriptDir & "\support\Icons\" & (GUICtrlRead($CM_Cascading_Radio) = $GUI_CHECKED? "cascading.jpg": "simple.jpg")
	GUICtrlSetImage($CM_Picture, $sPath)
EndFunc

; Close GUI and create context menu entries
Func GUI_ContextMenu_OK()
	Local $iSize = UBound($CM_Shells) - 1
	Sleep(100)
	GUISetState(@SW_HIDE)

	; Remove old associations
	GUI_ContextMenu_remove()

	Cout("Registering context menu entries")
	If GUICtrlRead($CM_Checkbox_enabled) == $GUI_CHECKED Then

		; Select registry key
		Global $reguser = GUICtrlRead($CM_Checkbox_allusers) == $GUI_CHECKED? $regall: $regcurrent

		; simple
		If GUICtrlRead($CM_Simple_Radio) == $GUI_CHECKED Then
			For $i = 0 To $iSize
				$command = '"' & @ScriptFullPath & '" "%1"' & $CM_Shells[$i][1]
				If GUICtrlRead($CM_Checkbox[$i]) == $GUI_CHECKED Then
					Local $sKey = $reguser & $CM_Shells[$i][0]
					RegWrite($sKey, "", "REG_SZ", t($CM_Shells[$i][2]))
					RegWrite($sKey & "\command", "", "REG_SZ", $command)
					If $CM_Shells[$i][3] Then RegWrite($sKey, "MultiSelectModel", "REG_SZ", $CM_Shells[$i][3])

					; Add icon to context menu, only works on win 7 or newer
					If $win7 Then RegWrite($sKey, "Icon", "REG_SZ", @ScriptFullPath & ",0")
				EndIf
			Next

		; cascading
		ElseIf $win7 And GUICtrlRead($CM_Cascading_Radio) == $GUI_CHECKED Then
			Local $sKey = $reguser & "uniextract"
			RegWrite($sKey, "MUIVerb", "REG_SZ", "Universal Extractor")
			RegWrite($sKey, "Icon", "REG_SZ", @ScriptFullPath & ",0")
			RegWrite($sKey, "SubCommands", "REG_SZ", "")
			RegWrite($sKey, "MultiSelectModel", "REG_SZ", "Player")

			For $i = 0 To $iSize
				$command = '"' & @ScriptFullPath & '" "%1"' & $CM_Shells[$i][1]
				$sKey = $reguser & "uniextract\Shell\" & $CM_Shells[$i][0]
				If GUICtrlRead($CM_Checkbox[$i]) == $GUI_CHECKED Then
					RegWrite($sKey, "", "REG_SZ", t($CM_Shells[$i][2]))
					RegWrite($sKey & "\command", "", "REG_SZ", $command)

					; Icon
					RegWrite($sKey, "Icon", "REG_SZ", @ScriptFullPath & ",0")
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
EndFunc

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
	_GUICtrlCreatePic($sLogoFile, 8, 312, 65, 65)
	GUICtrlCreateLabel($name, 8, 8, 488, 60, $SS_CENTER)
	GUICtrlSetFont(-1, 24, 800, 0, $FONT_ARIAL)
	GUICtrlCreateLabel(StringReplace(t('FIRSTSTART_TITLE'), "&", ""), 8, 50, 488, 60, $SS_CENTER)
	GUICtrlSetFont(-1, 14, 800, 0, $FONT_ARIAL)
	Global $FS_Section = GUICtrlCreateLabel("", 16, 85, 382, 28)
	GUICtrlSetFont(-1, 14, 800, 4, $FONT_ARIAL)
	Global $FS_Text = GUICtrlCreateLabel("", 16, 120, 468, 170)
	Global $FS_Next = GUICtrlCreateButton(t('NEXT_BUT'), 296, 344, 89, 25)
	Local $idExit = GUICtrlCreateButton(t('EXIT_BUT'), 400, 344, 89, 25)
	Global $FS_Prev = GUICtrlCreateButton(t('PREV_BUT'), 192, 344, 89, 25)
	GUICtrlSetState(-1, $GUI_HIDE)
	Global $FS_Button = GUICtrlCreateButton("", 187, 260, 129, 41)
	Global $FS_Progress = GUICtrlCreateLabel("", 80, 350, 21, 17)

	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_FirstStart_Exit")
	GUICtrlSetOnEvent($idExit, "GUI_FirstStart_Exit")
	GUICtrlSetOnEvent($FS_Next, "GUI_FirstStart_Next")
	GUICtrlSetOnEvent($FS_Prev, "GUI_FirstStart_Prev")

	Global $page = 1
	Global $FS_Sections = StringSplit(t('FIRSTSTART_PAGES'), "|")
	If @error Then
		SavePref("ID", "")
		If MsgBox(48+4, $title, "No language file found." & @CRLF & @CRLF & "Do you want Universal Extractor to download all missing files?") Then
			CheckUpdate($UPDATEMSG_SILENT, False, $UPDATE_HELPER)
			Restart()
		EndIf
		Exit 0
	EndIf
	Global $FS_Texts[UBound($FS_Sections)] = ["", t('FIRSTSTART_PAGE1'), t('FIRSTSTART_PAGE2'), t('FIRSTSTART_PAGE3')]

	GUISetState(@SW_SHOW)
	GUI_FirstStart_ShowPage()
EndFunc

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
		Case 1
			GUICtrlSetPos($FS_Text, 16, 120, 468, 170)
			GUICtrlSetState($FS_Button, $GUI_HIDE)
		Case 2
			GUICtrlSetPos($FS_Text, 16, 120, 468, 125)
			GUICtrlSetState($FS_Button, $GUI_SHOW)
			GUICtrlSetData($FS_Button, t('PREFS_TITLE_LABEL'))
			GUICtrlSetOnEvent($FS_Button, "GUI_Prefs")
		Case 3
			;GUICtrlSetState($FS_Button, $GUI_SHOW)
			GUICtrlSetData($FS_Button, t('CONTEXT_ENTRIES_LABEL'))
			GUICtrlSetOnEvent($FS_Button, "GUI_ContextMenu")
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

; Display UI for extraction method selection (compact)
Func GUI_MethodSelect($aData, $arcdisp)
	If $sMethodSelectOverride > 0 Then
		Cout("Method select override active, selected choice " & $sMethodSelectOverride)
		Return $sMethodSelectOverride
	EndIf

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
						Cout("Selected method: " & $i + 1)
						Return $i + 1
					EndIf
				Next
			; Exit if Cancel clicked or window closed
			Case $GUI_EVENT_CLOSE, $idCancel
				If $createdir Then DirRemove($outdir, 0)
				terminate($STATUS_SILENT)
		EndSwitch
	WEnd
EndFunc

; Display UI for extraction method selection (list-based)
Func GUI_MethodSelectList($aEntries, $sStandard = "", $sText = "METHOD_GAME_LABEL")
	If $sMethodSelectOverride > 0 Then
		Local $iLen = UBound($aEntries)
		Local $iIndex = $sMethodSelectOverride - 1
		If $iIndex < 1 Then Return 0

		If $iLen < $iIndex Then
			Cout("Invalid method select override: index is " & $iIndex & ", but only " & $iLen & " choices available")
		Else
			Cout("Method select override active, selected option " & $iIndex)
			Return $aEntries[$iIndex - 1]
		EndIf
	EndIf

	Local $sSelection = 0
	If $silentmode Then Return $sSelection

	Local $hGUI = GUICreate($title, 274, 460, -1, -1, BitOR($WS_SIZEBOX, $WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU))
	_GuiSetColor()
	GUICtrlCreateLabel(t($sText, CreateArray($filenamefull, $sStandard, t('CANCEL_BUT'))), 10, 8, 252, 144, $SS_CENTER)
	Local $idList = GUICtrlCreateList("", 24, 150, 225, 270, BitOR($WS_VSCROLL, $WS_HSCROLL, $LBS_NOINTEGRALHEIGHT))
	GUICtrlSetData(-1, $sStandard & '|' & _ArrayToString($aEntries))
	Local $idOk = GUICtrlCreateButton(t('OK_BUT'), 40, 427, 81, 25)
	Local $idCancel = GUICtrlCreateButton(t('CANCEL_BUT'), 152, 427, 81, 25)
	_GUICtrlListBox_UpdateHScroll($idList)
	_GUICtrlListBox_SetCurSel($idList, 0)
	GUISetState(@SW_SHOW)
	Opt("GUIOnEventMode", 0)

	While 1
		$nMsg = GUIGetMsg($hGUI)
		Switch $nMsg
			Case $idOk
				$sSelection = GUICtrlRead($idList)
				If $sSelection == $sStandard Then $sSelection = 0
				ExitLoop
			Case $GUI_EVENT_CLOSE, $idCancel
				$sSelection = -1
				ExitLoop
		EndSwitch
	WEnd

	GUIDelete($hGUI)
	Opt("GUIOnEventMode", 1)

	Return $sSelection
EndFunc

; Display file scan result
Func _GUI_FileScan()
	Opt("GUIOnEventMode", 0)

	Local $sFileType = _FiletypeGet(True, 48)
	Local $iCount = StringSplit($sFileType, @CR)[0]

	Local $hGUI = GUICreate($name, 454, 248)
	_GuiSetColor()
	GUICtrlCreateLabel(t('FILESCAN_TITLE'), 80, 10, 368, 17)
	GUICtrlSetFont(-1, 9, 600, 4)
	Local $idEdit = GUICtrlCreateEdit($sFileType, 81, 26, 367, 181, BitOR($ES_READONLY, $ES_MULTILINE, $iCount > 14? $WS_VSCROLL: 0), $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 8.5, 0, 0, "Courier New")
	Local $idOk = GUICtrlCreateButton(t('OK_BUT'), 362, 214, 81, 25)
	_GUICtrlCreatePic($sLogoFile, 4, 12, 73, 73)
	Local $idCopy = GUICtrlCreateButton(t('COPY_BUT'), 260, 214, 81, 25)

	GUICtrlSetBkColor($idEdit, $COLOR_WHITE)
	GUISetState(@SW_SHOWNORMAL)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $idOk
				ExitLoop
			Case $idCopy
				$aReturn = _GUICtrlEdit_GetSel($idEdit)
				$iLen = $aReturn[1] - $aReturn[0]
				ClipPut($iLen < 1? $sFileType: StringMid($sFileType, $aReturn[0], $iLen + 1))
		EndSwitch
	WEnd

	GUIDelete($hGUI)
	Opt("GUIOnEventMode", 1)
EndFunc

; Display unknown file type error message with file scan result box
Func GUI_Error_UnknownExt()
	If $silentmode Then Return

	Opt("GUIOnEventMode", 0)
	Local $idEdit, $idCopy

	Local $sFileType = _FiletypeGet(True, 50)
	Local Const $bHasResult = StringLen($sFileType) > 0
	Local Const $iHeight = $bHasResult? 290: 190
	Local Const $iPosY = $iHeight - 34

	Local $hGUI = GUICreate($name, 488, $iHeight)
	_GuiSetColor()

	GUICtrlCreateLabel(t('UNKNOWN_FILETYPE_TITLE'), 96, 10, 375, 28)
	GUICtrlSetFont(-1, 16, 400, 4, $FONT_ARIAL)
	GUICtrlCreateLabel(t('UNKNOWN_FILETYPE', $filenamefull), 96, 42, 375, 71)
	Local $idImage = _GUICtrlCreatePic($sLogoFile, 10, 26, 73, 73)

	If $bHasResult Then
		Local $iCount = StringSplit($sFileType, @CR)[0]
		GUICtrlCreateLabel(t('FILESCAN_TITLE'), 96, 118, 375, 17)
		GUICtrlSetFont(-1, 8.5, 0, 4, $FONT_ARIAL)
		$idEdit = GUICtrlCreateEdit($sFileType, 96, 134, 379, 115, BitOR($ES_READONLY, $ES_MULTILINE, $iCount > 7? $WS_VSCROLL: 0), $WS_EX_CLIENTEDGE)
		GUICtrlSetFont(-1, 8.5, 0, 0, "Courier New")
		GUICtrlSetBkColor($idEdit, $COLOR_WHITE)
		$idCopy = GUICtrlCreateButton(t('COPY_BUT'), 296, $iPosY, 81, 25)
	EndIf

	Local $idOk = GUICtrlCreateButton(t('OK_BUT'), 395, $iPosY, 81, 25)
	Local $idFeedback = GUICtrlCreateButton("Feedback", 95, $iPosY, 81, 25)

	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $idOk
				ExitLoop
			Case $idCopy
				$aReturn = _GUICtrlEdit_GetSel($idEdit)
				$iLen = $aReturn[1] - $aReturn[0]
				ClipPut($iLen < 1? $sFileType: StringMid($sFileType, $aReturn[0], $iLen + 1))
			Case $idFeedback
				GUIDelete($hGUI)
				GUI_Feedback()
				ExitLoop
			Case $idImage
				Run($exeinfope & ' "' & $file & '"', $filedir)
		EndSwitch
	WEnd

	GUIDelete($hGUI)
	Opt("GUIOnEventMode", 1)
EndFunc

; Custom update found message with changelog display
Func GUI_UpdatePrompt()
	Local $bChoice = False
	Opt("GUIOnEventMode", 0)

	Local $hGUI = GUICreate($name, 454, 314, -1, -1, -1, -1, $guimain)
	_GuiSetColor()
	GUICtrlCreateLabel(t('UPDATE_PROMPT', $name), 72, 12, 372, 40)
	GUICtrlCreateLabel(t('UPDATE_WHATS_NEW'), 8, 64, 372, 17)
	Local $idEdit = GUICtrlCreateEdit(t('TERM_LOADING'), 8, 80, 440, 193, BitOR($ES_READONLY, $WS_VSCROLL), $WS_EX_STATICEDGE)
	Local $idYes = GUICtrlCreateButton(t('YES_BUT'), 272, 280, 75, 25)
	Local $idNo = GUICtrlCreateButton(t('NO_BUT'), 368, 280, 75, 25)
	_GUICtrlCreatePic($sLogoFile, 8, 8, 48, 48)
	GUISetState(@SW_SHOW)

	Local $return = _INetGetSource($sUpdateURL & "news")
	If @error Then $return = t('DOWNLOAD_FAILED', "'" & t('UPDATE_WHATS_NEW') & "'")
	GUICtrlSetData($idEdit, $return)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $idNo
				ExitLoop
			Case $idYes
				$bChoice = True
				ExitLoop
		EndSwitch
	WEnd

	GUIDelete($hGUI)
	Opt("GUIOnEventMode", 1)
	Return $bChoice
EndFunc

; Create Plugin Manager GUI
Func GUI_Plugins($hParent = 0, $sSelection = 0)
	If @NumParams < 1 Then Local $hParent = $guimain, $sSelection = 0

	; Define plugins
	; executable|name|description|filetypes|filemask|extractionfilter|outdir|password
	Local $aPluginInfo[12][8] = [ _
		[$arc_conv, 'arc_conv', t('PLUGIN_ARC_CONV'), 'nsa, wolf, xp3, ypf', 'arc_conv_r*.7z', 'arc_conv.exe', '', 'I Agree'], _
		[$thinstall, 'h4sh3m Virtual Apps Dependency Extractor', t('PLUGIN_THINSTALL'), 'exe (Thinstall)', 'Extractor.rar', '', '', 'h4sh3m'], _
		[$iscab, 'iscab', t('PLUGIN_ISCAB'), 'cab', 'iscab.exe;ISTools.dll', '', '', 0], _
		[$unreal, 'Unreal Engine Resource Viewer', t('PLUGIN_UNREAL'), 'pak, u, uax, upk', 'umodel_win32.zip', 'umodel.exe|SDL2.dll', '', 0], _
		[$dcp, 'WinterMute Engine Unpacker', t('PLUGIN_WINTERMUTE'), 'dcp', $dcp, '', '', 0], _
		[$ci, 'CreateInstall Extractor', t('PLUGIN_CI', CreateArray("ci-extractor.exe", "gea.dll", "gentee.dll")), 'exe (CreateInstall)', 'ci-extractor.exe;gea.dll;gentee.dll', '', '', 0], _
		[$dgca, 'DGCA', t('PLUGIN_DGCA'), 'dgca', 'dgca_v*.zip', 'dgcac.exe', '', 0], _
		[$bootimg, 'bootimg', t('PLUGIN_BOOTIMG'), 'boot.img', 'unpack_repack_kernel_redmi1s.zip', 'bootimg.exe', '', 0], _
		[$is5cab, 'is5comp', t('PLUGIN_IS5COMP'), 'cab (InstallShield)', 'i5comp21.rar', 'I5comp.exe|ZD50149.DLL|ZD51145.DLL', '', 0], _
		[$sim, 'Smart Install Maker unpacker', t('PLUGIN_SIM'), 'exe (Smart Install Maker)', 'sim_unpacker.7z', $sim, '', 0], _
		[$wolf, 'WolfDec', t('PLUGIN_WOLF'), 'wolf (' & t('TERM_ENCRYPTED') & ')', $wolf, '', '', 0], _
		[$extsis, 'ExtSIS', t('PLUGIN_EXTSIS'), 'sis, sisx', 'siscontents*.zip', $extsis, '', 0] _
	]

	Local Const $sSupportedFileTypes = t('PLUGIN_SUPPORTED_FILETYPES')
	Local $current = -1, $sWorkingDir = @WorkingDir, $aReturn[0], $iIndex = -1
	If $sSelection Then $iIndex = _ArraySearch($aPluginInfo, $sSelection, 0, 0, 0, 0, 1, 0)
	FileChangeDir(@UserProfileDir)

	$GUI_Plugins = GUICreate($name, 410, 167, -1, -1, -1, -1, $hParent)
	_GuiSetColor()
	$GUI_Plugins_List = GUICtrlCreateList("", 8, 8, 209, 149)
	GUICtrlSetData(-1, _ArrayToString($aPluginInfo, "|", -1, -1, "|", 1, 1))
	If $iIndex > -1 Then GUICtrlSetData($GUI_Plugins_List, $aPluginInfo[$iIndex][1])
	$GUI_Plugins_SelectClose = GUICtrlCreateButton(t('FINISH_BUT'), 320, 132, 83, 25)
	$GUI_Plugins_Download = GUICtrlCreateButton(t('TERM_DOWNLOAD'), 224, 132, 83, 25)
	GUICtrlSetState(-1, $GUI_DISABLE)
	$GUI_Plugins_Description = GUICtrlCreateEdit("", 224, 8, 177, 85, BitOR($ES_AUTOVSCROLL, $ES_WANTRETURN, $ES_READONLY, $ES_NOHIDESEL,$ES_MULTILINE))
	$GUI_Plugins_FileTypes = GUICtrlCreateEdit("", 224, 96, 177, 33, BitOR($ES_AUTOVSCROLL, $ES_WANTRETURN, $ES_READONLY, $ES_NOHIDESEL,$ES_MULTILINE))
	$current = GUI_Plugins_Update($GUI_Plugins_List, $GUI_Plugins_FileTypes, $GUI_Plugins_Description, $GUI_Plugins_Download, $GUI_Plugins_SelectClose, $sSupportedFileTypes, $aPluginInfo)
	GUISetState(@SW_SHOW)

	Opt("GUIOnEventMode", 0)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $GUI_Plugins_List
				$current = GUI_Plugins_Update($GUI_Plugins_List, $GUI_Plugins_FileTypes, $GUI_Plugins_Description, $GUI_Plugins_Download, $GUI_Plugins_SelectClose, $sSupportedFileTypes, $aPluginInfo)
			Case $GUI_Plugins_SelectClose
				If $current = -1 Or HasPlugin($aPluginInfo[$current][0], True) Then ExitLoop

				Cout("Adding plugin " & $aPluginInfo[$current][1])
				$return = FileOpenDialog(t('OPEN_FILE'), _GetFileOpenDialogInitDir(), $aPluginInfo[$current][1] & " (" & $aPluginInfo[$current][4] & ")", 4+1, "", $GUI_Plugins)
				If @error Then ContinueLoop
				GUICtrlSetState($GUI_Plugins_SelectClose, $GUI_DISABLE)
				Cout("Plugin file selected: " & $return)
				If $aPluginInfo[$current][6] = "" Then $aPluginInfo[$current][6] = $bindir

				; Check permissions
				If Not CanAccess($aPluginInfo[$current][6]) Then
					If IsAdmin() Then MsgBox($iTopmost + 64, $title, t('ACCESS_DENIED'))
					MsgBox($iTopmost + 64, $title, t('ELEVATION_REQUIRED'))
					ShellExecute($sUpdater, "/pluginst")
					terminate($STATUS_SILENT)
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
						Else ; Multiple files are returned as path|file1|fileN
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

	FileChangeDir($sWorkingDir)	; Reset working dir in case it was changed by FileOpenDialog
	GUIDelete($GUI_Plugins)
	Opt("GUIOnEventMode", 1)
EndFunc

; Update Plugin Manager after list selecton has changed
Func GUI_Plugins_Update($GUI_Plugins_List, $GUI_Plugins_FileTypes, $GUI_Plugins_Description, $GUI_Plugins_Download, $GUI_Plugins_SelectClose, $sSupportedFileTypes, $aPluginInfo)
	GUICtrlSetData($GUI_Plugins_FileTypes, $sSupportedFileTypes)
	Local $iIndex = _GUICtrlListBox_GetCurSel($GUI_Plugins_List)
	If $iIndex < 0 Then Return -1

	$iIndex = _ArraySearch($aPluginInfo, _GUICtrlListBox_GetText($GUI_Plugins_List, $iIndex))
	If @error Then Return -1

	GUICtrlSetState($GUI_Plugins_Download, $GUI_DISABLE)
	GUICtrlSetData($GUI_Plugins_Description, $aPluginInfo[$iIndex][2])
	GUICtrlSetData($GUI_Plugins_FileTypes, $sSupportedFileTypes & " " & $aPluginInfo[$iIndex][3])

	If @Compiled And HasPlugin($aPluginInfo[$iIndex][0], True) Then
		GUICtrlSetData($GUI_Plugins_Download, t('TERM_INSTALLED'))
		GUICtrlSetData($GUI_Plugins_SelectClose, t('FINISH_BUT'))
	Else ; Not installed
		GUICtrlSetData($GUI_Plugins_Download, t('TERM_DOWNLOAD'))
		GUICtrlSetData($GUI_Plugins_SelectClose, t('SELECT_FILE'))
		GUICtrlSetState($GUI_Plugins_Download, $GUI_ENABLE)
	EndIf

	Return $iIndex
EndFunc

; Open most recent log file
Func GUI_OpenLastLog()
	Local $aFiles = _FileListToArray($logdir, "*.log", $FLTA_FILES, True)
	If @error Or $aFiles[0] < 1 Then Return

	Local $iIndex = UBound($aFiles) - 1
	ShellExecute($aFiles[$iIndex])
EndFunc

; Open log directory
Func GUI_OpenLogDir()
	ShellExecute($logdir)
EndFunc

; Option to delete all log files
Func GUI_DeleteLogs()
	Cout("Deleting log files")
	Local $hSearch, $return, $i = 0

	$hSearch = FileFindFirstFile($logdir & "*.log")
	If $hSearch == -1 Then Return

	While 1
		$return = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		FileDelete($logdir & $return)
		$i += 1
	WEnd

	FileClose($hSearch)
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

	Local $sTitle = StringReplace(t('MENU_HELP_STATS_LABEL'), "&", "")
	$GUI_Stats = GUICreate($sTitle, 730, 434, 315, 209)
	$GUI_Stats_Status_Pie = GUICtrlCreatePic("", 8, 72, 209, 209)
	$GUI_Stats_Types_Pie = GUICtrlCreatePic("", 368, 72, 353, 353)
	$GUI_Stats_Types_Legend = GUICtrlCreatePic("", 8, 312, 337, 113)
	$GUI_Stats_Status_Legend = GUICtrlCreatePic("", 232, 72, 113, 209)
	GUICtrlCreateLabel($sTitle, 8, 8, 715, 33, $SS_CENTER)
	GUICtrlSetFont(-1, 18, 500, 0, $FONT_ARIAL)
	GUICtrlCreateLabel(t('STATS_HEADER_STATUS'), 8, 48, 212, 24, $SS_CENTER)
	GUICtrlSetFont(-1, 12, 300, 0, $FONT_ARIAL)
	GUICtrlCreateLabel(t('STATS_HEADER_TYPE'), 368, 48, 354, 24, $SS_CENTER)
	GUICtrlSetFont(-1, 12, 300, 0, $FONT_ARIAL)
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
		Local $hFile = FileOpen($sPasswordFile, $FO_APPEND)
		FileClose($hFile)
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
	Local Const $iWidth = 437, $iHeight = 285
	Cout("Creating about GUI")

	Local $hGUI = _GUICreate($title & " " & $codename, $iWidth, $iHeight, -1, -1, -1, $exStyle, $guimain)
	_GuiSetColor()
	GUICtrlCreateLabel($name, 16, 16, $iWidth - 32, 52, $SS_CENTER)
	GUICtrlSetFont(-1, 25, 400, 0, $FONT_ARIAL)
	GUICtrlCreateLabel(t('ABOUT_VERSION', CreateArray($sVersion, FileGetVersion($sUniExtract, "Timestamp"))), 16, 72, $iWidth - 32, 17, $SS_CENTER)
	GUICtrlCreateLabel(t('ABOUT_INFO_LABEL', CreateArray("Jared Breland <jbreland@legroom.net>", "uniextract@bioruebe.com", "TrIDLib (C) 2008 - 2011 Marco Pontello" & @CRLF & "<http://mark0.net/code-tridlib-e.html>", "GNU GPLv2")), 16, 104, $iWidth - 32, -1, $SS_CENTER)
	GUICtrlCreateLabel($ID, 5, $iHeight - 15, 175, 15)
	GUICtrlSetFont(-1, 8, 800, 0, $FONT_ARIAL)
	_GUICtrlCreatePic(@ScriptDir & "\support\Icons\Bioruebe.png", $iWidth - 100 - 10, $iHeight - 48 - 10, 100, 48)
	Local $idOk = GUICtrlCreateButton(t('OK_BUT'), $iWidth / 2 - 45, $iHeight - 50, 90, 25)
	GUISetState(@SW_SHOW)

	GUICtrlSetOnEvent($idOk, "GUI_Close")
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_Close")
EndFunc

; Create a GUI and save window handle
Func _GUICreate($sTitle, $iWidth, $iHeight, $iLeft = -1, $iTop = -1, $iStyle = -1, $iExStyle = -1, $hParent = 0)
	Local $hGUI = GUICreate($sTitle, $iWidth, $iHeight, $iLeft, $iTop, $iStyle, $iExStyle, $hParent)
	_ArrayAdd($aGUIs, $hGUI)
	Return $hGUI
EndFunc

; Close active GUI
; This makes it possible to have multiple windows open and close the correct one
; via OnEventMode without having to create a wrapper function for each GUI
Func GUI_Close()
	For $hGUI In $aGUIs
		If WinActive($hGUI) Then ExitLoop
	Next

	; Fallback: use the last created GUI if finding the active one fails
	If Not $hGUI Then $hGUI = _ArrayPop($aGUIs)

	GUIDelete($hGUI)
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
	GUI_SavePosition()
	terminate($STATUS_SILENT)
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
		$bOptNoStatusBox = 0
		If $TBgui Then GUISetState(@SW_SHOW, $TBgui)
		TrayItemSetState($Tray_Statusbox, $TRAY_UNCHECKED)
	Else
		$bOptNoStatusBox = 1
		If $TBgui Then GUISetState(@SW_HIDE, $TBgui)
		TrayItemSetState($Tray_Statusbox, $TRAY_CHECKED)
	EndIf

	SavePref('nostatusbox', $bOptNoStatusBox)
EndFunc

; Exit and close helper binaries if necessary
Func Tray_Exit()
	Cout("Tray exit, helper PID: " & $run)
	KillHelper()
	GUI_SavePosition()

	If Not $guimain Then CreateLog("trayexit")

	terminate($STATUS_SILENT)
EndFunc
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=.\Support\Icons\uniextract_exe.ico
#AutoIt3Wrapper_Outfile=.\UniExtract.exe
#AutoIt3Wrapper_Outfile_x64=.\UniExtract64.exe
#AutoIt3Wrapper_Res_Comment=Compiled with AutoIt http://www.autoitscript.com/
#AutoIt3Wrapper_Res_Description=Universal Extractor
#AutoIt3Wrapper_Res_Fileversion=2.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=GNU General Public License v2
#AutoIt3Wrapper_Res_Field=Author|Jared Breland <jbreland@legroom.net>
#AutoIt3Wrapper_Res_Field=Homepage|http://www.legroom.net/software
#AutoIt3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_AU3Check_Parameters=-w 4 -w 5
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
#include <SQLite.dll.au3>
#include <StaticConstants.au3>
#include <String.au3>
#include <WinAPI.au3>
#include <WindowsConstants.au3>
#include "HexDump.au3"

Const $name = "Universal Extractor"
Const $version = "2.0.0 Beta 1"
Const $codename = '"Back from the grave"'
Const $title = $name & " v" & $version
Const $website = "http://www.legroom.net/software/uniextract"
Const $website2 = "http://bioruebe.com/uniextract"
Const $websiteGithub = "https://github.com/Bioruebe/UniExtract2"
Const $forum = "http://www.msfn.org/board/forum/159-universal-extractor/"
Const $updateURL = "http://update.bioruebe.com/uniextract/update.php"
Const $supportURL = "http://support.bioruebe.com/uniextract/upload.php"
Const $prefs = @ScriptDir & "\UniExtract.ini"
Const $unicodepattern = "(?i)(?m)^[\w\Q @!§$%&/\()=?,.-:+~'²³{[]}*#ß°^âëöäüîêôûïáéíóúàèìòù\E]+$"
;~ Const $cmd = @ComSpec & ' /d /k '
Const $cmd = @ComSpec & ' /d /c '
Opt("GUIOnEventMode", 1)
Opt("TrayOnEventMode", 1)
Opt("TrayMenuMode", 1 + 2)
Opt("TrayIconDebug", 1)
GUIRegisterMsg($WM_GETMINMAXINFO, "GUI_WM_GETMINMAXINFO") ; minimum window size for feedback GUI

; Preferences
Global $langdir = @ScriptDir & "\lang"
Global $batchQueue = @ScriptDir & "\batch.queue"
Global $fileScanLogFile = @ScriptDir & "\log\filescan.txt"
Global $batchEnabled = 0
Global $height = @DesktopHeight / 4
Global $globalprefs = 1
Global $language = ""
Global $history = 1
Global $appendext = 0
Global $warnexecute = 1
Global $freeSpaceCheck = 1
Global $NoBox = 0
Global $OpenOutDir = 0
Global $DeleteOrigFile = 2 ; 0 No  1 Yes  2 Ask
Global $Timeout = 60000 ; milliseconds
Global $updateinterval = 1 ; days
Global $lastupdate = "2010/12/05" ; release date of last version
Global $addassocenabled = 0
Global $addassocallusers = 0
Global $addassoc = ""
Global $ID = ""
Global $FB_ask = 0
Global $Opt_ConsoleOutput = 0
Global $Log = 0
Global $CheckGame = 1
Global $KeepOutdir = 0
Global $KeepOpen = 0
Global $silentmode = 0
Global $extract = 1
Global $checkUnicode = 1
Global $StoreGUIPosition = 0
Global $posx = -1, $posy = -1
Global $trayX = -1, $trayY = -1

; Global variables
Dim $file, $filename, $filedir, $fileext, $initoutdir, $outdir, $filetype = "", $initdirsize
Dim $prompt, $packed, $return, $Output, $langlist, $notpacked
Dim $gaDropFiles[1], $queueArray[1]
Dim $About, $Type, $win7, $silent, $bIsUnicode = False, $reg64 = ""
Dim $debug = "", $guimain = False, $success = False, $TBgui, $isofile = 0
Dim $test, $testarj, $testace, $test7z, $testzip, $testie, $testinno
Dim $innofailed, $arjfailed, $acefailed, $7zfailed, $zipfailed, $iefailed, $isfailed, $isofailed, $tridfailed = 0
Dim $oldpath, $oldoutdir, $sUnicodeName
Dim $createdir, $dirmtime
Dim $FB_finish = False, $TrayMsg_Status, $BatchBut, $Tray_File
Dim $isexe = False, $Message, $run = 0, $runtitle, $DeleteOrigFileOpt[3]
Dim $queueArray[0]
Dim $section = "UniExtract Preferences"

; Check if OS is 64 bit version
If @OSArch == "X64" Or @OSArch == "IA64" Then
	Global $OSArch = "x64"
	Global $reg64 = 64
Else
	Global $OSArch = "x86"
EndIf

; Extractors
Const $7z = '7z.exe' ;x64									;9.20  ; Problems with newer versions --> tee
Const $7zsplit = "7ZSplit.exe" 								;0.2
Const $ace = "xace.exe" 									;2.6
Const $alz = "unalz.exe" 									;0.64
Const $arc = "arc.exe" 										;5.21i
Const $arj = "arj.exe" 										;3.10
Const $aspack = "AspackDie.exe" 							;1.4.1
Const $daa = "daa2iso.exe" 									;0.1.7e
Const $dmg = "dmgextractor.jar" 							;0.70	;Java
Const $ethornell = "ethornell.exe" 							;unknown
Const $exeinfope = "exeinfope.exe" 							;0.0.3.7
Const $filetool = @ScriptDir & "\bin\file\bin\file.exe" 	;5.03
Const $flac = "flac.exe"									;1.3.1
Const $flv = "FLVExtractCL.exe" 							;1.6.2
Const $freearc = "unarc.exe"								;0.666
Const $fsb = "fsbext.exe" 									;0.3.3
Const $gcf = "GCFScape.exe" ;x64							;1.8.2
Const $hlp = "helpdeco.exe" 								;2.1
Const $img = "EXTRNT.EXE" 									;2.10
Const $inno = "innounp.exe" 								;0.40
Const $is6cab = "i6comp.exe" 								;0.2
Const $isxunp = "IsXunpack.exe" 							;0.99
Const $kgb = @ScriptDir & "\bin\kgb\kgb2_console.exe" 		;1.2.1.24
Const $lit = "clit.exe" 									;1.8
Const $lzo = "lzop.exe" 									;1.03
Const $lzx = "unlzx.exe" 									;1.21
Const $mht = "extractMHT.exe" 								;1.0
Const $msi_msix = "MsiX.exe" 								;1.0
Const $msi_jsmsix = "jsMSIx.exe" 							;1.11.0704
Const $nbh = "NBHextract.exe" 								;1.0
Const $pea = "pea.exe" 										;0.12/1.0
Const $peid = "peid.exe" 									;0.95   2012/04/24
Const $quickbms = "quickbms.exe" 							;0.5.12
Const $rai = "RAIU.EXE" 									;0.1a
Const $rar = "unrar.exe" 									;5.21
Const $sit = "Expander.exe" 								;6.0
Const $stix = "stix_d.exe" 									;2001/06/13
Const $swf = "swfextract.exe" 								;0.9.1
Const $trid = "trid.exe" 									;2.10	2012/05/06
Const $ttarch = "ttarchext.exe"								;0.2.4
Const $uharc = "UNUHARC06.EXE" 								;0.6b
Const $uharc04 = "UHARC04.EXE" 								;0.4
Const $uharc02 = "UHARC02.EXE" 								;0.2
Const $uif = "uif2iso.exe" 									;0.1.7c
Const $unity = "disunity.bat" 								;0.3.2
Const $unshield = "unshield.exe" 							;0.5
Const $upx = "upx.exe" 										;3.08w
Const $rpa = @ScriptDir & "\bin\unrpa\unrpa.exe"			;1.4 @Git-13 Dec 2014	;modified to include a progress indicator
Const $uu = "uudeview.exe" 									;0.5pl20
Const $wise_ewise = "e_wise_w.exe" 							;2002/07/01
Const $wise_wun = "wun.exe" 								;0.90A
Const $zip = "unzip.exe" 									;6.00
Const $zoo = "unzoo.exe" 									;4.5

; Plugins
Const $bms = "BMS.bms"
Const $dbx = "dbxplug.wcx"
Const $gaup = "gaup_pro.wcx"
Const $ie = "InstExpl.wcx"
Const $iso = "Iso_" & $OSArch & ".wcx" ;x64
Const $mht_plug = "MhtUnPack.wcx"
Const $msi_plug = "msi.wcx"
Const $sis = "PDunSIS.wcx"

; Other
Const $tee = "mtee.exe"
Const $mediainfo = "MediaInfo" & $reg64 & ".dll"			; 0.7.72

; Not included binaries
Const $ffmpeg = "ffmpeg.exe"	;x64
Const $iscab = "iscab.exe"
Const $thinstall = "Extractor.exe"
Const $rgss3 = "RPGDecrypter-spaces.ru.exe"
Const $arc_conv = "arc_conv.exe"
Const $dcp = "dcp_unpacker.exe"
Const $unreal = "extract.exe"
Const $crage = @ScriptDir & "\bin\crass-0.4.14.0\crage.exe"
Const $faad = "faad.exe"

If Not @Compiled Then HotKeySet("{^}", "Test")

; Define registry keys
Global Const $reg = "HKCU" & $reg64 & "\Software\UniExtract"
Global Const $regcurrent = "HKCU" & $reg64 & "\Software\Classes\*\shell\"
Global Const $regall = "HKCR" & $reg64 & "\*\shell\"
Global $reguser = $regcurrent

ReadPrefs()

; Define context menu commands
; On top to make remove via command line parameter possible
; After ReadPrefs!
; shell	| commandline parameter | translation
Global $CM_Shells[4][3] = [['uniextract_files', '', t('EXTRACT_FILES')], ['uniextract_here', ' .', t('EXTRACT_HERE')], _
		['uniextract_sub', ' /sub', t('EXTRACT_SUB')], ['uniextract_scan', ' /scan', t('SCAN_FILE')]]

Cout("Starting " & $name & " " & $version)

; Check commandline parameters
If $cmdline[0] = 0 Then
	$prompt = 1
Else
	ParseCommandLine()
EndIf

; Create tray menu items
$Tray_Statusbox = TrayCreateItem(t('PREFS_HIDE_STATUS_LABEL'))
If $NoBox Then TrayItemSetState(-1, $TRAY_CHECKED)
TrayCreateItem("")
$Tray_Exit = TrayCreateItem(t('MENU_FILE_QUIT_LABEL'))

TrayItemSetOnEvent($Tray_Statusbox, "Tray_Statusbox")
TrayItemSetOnEvent($Tray_Exit, "Tray_Exit")
TraySetToolTip($name)
TraySetClick(8)

; Set environment options for Windows NT, automatically reverted on exit
EnvSet("path", @ScriptDir & "\bin" & ';' & EnvGet("path"))
EnvSet("path", @ScriptDir & "\bin\" & $OSArch & ';' & EnvGet("path"))

; If no file passed, display GUI to select file and set options
If $prompt Then
	; Check if Universal Extractor is started the first time
	If $ID = "" Or StringIsSpace($ID) Then
		$ID = StringRight(String(_Crypt_EncryptData(Random(10000, 1000000), @ComputerName & Random(10000, 1000000), $CALG_AES_256)), 25)
		Cout("Created User ID: " & $ID)
		SavePref("ID", $ID)
		Global $FS_GUI = False
		GUI_FirstStart()
		While $FS_GUI
			Sleep(250)
		WEnd
	EndIf

	; Check for updates
	If _DateDiff("D", $lastupdate, _NowCalc()) >= $updateinterval Then CheckUpdate(True)

	CreateGUI()
	While 1
		If Not $guimain Then ExitLoop
		Sleep(100)
	WEnd
EndIf

; Update history
If $history Then
	WriteHist('file', $file)
	WriteHist('directory', $outdir)
EndIf

; Prevent multiple instances to avoid errors
; Only necessary when extract starts
; Do not do this in StartExtraction, the function can be called twice
If _Singleton($name & " " & $version, 1) = 0 Then
	AddToBatch()
	terminate("silent", '', '')
EndIf

StartExtraction($extract)

; -------------------------- Begin Custom Functions ---------------------------

; Start extraction process
Func StartExtraction($checkext = True)
	Cout("------------------------------------------------------------")

	$bIsUnicode = False

	; Parse filename
	FilenameParse($file)

	; Collect file information, for log/feedback only
	Cout("File size: " & Round(FileGetSize($file) / 1048576, 2) & " MB")
	Cout("Created " & FileGetTime($file, 1, 1) & ", modified " & FileGetTime($file, 0, 1))

	; Set full output directory
	If $outdir = '/sub' Then
		$outdir = $initoutdir
	ElseIf StringMid($outdir, 2, 1) <> ":" Then
		If StringLeft($outdir, 1) == '\' And StringMid($outdir, 2, 1) <> '\' Then
			$outdir = StringLeft($filedir, 2) & $outdir
		ElseIf StringLeft($outdir, 2) <> '\\' Then
			$outdir = _PathFull($filedir & '\' & $outdir)
		EndIf
	EndIf

	; Set filename as tray icon tooltip
	TraySetToolTip($filename & "." & $fileext)

	; Check for unicode characters in path
	If $checkUnicode And Not StringRegExp($file, $unicodepattern, 0) Then
		Cout("File seems to be unicode")
		$bIsUnicode = True
		$sUnicodeName = $filename
		$oldoutdir = $outdir
		$oldpath = $file
		$file = _TempFile($filedir, "Unicode_", $fileext)
		Cout('Renaming "' & $sUnicodeName & '" to "' & $file & '"')
		FileMove($oldpath, $file)
		FilenameParse($file)
		$outdir = $initoutdir
	EndIf

	; Register tray function
	TraySetOnEvent($TRAY_EVENT_PRIMARYUP, "Tray_ShowHide")

	; Reset variables
	$isexe = False
	$innofailed = False
	$arjfailed = False
	$acefailed = False
	$7zfailed = False
	$zipfailed = False
	$iefailed = False
	$isfailed = False
	$testinno = False
	$testarj = False
	$testace = False
	$test7z = False
	$testzip = False
	$testie = False
	$packed = False
	$filetype = ""

	; Extract contents from known file types
	;
	; UniExtract uses four methods of detection (in order):
	; 1. File extensions for special cases
	; 2. Binary file analysis of files using TrID if file extension is not .exe
	; 3. Binary file analysis of PE (executable) files using Exeinfo PE
	; 4. Extra analysis using PeID if executable is not recognized by Exeinfo PE
	; 5. Binary file analysis of files using TrID
	; 6. File extensions
	;
	; If detection fails, extraction is always attempted 7-Zip and InfoZip

	; First, check for file extensions that require special actions
	If $checkext Then
		;Cout("Extension check")
		; Compound compressed files that require multiple actions
		If $fileext = "ipk" Or $fileext = "tbz2" Or $fileext = "tgz" _
				Or $fileext = "tz" Or $fileext = "tlz" Or $fileext = "txz" Then
			extract("ctar", 'Compressed Tar ' & t('TERM_ARCHIVE'))
			; image files - TrID is not always reliable with these formats
		ElseIf $fileext = "bin" Or $fileext = "cue" Or $fileext = "cdi" Then
			CheckIso()
		ElseIf $fileext = "dmg" Then
			extract("dmg", 'DMG ' & t('TERM_IMAGE'))
		ElseIf $fileext = "iso" Then
			check7z()
			CheckIso()
		EndIf
	EndIf

	; If file is an .exe, scan with Exeinfo PE and PEiD
	If $fileext = "exe" Or $fileext = "dll" Then IsExe()

	; Scan file with TrID, if file is not an .exe
	If Not $tridfailed Then filescan($file, $extract)

	; Terminate if scan only mode is activated
	If Not $extract Then
		MediaFileScan($file)
		terminate("fileinfo", "", "")
	EndIf

	; Else perform additional extraction methods
	If Not $isofailed Then CheckIso()
	CheckGame()

	; Use file extension if signature not recognized
	If $checkext Then
		If $fileext = "1" Or $fileext = "lib" Then
			extract("is3arc", 'InstallShield 3.x ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "7z" Then
			extract("7z", '7-Zip ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "ace" Then
			extract("ace", 'ACE ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "arc" Then
			extract("arc", 'ARC ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "arj" Then
			extract("arj", 'ARJ ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "b64" Then
			extract("uu", 'Base64 ' & t('TERM_ENCODED'))

		ElseIf $fileext = "bz2" Then
			extract("bz2", 'bzip2 ' & t('TERM_COMPRESSED'))

		ElseIf $fileext = "cab" Then
			If StringInStr(FetchStdout($cmd & $7z & ' l "' & $file & '"', $filedir, @SW_HIDE), "Listing archive:", 0) Then
				extract("cab", 'Microsoft CAB ' & t('TERM_ARCHIVE'))
			Else
				extract("iscab", 'InstallShield CAB ' & t('TERM_ARCHIVE'))
			EndIf

		ElseIf $fileext = "chm" Then
			extract("chm", 'Compiled HTML ' & t('TERM_HELP'))

		ElseIf $fileext = "cpio" Then
			extract("7z", 'CPIO ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "dbx" Then
			extract('qbms', 'Outlook Express ' & t('TERM_ARCHIVE'), $dbx)

		ElseIf $fileext = "deb" Then
			extract("7z", 'Debian ' & t('TERM_PACKAGE'))

		ElseIf $fileext = "gz" Then
			extract("gz", 'gzip ' & t('TERM_COMPRESSED'))

		ElseIf $fileext = "hlp" Then
			extract("hlp", 'Windows ' & t('TERM_HELP'))

		ElseIf $fileext = "imf" Then
			extract("cab", 'IncrediMail ' & t('TERM_ECARD'))

		ElseIf $fileext = "img" Then
			extract("img", 'Floppy ' & t('TERM_DISK') & ' ' & t('TERM_IMAGE'))

		ElseIf ($fileext = "kgb" Or $fileext = "kge") Then
			extract("kgb", 'KGB ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "lit" Then
			extract("lit", 'Microsoft LIT ' & t('TERM_EBOOK'))

		ElseIf $fileext = "lzh" Or $fileext = "lha" Then
			extract("7z", 'LZH ' & t('TERM_COMPRESSED'))

			;elseif $fileext = "lzma" then
			;	extract("7z", 'LZMA ' & t('TERM_COMPRESSED'))

		ElseIf $fileext = "lzo" Then
			extract("lzo", 'LZO ' & t('TERM_COMPRESSED'))

		ElseIf $fileext = "lzx" Then
			extract("lzx", 'LZX ' & t('TERM_COMPRESSED'))

		ElseIf $fileext = "mht" Then
			extract("mht", 'MHTML ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "msi" Then
			extract("msi", 'Windows Installer (MSI) ' & t('TERM_PACKAGE'))

		ElseIf $fileext = "msm" Then
			extract("msm", 'Windows Installer (MSM) ' & t('TERM_MERGE_MODULE'))

		ElseIf $fileext = "msp" Then
			extract("msp", 'Windows Installer (MSP) ' & t('TERM_PATCH'))

		ElseIf $fileext = "nbh" Then
			extract("nbh", 'NBH ' & t('TERM_IMAGE'))

		ElseIf $fileext = "pea" Then
			extract("pea", 'Pea ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "rar" Or $fileext = "001" Or $fileext = "cbr" Then
			extract("rar", 'RAR ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "rpm" Then
			extract("7z", 'RPM ' & t('TERM_PACKAGE'))

		ElseIf $fileext = "sis" Then
			extract('qbms', 'SymbianOS ' & t('TERM_INSTALLER'), $sis)

		ElseIf $fileext = "sit" Then
			extract("sit", 'StuffIt ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "tar" Then
			extract("tar", 'Tar ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "uha" Then
			extract("uha", 'UHARC ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "uif" Then
			extract("uif", 'UIF ' & t('TERM_IMAGE'))

		ElseIf $fileext = "uu" Or $fileext = "uue" _
				Or $fileext = "xx" Or $fileext = "xxe" Then
			extract("uu", 'UUencode ' & t('TERM_ENCODED'))

		ElseIf $fileext = "vpk" Or $fileext = "gcf" Or $fileext = "ncf" _
				Or $fileext = "wad" Or $fileext = "xzp" Then
			extract("gcf", 'Valve ' & $fileext & " " & t('TERM_PACKAGE'))

		ElseIf $fileext = "wim" Then
			extract("7z", 'WIM ' & t('TERM_IMAGE'))

		ElseIf $fileext = "yenc" Or $fileext = "ntx" Then
			extract("uu", 'yEnc ' & t('TERM_ENCODED'))

		ElseIf $fileext = "z" Then
			If Not check7z() Then extract("is3arc", 'InstallShield 3.x ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "zip" Or $fileext = "cbz" Or $fileext = "jar" Or $fileext = "xpi" Or $fileext = "wz" Then
			extract("zip", 'ZIP ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "zoo" Then
			extract("zoo", 'ZOO ' & t('TERM_ARCHIVE'))

		ElseIf $fileext = "wolf" Then
			extract("arc_conv", "Wolf RPG Editor " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		ElseIf $fileext = "rgss" Or $fileext = "rgss2a" Then
			extract("arc_conv", "RPG Maker " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		ElseIf $fileext = "rgss3a" Then
			extract("rgss3", "RPG Maker VX Ace " & t('TERM_GAME') & t('TERM_ARCHIVE'))
		EndIf
	EndIf

	check7z()

	; Cannot determine filetype, all checks failed - abort
	_DeleteTrayMessageBox()
	terminate("unknownext", $file, "")
EndFunc   ;==>StartExtraction

; Extract if exe file detected
Func IsExe()
	Cout("File seems to be executable")

	; Just for fun
	If $file = @ScriptFullPath Then
		$filetype = $name
		terminate("notpacked", $file, "")
	EndIf

	; Check executable using Exeinfo PE
	advexescan($file, $extract)

	; Check executable using PEiD
	exescan($file, 'hard', $extract)

	If Not $extract Then Return

	; Both analyse programms fail -> perhaps no .exe file
	; Perform file analysis using TrID
	If Not $isexe Then Return

	; Perform additional tests if necessary
	If $testinno And Not $innofailed Then checkInno()
	If $testarj And Not $arjfailed Then checkArj()
	If $testace And Not $acefailed Then checkAce()
	If $testzip And Not $zipfailed Then checkZip()
	If $testie And Not $iefailed Then checkIE()
	If $test7z And Not $7zfailed Then check7z()

	If Not $iefailed Then checkIE()

	; Unpack (vs. extract) packed file
	If $packed Then unpack()

	; Try 7-Zip and Unzip if all else fails
	If Not $7zfailed Then check7z()
	If Not $zipfailed Then checkZip()

	If $fileext <> "exe" Or $isexe == False Then Return

	CheckGame()

	If $notpacked Then terminate("notpacked", $file, "")

	; Scan using TrID
	filescan($file)

	; Exit with unknown file type
	terminate("unknownexe", $file, $filetype)

EndFunc   ;==>IsExe

; Parse filename
Func FilenameParse($f)
	$filedir = StringLeft($f, StringInStr($f, '\', 0, -1) - 1)
	$filename = StringTrimLeft($f, StringInStr($f, '\', 0, -1))
	If StringInStr($filename, '.') Then
		$fileext = StringTrimLeft($filename, StringInStr($filename, '.', 0, -1))
		$filename = StringTrimRight($filename, StringLen($fileext) + 1)
		$initoutdir = $filedir & '\' & $filename
	Else
		$fileext = ''
		$initoutdir = $filedir & '\' & $filename & '_' & t('TERM_UNPACKED')
	EndIf
	Cout("FilenameParse: " & @CRLF & "Raw input: " & $f & @CRLF & "FileName: " & $filename & @CRLF & "FileExt: " & $fileext & @CRLF & "FileDir: " & $filedir & @CRLF & "InitOutDir: " & $initoutdir)
EndFunc   ;==>FilenameParse

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
	Local $ldir = @ScriptDir
	If $lang <> 'English' Then $ldir &= '\lang'
	$return = IniRead($ldir & '\' & $lang & '.ini', 'UniExtract', $t, '')
	If $return == '' Then
		Cout("Translation not found for term " & $t)
		$return = IniRead(@ScriptDir & '\English.ini', 'UniExtract', $t, '???')
		If $return = "???" Then Cout("Warning: term " & $t & " is not defined")
	EndIf

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
	Cout("Command line parameters: " & _ArrayToString($cmdline, " ", 1))
;~ 	_ArrayDisplay($cmdline)

	$extract = 1
	$OpenOutDir = 0
	If $cmdline[1] = "/help" Or $cmdline[1] = "/h" Or $cmdline[1] = "/?" Or $cmdline[1] = "-h" Or $cmdline[1] = "-?" Or $cmdline[1] = "--help" Then
		terminate("syntax", "", "")

	ElseIf $cmdline[1] = "/prefs" Then
		GUI_Prefs()
		$guiprefs = False
		While $guiprefs
			Sleep(250)
		WEnd
		terminate("silent", "", "")

	ElseIf $cmdline[1] = "/afterupdate" Then
		_AfterUpdate()
		$prompt = 1

	ElseIf $cmdline[1] = "/remove" Then
		; Completely delete registry entries, used by uninstaller
		If @OSVersion = "WIN_7" Or @OSVersion = "WIN_8" Or @OSVersion = "WIN_81" Then
			Global $win7 = True
		Else
			Global $win7 = False
		EndIf
		GUI_ContextMenu_remove()
		GUI_ContextMenu_fileassoc(0)

		terminate("silent", '', '')

	ElseIf $cmdline[1] = "/clear" Then
		GUI_Batch_Clear()
		terminate("silent", '', '')

	Else
		If FileExists($cmdline[1]) Then
			$file = _PathFull($cmdline[1])
		Else
			terminate("syntax", "", "")
		EndIf
		If $cmdline[0] > 1 Then
			; Scan only
			If $cmdline[2] = "/scan" Then
				$extract = False
				$Log = False
				If $cmdline[0] > 2 And $cmdline[3] = "/silent" Then $silentmode = 1
				; Silent mode
			ElseIf $cmdline[2] = "/silent" Then
				$silentmode = 1
				; Outdir specified
			Else
				$outdir = $cmdline[2]
				If $cmdline[0] > 2 And $cmdline[3] = "/silent" Then $silentmode = 1
			EndIf
		Else
			$prompt = 1
		EndIf

		If _ArraySearch($cmdline, "/batch") > -1 Then
			AddToBatch()
			terminate("silent", '', '')
		EndIf
	EndIf
EndFunc   ;==>ParseCommandLine

; Read complete preferences
Func ReadPrefs()
	LoadPref("globalprefs", $globalprefs)
	LoadPref("consoleoutput", $Opt_ConsoleOutput)
	LoadPref("language", $language, False)
	LoadPref("batchqueue", $batchQueue, False)
	If $batchQueue Then $batchQueue = _PathFull($batchQueue, @ScriptDir)
	LoadPref("filescanlogfile", $fileScanLogFile, False)
	If Not @error Then $fileScanLogFile = _PathFull($fileScanLogFile, @ScriptDir)
	LoadPref("batchenabled", $batchEnabled, 0)
	LoadPref("history", $history)
	LoadPref("appendext", $appendext)
	LoadPref("warnexecute", $warnexecute)
	LoadPref("nostatusbox", $NoBox)
	LoadPref("openfolderafterextr", $OpenOutDir)
	LoadPref("deletesourcefile", $DeleteOrigFile)
	LoadPref("freespacecheck", $freeSpaceCheck)

	LoadPref("timeout", $Timeout)
	$Timeout *= 1000
	If $Timeout < 10000 Then $Timeout = 60000

	LoadPref("keepoutputdir", $KeepOutdir)
	LoadPref("keepopen", $KeepOpen)
	LoadPref("feedbackprompt", $FB_ask)
	LoadPref("log", $Log)
	LoadPref("CheckGame", $CheckGame)
	LoadPref("extract", $extract)
	LoadPref("unicodecheck", $checkUnicode)
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
	LoadPref("updateinterval", $updateinterval)
	LoadPref("lastupdate", $lastupdate, False)
	LoadPref("ID", $ID, False)

	If $language == "" Then $language = _GetOSLanguage()

	If $silentmode Then Cout("Silent mode enabled")
	If $batchEnabled Then Cout("Batch mode enabled")
	If Not $extract Then Cout("Scan only mode enabled")

	; Read language files
	$langlist = ''
	Dim $langarr[1]
	$temp = FileFindFirstFile($langdir & '\*.ini')
	If $temp <> -1 Then
		Local $langarr = _FileListToArray($langdir, '*.ini', 1)
		FileClose($temp)
	EndIf
	$langarr[0] = 'English.ini'
	_ArraySort($langarr)
	For $i = 0 To UBound($langarr) - 1
		$langlist &= StringTrimRight($langarr[$i], 4) & '|'
	Next
	$langlist = StringTrimRight($langlist, 1)

	Cout("Finished loading preferences")
EndFunc   ;==>ReadPrefs

; Write complete preferences
Func WritePrefs()
	Cout("Saving preferences")

	SavePref('history', $history)
	SavePref('language', $language)
	SavePref('appendext', $appendext)
	SavePref('warnexecute', $warnexecute)
	SavePref('nostatusbox', $NoBox)
	SavePref('openfolderafterextr', $OpenOutDir)
	SavePref('deletesourcefile', $DeleteOrigFile)
	SavePref('freespacecheck', $freeSpaceCheck)
	SavePref('unicodecheck', $checkUnicode)
	SavePref('feedbackprompt', $FB_ask)
	SavePref('consoleoutput', $Opt_ConsoleOutput)
	SavePref('log', $Log)
	SavePref('CheckGame', $CheckGame)
	SavePref('storeguiposition', $StoreGUIPosition)
	SavePref('timeout', $Timeout / 1000)
	SavePref('updateinterval', $updateinterval)
EndFunc   ;==>WritePrefs

; Save single preference
Func SavePref($name, $value)
	If $globalprefs Then
		IniWrite($prefs, "UniExtract Preferences", $name, $value)
	Else
		RegWrite($reg, $name, 'REG_SZ', $value)
	EndIf
	Cout("Saving: " & $name & " = " & $value)
EndFunc   ;==>SavePref

; Load single preference
Func LoadPref($name, ByRef $value, $int = True)
	Local $return = ""

	If $globalprefs Then
		$return = IniRead($prefs, $section, $name, "")
	Else
		$return = RegRead($reg, $name)
	EndIf

	If @error Or $return = "" Then
		Cout("Error reading option " & $name & " --> " & $value)
		Return SetError(1, "", -1)
	EndIf

	If $int Then
		$value = Int($return)
	Else
		$value = $return
	EndIf

	Cout("Option: " & $name & " = " & $value)
EndFunc   ;==>LoadPref

; Read history
Func ReadHist($field)
	Local $items

	; Read from .ini file when globalprefs enabled
	If $globalprefs Then
		If $field = 'file' Then
			$section = "File History"
		ElseIf $field = 'directory' Then
			$section = "Directory History"
		Else
			Return
		EndIf
		For $i = 0 To 9
			$value = IniRead($prefs, $section, $i, "")
			If $value <> "" Then $items &= '|' & $value
		Next

		; Otherwise, read from HKCU registry
	Else
		If $field = 'file' Then
			$key = $reg & "\File"
		ElseIf $field = 'directory' Then
			$key = $reg & "\Directory"
		Else
			Return
		EndIf
		For $i = 0 To 9
			$value = RegRead($key, $i)
			If $value <> "" Then $items &= '|' & $value
		Next
	EndIf

	; return trimmed results
	Return StringTrimLeft($items, 1)
EndFunc   ;==>ReadHist

; Write history
Func WriteHist($field, $new)
	; Only update .ini file if globalprefs enabled
	If $globalprefs Then
		If $field = 'file' Then
			$section = "File History"
		ElseIf $field = 'directory' Then
			$section = "Directory History"
		Else
			Return
		EndIf
		$histarr = StringSplit(ReadHist($field), '|')
		IniWrite($prefs, $section, "0", $new)
		If $histarr[1] == "" Then Return
		For $i = 1 To $histarr[0]
			If $i > 9 Then ExitLoop
			If $histarr[$i] = $new Then
				IniDelete($prefs, $section, String($i))
				ContinueLoop
			EndIf
			IniWrite($prefs, $section, String($i), $histarr[$i])
		Next

		; Otherwise, write local settings to registry
	Else
		If $field = 'file' Then
			$key = $reg & "\File"
		ElseIf $field = 'directory' Then
			$key = $reg & "\Directory"
		Else
			Return
		EndIf
		$histarr = StringSplit(ReadHist($field), '|')
		RegWrite($key, "0", "REG_SZ", $new)
		If $histarr[1] == "" Then Return
		For $i = 1 To $histarr[0]
			If $i > 9 Then ExitLoop
			If $histarr[$i] = $new Then
				RegDelete($key, String($i))
				ContinueLoop
			EndIf
			RegWrite($key, String($i), "REG_SZ", $histarr[$i])
		Next
	EndIf
EndFunc   ;==>WriteHist

; Update saved statistics
Func SaveStat($status)
	Cout("Saving Statistics")
	If $globalprefs Then
		IniWrite($prefs, "Statistics", $status, Number(IniRead($prefs, "Statistics", $status, 0)) + 1)
	Else
		RegWrite($reg & "\Statistics", $name, 'REG_SZ', Number(RegRead($reg, $name)) + 1)
	EndIf
EndFunc   ;==>SaveStat

; Scan file using TrID
Func filescan($f, $analyze = 1)
	; Scan file using unix file tool
	advfilescan($f)

	_CreateTrayMessageBox(t('SCANNING_FILE', "TrID"))

	Cout("Start filescan using TrID")
	Local $filetype_curr = ""

	; Run TrID and fetch output
	$return = StringSplit(FetchStdout($cmd & $trid & ' "' & $f & '"' & ($appendext ? " -ce" : "") & ($analyze ? "" : " -v"), $filedir, @SW_HIDE), @CRLF)
	For $i = 1 To UBound($return) - 1
		If StringInStr($return[$i], "%") Or (Not $analyze And (StringInStr($return[$i], "Related URL") Or StringInStr($return[$i], "Remarks"))) Then
			$filetype_curr &= $return[$i] & @CRLF
		EndIf
	Next
	$return = ""

	; If appendext option enabled, trid fixes the file extension, so the paths may not exist anymore.
	; We need to find the new name and reparse the file.
	If $appendext And Not FileExists($f) Then
		Cout("File was renamed by TrID, trying to determine new filename")
		Sleep(2000)
		$return = _FileListToArray($filedir & "\", $filename & ".*", 1, True)
		If @error Then
			$error = True
		ElseIf $return[0] > 1 Then
			; There may be more files with the same name, but different extensions
			Local $new = t('TERM_UNKNOWN')
			Cout("More than one possible new file name found")
			Local $ret = _StringBetween($filetype_curr, "(", ")")
			If @error Then
				$error = True
			Else
				$new = $filename & $ret[0]
				Cout("Estimated filename: " & $new)
				Local $pos = _ArraySearch($return, $filedir & "\" & $new)
				If $pos > -1 Then
					$file = $return[$pos]
					$f = $return[$pos]
					Cout("New path: " & $file)
				Else
					$error = True
				EndIf
			EndIf

			If $error Then
				Cout("Failed to locate renamed file. Guessed name: " & $new)
				If Not $silentmode Then MsgBox(262144 + 16, $name, t('RENAME_NOTFOUND', CreateArray($new, StringReplace(t('PREFS_APPEND_EXT_LABEL'), "&", ""))))
				terminate('silent', '', '')
			EndIf
		Else
			$file = $return[1]
			$f = $return[1]
		EndIf

		FilenameParse($file)
	EndIf

	If $filetype_curr <> "" Then $filetype &= $filetype_curr & @CRLF

	_DeleteTrayMessageBox()

	; Return filetype without matching if specified
	If Not $analyze Then Return

	; Match known patterns
	Select
		Case StringInStr($filetype_curr, "7-Zip compressed archive", 0)
			extract("7z", '7-Zip ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "ACE compressed archive", 0) _
				Or StringInStr($filetype_curr, "ACE Self-Extracting Archive", 0)
			extract("ace", t('TERM_SFX') & ' ACE ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "ALZip compressed archive")
			CheckAlz()

		Case StringInStr($filetype_curr, "LZIP compressed archive")
			extract("lz", "LZIP " & t('TERM_COMPRESSED') & " " & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "FreeArc compressed archive", 0)
			extract("freearc", 'FreeArc ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "ARC Compressed archive", 0) And Not StringInStr($filetype_curr, "UHARC", 0)
			extract("arc", 'ARC ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "ARJ compressed archive", 0)
			extract("arj", 'ARJ ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "bzip2 compressed archive", 0)
			extract("bz2", 'bzip2 ' & t('TERM_COMPRESSED'))

		Case StringInStr($filetype_curr, "Microsoft Cabinet Archive", 0) _
				Or StringInStr($filetype_curr, "IncrediMail letter/ecard", 0)
			extract("cab", 'Microsoft CAB ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Magic ISO Universal Image Format", 0)
			extract("uif", 'UIF ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "CDImage", 0) Or StringInStr($filetype_curr, "null bytes", 0)
			If Not $isofailed Then CheckIso()
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

		Case StringInStr($filetype_curr, "Disk Image (Macintosh)", 0)
			extract("dmg", 'DMG ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "FLAC lossless", 0)
			extract("flac", 'FLAC ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))

		Case StringInStr($filetype_curr, "Flash Video", 0)
			extract("flv", 'Flash Video ' & t('TERM_CONTAINER'))

		Case StringInStr($filetype_curr, "FMOD Sample Bank Format")
			extract("fsb", 'FMOD ' & t('TERM_CONTAINER'))

		Case StringInStr($filetype_curr, "Gentee Installer executable", 0)
			extract("ie", 'Gentee ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "GZipped File", 0)
			extract("gz", 'gzip ' & t('TERM_COMPRESSED'))

		Case StringInStr($filetype_curr, "Windows Help File", 0)
			extract("hlp", 'Windows ' & t('TERM_HELP'))

		Case StringInStr($filetype_curr, "Generic PC disk image", 0)
			If Not $isofailed Then CheckIso()
			check7z()
			extract("img", 'Floppy ' & t('TERM_DISK') & ' ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "Inno Setup installer", 0)
			checkInno()

		Case StringInStr($filetype_curr, "Installer VISE executable", 0)
			extract("ie", 'Installer VISE ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "InstallShield archive", 0)
			extract("is3arc", 'InstallShield 3.x ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "InstallShield compressed archive", 0)
			extract("iscab", 'InstallShield CAB ' & t('TERM_ARCHIVE'))

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

		Case StringInStr($filetype_curr, "Microsoft Internet Explorer Web Archive", 0)
			extract("mht", 'MHTML ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Microsoft Reader eBook", 0)
			extract("lit", 'Microsoft LIT ' & t('TERM_EBOOK'))

		Case StringInStr($filetype_curr, "Microsoft Windows Installer merge module", 0)
			extract("msm", 'Windows Installer (MSM) ' & t('TERM_MERGE_MODULE'))

		Case StringInStr($filetype_curr, "(.MSI) Microsoft Windows Installer", 0)
			extract("msi", 'Windows Installer (MSI) ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Microsoft Windows Installer patch", 0)
			extract("msp", 'Windows Installer (MSP) ' & t('TERM_PATCH'))

		Case StringInStr($filetype_curr, "HTC NBH ROM Image", 0)
			extract("nbh", 'NBH ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "Outlook Express E-mail folder", 0)
			extract('qbms', 'Outlook Express ' & t('TERM_ARCHIVE'), $dbx)

		Case StringInStr($filetype_curr, "PEA archive", 0)
			extract("pea", 'Pea ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "RAR Archive", 0)
			extract("rar", 'RAR ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "RAR Self Extracting archive", 0)
			checkZip()
			extract("rar", 'RAR ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "RPG Maker VX Ace", 0)
			extract("rgss3", "RPG Maker VX Ace " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "NScripter archive", 0)
			extract("arc_conv", "NScripter " & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "RPG Maker", 0)
			extract("arc_conv", "RPG Maker " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Telltale Games ressource archive", 0)
			extract("ttarch", "Telltale " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Wolf RPG Editor", 0)
			extract("arc_conv", "Wolf RPG Editor " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "YU-RIS Script Engine", 0)
			extract("arc_conv", "YU-RIS Script Engine " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Ethornell", 0)
			extract("ethornell", "Ethornell Engine " & t('TERM_GAME') & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Reflexive Arcade installer wrapper", 0)
			extract("inno", 'Reflexive Arcade ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Ren'Py data file", 0)
			extract("rpa", "Ren'Py " & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "RPM Linux Package", 0)
			extract("7z", 'RPM ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Setup Factory 6.x Installer", 0)
			extract("ie", 'Setup Factory ' & t('TERM_ARCHIVE'))

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

		Case StringInStr($filetype_curr, "UUencoded file", 0) _
				Or StringInStr($filetype_curr, "XXencoded file", 0)
			extract("uu", 'UUencoded ' & t('TERM_ENCODED'))

		Case StringInStr($filetype_curr, "yEnc Encoded file", 0)
			extract("uu", 'yEnc ' & t('TERM_ENCODED'))

		Case StringInStr($filetype_curr, "Valve package", 0)
			extract("gcf", 'Valve ' & $fileext & " " & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Windows Imaging Format", 0)
			extract("7z", 'WIM ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "Wise Installer Executable", 0)
			extract("wise", 'Wise Installer ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "UNIX Compressed file", 0)
			extract("Z", 'LZW ' & t('TERM_COMPRESSED'))

		Case StringInStr($filetype_curr, "xz container", 0)
			extract("xz", 'XZ ' & t('TERM_COMPRESSED'))

		Case StringInStr($filetype_curr, "ZIP compressed archive", 0) _
				Or StringInStr($filetype_curr, "Winzip Win32 self-extracting archive", 0)
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

		Case StringInStr($filetype_curr, "InstallShield setup", 0)
			;extract("isexe", 'InstallShield ' & t('TERM_INSTALLER'))
			checkInstallShield()

		Case StringInStr($filetype_curr, "Video", 0) Or StringInStr($filetype_curr, "QuickTime Movie", 0) _
				Or StringInStr($filetype_curr, "Matroska", 0)
			extract("video", t('TERM_VIDEO') & ' ' & t('TERM_FILE'))

			; not supported filetypes
		Case StringInStr($filetype_curr, "Spoon Installer", 0)
			terminate("notsupported", $f, "")

		Case StringInStr($filetype_curr, "Long Range ZIP", 0)
			terminate("notsupported", $f, "")

			; Check for .exe file, only when fileext not .exe
		Case Not $isexe And (StringInStr($filetype_curr, 'Executable', 0) Or StringInStr($filetype_curr, '(.EXE)', 1))
			IsExe()

	EndSelect
	$tridfailed = True
EndFunc   ;==>filescan

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

	Select
		Case StringInStr($filetype_curr, "7 zip archive data") Or StringInStr($filetype_curr, "7-zip archive data")
			extract("7z", '7-Zip ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "lzip compressed data")
			extract("lz", "LZIP " & t('TERM_COMPRESSED') & " " & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "Zip archive data") And Not StringInStr($filetype_curr, "7")
			extract("zip", 'ZIP ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "StuffIt Archive")
			extract("sit", 'StuffIt ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "uuencoded", 0)
			extract("uu", 'UUencoded ' & t('TERM_ENCODED'))
		Case StringInStr($filetype_curr, "UHarc archive data", 0)
			extract("uha", 'UHARC ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "ARC archive data", 0)
			extract("arc", 'ARC ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "Symbian installation file", 0)
			extract('qbms', 'SymbianOS ' & t('TERM_INSTALLER'), $sis)
		Case StringInStr($filetype_curr, "Zoo archive data", 0)
			extract("zoo", 'ZOO ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "MS Outlook Express DBX file", 0)
			extract('qbms', 'Outlook Express ' & t('TERM_ARCHIVE'), $dbx)
		Case StringInStr($filetype_curr, "bzip2 compressed data", 0)
			extract("bz2", 'bzip2 ' & t('TERM_COMPRESSED'))
		Case StringInStr($filetype_curr, "ASCII cpio archive", 0)
			extract("7z", 'CPIO ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "LZX compressed archive", 0)
			extract("lzx", 'LZX ' & t('TERM_COMPRESSED'))
		Case StringInStr($filetype_curr, "ARJ archive data", 0)
			extract("arj", 'ARJ ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "POSIX tar archive", 0)
			extract("tar", 'Tar ' & t('TERM_ARCHIVE'))
		Case StringInStr($filetype_curr, "Microsoft Reader eBook Data", 0)
			extract("lit", 'Microsoft LIT ' & t('TERM_EBOOK'))
		Case StringInStr($filetype_curr, "LHa", 0) And StringInStr($filetype_curr, "archive data", 0)
			extract("7z", 'LZH ' & t('TERM_COMPRESSED'))
		Case StringInStr($filetype_curr, "Macromedia Flash data", 0)
			extract("swf", 'Shockwave Flash ' & t('TERM_CONTAINER'))
		Case StringInStr($filetype_curr, "PowerISO Direct-Access-Archive", 0)
			extract("daa", 'DAA/GBI ' & t('TERM_IMAGE'))
		Case StringInStr($filetype_curr, "video", 0) Or StringInStr($filetype_curr, "MPEG v", 0) Or _
			 StringInStr($filetype_curr, "MPEG sequence")
			extract("video", t('TERM_VIDEO') & ' ' & t('TERM_FILE'))
		Case StringInStr($filetype_curr, "AAC,")
			extract("aac", 'AAC ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))
		Case StringInStr($filetype_curr, "FLAC audio")
			extract("flac", 'FLAC ' & t('TERM_AUDIO') & ' ' & t('TERM_FILE'))
		Case StringInStr($filetype_curr, "ISO", 0) And StringInStr($filetype_curr, "filesystem", 0)
			CheckIso()

			; Not extractable filetypes
		Case (StringInStr($filetype_curr, "text", 0) And (StringInStr($filetype_curr, "CRLF", 0) Or _
			  StringInStr($filetype_curr, "long lines", 0) Or StringInStr($filetype_curr, "ASCII", 0)) Or _
			  StringInStr($filetype_curr, "batch file")) Or StringInStr($filetype_curr, "source text", 0) Or _
			  StringInStr($filetype_curr, "image", 0) Or StringInStr($filetype_curr, "icon resource", 0) Or _
			 (StringInStr($filetype_curr, "bitmap", 0) And Not StringInStr($filetype_curr, "MGR bitmap")) Or _
			  StringInStr($filetype_curr, "Audio file", 0) Or StringInStr($filetype_curr, "WAVE audio", 0) Or _
			  StringInStr($filetype_curr, "shortcut", 0) Or StringInStr($filetype_curr, "ogg", 0)
			terminate("notpacked", $file, "")
	EndSelect
EndFunc

; Scan .exe file using PEiD
Func exescan($f, $scantype, $analyze = 1)
	Local $filetype_curr = ""

	Cout("Start filescan using PEiD")

	_CreateTrayMessageBox(t('SCANNING_EXE', "PEiD"))

	; Backup existing PEiD options
	Local $exsig = RegRead("HKCU\Software\PEiD", "ExSig")
	Local $loadplugins = RegRead("HKCU\Software\PEiD", "LoadPlugins")
	Local $stayontop = RegRead("HKCU\Software\PEiD", "StayOnTop")

	; Set PEiD options
	RegWrite("HKCU\Software\PEiD", "ExSig", "REG_DWORD", 1)
	RegWrite("HKCU\Software\PEiD", "LoadPlugins", "REG_DWORD", 0)
	RegWrite("HKCU\Software\PEiD", "StayOnTop", "REG_DWORD", 0)

	; Analyze file
	Run($peid & ' -' & $scantype & ' "' & $f & '"', @ScriptDir, @SW_HIDE)
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
	If $exsig Then RegWrite("HKCU\Software\PEiD", "ExSig", "REG_DWORD", $exsig)
	If $loadplugins Then RegWrite("HKCU\Software\PEiD", "LoadPlugins", "REG_DWORD", $loadplugins)
	If $stayontop Then RegWrite("HKCU\Software\PEiD", "StayOnTop", "REG_DWORD", $stayontop)
	_DeleteTrayMessageBox()



	; Return filetype without matching if specified
	If Not $analyze Then Return $filetype_curr

	; Match known patterns
	Select
		; Check if packed first
		Case StringInStr($filetype_curr, "upx", 0) Or StringInStr($filetype_curr, "aspack", 0)
			unpack()
			;$packed = true

		Case StringInStr($filetype_curr, "ARJ SFX", 0)
			extract("arj", t('TERM_SFX') & ' ARJ ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Borland Delphi", 0) And Not StringInStr($filetype_curr, "RAR SFX", 0)
			$testinno = True
			$testzip = True

		Case StringInStr($filetype_curr, "Gentee Installer", 0)
			extract("ie", 'Gentee ' & t('TERM_INSTALLER'))

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

			; removed - not possible to access due to 7zip check after deep scan
			;case stringinstr($filetype_curr, "Microsoft Visual C++ 7.0", 0) AND stringinstr($filetype_curr, "Custom", 0) AND stringinstr($filetype_curr, "Hotfix", 0)
			;	extract("vssfxhotfix", 'Visual C++ ' & t('TERM_SFX') & ' ' & t('TERM_HOTFIX'))

		Case StringInStr($filetype_curr, "Microsoft Visual C++ 6.0", 0) And StringInStr($filetype_curr, "Custom", 0)
			extract("vssfxpath", 'Visual C++ ' & t('TERM_SFX') & '' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Netopsystems FEAD Optimizer", 0)
			extract("fead", 'Netopsystems FEAD ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Nullsoft PiMP SFX", 0)
			checkNSIS()

		Case StringInStr($filetype_curr, "PEtite", 0)
			$testarj = True
			$testace = True

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

		Case StringInStr($filetype_curr, "Unable to open file", 0)
			$isexe = False

	EndSelect
EndFunc   ;==>exescan

; Scan file using Exeinfo PE
Func advexescan($f, $analyze = 1)
	Cout("Start filescan using Exeinfo PE")

	_CreateTrayMessageBox(t('SCANNING_EXE', "Exeinfo PE"))

	Local $filetype_curr = ""

	; Backup existing Exeinfo PE options
	Local $exerr = RegRead("HKCU\Software\ExEi-pe", "ExeError")
	Local $scan = RegRead("HKCU\Software\ExEi-pe", "Scan")
	Local $stayontop = RegRead("HKCU\Software\ExEi-pe", "AllwaysOnTop")
	Local $skin = RegRead("HKCU\Software\ExEi-pe", "Skin")
	Local $shell = RegRead("HKCU\Software\ExEi-pe", "Shell_integr")
	Local $Log = RegRead("HKCU\Software\ExEi-pe", "Log")
	Local $biggui = RegRead("HKCU\Software\ExEi-pe", "Big_GUI")
	Local $lang = RegRead("HKCU\Software\ExEi-pe", "Lang")
	Local $close = RegRead("HKCU\Software\ExEi-pe", "closeExEi_whenExtRun")

	; Set Exeinfo PE options
	RegWrite("HKCU\Software\ExEi-pe", "ExeError", "REG_DWORD", 1)
	RegWrite("HKCU\Software\ExEi-pe", "Scan", "REG_DWORD", 1)
	RegWrite("HKCU\Software\ExEi-pe", "AllwaysOnTop", "REG_DWORD", 0)
	RegWrite("HKCU\Software\ExEi-pe", "Skin", "REG_DWORD", 0xFFFFFFFF)
	RegWrite("HKCU\Software\ExEi-pe", "Shell_integr", "REG_DWORD", 0)
	RegWrite("HKCU\Software\ExEi-pe", "Log", "REG_DWORD", 0xFFFFFFFF)
	RegWrite("HKCU\Software\ExEi-pe", "Big_GUI", "REG_DWORD", 0)
	RegWrite("HKCU\Software\ExEi-pe", "Lang", "REG_DWORD", 0xFFFFFFFF)
	RegWrite("HKCU\Software\ExEi-pe", "closeExEi_whenExtRun", "REG_DWORD", 0)

	; Analyze file
	Run($exeinfope & ' "' & $f & '"', @ScriptDir, @SW_MINIMIZE)
	WinWait("Exeinfo PE - ver")
	WinSetState("Exeinfo PE - ver", "", @SW_HIDE)
	$TimerStart = TimerInit()
	While ($filetype_curr = "") Or StringInStr($filetype_curr, "File too big") Or StringInStr($filetype_curr, "Antivirus may slow")
		Sleep(200)
		$filetype_curr = ControlGetText("Exeinfo PE - ver", "", "TEdit6")
		$TimerDiff = TimerDiff($TimerStart)
		If $TimerDiff > $Timeout Then ExitLoop
	WEnd

	; Get LamerInfo
	$filetype_curr &= @CRLF & @CRLF & ControlGetText("Exeinfo PE - ver", "", "TEdit5")

	WinClose("Exeinfo PE - ver")

	; Restore previous Exeinfo PE options
	If $exerr Then RegWrite("HKCU\Software\ExEi-pe", "ExeError", "REG_DWORD", $exerr)
	If $scan Then RegWrite("HKCU\Software\ExEi-pe", "Scan", "REG_DWORD", $scan)
	If $stayontop Then RegWrite("HKCU\Software\ExEi-pe", "AllwaysOnTop", "REG_DWORD", $scan)
	If $skin Then RegWrite("HKCU\Software\ExEi-pe", "Skin", "REG_DWORD", $skin)
	If $shell Then RegWrite("HKCU\Software\ExEi-pe", "Shell_integr", "REG_DWORD", $shell)
	If $Log Then RegWrite("HKCU\Software\ExEi-pe", "Log", "REG_DWORD", $Log)
	If $biggui Then RegWrite("HKCU\Software\ExEi-pe", "Big_GUI", "REG_DWORD", $biggui)
	If $lang Then RegWrite("HKCU\Software\ExEi-pe", "Lang", "REG_DWORD", $lang)
	If $close Then RegWrite("HKCU\Software\ExEi-pe", "closeExEi_whenExtRun", "REG_DWORD", $close)
	_DeleteTrayMessageBox()

	; Return if not .exe file
	If StringInStr($filetype_curr, "NOT EXE") Then Return

	; Return if file is too big
	If StringInStr($filetype_curr, "Skipped") Then Return

	If $filetype_curr <> "" Then $filetype &= $filetype_curr & @CRLF & @CRLF

	Cout($filetype_curr)

	; Otherwise continue exe file procedure
	$isexe = True

	; Return filetype without matching if specified
	If Not $analyze Then Return $filetype_curr

	; Match known patterns
	Select
		Case StringInStr($filetype_curr, "upx", 0) Or StringInStr($filetype_curr, "aspack", 0)
			$packed = True

		Case StringInStr($filetype_curr, "ARJ SFX", 0)
			extract("arj", t('TERM_SFX') & ' ARJ ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Gentee Installer", 0)
			extract("ie", 'Gentee ' & t('TERM_INSTALLER'))

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

			; removed - not possible to access due to 7zip check after deep scan
			;case stringinstr($filetype_curr, "Microsoft Visual C++ 7.0", 0) AND stringinstr($filetype_curr, "Custom", 0) AND stringinstr($filetype_curr, "Hotfix", 0)
			;	extract("vssfxhotfix", 'Visual C++ ' & t('TERM_SFX') & ' ' & t('TERM_HOTFIX'))

		Case StringInStr($filetype_curr, "Microsoft Visual C++ 6.0", 0) And StringInStr($filetype_curr, "Custom", 0)
			extract("vssfxpath", 'Visual C++ ' & t('TERM_SFX') & '' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Netopsystems FEAD Optimizer", 0)
			extract("fead", 'Netopsystems FEAD ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "Nullsoft", 0);
			checkNSIS()

		Case StringInStr($filetype_curr, "PEtite", 0)
			$testarj = True
			$testace = True

		Case StringInStr($filetype_curr, "RAR SFX", 0)
			extract("rar", t('TERM_SFX') & ' RAR ' & t('TERM_ARCHIVE'));

		Case StringInStr($filetype_curr, "Reflexive Arcade Installer", 0)
			extract("inno", 'Reflexive Arcade ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "RoboForm Installer", 0)
			extract("robo", 'RoboForm ' & t('TERM_INSTALLER'))

		Case StringInStr($filetype_curr, "Setup Factory 6.x", 0)
			extract("ie", 'Setup Factory ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Microsoft SFX CAB", 0) And StringInStr($filetype_curr, "rename file *.exe as *.cab", 0)
			Prompt(1, "FILE_COPY", $file, 1)
			Local $return = _TempFile($filedir, $filename & "_", ".cab")
			FileCopy($file, $return)
			$file = $return
			Global $DeleteOrigFile = 1
			check7z()

		Case StringInStr($filetype_curr, "SPx Method", 0) Or StringInStr($filetype_curr, "Microsoft SFX CAB", 0)
			extract("cab", t('TERM_SFX') & ' Microsoft CAB ' & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "SuperDAT", 0)
			extract("superdat", 'McAfee SuperDAT ' & t('TERM_UPDATER'))

		Case StringInStr($filetype_curr, "VMware ThinApp", 0) Or StringInStr($filetype_curr, "Thinstall", 0) Or StringInStr($filetype_curr, "ThinyApp Packager", 0)
			extract("thinstall", "ThinApp/Thinstall" & t('TERM_ARCHIVE'))

		Case StringInStr($filetype_curr, "Wise", 0) Or StringInStr($filetype_curr, "PEncrypt 4.0", 0)
			extract("wise", 'Wise Installer ' & t('TERM_PACKAGE'))

		Case StringInStr($filetype_curr, "ZIP SFX", 0)
			$testzip = True

		Case StringInStr($filetype_curr, "Borland Delphi", 0) And Not StringInStr($filetype_curr, "RAR SFX", 0)
			$testinno = True
			$testzip = True

		Case StringInStr($filetype_curr, ".dmg  Mac OS", 0)
			extract("dmg", 'DMG ' & t('TERM_IMAGE'))

		Case StringInStr($filetype_curr, "CAB archive", 0)
			$isexe = False

			; not supported filetypes
		Case StringInStr($filetype_curr, "Autoit", 0)
			terminate("notpacked", $file, "")

		Case StringInStr($filetype_curr, "Astrum InstallWizard", 0) Or StringInStr($filetype_curr, "clickteam", 0) Or _
			 StringInStr($filetype_curr, "CreateInstall", 0)
			terminate("notsupported", $file, "")

			; Terminate if file cannot be unpacked
		Case StringInStr($filetype_curr, "Not packed") And Not StringInStr($filetype_curr, "Microsoft Visual C++")
			$notpacked = True
	EndSelect
EndFunc   ;==>advexescan

; Scan file using MediaInfo dll, only used in scan only mode
Func MediaFileScan($f)
	Local $filetype_curr = ""
	Cout("Start filescan using MediaInfo dll")
	_CreateTrayMessageBox(t('SCANNING_FILE', "MediaInfo"))

	HasPlugin($mediainfo)
	$hDll = DllOpen("MediaInfo.dll")
	$hMI = DllCall($hDll, "ptr", "MediaInfo_New")

	$Open_Result = DllCall($hDll, "int", "MediaInfo_Open", "ptr", $hMI[0], "wstr", $f)
	$return = DllCall($hDll, "wstr", "MediaInfo_Inform", "ptr", $hMI[0], "int", 0)

	$hMI = DllCall($hDll, "none", "MediaInfo_Delete", "ptr", $hMI[0])
	DllClose($hDll)

	Cout($return[0])

	; Return if file is not a media file
	$return = StringSplit($return[0], @CRLF, 2)
	If UBound($return) < 10 Then
		_DeleteTrayMessageBox()
		Return
	EndIf

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
	Cout("Testing 7zip")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' 7-Zip ' & t('TERM_INSTALLER'))
	$return = FetchStdout($cmd & $7z & ' l "' & $file & '"', $filedir, @SW_HIDE)

	If StringInStr($return, "Listing archive:", 0) Then
		; failsafe in case TrID misidentifies MS SFX es
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
			extract("xz", 'XZ ' & t('TERM_COMPRESSED'))
		ElseIf $fileext = "z" Then
			extract("Z", 'LZW ' & t('TERM_COMPRESSED'))
		Else
			extract("7z", '7-Zip ' & t('TERM_ARCHIVE'))
		EndIf
	EndIf

	_DeleteTrayMessageBox()
	$7zfailed = True
	Return False
EndFunc   ;==>check7z

; Determine if file is self-extracting ACE archive
Func checkAce()
	; Ace testing handled by extract function
	extract("ace", t('TERM_SFX') & ' ACE ' & t('TERM_ARCHIVE'))
	$acefailed = True
	Return False
EndFunc   ;==>checkAce

; Determine if file is ALZip archive
Func CheckAlz()
	Cout("Testing ALZ")

	_CreateTrayMessageBox(t('TERM_TESTING') & ' ALZ ' & t('TERM_ARCHIVE'))
	$return = FetchStdout($cmd & $alz & ' -l "' & $file & '"', $filedir, @SW_HIDE)

	If StringInStr($return, "Listing archive:") And Not (StringInStr($return, "corrupted file") Or StringInStr($return, "file open error")) Then
		_DeleteTrayMessageBox()
		extract("alz", 'ALZ ' & t('TERM_ARCHIVE'))
	EndIf

	Return False
EndFunc   ;==>CheckAlz

; Determine if file is self-extracting ARJ archive
Func checkArj()
	Cout("Testing ARJ")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' SFX ARJ ' & t('TERM_ARCHIVE'))
	$return = FetchStdout($cmd & $arj & ' l "' & $file & '"', $filedir, @SW_HIDE)

	If StringInStr($return, "Archive created:", 0) Then
		_DeleteTrayMessageBox()
		extract("arj", t('TERM_SFX') & ' ARJ ' & t('TERM_ARCHIVE'))
	EndIf

	_DeleteTrayMessageBox()
	$arjfailed = True
	Return False
EndFunc   ;==>checkArj

; Determine if folder contains .bin files
Func checkBin()
	Cout("Searching additional .bin files to be extracted")

	$NSISbin = FileFindFirstFile($filedir & "\data*.bin")
	If $NSISbin == -1 Then Return
	If Prompt(64 + 1, "NSIS_BINFILES", CreateArray($file, $filename & "." & $fileext), 0) Then
		While 1
			$file = $filedir & "\" & FileFindNextFile($NSISbin)
			If @error Then ExitLoop
			;FilenameParse($file)
			extract("7z", ".bin " & t('TERM_ARCHIVE'), "", True, True)
		WEnd
	EndIf
	FileClose($NSISbin)
	terminate("success", "", "NSIS")
EndFunc   ;==>checkBin

; Determine if file is supported game archive
Func CheckGame()
	If Not $CheckGame Then Return

	Cout("Testing Game archive")

	_CreateTrayMessageBox(t('TERM_TESTING') & ' ' & t('TERM_GAME') & t('TERM_PACKAGE'))

	; Check GAUP first
	$return = FetchStdout($cmd & $quickbms & ' -l "' & @ScriptDir & '\bin\' & $gaup & '" "' & $file & '"', $filedir, @SW_HIDE, -1)

	If StringInStr($return, "Target directory:", 0) Or StringInStr($return, "0 files found", 0) Or StringInStr($return, "Error", 0) _
			Or StringInStr($return, "exception occured", 0) Or StringInStr($return, "not supported", 0) Or $return == "" Then

	Else
		_DeleteTrayMessageBox()
		extract("qbms", t('TERM_GAME') & t('TERM_PACKAGE'), $gaup, False, True)
	EndIf

	If $silentmode Then
		Cout("INFO: File may be extractable via BMS script, but user input is needed. Disable silent mode to try this method.")
		_DeleteTrayMessageBox()
		Return False
	EndIf

	; Check if game specific bms script is available
	_SQLite_Startup()
	If @error Then
		Cout("[ERROR] SQLite startup failed with code " & @error)
		Return
	EndIf
	_SQLite_Open(@ScriptDir & "\bin\BMS.db", $SQLITE_OPEN_READONLY)
	If @error Then
		Cout("[ERROR] Failed to open BMS database")
		Return
	EndIf

	Local $gameformat, $iRows, $iColumns, $game

	_SQLite_GetTable(-1, "SELECT n.Name FROM Names n, Scripts s, Extensions e WHERE s.SID = e.EID AND s.SID = n.NID AND e.Extension= '" _
			 & StringLower($fileext) & "' ORDER BY n.Name", $gameformat, $iRows, $iColumns)
 	_ArrayDelete($gameformat, 1)
;~ 	_ArrayDisplay($gameformat)

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
			$bmsScript = FileOpen(@ScriptDir & "\bin\" & $bms, 2)
			FileWrite($bmsScript, $gameformat[2])
			FileClose($bmsScript)
			$return = FetchStdout($cmd & $quickbms & ' -l "' & @ScriptDir & '\bin\' & $bms & '" "' & $file & '"', $filedir, @SW_HIDE, -1)
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
EndFunc   ;==>CheckGame

; Determine if InstallExplorer can extract the file
Func checkIE()
	Cout("Testing InstallExplorer")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' InstallExplorer ' & t('TERM_INSTALLER'))
	$return = FetchStdout($cmd & $quickbms & ' -l "' & @ScriptDir & '\bin\' & $ie & '" "' & $file & '"', $filedir, @SW_HIDE)
	_DeleteTrayMessageBox()
	If StringInStr($return, "Target directory:", 0) Or StringInStr($return, "0 files found", 0) Or StringInStr($return, "Error", 0) _
			Or StringInStr($return, "exception occured", 0) Or StringInStr($return, "not supported", 0) Or StringInStr($return, "crash occurred", 0) _
			Or $return == "" Then
		$iefailed = True
		Return False
	Else
		extract("qbms", 'InstallExplorer ' & t('TERM_INSTALLER'), $ie)
	EndIf
EndFunc   ;==>checkIE

; Determine if file is Inno Setup installer
Func checkInno()
	Cout("Testing Inno Setup")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' Inno Setup ' & t('TERM_INSTALLER'))

	$return = FetchStdout($cmd & $inno & ' "' & $file & '"', $filedir, @SW_HIDE)
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
	Cout("Testing image file")
	_CreateTrayMessageBox(t('TERM_TESTING') & " " & t('TERM_IMAGE') & " " & t('TERM_FILE'))
	If $fileext = "cue" And FileExists(StringTrimRight($file, 3) & "bin") Then $file = StringTrimRight($file, 3) & "bin"
	$return = FetchStdout($cmd & $quickbms & ' -l "' & @ScriptDir & '\bin\' & $iso & '" "' & $file & '"', $filedir, @SW_HIDE, -1)
	_DeleteTrayMessageBox()
	If StringInStr($return, "Target directory:", 0) Or StringInStr($return, "0 files found", 0) Or StringInStr($return, "Error", 0) _
			Or StringInStr($return, "exception occured", 0) Or StringInStr($return, "not supported", 0) Or $return == "" Then
		$isofailed = True
		Return False
	Else
		extract("qbms", t('TERM_IMAGE') & " " & t('TERM_FILE'), $iso)
	EndIf
EndFunc   ;==>CheckIso

; Determine if file is NSIS installer
Func checkNSIS()
	Cout("Testing NSIS")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' NSIS ' & t('TERM_INSTALLER'))

	$return = FetchStdout($cmd & $7z & ' l "' & $file & '"', $filedir, @SW_HIDE)
	If StringInStr($return, "Listing archive:", 0) Then
		_DeleteTrayMessageBox()
		extract("NSIS", 'NSIS ' & t('TERM_INSTALLER'))
	EndIf

	_DeleteTrayMessageBox()
	checkIE()
	Return False
EndFunc   ;==>checkNSIS

; Determine if file is self-extracting Zip archive
Func checkZip()
	Cout("Testing Zip")
	_CreateTrayMessageBox(t('TERM_TESTING') & ' SFX ZIP ' & t('TERM_ARCHIVE'))
	$return = FetchStdout($cmd & $zip & ' -l "' & $file & '"', $filedir, @SW_HIDE)
	If Not StringInStr($return, "signature not found", 0) Then
		_DeleteTrayMessageBox()
		extract("zip", t('TERM_SFX') & ' ZIP ' & t('TERM_ARCHIVE'))
	EndIf

	_DeleteTrayMessageBox()
	$zipfailed = True
	Return False
EndFunc   ;==>checkZip

; Extract from known archive format
Func extract($arctype, $arcdisp, $additionalParameters = "", $returnSuccess = False, $returnFail = False)
	$success = False

	Cout("Starting " & $arctype & " extraction")

	; Display banner and create subdirectory
	_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & $arcdisp)
	If StringRight($outdir, 1) = '\' Then $outdir = StringTrimRight($outdir, 1)
	If Not FileExists($outdir) Then
		If Not DirCreate($outdir) Then terminate("invaliddir", $outdir, "")
		$createdir = True
	Else
		$dirmtime = FileGetTime($outdir, 0, 1)
	EndIf

	HasFreeSpace()

	$initdirsize = DirGetSize($outdir)
	$tempoutdir = _TempFile($outdir, 'uni_', '')

	; Extract archive based on filetype
	Switch $arctype
		Case "7z"
			_Run($cmd & $7z & ' x "' & $file & '"', $outdir)

			; Extract inner CPIO for RPMs
			If StringInStr($filetype, 'RPM Linux Package', 0) Then
				If FileExists($outdir & '\' & $filename & '.cpio') Then
					RunWait($cmd & $7z & ' x "' & $outdir & '\' & $filename & '.cpio"', $outdir)
					FileDelete($outdir & '\' & $filename & '.cpio')
				EndIf

				; Extract inner tarball for DEBs
			ElseIf StringInStr($filetype, 'Debian Linux Package', 0) Then
				If FileExists($outdir & '\data.tar') Then
					RunWait($cmd & $7z & ' x "' & $outdir & '\data.tar"', $outdir)
					FileDelete($outdir & '\data.tar')
				EndIf
			EndIf

		Case "aac"
			HasPlugin($faad)
			_Run($cmd & $faad & ' -o "' & $outdir & '\' & $filename & '.wav" "' & $file & '"', $outdir, @SW_HIDE)

		Case "ace"
			Opt("WinTitleMatchMode", 3)
			$pid = Run($ace & ' -x "' & $file & '" "' & $outdir & '"', $filedir)
			Sleep(100)
			While ProcessExists($pid)
				If WinExists("Error") Then
					ProcessClose($pid)
					ExitLoop
				EndIf
				Sleep(50)
			WEnd

		Case "alz"
			_Run($cmd & $alz & ' -d "' & $outdir & '" "' & $file & '"', $outdir)
			If @error Then terminate("failed", $file, $arcdisp)

		Case "arc"
			_Run($cmd & $arc & ' x "' & $file & '"', $outdir, @SW_HIDE, True, False, False)

		Case "arc_conv"
			If Not HasPlugin($arc_conv, $returnFail) Then Return

			Run($cmd & $arc_conv & ' "' & $file & '"', $outdir, @SW_HIDE)
			Local $hWnd = WinWait("arc_conv", "", $Timeout)
			If $hWnd == 0 Then terminate("timeout", $file, $arcdisp)
			Local $current = "", $last = ""
			; Hide not possible as window text has to be read
			WinSetState("arc_conv", "", @SW_MINIMIZE)
			While WinExists("arc_conv")
				$current = WinGetText("arc_conv")
				If $current <> $last Then
					If StringInStr($current, "/") Then
						GUICtrlSetData($TrayMsg_Status, $current)
					Else
						GUICtrlSetData($TrayMsg_Status, t('TERM_FILE') & " #" & $current)
					EndIf
					$last = $current
					Sleep(10)
				EndIf
			WEnd
			MoveFiles($file & "~", $outdir, True, "", True)

		Case "arj"
			_Run($cmd & $arj & ' x "' & $file & '"', $outdir)

		Case "bz2"
			_Run($cmd & $7z & ' x "' & $file & '"', $outdir)
			If FileExists($outdir & '\' & $filename) Then
				_Run($cmd & $7z & ' x "' & $outdir & '\' & $filename & '"', $outdir)
				FileDelete($outdir & '\' & $filename)
			EndIf

		Case "cab"
			If StringInStr($filetype, 'Type 1', 0) Then
				If $warnexecute Then Warn_Execute($filename & '.exe /q /x:"<outdir>"')
				RunWait('"' & $file & '" /q /x:"' & $outdir & '"', $outdir)
			Else
				check7z()
			EndIf

		Case "chm"
			RunWait($cmd & $7z & ' x "' & $file & '"', $outdir)
			FileDelete($outdir & '\#*')
			FileDelete($outdir & '\$*')
			$dirs = FileFindFirstFile($outdir & '\*')
			If $dirs <> -1 Then
				$dir = FileFindNextFile($dirs)
				Do
					If StringLeft($dir, 1) == '#' Or StringLeft($dir, 1) == '$' Then DirRemove($outdir & '\' & $dir, 1)
					$dir = FileFindNextFile($dirs)
				Until @error
			EndIf
			FileClose($dirs)

		Case "crage"
			HasPlugin($crage)
			_Run($cmd & 'crage.exe -p "' & $file & '" -o "' & $outdir & '" -v', @ScriptDir & "\bin\crass-0.4.14.0", @SW_SHOW, False)

		Case "ctar"
			; Get existing files in $outdir
			$oldfiles = ReturnFiles($outdir)

			; Decompress archive with 7-zip
			_Run($cmd & $7z & ' x "' & $file & '"', $outdir)

			; Check for new files
			$handle = FileFindFirstFile($outdir & "\*")
			If Not @error Then
				While 1
					$fname = FileFindNextFile($handle)
					If @error Then ExitLoop
					If Not StringInStr($oldfiles, $fname) Then

						; Check for supported archive format
						$return = FetchStdout($cmd & $7z & ' l "' & $outdir & '\' & $fname & '"', $outdir, @SW_HIDE)
						If StringInStr($return, "Listing archive:", 0) Then
							_Run($cmd & $7z & ' x "' & $outdir & '\' & $fname & '"', $outdir, @SW_HIDE)
							FileDelete($outdir & '\' & $fname)
						EndIf
					EndIf
				WEnd
			EndIf
			FileClose($handle)

		Case "dmg"
			_DeleteTrayMessageBox()

			IsJavaInstalled()

			Prompt(32 + 4, 'CONVERT_CDROM', 'DMG', 1)

			; Begin conversion to .iso format
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'DMG ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 1)')
			$isofile = $filedir & '\' & $filename & '.iso'
			Cout('Executing: ' & $cmd & 'start javaw -jar "' & @ScriptDir & "\bin\" & $dmg & '" "' & $file & '" "' & $isofile & '"')
			$pid = Run($cmd & 'start javaw -jar "' & @ScriptDir & "\bin\" & $dmg & '" "' & $file & '" "' & $isofile & '"', $filedir, @SW_HIDE, False)

			$handle = WinWait("DMGExtractor 0.70", "", $Timeout)
			If $handle = 0 Then terminate("timeout", $file, $arcdisp)
			WinActivate("DMGExtractor 0.70")
			Send("{ENTER}")
			Opt("WinTitleMatchMode", 3)

			Do
				If WinExists("DMGExtractor 0.70: Error") Then
					Cout("Error detected")
					WinActivate("DMGExtractor 0.70: Error")
					Send("{ENTER}")
					If FileExists($isofile) Then FileDelete($isofile)
					_DeleteTrayMessageBox()
					Return
				EndIf
				Sleep(500)
			Until WinExists("DMGExtractor 0.70") ;OR NOT ProcessExists($pid)

			If WinExists("DMGExtractor 0.70") Then
				WinActivate("DMGExtractor 0.70")
				Send("{ENTER}")
			Else
				If FileExists($isofile) Then FileDelete($isofile)
				Return
			EndIf

			_DeleteTrayMessageBox()
			; Begin extraction from .iso
			If FileExists($isofile) And FileGetSize($isofile) > 0 Then
				_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'DMG ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 2)')
				$file = $isofile
				If Not CheckIso() Then _Run($cmd & $7z & ' x "' & $isofile & '"', $outdir)
			Else ; Exit if conversion failed
				Prompt(16, 'CONVERT_CDROM_STAGE1_FAILED', "", 0)
				If FileExists($isofile) Then FileDelete($isofile)
				check7z()
				terminate("failed", $file, $arcdisp)
			EndIf

		Case "daa"
			; Prompt user to continue
			_DeleteTrayMessageBox()
			Prompt(32 + 4, 'CONVERT_CDROM', 'DAA/GBI', 1)

			; Begin conversion to .iso format
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'DAA/GBI ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 1)')
			$isofile = $filedir & '\' & $filename & '.iso'
			_Run($cmd & $daa & ' "' & $file & '" "' & $isofile & '"', $filedir)

			; Begin extraction from .iso
			If FileExists($isofile) And FileGetSize($isofile) > 0 Then
				_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'DAA/GBI ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 2)')
				$file = $isofile
				If Not CheckIso() Then _Run($cmd & $7z & ' x "' & $isofile & '"', $outdir)

				; Exit if conversion failed
			Else
				Prompt(16, 'CONVERT_CDROM_STAGE1_FAILED', "", 0)
				If FileExists($isofile) Then FileDelete($isofile)
				check7z()
				terminate("failed", $file, $arcdisp)
			EndIf

		Case "dcp"
			HasPlugin($dcp)
			_Run($cmd & $dcp & ' "' & $file & '"', $outdir)

		Case "ethornell"
			_Run($cmd & $ethornell & ' "' & $file & '" "' & $outdir & '"', $outdir)

		Case "fead"
			If $warnexecute Then Warn_Execute($filename & '.exe /s -nos_ne -nos_o"<outdir>"')
			RunWait($file & ' /s -nos_ne -nos_o"' & $tempoutdir & '"', $filedir)
			FileSetAttrib($tempoutdir & '\*', '-R', 1)
			MoveFiles($tempoutdir, $outdir, False, "", True)
			DirRemove($tempoutdir)

		Case "flac"
			_Run($cmd & $flac & ' -o "' & $outdir & '\' & $filename & '.wav" -d "' & $file & '"', $outdir, @SW_HIDE, True, False)

		Case "flv"
			_Run($cmd & $flv & ' -v -a -t -d "' & $outdir & '" "' & $file & '"', $filedir)

		Case "freearc"
			_Run($cmd & $freearc & ' x -dp"' & $outdir & '" "' & $file & '"', $filedir, @SW_HIDE, True, False, False)

		Case "fsb"
			_Run($cmd & $fsb & ' -d "' & $outdir & '" "' & $file & '"', $filedir)

		Case "gcf"
			Prompt(48 + 1, 'PACKAGE_EXPLORER', $file, 1)
			Run($gcf & ' "' & $file & '"')
			terminate("silent", "", "")

		Case "gz"
			_Run($cmd & $7z & ' x "' & $file & '"', $outdir)
			If FileExists($outdir & '\' & $filename) And StringTrimLeft($filename, StringInStr($filename, '.', 0, -1)) = "tar" Then
				_Run($cmd & $7z & ' x "' & $outdir & '\' & $filename & '"', $outdir)
				FileDelete($outdir & '\' & $filename)
			EndIf

		Case "hlp"
			RunWait($cmd & $hlp & ' "' & $file & '"', $outdir)
			If DirGetSize($outdir) > $initdirsize Then
				DirCreate($tempoutdir)
				_Run($cmd & $hlp & ' /r /n "' & $file & '"', $tempoutdir)
				FileMove($tempoutdir & '\' & $filename & '.rtf', $outdir & '\' & $filename & '_Reconstructed.rtf')
				DirRemove($tempoutdir, 1)
			EndIf

			; failsafe in case TrID misidentifies MS SFX hotfixes
		Case "hotfix"
			If $warnexecute Then Warn_Execute($filename & '.exe /q /x:"<outdir>"')
			RunWait('"' & $file & '" /q /x:"' & $outdir & '"', $outdir)

		Case "img"
			_Run($cmd & $img & ' -x "' & $file & '"', $outdir)

		Case "inno"
			If StringInStr($filetype, "Reflexive Arcade", 0) Then
				DirCreate($tempoutdir)
				_Run($cmd & $rai & ' "' & $file & '" "' & $tempoutdir & '\' & $filename & '.exe"', $filedir)
				_Run($cmd & $inno & ' -x -m -a "' & $tempoutdir & '\' & $filename & '.exe"', $outdir)
				FileDelete($tempoutdir & '\' & $filename & '.exe')
				DirRemove($tempoutdir)
			Else
				_Run($cmd & $inno & ' -x -m -a "' & $file & '"', $outdir)
			EndIf

		Case "is3arc"
			$choice = MethodSelect($arctype, $arcdisp)

			; Extract using i3comp
			; Removed due to license problems with .dll files
			;if $choice == 'i3comp' then
			;	runwait($cmd & $is3arc & ' "' & $file & '" *.* -d -i' & $output, $outdir)

			; Extract using unshield
			If $choice == 'unshield' Then
				_Run($cmd & $unshield & ' -d "' & $outdir & '" x "' & $file & '"', $outdir)

				; Extract using STIX
			ElseIf $choice == 'STIX' Then
				_Run($cmd & $stix & ' ' & FileGetShortName($file) & ' ' & FileGetShortName($outdir), $filedir)
			EndIf

		Case "iscab"

			$choice = "is6comp"
			If FileExists(@ScriptDir & "\bin\" & $iscab) Then $choice = MethodSelect($arctype, $arcdisp)

			If $choice == "is6comp" Then
				; List contents of archive
				$return = FetchStdout($cmd & $is6cab & ' l "' & $file & '"', $filedir, @SW_HIDE)
				$return = _StringBetween(StringRight($return, 22), " ", " file(s) total")
				If Not @error Then $return = Number(StringStripWS($return[0], 8))
				;MsgBox(1,"",$return)

				; If successful, extract contents of InstallShield cabs file-by-file
				If $return > 0 Then
					RunWait($cmd & $is6cab & ' x "' & $file & '"', $outdir, @SW_MINIMIZE)
				Else
					; Otherwise, attempt to extract with unshield
					_Run($cmd & $unshield & ' -d "' & $outdir & '" x "' & $file & '"', $outdir)
				EndIf

			ElseIf $choice == "iscab" Then
				RunWait($cmd & $iscab & ' "' & $file & '" -i"files.ini" -lx', $outdir, @SW_HIDE)
				;MsgBox(1,"",$cmd & $iscab & ' "' & $file & '" -i"files.ini" -x')
				RunWait($cmd & $iscab & ' "' & $file & '" -i"files.ini" -x', $outdir, @SW_MINIMIZE)
				FileDelete($outdir & "\files.ini")
			EndIf

		Case "isexe"
			exescan($file, 'ext', 0)
			If StringInStr($filetype, "3.x", 0) Then
				; Extract 3.x SFX installer using stix
				_Run($cmd & $stix & ' ' & FileGetShortName($file) & ' ' & FileGetShortName($outdir), $filedir)

			Else
				$choice = MethodSelect($arctype, $arcdisp)

				; User-specified false positive; return for additional analysis
				If $choice == 'not InstallShield' Then
					$isfailed = True
					Return False

					; Extract using isxunpack
				ElseIf $choice == 'isxunpack' Then
					FileMove($file, $outdir)
					Run($cmd & $isxunp & ' "' & $outdir & '\' & $filename & '.' & $fileext & '"', $outdir)
					WinWait(@ComSpec)
					WinActivate(@ComSpec)
					Send("{ENTER}")
					ProcessWaitClose($isxunp)
					FileMove($outdir & '\' & $filename & '.' & $fileext, $filedir)

					; Try to extract using unshield
				ElseIf $choice == 'unshield' Then
					_Run($cmd & $unshield & ' -d "' & $outdir & '" x "' & $file & '"', $outdir)

					; Try to extract MSI using cache switch
				ElseIf $choice == 'InstallShield /b' Then
					; Run installer and wait for temp files to be copied
					If $warnexecute Then Warn_Execute('"' & $file & '" /b"' & $tempoutdir & '"')
					_CreateTrayMessageBox(t('INIT_WAIT'))

					If $Log Then
						_Run('"' & $file & '" /b"' & $tempoutdir & '" /v"/l "' & @ScriptDir & '\log\teelog.txt""', $filedir, @SW_SHOW, False)
					Else
						RunWait('"' & $file & '" /b"' & $tempoutdir & '"', $filedir)
					EndIf

					; Wait for matching windows for up to 30 seconds (60 * .5)
					Opt("WinTitleMatchMode", 4)
					Local $success
					For $i = 1 To 60
						If Not WinExists("classname=MsiDialogCloseClass") Then
							Sleep(500)

						Else
							; Search temp directory for MSI support and copy to tempoutdir
							$msihandle = FileFindFirstFile($tempoutdir & "\*.msi")
							If Not @error Then
								While 1
									$msiname = FileFindNextFile($msihandle)
									If @error Then ExitLoop
									$tsearch = FileSearch(EnvGet("temp") & "\" & $msiname)
									If Not @error Then
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
									EndIf
								WEnd
								FileClose($msihandle)
							EndIf

							; Move files to outdir
							_DeleteTrayMessageBox()
							Prompt(64, 'INIT_COMPLETE', 0)
							MoveFiles($tempoutdir, $outdir, False, "", True)
							$success = True
							ExitLoop
						EndIf
					Next

					; Not a supported installer
					If Not $success Then
						_DeleteTrayMessageBox()
						Prompt(16, 'INIT_COMPLETE', 0)
					EndIf
				EndIf
			EndIf

		Case "kgb"
			_Run($cmd & $kgb & ' "' & $file & '"', $outdir, @SW_SHOW, False)
			#cs
				$show_stats = regread("HKCU\Software\KGB Archiver", "show_stats")
				regwrite("HKCU\Software\KGB Archiver", "show_stats", "REG_DWORD", 0)
				runwait($kgb & ' /s "' & $file & '" "' & $outdir & '"', $outdir)
				if $show_stats == "" then
				regdelete("HKCU\Software\KGB Archiver")
				else
				regwrite("HKCU\Software\KGB Archiver", "show_stats", "REG_DWORD", $show_stats)
				endif
			#ce

		Case "lit"
			_Run($cmd & $lit & ' "' & $file & '" "' & $outdir & '"', $outdir)

		Case "lzo"
			_Run($cmd & $lzo & ' -d -p"' & $outdir & '" "' & $file & '"', $filedir)

		Case "lzx"
			_Run($cmd & $lzx & ' -x "' & $file & '"', $outdir)

		Case "mht"
			$choice = MethodSelect($arctype, $arcdisp)

			; Extract using ExtractMHT
			If $choice == 'ExtractMHT' Then
				_Run($mht & ' "' & $file & '" "' & $outdir & '"', $outdir, @SW_MINIMIZE, False)
			ElseIf $choice == 'MhtUnPack' Then
				extract('qbms', $arcdisp, $mht_plug)
			EndIf

		Case "msi"
			$choice = MethodSelect($arctype, $arcdisp)

			; Extract using administrative install
			If $choice == 'MSI' Then
				If $warnexecute Then Warn_Execute('msiexec.exe /a "' & $file & '" /qb TARGETDIR="' & $outdir & '"')
				RunWait('msiexec.exe /a "' & $file & '" /qb TARGETDIR="' & $outdir & '"', $filedir, @SW_SHOW)

				; Extract with MsiX
			ElseIf $choice == 'MsiX' Then
				Local $appendargs = ''
				If $appendext Then $appendargs = '/ext'
				_Run($cmd & $msi_msix & ' "' & $file & '" /out "' & $outdir & '" ' & $appendargs, $filedir)

				; Extract with jsMSI Unpacker
			ElseIf $choice == "jsMSI Unpacker" Then
				_Run($msi_jsmsix & ' "' & $file & '"|"' & $outdir & '"', $filedir, @SW_SHOW, False)
				If $Log Then
					Cout("jsMSI log:")
					$handle = FileOpen($outdir & "\MSI Unpack.log")
					$debug &= FileRead($handle)
					FileClose($handle)
				EndIf
				FileDelete($outdir & "\MSI Unpack.log")

				; Extract with MSI Total Commander plugin
			ElseIf $choice == 'MSI TC Packer' Then
				;dircreate($tempoutdir)
				extract('qbms', $arcdisp, $msi_plug)

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
				If $appendext Then
					AppendExtensions($tempoutdir)
				EndIf

				; Move files to output directory and remove tempdir
				MoveFiles($tempoutdir, $outdir, False, "", True)
			EndIf

		Case "msm"
			Local $appendargs = ''
			If $appendext Then $appendargs = '/ext'
			_Run($cmd & $msi_msix & ' "' & $file & '" /out "' & $outdir & '" ' & $appendargs, $filedir)

		Case "msp"
			$choice = MethodSelect($arctype, $arcdisp)

			; Extract using TC MSI
			DirCreate($tempoutdir)
			If $choice == 'MSI TC Packer' Then
				extract('qbms', $arcdisp, $msi_plug)

				; Extract with MsiX
			ElseIf $choice == 'MsiX' Then
				Run($cmd & $msi_msix & ' "' & $file & '" /out "' & $tempoutdir & '"', $filedir)

				; Extract using 7-Zip
			ElseIf $choice == '7-Zip' Then
				_Run($cmd & $7z & ' x "' & $file & '"', $outdir)

			EndIf

			; Regardless of method, extract files from extracted CABs
			$cabfiles = FileSearch($tempoutdir)
			For $i = 1 To $cabfiles[0]
				filescan($cabfiles[$i], 0)
				If StringInStr($filetype, "Microsoft Cabinet Archive", 0) Then
					_Run($cmd & $7z & ' x "' & $cabfiles[$i] & '"', $outdir)
					FileDelete($cabfiles[$i])
				EndIf
			Next

			; Append missing file extensions
			If $appendext Then
				AppendExtensions($tempoutdir)
			EndIf

			; Move files to output directory and remove tempdir
			MoveFiles($tempoutdir, $outdir, False, "", True)

		Case "nbh"
			RunWait($cmd & $nbh & ' "' & $file & '"', $outdir)

		Case "NSIS"
			; Rename duplicates and extract
			_Run($cmd & $7z & ' x -aou' & ' "' & $file & '"', $outdir)

			; Determine if there are .bin files in filedir
			checkBin()

		Case "pea"
			Local $pid = Run($pea & ' UNPEA "' & $file & '" "' & $tempoutdir & '" RESETDATE SETATTR EXTRACT2DIR INTERACTIVE', $filedir)
			While ProcessExists($pid)
				$return = ControlGetText(_WinGetByPID($pid), '', 'Button1')
				If StringLeft($return, 4) = 'Done' Then ProcessClose($pid)
				Sleep(10)
			WEnd
			MoveFiles($tempoutdir, $outdir, False, "", True)

		Case "qbms"
			_Run($cmd & $quickbms & ' "' & @ScriptDir & '\bin\' & $additionalParameters & '" "' & $file & '" "' & $outdir & '"', $outdir, @SW_MINIMIZE, False)
			If FileExists(@ScriptDir & "\bin\" & $bms) Then FileDelete(@ScriptDir & "\bin\" & $bms)

		Case "rar"
			_Run($cmd & $rar & ' x "' & $file & '"', $outdir, @SW_SHOW)

		Case "rgss3"
			HasPlugin($rgss3)
			Run($rgss3, $outdir, @SW_HIDE)
			Local $handle = WinWait("RPGMaker Decrypter", "", $Timeout)
			If $handle = 0 Then terminate("timeout", $file, $arcdisp)
			ControlSetText($handle, "", "Edit1", $file)
			ControlSetText($handle, "", "Edit2", $outdir)
			ControlClick($handle, "", "Button4")
			Local $handle2 = WinWait("[CLASS:#32770]", "Extraction completed")
			WinClose($handle2)
			WinClose($handle)
			Local $return = FileRead($outdir & "\extract.log")
			If @error Then terminate("failed", $file, $arcdisp)
			FileDelete($outdir & "\extract.log")
			Cout($return)
			$success = True

		Case "robo"
			If $warnexecute Then Warn_Execute($filename & '.exe /unpack="<outdir>"')
			RunWait($file & ' /unpack="' & $outdir & '"', $filedir)

		Case "rpa"
			_Run($cmd & $rpa & ' -m -v -p "' & $outdir & '" "' & $file & '"', @ScriptDir)

		Case "sit"
			DirCreate($tempoutdir)
			FileMove($file, $tempoutdir)
			_Run($sit & ' "' & $tempoutdir & '\' & $filename & '.' & $fileext & '"', $tempoutdir, @SW_SHOW, False)
			FileMove($tempoutdir & '\' & $filename & '.' & $fileext, $file)
			MoveFiles($tempoutdir & '\', $outdir, True, "", True)

		Case "superdat"
			If $warnexecute Then Warn_Execute($filename & '.exe /e "<outdir>"')
			RunWait($file & ' /e "' & $outdir & '"', $outdir)
			Cout(FileRead($filedir & '\SuperDAT.log'))
			FileDelete($filedir & '\SuperDAT.log')

		Case "swf"
			; Run swfextract to get list of contents
			$return = StringSplit(FetchStdout($cmd & $swf & ' "' & $file & '"', $filedir, @SW_HIDE), @CRLF)
			;_ArrayDisplay($return)
			For $i = 2 To $return[0]
				$line = $return[$i]
				; Extract files
				If StringInStr($line, "MP3 Soundstream") Then
					_Run($cmd & $swf & ' -m "' & $file & '"', $outdir, @SW_HIDE, True, False, False)
					If FileExists($outdir & "\output.mp3") Then FileMove($outdir & "\output.mp3", $outdir & "\MP3 Soundstream\soundstream.mp3", 8 + 1)
				ElseIf $line <> "" Then
					$swf_arr = StringSplit(StringRegExpReplace(StringStripWS($line, 8), '(?i)\[(-\w)\]\d+(.+):(.*?)\)', "$1,$2,"), ",")
;~ 					_ArrayDisplay($swf_arr)
					$j = 3
					Do
						;Cout("$j = " & $j & @TAB & $swf_arr[$j])
						$swf_obj = 0
						$swf_obj = StringInStr($swf_arr[$j], "-")
						If $swf_obj Then
							For $k = StringMid($swf_arr[$j], 1, $swf_obj - 1) To StringMid($swf_arr[$j], $swf_obj + 1)
								_ArrayAdd($swf_arr, $k)
							Next
							$swf_arr[0] = UBound($swf_arr) - 1
;~ 							_ArrayDisplay($swf_arr)
						Else
							_Run($cmd & $swf & " " & $swf_arr[1] & " " & StringStripWS($swf_arr[$j], 1) & ' "' & $file & '"', $outdir, @SW_HIDE, True, False, False)
							; Rename and move file to subfolder
							If $swf_arr[2] = "Sound" Then
								FileMove($outdir & "\output.mp3", $outdir & "\" & $swf_arr[2] & "\" & $swf_arr[$j] & ".mp3", 8 + 1)
							ElseIf $swf_arr[2] = "PNGs" Then
								FileMove($outdir & "\output.png", $outdir & "\" & $swf_arr[2] & "\" & $swf_arr[$j] & ".png", 8 + 1)
							ElseIf $swf_arr[2] = "JPEG" Then
								FileMove($outdir & "\output.jpg", $outdir & "\" & $swf_arr[2] & "\" & $swf_arr[$j] & ".jpg", 8 + 1)
							Else
								FileMove($outdir & "\output.swf", $outdir & "\" & $swf_arr[2] & "\" & $swf_arr[$j] & ".swf", 8 + 1)
							EndIf
						EndIf
						$j += 1
					Until $j = $swf_arr[0] + 1
				EndIf
			Next

		Case "tar"
			If $fileext = "tar" Then
				_Run($cmd & $7z & ' x "' & $file & '"', $outdir)
			Else
				_Run($cmd & $7z & ' x "' & $file & '"', $outdir)
				_Run($cmd & $7z & ' x "' & $outdir & '\' & $filename & '.tar"', $outdir)
				FileDelete($outdir & '\' & $filename & '.tar')
			EndIf

		Case "thinstall"
			HasPlugin($thinstall)

			If $warnexecute Then Warn_Execute($file)
			$pid = Run($file, $filedir)
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

		Case "ttarch"
			; Get all supported games
			$aReturn = _StringBetween(FetchStdout($ttarch, @ScriptDir, @SW_HIDE), "Games", "Examples")
			If @error Then terminate("failed", $file, $arcdisp)
			$gameformat = StringRegExp($aReturn[0], "\d+ +(.+)", 3)
;~ 			_ArrayDisplay($gameformat)

			; Display game select GUI
			$game = GameSelect(_ArrayToString($gameformat), t('METHOD_GAME_NOGAME'))
			If $game Then
				$game = _ArraySearch($gameformat, $game)
				If $game > -1 Then _Run($cmd & $ttarch & ' -m ' & $game & ' "' & $file & '" "' & $outdir & '"', $outdir, @SW_HIDE)
			Else ; Delete outdir and return
				$returnFail = True
			EndIf

		Case "uif"
			_DeleteTrayMessageBox()
			Prompt(32 + 4, 'CONVERT_CDROM', 'UIF', 1)

			; Begin conversion, format selected by uif2iso
			_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'UIF ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 1)')
			_Run($cmd & $uif & ' "' & $file & '" "' & $outdir & "\" & $filename & '"', $filedir)
			$handle = FileFindFirstFile($outdir & "\" & $filename & ".*")
			$isofile = $outdir & "\" & FileFindNextFile($handle)
			FileClose($handle)

			; Begin extraction from .iso
			If FileGetSize($isofile) > 0 Then
				_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & 'UIF ' & t('TERM_IMAGE') & ' (' & t('TERM_STAGE') & ' 2)')
				$file = $isofile
				If Not CheckIso() Then _Run($cmd & $7z & ' x "' & $isofile & '"', $outdir)

				; Exit if conversion failed
			Else
				Prompt(16, 'CONVERT_CDROM_STAGE1_FAILED', '', 0)
				If FileExists($isofile) Then FileDelete($isofile)
				check7z()
				terminate("failed", $file, $arcdisp)
			EndIf

		Case "unity"
			IsJavaInstalled()
			_Run($unity & ' extract "' & $file & '"', $filedir, @SW_MINIMIZE, True, False)

		Case "unreal"
			HasPlugin($unreal)
			_Run($cmd & $unreal & ' -out="' & $outdir & '" "' & $file & '"', $outdir)

		Case "video"
			; Prompt to download FFmpeg if file not found
			If Not FileExists(@ScriptDir & "\bin\" & $OSArch & "\" & $ffmpeg) Then
				Prompt(48 + 4, 'FFMPEG_NEEDED', CreateArray($file, "http://ffmpeg.org/legal.html"), 1)
				GetFFmpeg()
				If @error Then terminate("silent", "", "")
			EndIf

			; Collect information about number of tracks
			$command = $cmd & $ffmpeg & ' -i "' & $file & '"'
			$return = FetchStdout($command, $outdir, @SW_HIDE)

			; Terminate if file could not be read by FFmpeg
			If StringInStr($return, "Invalid data found when processing input") Or Not StringInStr($return, "Stream") Then _
				terminate("failed", $file, $arcdisp)

			; Otherwise, extract all tracks
			$Streams = StringSplit($return, "Stream", 1)
			;_ArrayDisplay($Streams)
			Local $iVideo = 0, $iAudio = 0
			For $i = 2 To $Streams[0]
				$Streams[$i] = StringRegExpReplace($Streams[$i], "(?i)(?s).*?#(\d:\d)(.*?): (\w+): (\w+).*", "$3,$4,$1,$2")
;~ 				_ArrayDisplay($Streams)
				$StreamType = StringSplit($Streams[$i], ",")
				If $StreamType[1] == "Video" Then
					$iVideo += 1
					If $StreamType[2] == "h264" Then
						_Run($command & ' -vcodec copy -an -bsf:v h264_mp4toannexb -map ' & $StreamType[3] & ' "' & ($bIsUnicode? $sUnicodeName: $filename) & "_" & t('TERM_VIDEO') & StringFormat("_%02s", $iVideo) & $StreamType[4] & "." & $StreamType[2] & '"', $outdir, @SW_HIDE, True, False)
					Else
						; Special cases
						If StringInStr($StreamType[2], "wmv") Then
							$StreamType[2] = "wmv" ;wmv3
						ElseIf StringInStr($StreamType[2], "mpeg") Then
							$StreamType[2] = "mpeg" ;mpeg1video
						ElseIf StringInStr($StreamType[2], "v8") Then
							$StreamType[2] = "webm"
						EndIf
						_Run($command & ' -vcodec copy -an -map ' & $StreamType[3] & ' "' & ($bIsUnicode? $sUnicodeName: $filename) & "_" & t('TERM_VIDEO') & StringFormat("_%02s", $iVideo) & $StreamType[4] & "." & $StreamType[2] & '"', $outdir, @SW_HIDE, True, False)
					EndIf
				ElseIf $StreamType[1] == "Audio" Then
					$iAudio += 1
					; Special cases
					If StringInStr($StreamType[2], "wma") Then
						$StreamType[2] = "wma" ;wmav2
					ElseIf StringInStr($StreamType[2], "vorbis") Then
						$StreamType[2] = "ogg"
					EndIf

					_Run($command & ' -acodec copy -vn -map ' & $StreamType[3] & ' "' & ($bIsUnicode? $sUnicodeName: $filename) & "_" & t('TERM_AUDIO') & StringFormat("_%02s", $iAudio) & $StreamType[4] & "." & $StreamType[2] & '"', $outdir, @SW_HIDE, True, False)
				Else
					Cout("Unknown stream type: " & $StreamType[1])
				EndIf
			Next

		Case "vssfx"
			If $warnexecute Then Warn_Execute($filename & '.exe /extract')
			FileMove($file, $outdir)
			RunWait($outdir & '\' & $filename & '.' & $fileext & ' /extract', $outdir)
			FileMove($outdir & '\' & $filename & '.' & $fileext, $filedir)

			; removed - not possible to access due to 7zip check after deep scan
			;case "vssfxhotfix"
			;	if $warnexecute then Warn_Execute($filename & '.exe /xp:"<outdir>" /q')
			;	runwait($file & ' /xp:"' & $outdir & '" /q', $outdir)

		Case "vssfxpath"
			If $warnexecute Then Warn_Execute($filename & '.exe /extract:"<outdir>" /quiet')
			RunWait($file & ' /extract:"' & $outdir & '" /quiet', $outdir)

		Case "wise"
			$choice = MethodSelect($arctype, $arcdisp)

			; Extract with E_WISE
			If $choice == 'E_Wise' Then
				_Run($cmd & $wise_ewise & ' "' & $file & '" "' & $outdir & '"', $filedir)
				If DirGetSize($outdir) > $initdirsize Then
					RunWait($cmd & '00000000.BAT', $outdir, @SW_HIDE)
					FileDelete($outdir & '\00000000.BAT')
				EndIf

				; Extract with WUN
			ElseIf $choice == 'WUN' Then
				RunWait($cmd & $wise_wun & ' "' & $filename & '" "' & $tempoutdir & '"', $filedir)
				Local $removetemp = 1
				LoadPref("removetemp", $removetemp)
				If $removetemp Then
					FileDelete($tempoutdir & "\INST0*")
					FileDelete($tempoutdir & "\WISE0*")
				Else
					FileMove($tempoutdir & "\INST0*", $outdir)
					FileMove($tempoutdir & "\WISE0*", $outdir)
				EndIf
				MoveFiles($tempoutdir, $outdir, False, "", True)

				; Extract using the /x switch
			ElseIf $choice == 'Wise Installer /x' Then
				If $warnexecute Then Warn_Execute($filename & '.exe /x <outdir>')
				RunWait($file & ' /x ' & $outdir, $filedir)

				; Attempt to extract MSI
			ElseIf $choice == 'Wise MSI' Then

				; Prompt to continue
				_DeleteTrayMessageBox()
				Prompt(48 + 4, 'WISE_MSI_PROMPT', 1)

				; First, check for any files that are already in extraction dir
				If $warnexecute Then Warn_Execute($filename & '.exe /?')
				_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & $arcdisp)
				$oldfiles = ReturnFiles(@CommonFilesDir & "\Wise Installation Wizard")

				; Run installer
				Opt("WinTitleMatchMode", 3)
				$pid = Run($file & ' /?', $filedir)
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
			ElseIf $choice == 'Unzip' Then
				$return = RunWait($cmd & $zip & ' -x "' & $file & '"', $outdir)
				If $return <> 0 Then
					_Run($cmd & $7z & ' x "' & $file & '"', $outdir)
				EndIf
			EndIf

			; Append missing file extensions
			If $appendext Then
				AppendExtensions($outdir)
			EndIf

		Case "uha"
			_Run($cmd & $uharc & ' x -t"' & $outdir & '" "' & $file & '"', $outdir)
			If Not $success And DirGetSize($outdir) <= $initdirsize Then
				_Run($cmd & $uharc04 & ' x -t"' & $outdir & '" "' & $file & '"', $outdir)
				If Not $success And DirGetSize($outdir) <= $initdirsize Then _
						_Run($cmd & $uharc02 & ' x -t' & FileGetShortName($outdir) & ' ' & FileGetShortName($file), $outdir)
			EndIf

		Case "uu"
			_Run($cmd & $uu & ' -p "' & $outdir & '" -i "' & $file & '"', $filedir)

		Case "xz"
			RunWait($cmd & $7z & ' x "' & $file & '"', $outdir)
			If FileExists($outdir & '\' & $filename) And StringTrimLeft($filename, StringInStr($filename, '.', 0, -1)) = "tar" Then
				RunWait($cmd & $7z & ' x "' & $outdir & '\' & $filename & '"', $outdir)
				FileDelete($outdir & '\' & $filename)
			EndIf

		Case "Z"
			_Run($cmd & $7z & ' x "' & $file & '"', $outdir)
			If FileExists($outdir & '\' & $filename) And StringTrimLeft($filename, StringInStr($filename, '.', 0, -1)) = "tar" Then
				RunWait($cmd & $7z & ' x "' & $outdir & '\' & $filename & '"', $outdir)
				FileDelete($outdir & '\' & $filename)
			EndIf

		Case "zip"
			_Run($cmd & $7z & ' x "' & $file & '"', $outdir)
			If Not $success Then _Run($cmd & $zip & ' -x "' & $file & '"', $outdir, @SW_MINIMIZE, False)

		Case "zoo"
			DirCreate($tempoutdir)
			FileMove($file, $tempoutdir)
			_Run($cmd & $zoo & ' -x ' & $filename & '.' & $fileext, $tempoutdir, @SW_HIDE)
			FileMove($tempoutdir & '\' & $filename & '.' & $fileext, $file)
			MoveFiles($tempoutdir, $outdir, False, "", True)

		Case Else
			Cout("Unknown arctype: " & $arctype & ". Feature not implemented!")
	EndSwitch

	_DeleteTrayMessageBox()

	; -----Success evaluation----- ;

	; Exit if success returned by _Run function
	If $success Then
		; Special actions for 7zip extraction
		If $arctype == "7z" And ($fileext == "exe" Or StringInStr($filetype, "SFX")) Then
			; Check if sfx archive and extract sfx script using 7ZSplit if possible
			Cout("Trying to extract sfx script")
			Run($7zsplit & ' "' & $file & '"', $outdir, @SW_HIDE)
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

			; Check generic .exe ressource extraction
			If FileExists($outdir & "\[0]") Then
				; Try to find correct file extensions for unpacked files
				If Prompt(48 + 1, "UNPACK_GENERIC_ZIP", CreateArray($file, "7Zip", $outdir & "\[0]"), 0) Then
					$file = $outdir & "\[0]"
					$outdir = $file & "\"
					FilenameParse($file)
					; Try to extract unpacked File (skip file extension checks)
					StartExtraction(False)
				Else
					AppendExtensions($outdir)
					If $returnSuccess Then Return True
					terminate("success", "", $arctype)
				EndIf
			EndIf
		Else
			If $returnSuccess Then Return True
			terminate("success", "", $arctype)
		EndIf
	EndIf

	; Otherwise, check directory size
	If DirGetSize($outdir) <= $initdirsize Then
		If $arctype = "ace" And $fileext = "exe" Then Return False
		If FileExists($isofile) Then
			If Prompt(16 + 4, 'CONVERT_CDROM_STAGE2_FAILED', '', 0) Then FileDelete($isofile)
		EndIf
		If $returnFail Then Return False
		terminate("failed", $file, $arcdisp)
	ElseIf FileGetTime($outdir, 0, 1) == $dirmtime Then
		If $returnFail Then Return False
		terminate("failed", $file, $arcdisp)
	EndIf

	If $isofile Then
		If FileExists($isofile) Then FileDelete($isofile)
		If FileExists($tempoutdir) Then DirRemove($tempoutdir, 1)
	EndIf

	If $returnSuccess Then Return True
	terminate("success", "", $arctype)
EndFunc   ;==>extract

; Unpack packed executable
Func unpack()
	Local $packer
	If StringInStr($filetype, "UPX", 0) Or $fileext = "dll" Then
		$packer = "UPX"
	ElseIf StringInStr($filetype, "ASPack", 0) Then
		$packer = "ASPack"
	EndIf

	; Prompt to continue
	If Not Prompt(32 + 4, 'UNPACK_PROMPT', CreateArray($packer, $filedir & "\" & $filename & "_" & t('TERM_UNPACKED') _
			 & "." & $fileext), 0) Then Return

	; Unpack file
	If $packer == "UPX" Then
		_Run($cmd & $upx & ' -d -k "' & $file & '"', $filedir)
		$tempext = StringTrimRight($fileext, 1) & '~'
		If FileExists($filedir & "\" & $filename & "." & $tempext) Then
			FileMove($file, $filedir & "\" & $filename & "_" & t('TERM_UNPACKED') & "." & $fileext)
			FileMove($filedir & "\" & $filename & "." & $tempext, $file)
		EndIf
	ElseIf $packer == "ASPack" Then
		RunWait($cmd & $aspack & ' "' & $file & '" "' & $filedir & '\' & $filename & '_' & t('TERM_UNPACKED') & _
				$fileext & '" /NO_PROMPT', $filedir)
	EndIf

	; Success evaluation
	If FileExists($filedir & "\" & $filename & "_" & t('TERM_UNPACKED') & "." & $fileext) Then
		; Prompt if unpacked file should be scanned
		If Prompt(32 + 4, 'UNPACK_AGAIN', CreateArray($file, $filename & '_' & t('TERM_UNPACKED') & "." & $fileext), 0) Then
			$file = $filedir & "\" & $filename & "_" & t('TERM_UNPACKED') & "." & $fileext
			$outdir = $filedir & "\" & $filename & "_" & t('TERM_UNPACKED') & "\"
			StartExtraction()
		Else
			terminate("success", "", $packer)
		EndIf
	Else
		Prompt(16, 'UNPACK_FAILED', $file, 1)
	EndIf
EndFunc   ;==>unpack

; Terminate if specified plugin was not found
Func HasPlugin($plugin, $returnFail = False)
	Cout("Searching for plugin " & $plugin)
	If Not StringInStr($plugin, "\bin\") Then $plugin = @ScriptDir & "\bin\" & $plugin
	If Not FileExists(_PathFull($plugin, @ScriptDir)) Then
		Cout("Error: plugin not found")
		If $returnFail Then Return False
		terminate('missingexe', $file, $plugin)
	EndIf
	Return True
EndFunc

; Check if enough free space is available
Func HasFreeSpace()
	Local $freeSpace = DriveSpaceFree($outdir)
	Local $fileSize = FileGetSize($file) / 1048576 ; MB

	If $freeSpace < $fileSize Then
		Local $diff = $freeSpace - $fileSize
		Cout("Not enough free space available: " & $freeSpace & " MB, needed: " & $fileSize & " MB, difference: " & $diff & " MB.")

		If $silentmode Then terminate("failed", '', "Not enough free space available: " & $freeSpace & " MB, needed: " & $fileSize & " MB, difference: " & $diff & " MB.")

		$return = MsgBox(262144 + 48 + 2, $name, t('NO_FREE_SPACE', CreateArray($freeSpace, Round($fileSize, 1), Round($diff, 2))))
		If $return = 4 Then ; Retry
			HasFreeSpace()
			Return
		ElseIf $return = 2 Then ; Cancel
			If $createdir Then DirRemove($outdir, 0)
			terminate("silent", '', '')
		EndIf
	EndIf

EndFunc   ;==>HasFreeSpace

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
	DirCreate($dest)
	$handle = FileFindFirstFile($source & "\*")
	If Not @error Then
		While 1
			$fname = FileFindNextFile($handle)
			If @error Then
				ExitLoop
			ElseIf StringInStr($omit, $fname) Then
				ContinueLoop
			Else
				If StringInStr(FileGetAttrib($source & '\' & $fname), 'D') Then
					DirMove($source & '\' & $fname, $dest, 1)
				Else
					FileMove($source & '\' & $fname, $dest, $force)
				EndIf
			EndIf
		WEnd
		FileClose($handle)
	Else
		SetError(1)
		Return
	EndIf

	If $removeSourceDir Then Return DirRemove($source, ($omit = ""? 1: 0))
EndFunc   ;==>MoveFiles

; Append missing file extensions using TrID
Func AppendExtensions($dir)
	Local $files
	$files = FileSearch($dir)
	If $files[1] <> '' Then
		For $i = 1 To $files[0]
			If Not StringInStr(FileGetAttrib($files[$i]), 'D') Then
				$filename = StringTrimLeft($files[$i], StringInStr($files[$i], '\', 0, -1))
				If Not StringInStr($filename, '.') _
						Or StringLeft($filename, 7) = 'Binary.' _
						Or StringRight($filename, 4) = '.bin' Then
					RunWait($cmd & $trid & ' "' & $files[$i] & '" -ae', $dir, @SW_HIDE)
				EndIf
			EndIf
		Next
	EndIf
EndFunc   ;==>AppendExtensions

; Recursively search for given pattern
; code by w0uter (http://www.autoitscript.com/forum/index.php?showtopic=16421)
Func FileSearch($s_Mask = '', $i_Recurse = 1)
	Local $s_Buf = ''
	If $i_Recurse Then
		Local $s_Command = ' /c dir /B /S "'
	Else
		Local $s_Command = ' /c dir /B "'
	EndIf
	$i_Pid = Run(@ComSpec & $s_Command & $s_Mask & '"', @WorkingDir, @SW_HIDE, 2 + 4)
	While Not @error
		$s_Buf &= StdoutRead($i_Pid)
	WEnd
	$s_Buf = StringSplit(StringTrimRight($s_Buf, 2), @CRLF, 1)
	ProcessClose($i_Pid)
	If UBound($s_Buf) = 2 And $s_Buf[1] = '' Then SetError(1)
	Return $s_Buf
EndFunc   ;==>FileSearch

; Handle program termination with appropriate error message
Func terminate($status, $fname, $ID)
	Local $LogFile = 0, $exitcode = 0, $shortStatus = ($status = "success")? $ID: $status

	; Rename unicode file
	If $bIsUnicode Then
		Cout("Renaming unicode file")
		FileMove($file, $oldpath, 1)
		DirMove($outdir, $oldoutdir)
		$fname = $oldpath
	EndIf

	Cout("Terminating - Status: " & $status)

	; When multiple files are selected and executed via command line, they are added to batch queue, but the working instance uses in-memory data.
	; So we need to look for changes in the batch queue file, so batch mode could be enabled if necessary.
	If Not $silentmode And GetBatchQueue() Then $silentmode = True

	; Create log file if enabled in options
	If $Log And Not ($status = "silent" Or $status = "syntax" Or $status = "fileinfo" Or $status = "notpacked") Or ($status = "fileinfo" And $silentmode) Then _
		$LogFile = CreateLog($shortStatus)

	SaveStat($shortStatus)

	Select
		; Display usage information and exit
		Case $status == "syntax"
			$syntax = t('HELP_SUMMARY')
			$syntax &= t('HELP_SYNTAX', @ScriptName)
			$syntax &= t('HELP_ARGUMENTS')
			$syntax &= t('HELP_HELP')
			$syntax &= t('HELP_PREFS')
			$syntax &= t('HELP_REMOVE')
			$syntax &= t('HELP_CLEAR')
			$syntax &= t('HELP_FILENAME')
			$syntax &= t('HELP_DESTINATION')
			$syntax &= t('HELP_SCAN')
			$syntax &= t('HELP_SILENT')
			$syntax &= t('HELP_BATCH')
			$syntax &= t('HELP_SUB')
			$syntax &= t('HELP_EXAMPLE1')
			$syntax &= t('HELP_EXAMPLE2', @ScriptName)
			$syntax &= t('HELP_NOARGS')
			MsgBox(262144 + 32, $title, $syntax, 15)

			; Display file type information and exit
		Case $status == "fileinfo"
			If $filetype == "" Then
				$exitcode = 4
				$filetype = t('UNKNOWN_EXT', CreateArray($file, ""))
			EndIf
			If $silentmode Then ; Save info to result file if in silent mode
				$handle = FileOpen($fileScanLogFile, 8 + 1)
				FileWrite($handle, $file & @CRLF & @CRLF & $filetype & @CRLF & "------------------------------------------------------------" & @CRLF)
				FileClose($handle)
			Else
				MsgBox(262144 + 64, $title, $filetype)
			EndIf

			; Display error information and exit
		Case $status == "unknownexe"
			If Not $silentmode And Prompt(256 + 16 + 4, 'CANNOT_EXTRACT', CreateArray($file, $filetype), 0) Then Run($exeinfope & ' "' & $file & '"', $filedir)
			$exitcode = 3
		Case $status == "unknownext"
			Prompt(16, 'UNKNOWN_EXT', CreateArray($file, $filetype), 0)
			$exitcode = 4
		Case $status == "invaliddir"
			Prompt(16, 'INVALID_DIR', $fname, 0)
			$exitcode = 5
		Case $status == "notpacked"
			Prompt(48, 'NOT_PACKED', CreateArray($file, $filetype), 0)
			$exitcode = 6
		Case $status == "notsupported"
			Prompt(16, 'NOT_SUPPORTED', CreateArray($file, $filetype), 0)
			$exitcode = 7
		Case $status == "missingexe"
			Prompt(48, 'MISSING_EXE', CreateArray($file, $ID))
			$exitcode = 8
		Case $status == "timeout"
			Prompt(48, 'EXTRACT_TIMEOUT', $file)
			$exitcode = 9

			; Display failed attempt information and exit
		Case $status == "failed"
			If Not $silentmode And Prompt(256 + 16 + 4, 'EXTRACT_FAILED', CreateArray($file, $ID), 0) Then
				If $LogFile Then
					ShellExecute($LogFile)
				Else
					ShellExecute(CreateLog($status))
				EndIf
			EndIf
			$exitcode = 1

			; Exit successfully
		Case $status == "success"
			If $DeleteOrigFile = 1 Then
				FileDelete($file)
			ElseIf $DeleteOrigFile = 2 Then
				If Not $silentmode And Prompt(32 + 4, 'FILE_DELETE', $file) Then FileDelete($file)
			EndIf
			If $OpenOutDir And Not $silentmode Then Run("explorer.exe /e, " & $outdir)
	EndSelect

	; Write error log if in batchmode
	If $exitcode <> 0 And $silentmode And $extract Then
		$handle = FileOpen(@ScriptDir & "\log\errorlog.txt", 8 + 1)
		FileWrite($handle, $filename & "." & $fileext & " (" & StringUpper($status) & ")" & @CRLF & @TAB & $ID & @CRLF)
		FileClose($handle)
	EndIf

	; Delete empty output directory if failed
	If $createdir And $status <> "success" And DirGetSize($outdir) = 0 Then DirRemove($outdir, 0)

	If $exitcode == 1 Or $exitcode == 3 Or $exitcode == 4 Or $exitcode == 7 Then
		If $FB_ask And $extract And Not $silentmode And Prompt(4, 'FEEDBACK_PROMPT', $file, 0) Then
			; Attach input file's first bytes for debug purpose
			Cout("--------------------------------------------------File dump--------------------------------------------------" & _
				 @CRLF & _HexDump($file, 1024))
			Cout("------------------------------------------------File metadata------------------------------------------------" & _
				 @CRLF & _ArrayToString(_GetExtProperty($file), @CRLF))
			; Prompt to send feedback
			GUI_Feedback($status, $file, $debug)
			Do
				Sleep(500)
			Until $FB_finish
		EndIf
	EndIf

	If $batchEnabled = 1 And $status <> "silent" Then ; Don't start batch if gui is closed
		; Start next extraction
		BatchQueuePop()
	ElseIf $KeepOpen And $cmdline[0] = 0 And $status <> "silent" Then
		Run(@ScriptFullPath)
	EndIf

	; Check for updates
	If Not $silentmode And $status <> "silent" And _DateDiff("D", $lastupdate, _NowCalc()) >= $updateinterval Then CheckUpdate(True)

	Exit $exitcode
EndFunc

; Function to prompt user for choice of extraction method
Func MethodSelect($format, $splashdisp)
	; Set info base on format
	_DeleteTrayMessageBox()
	Local $base_height = 130
	Local $base_radio = 100
	If $format == 'wise' Then
		$select_type = 'Wise Installer'
		Dim $method[5][2], $select[5]
		$method[0][0] = 'E_Wise'
		$method[0][1] = 'METHOD_UNPACKER_RADIO'
		$method[1][0] = 'WUN'
		$method[1][1] = 'METHOD_UNPACKER_RADIO'
		$method[2][0] = 'Wise Installer /x'
		$method[2][1] = 'METHOD_SWITCH_RADIO'
		$method[3][0] = 'Wise MSI'
		$method[3][1] = 'METHOD_EXTRACTION_RADIO'
		$method[4][0] = 'Unzip'
		$method[4][1] = 'METHOD_EXTRACTION_RADIO'
		;$base_height += 45
	ElseIf $format == 'msi' Then
		$select_type = 'MSI Installer'
		Dim $method[4][2], $select[4]
		$method[0][0] = 'jsMSI Unpacker'
		$method[0][1] = 'METHOD_EXTRACTION_RADIO'
		$method[1][0] = 'MsiX'
		$method[1][1] = 'METHOD_EXTRACTION_RADIO'
		$method[2][0] = 'MSI TC Packer'
		$method[2][1] = 'METHOD_EXTRACTION_RADIO'
		$method[3][0] = 'MSI'
		$method[3][1] = 'METHOD_ADMIN_RADIO'
		;$base_height += 15
		;$base_radio += 15
	ElseIf $format == 'msp' Then
		$select_type = 'MSP Package'
		Dim $method[3][2], $select[3]
		$method[0][0] = 'MSI TC Packer'
		$method[0][1] = 'METHOD_EXTRACTION_RADIO'
		$method[1][0] = 'MsiX'
		$method[1][1] = 'METHOD_EXTRACTION_RADIO'
		$method[2][0] = '7-Zip'
		$method[2][1] = 'METHOD_EXTRACTION_RADIO'
	ElseIf $format == 'mht' Then
		$select_type = 'MHTML Archive'
		Dim $method[2][2], $select[2]
		$method[0][0] = 'ExtractMHT'
		$method[0][1] = 'METHOD_EXTRACTION_RADIO'
		$method[1][0] = 'MhtUnPack'
		$method[1][1] = 'METHOD_EXTRACTION_RADIO'
	ElseIf $format == 'is3arc' Then
		$select_type = 'InstallShield 3.x Archive'
		Dim $method[2][2], $select[2]
		$method[0][0] = 'STIX'
		$method[0][1] = 'METHOD_EXTRACTION_RADIO'
		$method[1][0] = 'unshield'
		$method[1][1] = 'METHOD_EXTRACTION_RADIO'
	ElseIf $format == 'isexe' Then
		$select_type = 'InstallShield Installer'
		Dim $method[4][2], $select[4]
		$method[0][0] = 'isxunpack'
		$method[0][1] = 'METHOD_EXTRACTION_RADIO'
		$method[1][0] = 'unshield'
		$method[1][1] = 'METHOD_EXTRACTION_RADIO'
		$method[2][0] = 'InstallShield /b'
		$method[2][1] = 'METHOD_SWITCH_RADIO'
		$method[3][0] = 'not InstallShield'
		$method[3][1] = 'METHOD_NOTIS_RADIO'
	ElseIf $format == 'iscab' Then
		$select_type = 'InstallShield Cabinet'
		Dim $method[2][2], $select[2]
		$method[0][0] = 'iscab'
		$method[0][1] = 'METHOD_EXTRACTION_RADIO'
		$method[1][0] = 'is6comp'
		$method[1][1] = 'METHOD_EXTRACTION_RADIO'
	EndIf

	; Auto choose first extraction method in silent mode
	If $silentmode Then
		Cout("Extractor selected automatically - run again in normal mode if not extracted correctly")
		_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & $splashdisp)
		Return $method[0][0]
	EndIf

	; Create GUI and set header information
	Opt("GUIOnEventMode", 0)
	Local $guimethod = GUICreate($title, 330, $base_height + (UBound($method) * 20))
	$header = GUICtrlCreateLabel(t('METHOD_HEADER', $select_type), 5, 5, 320, 20)
	GUICtrlCreateLabel(t('METHOD_TEXT_LABEL', $select_type), 5, 25, 320, 65, $SS_LEFT)

	; Create radio selection options
	GUICtrlCreateGroup(t('METHOD_RADIO_LABEL'), 5, $base_radio, 215, 25 + (UBound($method) * 20))
	For $i = 0 To UBound($method) - 1
		$select[$i] = GUICtrlCreateRadio(t($method[$i][1], $method[$i][0]), 10, $base_radio + 20 + ($i * 20), 205, 20)
	Next
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; Create buttons
	$ok = GUICtrlCreateButton(t('OK_BUT'), 235, $base_radio - 10 + (UBound($method) * 10), 80, 20)
	$cancel = GUICtrlCreateButton(t('CANCEL_BUT'), 235, $base_radio - 10 + (UBound($method) * 10) + 30, 80, 20)

	; Set properties
	GUICtrlSetFont($header, -1, 1200)
	GUICtrlSetState($select[0], $GUI_CHECKED)
	GUICtrlSetState($ok, $GUI_DEFBUTTON)

	; Display GUI and wait for action
	GUISetState(@SW_SHOW)
	While 1
		$action = GUIGetMsg()
		Select
			; Set extract command
			Case $action == $ok
				For $i = 0 To UBound($method) - 1
					If GUICtrlRead($select[$i]) == $GUI_CHECKED Then
						GUIDelete($guimethod)
						_CreateTrayMessageBox(t('EXTRACTING') & @CRLF & $splashdisp)
						Return $method[$i][0]
					EndIf
				Next

				; Exit if Cancel clicked or window closed
			Case $action == $GUI_EVENT_CLOSE Or $action == $cancel
				If $createdir Then DirRemove($outdir, 0)
				terminate("silent", '', '')

		EndSelect
	WEnd
EndFunc   ;==>MethodSelect

; Create GUI to select game and return selection
Func GameSelect($sEntries, $sStandard)
	Local $sSelection = 0
	$GameSelectGUI = GUICreate($title, 274, 460, -1, -1, BitOR($WS_SIZEBOX, $WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU))
	$GameSelectLabel = GUICtrlCreateLabel(t('METHOD_GAME_LABEL', CreateArray($filename & "." & $fileext, $sStandard)), 10, 8, 252, 210, $SS_CENTER)
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
	Return $sSelection
EndFunc

; Warn user before executing files for extraction
Func Warn_Execute($command)
	Cout("Warning: In order to extract files of this type, the file must be directly executed.")
	$prompt = MsgBox(262144 + 49, $title, t('WARN_EXECUTE', $command))
	If $prompt <> 1 Then
		If $createdir Then DirRemove($outdir, 0)
		terminate('silent', '', '')
	EndIf
	Return True
EndFunc   ;==>Warn_Execute

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
	Local $return = MsgBox(262144 + $show_flag, $title, t($Message, $vars))
	If $return == 1 Or $return == 6 Then
		Return 1
	Else
		If Not $terminate Then Return 0
		If $createdir Then DirRemove($outdir, 0)
		terminate("silent", '', '')
	EndIf
EndFunc

; Show Tray Message
; Based on work by Valuater (http://www.autoitscript.com/forum/topic/85977-system-tray-message-box-udf/)
Func _CreateTrayMessageBox($TBText)
	_DeleteTrayMessageBox()

	Local $iSpace = -1
	Local Const $TBwidth = 225, $TBheight = 100, $left = 15, $top = 15, $width = 195, $iBetween = 5, $iMaxCharCount = 28
	If $NoBox = 1 Then Return

	; Determine taskbar size
	Local $pos = WinGetPos("[CLASS:Shell_TrayWnd]")
	If @error Then Local $pos[4] = [0, 0, @DesktopWidth, 30]
	If $pos[0] = $pos[1] Then
		$iSpace = $pos[3] + $iBetween
	Else
		$iSpace = $pos[1] - $TBheight - $iBetween
	EndIf

	If $iSpace < 0 Or $iSpace > @DesktopHeight Then $iSpace = @DesktopHeight - $TBheight - $iBetween

	; Create GUI
	Global $TBgui = GUICreate($name, $TBwidth, $TBheight, $trayX > -1 ? $trayX : @DesktopWidth - ($TBwidth + $iBetween), $trayY > -1 ? $trayY : $iSpace, _
			$WS_POPUP, BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
	_GuiRoundCorners($TBgui, 0, 0, 30, 30)
	If $filename = "" Then
		Global $Tray_File = GUICtrlCreateLabel($TBText, $left, $top, $width, 80)
	ElseIf StringLen(($bIsUnicode? $sUnicodeName: $filename) & "." & $fileext) > $iMaxCharCount Then
		Global $Tray_File = GUICtrlCreateLabel(StringLeft(($bIsUnicode? $sUnicodeName: $filename) & "." & $fileext, $iMaxCharCount) & " [...]" & @CRLF & @CRLF & $TBText, $left, $top, $width, 80)
	Else
		Global $Tray_File = GUICtrlCreateLabel(($bIsUnicode? $sUnicodeName: $filename) & "." & $fileext & @CRLF & @CRLF & $TBText, $left, $top, $width, 80)
	EndIf
	Global $TrayMsg_Status = GUICtrlCreateLabel("", $left, 74, $width, 20, $SS_CENTER)
;~     DllCall ( "user32.dll", "int", "AnimateWindow", "hwnd", $TBgui, "int", 250, "long", 0x00080000 )
	GUISetState(@SW_SHOWNOACTIVATE)

	; Workaround to keep corners round while fading in
	For $i = 0 To 255 Step 10
		WinSetTrans($TBgui, "", $i)
		Sleep(1)
	Next
EndFunc   ;==>_CreateTrayMessageBox

; Close Tray Message
; Based on work by Valuater (http://www.autoitscript.com/forum/topic/85977-system-tray-message-box-udf/)
Func _DeleteTrayMessageBox()
	If Not $TBgui Then Return
	;DllCall ( "user32.dll", "int", "AnimateWindow", "hwnd", $TBgui, "int", 300, "long", 0x00090000 )

	; Workaround to keep corners round while fading out
	For $i = 255 To 0 Step -10
		WinSetTrans($TBgui, "", $i)
		Sleep(1)
	Next

	GUIDelete($TBgui)
	$TBgui = 0
EndFunc   ;==>_DeleteTrayMessageBox

; Create command line for current file
Func GetCmd($silent = True)
	If Not $file Then Return
	Local $return = '"' & $file & '"'

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

	If StringInStr($return, $cmdline) And Not Prompt(32 + 4, 'BATCH_DUPLICATE', $file, 0) Then
		Cout("Not adding duplicate file " & $cmdline)
		FileClose($handle)
		Return
	EndIf
	FileWriteLine($handle, $cmdline)
	FileClose($handle)
	Cout("File added to batch queue: " & $cmdline)
	GetBatchQueue()
EndFunc   ;==>AddToBatch

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
	Cout("Fetching next queue element")
;~ 	_ArrayDisplay($queueArray)
;~ 	MsgBox(0, "", "UBound: " & UBound($queueArray))
	If Not IsArray($queueArray) Or UBound($queueArray) = 0 Then GetBatchQueue()

	If Not IsArray($queueArray) Or UBound($queueArray) = 0 Or $queueArray[0] = 0 Then ; Queue empty
		Cout("Queue empty")
		EnableBatchMode(False)

		If FileExists($fileScanLogFile) Then ShellExecute($fileScanLogFile)

		$handle = FileOpen(@ScriptDir & "\log\errorlog.txt")
		If $handle = -1 Then Return
		Local $return = FileRead($handle)
		FileClose($handle)

		FileDelete(@ScriptDir & "\log\errorlog.txt")
		If $return <> "" Then MsgBox(262144 + 48, $name, t('BATCH_FINISH', $return))
	Else ; Get next command and execute it
		Local $element = $queueArray[1]
		_ArrayDelete($queueArray, 1)
		$queueArray[0] -= 1
		Cout($element)
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

	Local $aLanguage[36] = [35, "English", "Chinese", "Danish", "Dutch", "Estonian", "Finnish", "French", "German", "Hungarian", "Italian", _
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
EndFunc   ;==>_GetOSLanguage

; Determine whether JRE is installed or not and terminate if not found
Func IsJavaInstalled()
	Local $return = FetchStdout($cmd & "java", @ScriptDir, @SW_HIDE)
	If StringInStr($return, "java [-options]") Then
		Cout("Java is installed")
		Return True
	EndIf
	terminate('missingexe', $file, "Java Runtime Environment")
EndFunc   ;==>IsJavaInstalled

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
	If $iLine = -1 Then Return StringTrimLeft($sString, StringInStr($sString, @CRLF, 0, -2))
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
; Requirement(s):   File specified in $spath must exist.
; Return Value(s):  On Success - The extended file property, or if $iProp = -1 then an array with all properties
;                   On Failure - 0, @Error - 1 (If file does not exist)
; Author(s):        Simucal (Simucal@gmail.com)
; Note(s):
;
;===============================================================================
Func _GetExtProperty($sPath, $iProp = -1)
    Local $iExist, $sFile, $sDir, $oShellApp, $oDir, $oFile, $aProperty, $sProperty
    $iExist = FileExists($sPath)
    If $iExist = 0 Then
        SetError(1)
        Return 0
    Else
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
            $sProperty = $oDir.GetDetailsOf ($oFile, $iProp)
            If $sProperty = "" Then
                Return 0
            Else
                Return $sProperty
            EndIf
        EndIf
    EndIf
EndFunc   ;==>_GetExtProperty

; Dump complete debug content to log file
Func CreateLog($status)
	Local $name = @ScriptDir & "\log\" & @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & "-" & @MIN & "-" & @SEC & "_"
	If $status <> "success" Then $name &= StringUpper($status) & "_"
	$name &= ($bIsUnicode? $sUnicodeName: $filename) & "." & $fileext & ".log"
	$handle = FileOpen($name, 32 + 8 + 2)
	FileWrite($handle, $debug)
	FileClose($handle)
	Return $name
EndFunc

; Executes a program and log output using tee
Func _Run($f, $workingdir, $show_flag = @SW_MINIMIZE, $useTee = True, $patternSearch = True, $initialShow = True)
	Local $teeCmd = ' 2>&1 | ' & $tee & ' "' & @ScriptDir & '\log\teelog.txt"'
	Cout("Executing: " & $f & ($useTee? $teeCmd: "") & " with options: useTee = " & $useTee & ", patternSearch = " & $patternSearch)
	Global $run = 0, $runtitle = 0
	Local $return = "", $pos = 0, $size = 1, $lastSize = 0

	; Create log
	If $useTee Then
		HasPlugin($tee)
		If Not FileExists(@ScriptDir & "\log\") Then DirCreate(@ScriptDir & "\log\")
		$run = Run($f & $teeCmd, $workingdir, $initialShow? @SW_MINIMIZE: $show_flag)
		If @error Then Return
		Local $TimerStart = TimerInit()
		Cout("Pid: " & $run)
		Do
			Sleep(1)
			If TimerDiff($TimerStart) > 5000 Then ExitLoop
		Until ProcessExists($run)

		$runtitle = _WinGetByPID($run)
		WinSetState($runtitle, "", $show_flag)
		Cout("Runtitle: " & $runtitle)

		; Wait until logfile exists
		$TimerStart = TimerInit()

		Do
			Sleep(10)
			If TimerDiff($TimerStart) > 5000 Then ExitLoop
		Until FileExists(@ScriptDir & "\log\teelog.txt")
		$handle = FileOpen(@ScriptDir & "\log\teelog.txt")
		$state = ""

		; Show progress (percentage) in status box
		While ProcessExists($run)
			$return = FileRead($handle)
;~ 			Cout($return)
			If $return <> $state Then
				$state = $return
				; Automatically show cmd window when user input needed
				If StringInStr($return, "password", 0) Or StringInStr($return, "already exist", 0) Or _
						StringInStr($return, "overwrite", 0) Or StringInStr($return, "replace", 0) Then
					WinSetState($runtitle, "", @SW_SHOW)
					GUICtrlSetFont($TrayMsg_Status, 8.5, 900)
					GUICtrlSetData($TrayMsg_Status, t('INPUT_NEEDED'))
					WinActivate($runtitle)
					$lastSize = Round((DirGetSize($outdir) - $initdirsize) / 1024 / 1024, 3)
					ContinueLoop
				EndIf
				; Percentage indicator
				If $patternSearch = True And $TBgui Then
					Cout("PATTERN")
					If StringInStr($return, "%", 0, -1) Then ; x %
						$aReturn = StringRegExp($return, "(\d{1,3})[\d\.,]*%", 1)
						If UBound($aReturn) > 0 Then
							$size = -1
							GUICtrlSetData($TrayMsg_Status, _ArrayPop($aReturn) & "%")
						EndIf
					ElseIf StringInStr($return, "/", 0, -1) Then
						$aReturn = StringRegExp($return, " (\d+)/(\d+)", 1) ; x/y
						If UBound($aReturn) > 1 Then
							$size = -1
							$Num = _ArrayPop($aReturn)
							GUICtrlSetData($TrayMsg_Status, t('TERM_FILE') & " " & _ArrayPop($aReturn) & "/" & $Num)
						EndIf
					Else
						$aReturn = StringRegExp($return, "\[(\d+) on (\d+)\]", 1) ; [x on y]
						If UBound($aReturn) > 1 Then
							$size = -1
							$Num = _ArrayPop($aReturn)
							GUICtrlSetData($TrayMsg_Status, t('TERM_FILE') & " " & _ArrayPop($aReturn) & "/" & $Num)
						Else ; # x
							$pos = StringInStr($return, "#", 0, -1)
							If $pos Then
								$Num = Number(StringMid($return, $pos + 1), 1)
								If $Num > 0 Then
									$size = -1
									GUICtrlSetData($TrayMsg_Status, t('TERM_FILE') & " #" & $Num)
								EndIf
							EndIf
							Sleep(50)
						EndIf
					EndIf
				EndIf
			EndIf
			; Size of extracted file(s)
			If $size > -1 Then
				$size = Round((DirGetSize($outdir) - $initdirsize) / 1024 / 1024, 3)
;~ 				Cout("Size: " & $size & @TAB & $lastSize)
				If $size > 0 And $size <> $lastSize Then
					Cout("Size: " & $size & @TAB & $lastSize)
					GUICtrlSetData($TrayMsg_Status, $size & " MB")
				EndIf
				$lastSize = $size
				Sleep(50)
			EndIf
			Sleep(200)
		WEnd
		; Write tee log to UniExtract log file
		FileSetPos($handle, 0, $FILE_BEGIN)
		$return = FileRead($handle)
		If Not StringIsSpace($return) Then Cout("Teelog:" & @CRLF & $return)
		FileClose($handle)
		FileDelete(@ScriptDir & "\log\teelog.txt")

		; Check for success indicator in log
		If StringInStr($return, "err code(", 1) OR StringInStr($return, "stacktrace", 1) Then
			SetError(1)
		ElseIf StringInStr($return, "Everything is Ok") Or StringInStr($return, "Break signaled") _
				Or StringInStr($return, "0 failed") Or StringInStr($return, "All files OK") _
				Or StringInStr($return, "All OK") Or StringInStr($return, "done.") _
				Or StringInStr($return, "Done ...") Or StringInStr($return, ": done") Then
			Cout("Success evaluation passed")
			$success = True
		EndIf

	; Do not create log
	Else
		Cout("Extraction cannot be logged")
		$run = Run($f, $workingdir, $show_flag)
		If @error Then Return

		Do
			Sleep(10)
		Until ProcessExists($run)

		$runtitle = _WinGetByPID($run)
		WinSetState($runtitle, "", @SW_HIDE)
		$TimerStart = TimerInit()

		; Size of extracted file(s)
		While ProcessExists($run)
			$size = Round((DirGetSize($outdir) - $initdirsize) / 1024 / 1024, 3)
			If $size > 0 Then
				If $TBgui Then GUICtrlSetData($TrayMsg_Status, $size & " MB")
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
EndFunc

; Run a program and return stdout/stderr stream
Func FetchStdout($f, $workingdir, $show_flag, $line = 0)
	Cout("Executing: " & $f)
	Global $run = 0, $return = ""
	$run = Run($f, $workingdir, $show_flag, $STDERR_MERGED)
	$runtitle = _WinGetByPID($run)
;~ 	Cout("PID: " & $run)
	Do
		Sleep(1)
		$return &= StdoutRead($run)
	Until @error
	Cout($return)
	$run = 0
	If $line <> 0 Then Return _StringGetLine($return, $line)
	Return $return
EndFunc

; Stop running helper process
Func KillHelper()
	If Not $run Then Return
	$runtitle = _WinGetByPID($run)

	If Not @error And Not StringIsSpace($runtitle) Then
		Cout("Runtitle: " & $runtitle)
		; Send SIGINT to console to terminate child processes
		WinActivate($runtitle)
		Send("^c")
		; Close console
		WinClose($runtitle)
	EndIf

	; Force termination if other close commands failed
	If ProcessExists($run) Then ProcessClose($run)
EndFunc

; Write data to stdout stream if enabled in options
Func Cout($Data)
	Local $Output = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & ":" & @MSEC & @TAB & $Data & @CRLF; & @CRLF
	If $Opt_ConsoleOutput == 1 Then ConsoleWrite($Output)
	$debug &= $Output
EndFunc

; Check for new version
Func CheckUpdate($silent = False)
	Local $return = 0, $found = False
	Cout("Checking for update")

	; Save date of last check for update
	SavePref('lastupdate', @YEAR & "/" & @MON & "/" & @MDAY)

	; Universal Extractor
	$return = _INetGetSource($updateURL & "?get=version&id=" & $ID)
	If @error Then
		If Not $silent Then MsgBox(262144 + 48, $title, t('UPDATE_FAILED'))
		Return
	EndIf

	Cout("Local: " & $version)
	Cout("Server: " & $return)

	If $return <> $version Then
		Cout("Update available")
		$found = True
		If Prompt(48 + 4, 'UPDATE_PROMPT', CreateArray($name, $version, $return), 0) Then
			$UEURL = _INetGetSource($updateURL & "?get=uniextract&version=" & $version & "&id=" & $ID)
			If @error Or $UEURL = "" Then
				If Not $silent Then MsgBox(262144 + 48, $title, t('UPDATE_FAILED'))
				Return
			EndIf
			$return = Download($UEURL)
			If $return == 0 Then Return
			$handle = FileOpen(@ScriptDir & "\Update.bat", 2)
			FileWrite($handle, "@ping -n 3 localhost> nul" & @CRLF & "taskkill -f -im UniExtract.exe" & @CRLF & '"' & @ScriptDir & "\bin\" & $OSArch & "\" & $7z & '" x -y -o"' & @ScriptDir & '" "' & $return & '"' & _
					@CRLF & @CRLF & "@ping -n 3 localhost> nul" & @CRLF & 'del "' & $return & '"' & @CRLF & 'start "" ".\' & @ScriptName & '" /afterupdate' & @CRLF & "del Update.bat" & @CRLF & "exit")
			FileClose($handle)
			Run(@ScriptDir & "\Update.bat", @ScriptDir)
			Exit
		EndIf
	EndIf

	; If ffmpeg.exe exists check for new ffmpeg version
	If FileExists(@ScriptDir & "\bin\" & $OSArch & "\" & $ffmpeg) Then
		; Determine FFmpeg version
		$return = FetchStdout(@ScriptDir & "\bin\" & $OSArch & "\" & $ffmpeg, @ScriptDir, @SW_HIDE)
		$ffmpegvers = _StringBetween($return, "ffmpeg version ", " Copyright")
		$return = _INetGetSource($updateURL & "?get=ffversion")
		; Download new
		If $return > $ffmpegvers[0] Then
			$found = True
			Cout("Found update for FFmpeg")
			If Prompt(48 + 4, 'UPDATE_PROMPT', CreateArray("FFmpeg", $ffmpegvers[0], $return), 0) Then GetFFmpeg()
		EndIf
	EndIf

	If $found = False And $silent = False Then MsgBox(262144 + 64, $name, t('UPDATE_CURRENT'))

	Cout("Check for updates finished")
EndFunc

; Perform special actions after update, e.g. delete files
Func _AfterUpdate()
	; Open most recent changelog
	$1 = FileGetTime("changelog_minor.txt", 0, 1)
	$2 = FileGetTime("changelog.txt", 0, 1)
	If $1 > $2 Then
		ShellExecute("changelog_minor.txt")
	Else
		ShellExecute("changelog.txt")
	EndIf

	; Remove unused files
	If FileExists(@ScriptDir & "\bin\plugins\kanal_for_Exeinfo.dll") Then FileDelete(@ScriptDir & "\bin\plugins\kanal_for_Exeinfo.dll")
	If FileExists(@ScriptDir & "\bin\tee.exe") Then FileDelete(@ScriptDir & "\bin\tee.exe")
	If FileExists(@ScriptDir & "\bin\BOOZ.EXE") Then FileDelete(@ScriptDir & "\bin\BOOZ.EXE")

	; Add new options to ini file (for options without corresponding GUI control)
	SavePref("unicodecheck", $checkUnicode)
	SavePref("filescanlogfile", $fileScanLogFile)
	SavePref("statusposx", $trayX)
	SavePref("statusposy", $trayY)
EndFunc   ;==>_AfterUpdate

; Download FFmpeg and move needed files to Universal Extractor directory
Func GetFFmpeg()
	$FFmpegURL = _INetGetSource($updateURL & "?get=ffmpeg&OSarch=" & @OSArch & "&id=" & $ID)
	$return = Download($FFmpegURL)
	If @error Then Return SetError(1, 0, 0)

	; Extract files, move them to scriptdir and delete files from tempdir
	Cout("Extracting FFmpeg")
	Local $success = RunWait($cmd & $7z & ' x "' & $return & '"', @TempDir)
	FileDelete($return)
	If $success <> 0 Then
		MsgBox(262144 + 48 + 1, $title, t('EXTRACT_FAILED', CreateArray($return, "7Zip")))
		Return SetError(1, 0, 0)
	EndIf

	Cout("Moving FFmpeg files")
	Local $dir = StringTrimRight($return, 3)
	FileMove($dir & "\bin\ffmpeg.exe", @ScriptDir & "\bin\" & $OSArch & "\ffmpeg.exe", 1)
	MoveFiles($dir & "\licenses", @ScriptDir & "\docs\FFmpeg", True)
	DirRemove($dir, 1)

	; Download license information
	If Not FileExists(@ScriptDir & "\docs\FFmpeg\FFmpeg_license.html") Then
		$return = Download("http://ffmpeg.org/legal.html")
		If Not @error Then FileMove($return, @ScriptDir & "\docs\FFmpeg\FFmpeg_license.html", 8 + 1)
	EndIf

	Cout("FFmpeg succesfully loaded")
	Return 1
EndFunc

; Download a file
Func Download($f)
	Cout("Downloading: " & $f)

	; Create GUI with progressbar
	Local $DownloadGUI = GUICreate(t('TERM_DOWNLOADING'), 466, 109, -1, -1, $WS_POPUPWINDOW, -1, $guimain)
	GUICtrlCreateLabel($f, 8, 16, 446, 17, $SS_CENTER)
	Local $DownloadProgress = GUICtrlCreateProgress(8, 46, 446, 25)
	GUISetState(@SW_SHOW)

	; Get file size
	Local $BytesReceived = 0
	Local $BytesNeeded = InetGetSize($f)
	Local $DownloadSizeLabel = GUICtrlCreateLabel($BytesReceived & "/" & $BytesNeeded & " kb", 8, 76, 446, 17, $SS_CENTER)

	; Download File
	Local $DownloadedFile = @TempDir & "\" & StringTrimLeft($f, StringInStr($f, "/", 0, -1))
	Local $Download = InetGet($f, $DownloadedFile, 1, 1)

	; Update progress bar
	While Not InetGetInfo($Download, 2)
		Sleep(50)
		If InetGetInfo($Download, 4) <> 0 Then
			Cout("Download failed")
			Prompt(48, 'DOWNLOAD_FAILED', $f, 0)
			GUIDelete($DownloadGUI)
			Return SetError(1, 0, 0)
		EndIf
		$BytesReceived = InetGetInfo($Download, 0)
		GUICtrlSetData($DownloadProgress, Int($BytesReceived / $BytesNeeded * 100))
		GUICtrlSetData($DownloadSizeLabel, $BytesReceived & "/" & $BytesNeeded & " kb")
	WEnd

	; Close GUI
	GUIDelete($DownloadGUI)
	Cout("Download finished")
	If Not FileExists($DownloadedFile) Then
		Cout("Downloaded file does not exist")
		Return SetError(1, 0, 0)
	EndIf
	Return $DownloadedFile
EndFunc

; ------------------------ Begin GUI Control Functions ------------------------

; Build and display GUI if necessary (moved to function to allow on-the-fly language changes)
Func CreateGUI()
	Cout("Creating main GUI")
	GUIRegisterMsg($WM_DROPFILES, "WM_DROPFILES_UNICODE_FUNC")

	; Create GUI
	If $StoreGUIPosition Then
		Global $guimain = GUICreate($title, 300, 135, $posx, $posy, -1, BitOR($WS_EX_ACCEPTFILES, $WS_EX_TOPMOST))
	Else
		Global $guimain = GUICreate($title, 300, 135, -1, -1, -1, BitOR($WS_EX_ACCEPTFILES, $WS_EX_TOPMOST))
	EndIf
	Local $dropzone = GUICtrlCreateLabel("", 0, 0, 300, 135)

	; Menu controls
	Local $filemenu = GUICtrlCreateMenu(t('MENU_FILE_LABEL'))
	Local $openitem = GUICtrlCreateMenuItem(t('MENU_FILE_OPEN_LABEL'), $filemenu)
	GUICtrlCreateMenuItem("", $filemenu)
	Global $keepopenitem = GUICtrlCreateMenuItem(t('MENU_FILE_KEEP_OPEN_LABEL'), $filemenu)
	GUICtrlCreateMenuItem("", $filemenu)
	Global $showitem = GUICtrlCreateMenuItem(t('MENU_FILE_SHOW_LABEL'), $filemenu)
	Global $clearitem = GUICtrlCreateMenuItem(t('MENU_FILE_CLEAR_LABEL'), $filemenu)
	GUICtrlCreateMenuItem("", $filemenu)
	Global $logitem = GUICtrlCreateMenuItem("DUMMY", $filemenu)
	GUICtrlCreateMenuItem("", $filemenu)
	Local $quititem = GUICtrlCreateMenuItem(t('MENU_FILE_QUIT_LABEL'), $filemenu)
	Local $editmenu = GUICtrlCreateMenu(t('MENU_EDIT_LABEL'))
	Global $keepitem = GUICtrlCreateMenuItem(t('MENU_EDIT_KEEP_LABEL'), $editmenu)
	Global $silentitem = GUICtrlCreateMenuItem(t('MENU_EDIT_SILENT_MODE_LABEL'), $editmenu)
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
	Local $forumitem = GUICtrlCreateMenuItem(t('MENU_HELP_FORUM_LABEL'), $helpmenu)
	GUICtrlCreateMenuItem("", $helpmenu)
	Local $statsitem = GUICtrlCreateMenuItem(t('MENU_HELP_STATS_LABEL'), $helpmenu)
	GUICtrlSetState(-1, $GUI_DISABLE)
	Local $programdiritem = GUICtrlCreateMenuItem(t('MENU_HELP_PROGDIR_LABEL'), $helpmenu)
	GUICtrlCreateMenuItem("", $helpmenu)
	Local $aboutitem = GUICtrlCreateMenuItem(t('MENU_HELP_ABOUT_LABEL'), $helpmenu)
	GUI_UpdateLogItem()

	; File controls
	Local $filelabel = GUICtrlCreateLabel(t('MAIN_FILE_LABEL'), 5, 4, -1, 15)
	Global $GUI_Main_Extract = GUICtrlCreateRadio(t('TERM_EXTRACT'), GetPos($guimain, $filelabel, 5), 3, Default, 15)
	Global $GUI_Main_Scan = GUICtrlCreateRadio(t('TERM_SCAN'), GetPos($guimain, $GUI_Main_Extract, 10), 3, Default, 15)

	If $extract Then
		GUICtrlSetState($GUI_Main_Extract, $GUI_CHECKED)
	Else
		GUICtrlSetState($GUI_Main_Scan, $GUI_CHECKED)
	EndIf

	If $history Then
		Global $filecont = GUICtrlCreateCombo("", 5, 20, 260, 20)
	Else
		Global $filecont = GUICtrlCreateInput("", 5, 20, 260, 20)
	EndIf
	Local $filebut = GUICtrlCreateButton("...", 270, 20, 25, 20)

	; Directory controls
	GUICtrlCreateLabel(t('MAIN_DEST_DIR_LABEL'), 5, 45, -1, 15)
	If $history Then
		Global $dircont = GUICtrlCreateCombo("", 5, 60, 260, 20)
	Else
		Global $dircont = GUICtrlCreateInput("", 5, 60, 260, 20)
	EndIf
	Local $dirbut = GUICtrlCreateButton("...", 270, 60, 25, 20)

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
	If $KeepOutdir Then
		GUICtrlSetState($keepitem, $GUI_CHECKED)
	Else
		GUICtrlSetState($keepitem, $GUI_UNCHECKED)
	EndIf
	If $KeepOpen Then
		GUICtrlSetState($keepopenitem, $GUI_CHECKED)
	Else
		GUICtrlSetState($keepopenitem, $GUI_UNCHECKED)
	EndIf
	If $silentmode Then
		GUICtrlSetState($silentitem, $GUI_CHECKED)
	Else
		GUICtrlSetState($silentitem, $GUI_UNCHECKED)
	EndIf
	If $batchEnabled = 0 Then
		GUICtrlSetState($showitem, $GUI_DISABLE)
		GUICtrlSetState($clearitem, $GUI_DISABLE)
	EndIf
	If $file <> "" Then
		FilenameParse($file)
		If $history Then
			$filelist = '|' & $file & '|' & ReadHist('file')
			GUICtrlSetData($filecont, $filelist, $file)
			$dirlist = '|' & $initoutdir & '|' & ReadHist('directory')
			GUICtrlSetData($dircont, $dirlist, $initoutdir)
		Else
			GUICtrlSetData($filecont, $file)
			GUICtrlSetData($dircont, $initoutdir)
		EndIf
		GUICtrlSetState($dircont, $GUI_FOCUS)
	ElseIf $history Then
		GUICtrlSetData($filecont, ReadHist('file'))
		GUICtrlSetData($dircont, ReadHist('directory'))
	EndIf

	; Set events
	GUISetOnEvent($GUI_EVENT_DROPPED, "GUI_Drop")
	GUICtrlSetOnEvent($filebut, "GUI_File")
	GUICtrlSetOnEvent($dirbut, "GUI_Directory")
	GUICtrlSetOnEvent($openitem, "GUI_File")
	GUICtrlSetOnEvent($keepopenitem, "GUI_KeepOpen")
	GUICtrlSetOnEvent($showitem, "GUI_Batch_Show")
	GUICtrlSetOnEvent($clearitem, "GUI_Batch_Clear")
	GUICtrlSetOnEvent($logitem, "GUI_DeleteLogs")
	GUICtrlSetOnEvent($keepitem, "GUI_KeepOutdir")
	GUICtrlSetOnEvent($GUI_Main_Extract, "GUI_ScanOnly")
	GUICtrlSetOnEvent($GUI_Main_Scan, "GUI_ScanOnly")
	GUICtrlSetOnEvent($silentitem, "GUI_Silent")
	GUICtrlSetOnEvent($contextitem, "GUI_ContextMenu")
	GUICtrlSetOnEvent($prefsitem, "GUI_Prefs")
	GUICtrlSetOnEvent($pluginsitem, "GUI_Plugins")
	GUICtrlSetOnEvent($feedbackitem, "GUI_Feedback")
	GUICtrlSetOnEvent($firststartitem, "GUI_FirstStart")
	GUICtrlSetOnEvent($updateitem, "CheckUpdate")
	GUICtrlSetOnEvent($webitem, "GUI_Website")
	GUICtrlSetOnEvent($web2item, "GUI_Website2")
	GUICtrlSetOnEvent($gititem, "GUI_Website_Github")
	GUICtrlSetOnEvent($forumitem, "GUI_Forum")
	GUICtrlSetOnEvent($statsitem, "GUI_Stats")
	GUICtrlSetOnEvent($programdiritem, "GUI_ProgDir")
	GUICtrlSetOnEvent($aboutitem, "GUI_About")
	GUICtrlSetOnEvent($ok, "GUI_Ok")
	GUICtrlSetOnEvent($cancel, "GUI_Exit")
	GUICtrlSetOnEvent($BatchBut, "GUI_Batch")
	GUICtrlSetOnEvent($quititem, "GUI_Exit")
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_Exit")

	GetBatchQueue()

	; Display GUI and wait for action
	GUISetState(@SW_SHOW)
EndFunc   ;==>CreateGUI

; Return control width (for dynamic positioning)
Func GetPos($hGUI, $hControl, $iOffset = 0)
	$return = ControlGetPos($hGUI, '', $hControl)
	If @error Then Return SetError(1, '', $iOffset)
	Return $return[0] + $return[2] + $iOffset
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
EndFunc   ;==>GUI_File

; Prompt user for directory
Func GUI_Directory()
	If FileExists(GUICtrlRead($dircont)) Then
		$defdir = GUICtrlRead($dircont)
	ElseIf FileExists(GUICtrlRead($filecont)) Then
		$defdir = StringLeft(GUICtrlRead($filecont), StringInStr(GUICtrlRead($filecont), '\', 0, -1) - 1)
	Else
		$defdir = ''
	EndIf
	$outdir = FileSelectFolder(t('EXTRACT_TO'), "", 3, $defdir, $guimain)
	If Not @error Then
		If $history Then
			$dirlist = '|' & $outdir & '|' & ReadHist('directory')
			GUICtrlSetData($dircont, $dirlist, $outdir)
		Else
			GUICtrlSetData($dircont, $outdir)
		EndIf
	EndIf
EndFunc   ;==>GUI_Directory

; Option to keep the destination directory
Func GUI_KeepOutdir()
	If BitAND(GUICtrlRead($keepitem), $GUI_CHECKED) = $GUI_CHECKED Then
		GUICtrlSetState($keepitem, $GUI_UNCHECKED)
		$KeepOutdir = 0
	Else
		GUICtrlSetState($keepitem, $GUI_CHECKED)
		$KeepOutdir = 1
	EndIf

	SavePref('keepoutputdir', $KeepOutdir)
EndFunc   ;==>GUI_KeepOutdir

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
EndFunc   ;==>GUI_KeepOpen

; Build and display preferences GUI
Func GUI_Prefs()
	Cout("Creating preferences GUI")
	; Create GUI
	If $guimain Then
		Global $guiprefs = GUICreate(t('PREFS_TITLE_LABEL'), 250, 430, -1, -1, -1, -1, $guimain)
	Else
		Global $guiprefs = GUICreate(t('PREFS_TITLE_LABEL'), 250, 430)
	EndIf

	; Universal prefs box
	GUICtrlCreateGroup(t('PREFS_UNIEXTRACT_OPTS_LABEL'), 5, 5, 240, 122)

	; History option
	Global $historyopt = GUICtrlCreateCheckbox(t('PREFS_HISTORY_LABEL'), 10, 20, -1, 20)

	; Language controls
	Local $langlabel = GUICtrlCreateLabel(t('PREFS_LANG_LABEL'), 10, 45, -1, 15)
	Local $langselectpos = GetPos($guiprefs, $langlabel, -8)
	Global $langselect = GUICtrlCreateCombo("", $langselectpos, 42, 245 - $langselectpos - 4, 20 * CharCount($langlist, '|'), $CBS_DROPDOWNLIST)

	; Timeout controls
	Local $TimeoutLabel = GUICtrlCreateLabel(t('PREFS_TIMEOUT_LABEL'), 10, 72, -1, 15)
	Global $TimeoutCont = GUICtrlCreateInput($Timeout / 1000, GetPos($guiprefs, $TimeoutLabel), 70, 35, 20, $ES_NUMBER)
	GUICtrlCreateLabel(t('PREFS_SECONDS_LABEL'), GetPos($guiprefs, $TimeoutCont, 5), 72, -1, 15)

	; Update intervall controls
	GUICtrlCreateLabel(t('PREFS_UPDATEINTERVAL_LABEL'), 10, 102, -1, 15)
	Global $IntervalCont = GUICtrlCreateInput($updateinterval, GetPos($guiprefs, $TimeoutLabel), 100, 35, 20, $ES_NUMBER)
	GUICtrlCreateLabel(t('PREFS_DAYS_LABEL'), GetPos($guiprefs, $IntervalCont, 5), 102, -1, 15)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; Format-specific preferences
	GUICtrlCreateGroup(t('PREFS_FORMAT_OPTS_LABEL'), 5, 130, 240, 220)
	Global $warnexecuteopt = GUICtrlCreateCheckbox(t('PREFS_WARN_EXECUTE_LABEL'), 10, 145, -1, 20)
	Global $freeSpaceCheckOpt = GUICtrlCreateCheckbox(t('PREFS_CHECK_FREE_SPACE_LABEL'), 10, 165, -1, 20)
	Global $unicodecheckopt = GUICtrlCreateCheckbox(t('PREFS_CHECK_UNICODE_LABEL'), 10, 185, -1, 20)
	Global $appendextopt = GUICtrlCreateCheckbox(t('PREFS_APPEND_EXT_LABEL'), 10, 205, -1, 20)
	Global $NoBoxOpt = GUICtrlCreateCheckbox(t('PREFS_HIDE_STATUS_LABEL'), 10, 225, -1, 20)
	Global $OpenOutDirOpt = GUICtrlCreateCheckbox(t('PREFS_OPEN_FOLDER_LABEL'), 10, 245, -1, 20)
	Global $FeedbackPromptOpt = GUICtrlCreateCheckbox(t('PREFS_FEEDBACK_PROMPT_LABEL'), 10, 265, -1, 20)
	Global $StoreGUIPositionOpt = GUICtrlCreateCheckbox(t('PREFS_WINDOW_POSITION_LABEL'), 10, 285, -1, 20)
	Global $CheckGameOpt = GUICtrlCreateCheckbox(t('PREFS_CHECK_GAME_LABEL'), 10, 305, -1, 20)
	Global $LogOpt = GUICtrlCreateCheckbox(t('PREFS_LOG_LABEL'), 10, 325, -1, 20)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; Source file options
	GUICtrlCreateGroup(t('PREFS_SOURCE_FILES_LABEL'), 5, 355, 240, 40)
	$DeleteOrigFileOpt[0] = GUICtrlCreateRadio(t('PREFS_SOURCE_FILES_OPT_KEEP'), 10, 370)
	$DeleteOrigFileOpt[1] = GUICtrlCreateRadio(t('PREFS_SOURCE_FILES_OPT_DELETE'), GetPos($guiprefs, $DeleteOrigFileOpt[0], 20), 370)
	$DeleteOrigFileOpt[2] = GUICtrlCreateRadio(t('PREFS_SOURCE_FILES_OPT_ASK'), GetPos($guiprefs, $DeleteOrigFileOpt[1], 20), 370)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; Buttons
	Local $prefsok = GUICtrlCreateButton(t('OK_BUT'), 55, 403, 60, 20)
	Local $prefscancel = GUICtrlCreateButton(t('CANCEL_BUT'), 135, 403, 60, 20)

	; Set properties
	GUICtrlSetState($prefsok, $GUI_DEFBUTTON)
	If $history Then GUICtrlSetState($historyopt, $GUI_CHECKED)
	If $warnexecute Then GUICtrlSetState($warnexecuteopt, $GUI_CHECKED)
	If $freeSpaceCheck Then GUICtrlSetState($freeSpaceCheckOpt, $GUI_CHECKED)
	If $checkUnicode Then GUICtrlSetState($unicodecheckopt, $GUI_CHECKED)
	If $appendext Then GUICtrlSetState($appendextopt, $GUI_CHECKED)
	If $NoBox Then GUICtrlSetState($NoBoxOpt, $GUI_CHECKED)
	If $OpenOutDir Then GUICtrlSetState($OpenOutDirOpt, $GUI_CHECKED)
	If $FB_ask Then GUICtrlSetState($FeedbackPromptOpt, $GUI_CHECKED)
	If $StoreGUIPosition Then GUICtrlSetState($StoreGUIPositionOpt, $GUI_CHECKED)
	If $CheckGame Then GUICtrlSetState($CheckGameOpt, $GUI_CHECKED)
	If $Log Then GUICtrlSetState($LogOpt, $GUI_CHECKED)
	GUICtrlSetState($DeleteOrigFileOpt[$DeleteOrigFile], $GUI_CHECKED)

	GUICtrlSetData($langselect, $langlist, $language)

	; Set events
	GUICtrlSetOnEvent($prefsok, "GUI_Prefs_Ok")
	GUICtrlSetOnEvent($prefscancel, "GUI_Prefs_Exit")
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_Prefs_Exit")

	; Display GUI and wait for action
	GUISetState(@SW_SHOW)
EndFunc   ;==>GUI_Prefs

; Exit preferences GUI if Cancel clicked or window closed
Func GUI_Prefs_Exit()
	GUIDelete($guiprefs)
	$guiprefs = False
	Cout("Closing preferences GUI")
EndFunc   ;==>GUI_Prefs_Exit

; Exit preferences GUI if OK clicked
Func GUI_Prefs_OK()
	; universal preferences
	$redrawgui = False

	If GUICtrlRead($historyopt) == $GUI_CHECKED Then
		If $history == 0 Then
			$history = 1
			$redrawgui = True
		EndIf
	Else
		If $history == 1 Then
			$history = 0
			If $globalprefs Then
				IniDelete($prefs, "File History")
				IniDelete($prefs, "Directory History")
			Else
				RegDelete($reg & '\File')
				RegDelete($reg & '\Directory')
			EndIf
			$redrawgui = True
		EndIf
	EndIf
	If $language <> GUICtrlRead($langselect) Then
		$language = GUICtrlRead($langselect)
		$redrawgui = True
	EndIf
	If $Timeout / 1000 <> GUICtrlRead($TimeoutCont) And Int(GUICtrlRead($TimeoutCont)) > 9 Then $Timeout = Int(GUICtrlRead($TimeoutCont)) * 1000

	If $updateinterval <> GUICtrlRead($IntervalCont) Then $updateinterval = GUICtrlRead($IntervalCont)

	; format-specific preferences
	If GUICtrlRead($warnexecuteopt) == $GUI_CHECKED Then
		$warnexecute = 1
	Else
		$warnexecute = 0
	EndIf
	If GUICtrlRead($unicodecheckopt) == $GUI_CHECKED Then
		$checkUnicode = 1
	Else
		$checkUnicode = 0
	EndIf

	If GUICtrlRead($freeSpaceCheckOpt) == $GUI_CHECKED Then
		$freeSpaceCheck = 1
	Else
		$freeSpaceCheck = 0
	EndIf
	If GUICtrlRead($appendextopt) == $GUI_CHECKED Then
		$appendext = 1
	Else
		$appendext = 0
	EndIf
	If GUICtrlRead($NoBoxOpt) == $GUI_CHECKED Then
		$NoBox = 1
		TrayItemSetState($Tray_Statusbox, $TRAY_CHECKED)
	Else
		$NoBox = 0
		TrayItemSetState($Tray_Statusbox, $TRAY_UNCHECKED)
	EndIf
	If GUICtrlRead($OpenOutDirOpt) == $GUI_CHECKED Then
		$OpenOutDir = 1
	Else
		$OpenOutDir = 0
	EndIf
	If GUICtrlRead($FeedbackPromptOpt) == $GUI_CHECKED Then
		$FB_ask = 1
	Else
		$FB_ask = 0
	EndIf
	If GUICtrlRead($StoreGUIPositionOpt) == $GUI_CHECKED Then
		$StoreGUIPosition = 1
	Else
		$StoreGUIPosition = 0
	EndIf
	If GUICtrlRead($LogOpt) == $GUI_CHECKED Then
		$Log = 1
	Else
		$Log = 0
	EndIf
	For $i = 0 To 2
		If GUICtrlRead($DeleteOrigFileOpt[$i]) == $GUI_CHECKED Then $DeleteOrigFile = $i
	Next
	WritePrefs()
	GUIDelete($guiprefs)
	$guiprefs = False
	If $redrawgui Then
		GUIDelete($guimain)
		CreateGUI()
	EndIf
EndFunc   ;==>GUI_Prefs_OK

; Handle click on OK
Func GUI_OK()
	If Not GUI_OK_Set(True) Then Return
	GUIGetPosition()
	GUIDelete($guimain)
	$guimain = False
	Cout("Closing main GUI")
EndFunc   ;==>GUI_OK

; Set file to extract and target directory
Func GUI_OK_Set($Msg = False)
	$file = EnvParse(GUICtrlRead($filecont))
	If FileExists($file) Then
		If EnvParse(GUICtrlRead($dircont)) == "" Then
			$outdir = '/sub'
		Else
			$outdir = EnvParse(GUICtrlRead($dircont))
		EndIf
		Return 1
	ElseIf $Msg Then
		If $file <> '' Then $file &= ' ' & t('DOES_NOT_EXIST')
		MsgBox(262144 + 48, $title, t('INVALID_FILE_SELECTED', $file))
	EndIf
	Return 0
EndFunc   ;==>GUI_OK_Set

; Add file to batch queue
Func GUI_Batch()
	If GUI_OK_Set() Then
		AddToBatch()
		GUICtrlSetData($filecont, "")
		If Not $KeepOutdir Then GUICtrlSetData($dircont, "")
	Else
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

	terminate("batch", '', '')
EndFunc   ;==>GUI_Batch_OK

; Display batch queue and allow changes
Func GUI_Batch_Show()
	Cout("Opening batch queue edit GUI")
	$GUI_Batch = GUICreate($name, 418, 267, 476, 262, -1, -1, $guimain)
	$GUI_Batch_List = GUICtrlCreateList("", 8, 8, 401, 201)
	GUICtrlSetData(-1, _ArrayToString($queueArray, "|", 1))
	$GUI_Batch_OK = GUICtrlCreateButton("&OK", 40, 225, 75, 25)
	$GUI_Batch_Cancel = GUICtrlCreateButton("&Cancel", 171, 225, 75, 25)
	$GUI_Batch_Delete = GUICtrlCreateButton("Delete", 304, 224, 73, 25)
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
		EndSwitch
	WEnd

	GUIDelete($GUI_Batch)
	Opt("GUIOnEventMode", 1)
EndFunc   ;==>GUI_Batch_Show

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
		$filelist = '|' & $file & '|' & ReadHist('file')
		GUICtrlSetData($filecont, $filelist, $file)
	Else
		GUICtrlSetData($filecont, $file)
	EndIf

	If GUICtrlRead($dircont) == "" Or Not $KeepOutdir Then
		FilenameParse($file)
		If $history Then
			$dirlist = '|' & $initoutdir & '|' & ReadHist('directory')
			GUICtrlSetData($dircont, $dirlist, $initoutdir)
		Else
			GUICtrlSetData($dircont, $initoutdir)
		EndIf
	EndIf
EndFunc   ;==>GUI_Drop_Parse

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
Func GUI_Feedback($Type = "", $file = "", $Output = "")
	Cout("Creating feedback GUI")
	Global $FB_finish = False
	; Create Feedback GUI
	If $guimain Then
		Global $FB_GUI = GUICreate(t('FEEDBACK_TITLE_LABEL'), 251, 370, -1, -1, BitOR($WS_SIZEBOX, $WS_SYSMENU), -1, $guimain)
	Else
		Global $FB_GUI = GUICreate(t('FEEDBACK_TITLE_LABEL'), 251, 370, -1, -1, BitOR($WS_SIZEBOX, $WS_SYSMENU), -1)
	EndIf

	GUICtrlCreateLabel(t('FEEDBACK_TYPE_LABEL'), 8, 8, -1, 15)
	If $Type == "" Then
		Global $FB_TypeCont = GUICtrlCreateCombo(t('FEEDBACK_TYPE_STANDARD'), 8, 24, 105, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
	Else
		Global $FB_TypeCont = GUICtrlCreateCombo($Type, 8, 24, 105, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
		GUICtrlSetData(-1, t('FEEDBACK_TYPE_STANDARD'))
	EndIf
	GUICtrlSetData(-1, t('FEEDBACK_TYPE_OPTIONS'))
	;GUICtrlSetData(-1, $Type)
	GUICtrlCreateLabel(t('FEEDBACK_SYSINFO_LABEL'), 120, 8, -1, 15)
	Global $FB_SysCont = GUICtrlCreateInput(@OSVersion & " " & @OSArch & " " & @OSServicePack & ", Lang: " & @OSLang & ", UE: " & $language, 120, 24, 121, 21)
	GUICtrlCreateLabel(t('FEEDBACK_FILE_LABEL'), 8, 56, -1, 15)
	GUICtrlSetTip(-1, t('FEEDBACK_FILE_TOOLTIP'), "", 0, 1)
	Global $FB_FileCont = GUICtrlCreateInput($file, 8, 72, 233, 21)
	GUICtrlSetTip(-1, t('FEEDBACK_FILE_TOOLTIP'), "", 0, 1)
	GUICtrlCreateLabel(t('FEEDBACK_OUTPUT_LABEL'), 8, 104, -1, 15)
	GUICtrlSetTip(-1, t('FEEDBACK_OUTPUT_TOOLTIP'), "", 0, 1)
	Global $FB_OutputCont = GUICtrlCreateEdit("", 8, 120, 233, 49, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_WANTRETURN, $WS_VSCROLL))
	GUICtrlSetData(-1, $Output)
	GUICtrlSetTip(-1, t('FEEDBACK_OUTPUT_TOOLTIP'), "", 0, 1)
	GUICtrlCreateLabel(t('FEEDBACK_MESSAGE_LABEL'), 8, 176, -1, 15)
	GUICtrlSetTip(-1, t('FEEDBACK_MESSAGE_TOOLTIP'), "", 0, 1)
	Global $FB_MessageCont = GUICtrlCreateEdit("", 8, 192, 233, 73, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_WANTRETURN, $WS_VSCROLL))
	GUICtrlSetTip(-1, t('FEEDBACK_MESSAGE_TOOLTIP'), "", 0, 1)
	GUICtrlCreateLabel(t('FEEDBACK_EMAIL_LABEL'), 8, 272, -1, 15)
	GUICtrlSetTip(-1, t('FEEDBACK_EMAIL_TOOLTIP'), "", 0, 1)
	Global $FB_MailCont = GUICtrlCreateInput("", 8, 288, 233, 21)
	GUICtrlSetTip(-1, t('FEEDBACK_EMAIL_TOOLTIP'), "", 0, 1)
	$FB_Send = GUICtrlCreateButton(t('SEND_BUT'), 55, 318, 60, 20)
	$FB_Cancel = GUICtrlCreateButton(t('CANCEL_BUT'), 135, 318, 60, 20)
	; Dummy for accelerator key to activate
	$hSelectAll = GUICtrlCreateDummy()
	GUISetState(@SW_SHOW)
	GUICtrlSetState($FB_MessageCont, $GUI_FOCUS)
	Local $accelKeys[1][2] = [["^a", $hSelectAll]]
	GUISetAccelerators($accelKeys)

	; Set events
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_Feedback_Exit")
	GUICtrlSetOnEvent($FB_Cancel, "GUI_Feedback_Exit")
	GUICtrlSetOnEvent($FB_Send, "GUI_Feedback_Send")
	GUICtrlSetOnEvent($hSelectAll, "GUI_Edit_SelectAll")
EndFunc   ;==>GUI_Feedback

; Exit feedback GUI if OK clicked
Func GUI_Feedback_Send()
	$FB_Type = _GUICtrlComboBox_GetCurSel($FB_TypeCont)
	$FB_Sys = GUICtrlRead($FB_SysCont)
	$FB_File = GUICtrlRead($FB_FileCont)
	$FB_Output = GUICtrlRead($FB_OutputCont)
	$FB_Message = GUICtrlRead($FB_MessageCont)
	$FB_Mail = GUICtrlRead($FB_MailCont)

	If $FB_File = "" And $FB_Output = "" And $FB_Message = "" Then
		MsgBox(262144 + 16, $name, t('FEEDBACK_EMPTY'))
		Return
	EndIf

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
			 & @CRLF & "System Information: " & $FB_Sys & @CRLF & @CRLF & "Sample File: " & $FB_File & @CRLF _
			 & @CRLF & $name & " Output:" & @CRLF & $FB_Output & @CRLF & @CRLF & _
			"------------------------------------------------------------------------------------------------" _
			 & @CRLF & $FB_Message & @CRLF & @CRLF & "Sent by: " & @CRLF & $ID

	If StringInStr($FB_Mail, "@") Then $FB_Text &= @CRLF & $FB_Mail

	Const $boundary = "--UniExtractLog"

	$http = ObjCreate("winhttp.winhttprequest.5.1")
	$http.Open("POST", $supportURL, False)
	$http.SetRequestHeader("Content-Type", "multipart/form-data; boundary=" & StringTrimLeft($boundary, 2))

	Local $Data = $boundary & @CRLF & 'Content-Disposition: form-data; name="file"; filename="UE_Feedback"' & @CRLF & 'Content-Type: text/plain' & @CRLF & @CRLF & $FB_Text & @CRLF & $boundary & @CRLF & 'Content-Disposition: form-data; name="id"' & @CRLF & @CRLF & $ID & @CRLF & $boundary & '--'

	$http.Send($Data)
	$http.WaitForResponse()
	Local $sResponse = $http.ResponseText()

	_DeleteTrayMessageBox()

	If $sResponse = "1" Then
		Cout("Feedback successfully sent")
		MsgBox(262144, $title, t('FEEDBACK_SUCCESS'))
	Else
		Cout("Error sending feedback")
		MsgBox(262144+16, $title, t('FEEDBACK_ERROR'))
	EndIf

	GUISetState(@SW_SHOW, $guimain)
	$FB_finish = True
EndFunc   ;==>GUI_Feedback_Send

; Exit feedback GUI if Cancel clicked or window closed
Func GUI_Feedback_Exit()
	GUIDelete($FB_GUI)
	$FB_finish = True
	Cout("Closing feedback GUI")
EndFunc   ;==>GUI_Feedback_Exit

; Due to a bug in the Windows API, ctrl+a does not work for edit controls
; Workaround by Zedna (http://www.autoitscript.com/forum/topic/97473-hotkey-ctrla-for-select-all-in-the-selected-edit-box/#entry937287)
Func GUI_Edit_SelectAll()
	$hWnd = _WinAPI_GetFocus()
	$class = _WinAPI_GetClassName($hWnd)
	If $class = 'Edit' Then _GUICtrlEdit_SetSel($hWnd, 0, -1)
EndFunc   ;==>GUI_Edit_SelectAll

; Saves current position of main GUI
Func GUIGetPosition()
	If Not $StoreGUIPosition Then Return
	$pos = WinGetPos($guimain)
	SavePref('posx', $pos[0])
	SavePref('posy', $pos[1])
EndFunc   ;==>GUIGetPosition

; Set minimal size of feedback GUI for resizing
Func GUI_WM_GETMINMAXINFO($hWnd, $Msg, $wParam, $lParam)
	$tagMaxinfo = DllStructCreate("int;int;int;int;int;int;int;int;int;int", $lParam)
	DllStructSetData($tagMaxinfo, 7, 270) ; min width
	DllStructSetData($tagMaxinfo, 8, 380) ; min height
	;DllStructSetData($tagMaxinfo,  9, ) ; max width
	;DllStructSetData($tagMaxinfo, 10, )  ; max height
	Return 0
EndFunc   ;==>GUI_WM_GETMINMAXINFO

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
	Global $CM_Checkbox[4], $CM_Picture = False

	If $guimain Then
		Global $CM_GUI = GUICreate(t('PREFS_TITLE_LABEL'), 450, 600, -1, -1, -1, -1, $guimain)
	Else
		Global $CM_GUI = GUICreate(t('PREFS_TITLE_LABEL'), 450, 600)
	EndIf

	GUICtrlCreateGroup(t('CONTEXT_ENTRIES_LABEL'), 5, 0, 440, 470)
	Global $CM_Checkbox_enabled = GUICtrlCreateCheckbox(t('CONTEXT_ENABLED_LABEL'), 24, 16, -1, 17)
	Global $CM_Checkbox_allusers = GUICtrlCreateCheckbox(t('CONTEXT_ALL_USERS_LABEL'), GetPos($CM_GUI, $CM_Checkbox_enabled, 25), 16, -1, 17)
	GUICtrlSetState(-1, $GUI_DISABLE)
	Global $CM_Simple_Radio = GUICtrlCreateRadio(t('CONTEXT_SIMPLE_RADIO'), 96, 48, 145, 17)
	Global $CM_Cascading_Radio = GUICtrlCreateRadio(t('CONTEXT_CASCADING_RADIO'), 296, 48, 137, 17)
	$CM_Checkbox[0] = GUICtrlCreateCheckbox(t('EXTRACT_FILES'), 25, 426, 113, 17)
	$CM_Checkbox[1] = GUICtrlCreateCheckbox(t('EXTRACT_HERE'), GetPos($CM_GUI, $CM_Checkbox[0], 150), 426, 113, 17)
	$CM_Checkbox[2] = GUICtrlCreateCheckbox(t('EXTRACT_SUB'), 25, 446, 129, 17)
	$CM_Checkbox[3] = GUICtrlCreateCheckbox(t('SCAN_FILE'), GetPos($CM_GUI, $CM_Checkbox[0], 150), 446, 113, 17)
	Global $CM_Picture = GUICtrlCreatePic("", 55, 75, 0, 0, -1, $WS_EX_LAYERED)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	GUICtrlCreateGroup(t('CONTEXT_FILE_ASSOC_LABEL'), 5, 475, 440, 80)
	Global $CM_Checkbox_add = GUICtrlCreateCheckbox(t('CONTEXT_ENABLED_LABEL'), 24, 495, -1, 17)
	Global $CM_Checkbox_allusers2 = GUICtrlCreateCheckbox(t('CONTEXT_ALL_USERS_LABEL'), GetPos($CM_GUI, $CM_Checkbox_enabled, 25), 495, -1, 17)
	GUICtrlSetState(-1, $GUI_DISABLE)
	Global $CM_add_input = GUICtrlCreateInput("", 24, 520, 401, 21)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	Local $CM_OK = GUICtrlCreateButton(t('OK_BUT'), 112, 565, 89, 25)
	Local $CM_Cancel = GUICtrlCreateButton(t('CANCEL_BUT'), 232, 565, 89, 25)
	GUICtrlSetState($CM_Simple_Radio, $GUI_CHECKED)
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_ContextMenu_Exit")
	GUICtrlSetOnEvent($CM_Cancel, "GUI_ContextMenu_Exit")
	GUICtrlSetOnEvent($CM_OK, "GUI_ContextMenu_OK")
	GUICtrlSetOnEvent($CM_Checkbox_enabled, "GUI_ContextMenu_activate")
	GUICtrlSetOnEvent($CM_Checkbox_add, "GUI_ContextMenu_activate")

	;_ArrayDisplay($CM_Shells)

	; Check which commands are activated
	For $i = 0 To 3
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

	; Disable Cascading context menu for non Win 7 / Win 8 users as it is not supported
	If @OSVersion = "WIN_7" Or @OSVersion = "WIN_8" Or @OSVersion = "WIN_81" Then
		Global $win7 = True
		; Check if Cascading context menu entries are enabled
		For $i = 0 To 3
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
		Global $win7 = False
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
	Sleep(100)
	GUISetState(@SW_HIDE)

	; Remove old associations
	GUI_ContextMenu_remove()

	;If NOT RegExists($reguser) Then RegWrite($reguser)

	Cout("Registering context menu entries")
	If GUICtrlRead($CM_Checkbox_enabled) == $GUI_CHECKED Then

		; Select registry key
		If GUICtrlRead($CM_Checkbox_allusers) == $GUI_CHECKED Then
			Global $reguser = $regall
		Else
			Global $reguser = $regcurrent
		EndIf

		; simple
		If GUICtrlRead($CM_Simple_Radio) == $GUI_CHECKED Then
			For $i = 0 To 3
				$command = '"' & @ScriptFullPath & '" "%1"' & $CM_Shells[$i][1]
				If GUICtrlRead($CM_Checkbox[$i]) == $GUI_CHECKED Then
					RegWrite($reguser & $CM_Shells[$i][0], "", "REG_SZ", $CM_Shells[$i][2])
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

			For $i = 0 To 3
				$command = '"' & @ScriptFullPath & '" "%1"' & $CM_Shells[$i][1]
				If GUICtrlRead($CM_Checkbox[$i]) == $GUI_CHECKED Then
					RegWrite($reguser & "Uniextract\Shell\" & $CM_Shells[$i][0], "", "REG_SZ", $CM_Shells[$i][2])

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
		$return = MsgBox(262144 + 48 + 4, $name, t('CONTEXT_DANGEROUS'))
		If $return == 6 And ($addassocenabled == 0 Or ($addassocenabled = 1 And $addassoc <> GUICtrlRead($CM_add_input))) Then GUI_ContextMenu_fileassoc(1)
	Else
		If $addassocenabled Then GUI_ContextMenu_fileassoc(0)
	EndIf
	GUIDelete($CM_GUI)
EndFunc   ;==>GUI_ContextMenu_OK

; (De)activate controls if enabled but not checked
Func GUI_ContextMenu_activate()
	; Set state according to main enable checkbox
	If GUICtrlRead($CM_Checkbox_enabled) = $GUI_CHECKED Then
		Local $state = $GUI_ENABLE
	Else
		Local $state = $GUI_DISABLE
	EndIf

	If IsAdmin() Then GUICtrlSetState($CM_Checkbox_allusers, $state)
	GUICtrlSetState($CM_Simple_Radio, $state)

	For $i = 0 To 3
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
	For $i = 0 To 3
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

; Close context menu GUI
Func GUI_ContextMenu_Exit()
	Cout("Closing context menu GUI")
	GUIDelete($CM_GUI)
EndFunc   ;==>GUI_ContextMenu_Exit

; Perform special actions if Universal Extractor is started the first time
Func GUI_FirstStart()
	Cout("Creating first start assistant")
	GUISetState(@SW_HIDE, $guimain)
	; Create GUI
	Global $FS_GUI = GUICreate($title, 504, 387)
	GUICtrlCreatePic(".\support\Icons\uniextract_inno.bmp", 8, 312, 65, 65)
	GUICtrlCreateLabel($name, 8, 8, 488, 60, $SS_CENTER)
	GUICtrlSetFont(-1, 24, 800, 0, "MS Sans Serif")
	GUICtrlCreateLabel(t('FIRSTSTART_TITLE'), 8, 50, 488, 60, $SS_CENTER)
	GUICtrlSetFont(-1, 14, 800, 0, "MS Sans Serif")
	Global $FS_Section = GUICtrlCreateLabel("", 16, 85, 382, 28)
	GUICtrlSetFont(-1, 14, 800, 4, "MS Sans Serif")
	Global $FS_Next = GUICtrlCreateButton(t('NEXT_BUT'), 296, 344, 89, 25)
	Local $FS_Cancel = GUICtrlCreateButton(t('CANCEL_BUT'), 400, 344, 89, 25)
	Global $FS_Prev = GUICtrlCreateButton(t('PREV_BUT'), 192, 344, 89, 25)
	GUICtrlSetState(-1, $GUI_HIDE)
	Global $FS_Button = GUICtrlCreateButton("", 187, 260, 129, 41)
	Global $FS_Text = GUICtrlCreateLabel("", 16, 120, 468, 125)
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
		Exit 99
	EndIf
	Global $FS_Texts[UBound($FS_Sections)] = [ _
				"", t('FIRSTSTART_PAGE1'), t('FIRSTSTART_PAGE2'), t('FIRSTSTART_PAGE3'), t('FIRSTSTART_PAGE4'), _
				t('FIRSTSTART_PAGE5', @ScriptDir & "\bin"), t('FIRSTSTART_PAGE6') _
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
EndFunc   ;==>GUI_FirstStart_Prev

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
EndFunc   ;==>GUI_FirstStart_Next

; Load a page of the first start GUI
Func GUI_FirstStart_ShowPage()
	GUICtrlSetData($FS_Progress, $page & "/" & $FS_Sections[0])
	GUICtrlSetData($FS_Section, $FS_Sections[$page])
	GUICtrlSetData($FS_Text, $FS_Texts[$page])
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
			GUICtrlSetData($FS_Button, t('TERM_DOWNLOAD'))
			GUICtrlSetOnEvent($FS_Button, "GetFFmpeg")
		Case 6
			GUICtrlSetState($FS_Button, $GUI_HIDE)
			GUICtrlSetPos($FS_Text, 16, 120, 468, 170)
		Case Else
			GUICtrlSetState($FS_Button, $GUI_HIDE)
	EndSwitch
EndFunc   ;==>GUI_FirstStart_ShowPage

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
	; executable|name|description|filetypes|url
	Local $aPluginInfo[8][5] = [ _
		[$arc_conv, 'arc_conv', t('PLUGIN_ARC_CONV'), 'nsa, rgss2a, rgssad, wolf, xp3, ypf', 'http://honyaku-subs.ru/forums/viewtopic.php?f=17&t=470'], _
		[$thinstall, 'h4sh3m Virtual Apps Dependency Extractor', t('PLUGIN_THINSTALL'), 'exe (Thinstall)', 'http://hashem20.persiangig.com/crack%20tools/Extractor.rar'], _
		[$iscab, 'iscab', t('PLUGIN_ISCAB'), 'cab', False], _
		[$rgss3, 'RPGMaker Decrypter', t('PLUGIN_RPGMAKER'), 'rgss3a', 'http://cs10.userfiles.me/f/0/1411802749/48096296/0/a18c40637226ca066f472fc6d69fd877/RPGDecrypter-spaces.ru.exe'], _
		[$unreal, 'Unreal Engine package extractor', t('PLUGIN_UNREAL'), 'u, uax, upk', 'http://www.gildor.org/downloads'], _
		[$dcp, 'WinterMute Engine Unpacker', t('PLUGIN_WINTERMUTE'), 'dcp', 'http://forum.xentax.com/viewtopic.php?f=32&t=9625'], _
		[$crage, 'Crass/Crage', t('PLUGIN_CRAGE'), 'exe (Livemaker)', 'http://tlwiki.org/images/8/8a/Crass-0.4.14.0.bin.7z'], _
		[$faad, 'FAAD2', t('PLUGIN_FAAD'), 'aac', 'http://www.rarewares.org/files/aac/faad2-20100614.zip'] _
	]
	Local Const $sSupportedFileTypes = t('PLUGIN_SUPPORTED_FILETYPES')
	Local $current = -1
;~ 	_ArrayDisplay($aPluginInfo)

	$GUI_Plugins = GUICreate($name, 410, 167, -1, -1, -1, -1, $guimain)
	$GUI_Plugins_List = GUICtrlCreateList("", 8, 8, 209, 149)
	GUICtrlSetData(-1, _ArrayToString($aPluginInfo, "|", 0, 0, "|", 1, 1))
	$GUI_Plugins_Close = GUICtrlCreateButton(t('FINISH_BUT'), 320, 132, 83, 25)
	$GUI_Plugins_Download = GUICtrlCreateButton(t('TERM_DOWNLOAD'), 224, 132, 83, 25)
	GUICtrlSetState(-1, $GUI_DISABLE)
	$GUI_Plugins_Description = GUICtrlCreateEdit("", 224, 8, 177, 85, BitOR($ES_AUTOVSCROLL, $ES_WANTRETURN, $ES_READONLY, $ES_NOHIDESEL,$ES_MULTILINE))
	$GUI_Plugins_FileTypes = GUICtrlCreateEdit("", 224, 96, 177, 33, BitOR($ES_AUTOVSCROLL, $ES_WANTRETURN, $ES_READONLY, $ES_NOHIDESEL,$ES_MULTILINE))
	GUISetState(@SW_SHOW)

	Opt("GUIOnEventMode", 0)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $GUI_Plugins_Close
				ExitLoop
			Case $GUI_Plugins_List
				GUICtrlSetData($GUI_Plugins_FileTypes, $sSupportedFileTypes)
				Local $pos = _GUICtrlListBox_GetCurSel($GUI_Plugins_List)
				If $pos > -1 Then
					Local $return = _GUICtrlListBox_GetText($GUI_Plugins_List, $pos)
					$pos = _ArraySearch($aPluginInfo, $return)
					If @error Then ContinueLoop
					$current = $pos
					If Not StringInStr($aPluginInfo[$pos][0], "\bin\") Then $aPluginInfo[$pos][0] = @ScriptDir & "\bin\" & $aPluginInfo[$pos][0]
					If FileExists($aPluginInfo[$pos][0]) Then ; Installed
						GUICtrlSetData($GUI_Plugins_Download, t('TERM_INSTALLED'))
						GUICtrlSetState($GUI_Plugins_Download, $GUI_DISABLE)
					ElseIf $aPluginInfo[$pos][4] = False Then ; No download url available
						GUICtrlSetData($GUI_Plugins_Download, t('TERM_DOWNLOAD'))
						GUICtrlSetState($GUI_Plugins_Download, $GUI_DISABLE)
					Else
						GUICtrlSetData($GUI_Plugins_Download, t('TERM_DOWNLOAD'))
						GUICtrlSetState($GUI_Plugins_Download, $GUI_ENABLE)
					EndIf
					GUICtrlSetData($GUI_Plugins_Description, $aPluginInfo[$pos][2])
					GUICtrlSetData($GUI_Plugins_FileTypes, $sSupportedFileTypes & " " & $aPluginInfo[$pos][3])
				EndIf
			Case $GUI_Plugins_Download
				If $current = -1 Then ContinueLoop
				Cout("Download clicked for plugin " & $aPluginInfo[$current][1])
				ShellExecute($aPluginInfo[$current][4])
		EndSwitch
	WEnd

	GUIDelete($GUI_Plugins)
	Opt("GUIOnEventMode", 1)
EndFunc   ;==>GUI_Plugins

; Option to delete all log files
Func GUI_DeleteLogs()
	Cout("Deleting log files")
	Local $handle, $return, $i

	$handle = FileFindFirstFile(@ScriptDir & "\log\*.log")
	If $handle == -1 Then Return

	While 1
		$return = FileFindNextFile($handle)
		If @error Then ExitLoop
		FileDelete(@ScriptDir & "\log\" & $return)
		$i += 1
	WEnd

	FileClose($handle)

	GUI_UpdateLogItem()

	Cout("Deleted a total of " & $i & " files")
EndFunc   ;==>GUI_DeleteLogs

; Update log directory size in menu entry after deleting log files
Func GUI_UpdateLogItem()
	Local $size = Round(DirGetSize(@ScriptDir & "\log") / 1024 / 1024, 2) & " MB"
	GUICtrlSetData($logitem, t('MENU_FILE_LOG_LABEL', $size))
EndFunc   ;==>GUI_UpdateLogItem

; Display use statistics
Func GUI_Stats()


EndFunc   ;==>GUI_Stats

; Open program directory
Func GUI_ProgDir()
	ShellExecute(@ScriptDir)
EndFunc

; Create about GUI
Func GUI_About()
	Cout("Creating about GUI")
	$About = GUICreate($title & " " & $codename, 397, 270, -1, -1, -1, -1, $guimain)
	GUICtrlCreateLabel($name, 24, 16, 348, 52, $SS_CENTER)
	GUICtrlSetFont(-1, 30, 400, 0, "MS Sans Serif")
	GUICtrlCreateLabel(t('ABOUT_VERSION', $version), 16, 72, 362, 17, $SS_CENTER)
	GUICtrlCreateLabel(t('ABOUT_INFO_LABEL', CreateArray("Jared Breland <jbreland@legroom.net>", "uniextract@bioruebe.com", $website, "GNU GPLv2")), 16, 104, 364, 113, $SS_CENTER)
	GUICtrlCreateLabel($ID, 5, 255, 175, 15)
	GUICtrlSetFont(-1, 8, 800, 0, "Arial")
	GUICtrlCreatePic(".\support\Icons\Bioruebe.jpg", 295, 212, 89, 50)
	$About_OK = GUICtrlCreateButton(t('OK_BUT'), 154, 225, 89, 25)
	GUISetState(@SW_SHOW)

	GUICtrlSetOnEvent($About_OK, "GUI_About_Exit")
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_About_Exit")
EndFunc   ;==>GUI_About

; Exit about GUI if OK clicked or window closed
Func GUI_About_Exit()
	Cout("Closing about GUI")
	GUIDelete($About)
EndFunc   ;==>GUI_About_Exit

; Launch Universal Extractor website if help menu item clicked
Func GUI_Website()
	Cout("Opening website")
	ShellExecute($website)
EndFunc

; Launch Universal Extractor 2 website if help menu item clicked
Func GUI_Website2()
	Cout("Opening version 2 website")
	ShellExecute($website2)
EndFunc

; Launch Universal Extractor 2 Github website if help menu item clicked
Func GUI_Website_Github()
	Cout("Opening Github website")
	ShellExecute($websiteGithub)
EndFunc

; Launch developer forum website if help menu item selected
Func GUI_Forum()
	Cout("Opening forum")
	ShellExecute($forum)
EndFunc   ;==>GUI_Forum

; Exit if Cancel clicked or window closed
Func GUI_Exit()
	GUIGetPosition()
	terminate("silent", '', '')
EndFunc   ;==>GUI_Exit

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

	terminate("silent", '', '')
EndFunc

; Test specific functions
Func Test()
	;Sleep(50000)
	;GUI_FirstStart()
	;terminate("syntax", "", "")
	;check7z()
	;_GetOSLanguage()
	;GUI_ContextMenu()
	;CreateLog("debug")
;~ 	BatchQueuePop()
;~ 	IsJavaInstalled()
;~ 	_CreateTrayMessageBox("Test")
	;GUI_ContextMenu_remove()
	;GUIGetPosition()
	;$file = "C:\Users\William\Desktop\Documents\Unpacker Test\wzipse40.exe"
	;$outdir = "C:\Users\William\Desktop\Documents\Unpacker Test\wzipse40"
	;extract("isexe", 'InstallShield' & t('TERM_ARCHIVE'))
	;MsgBox(1,"", $unshield & ' -d "' & $outdir & '" x "' & $file & '"')
	;extract("rgss", "")
	;GetFFmpeg()
;~ 	_AfterUpdate()
EndFunc   ;==>Test

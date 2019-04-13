#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=support\Icons\uniextract_exe.ico
#AutoIt3Wrapper_Outfile=UniExtractUpdater_NoAdmin.exe
#AutoIt3Wrapper_Res_Description=Update utility for Universal Extractor
#AutoIt3Wrapper_Res_Fileversion=2.0.0.0
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Bioruebe

 Script Function:
	Auto-updater for Universal Extractor

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <Array.au3>
#include <GUIConstants.au3>
#include <Inet.au3>

Const $sUpdaterTitle = "Universal Extractor Updater"
Const $sMainUpdateURL = "https://update.bioruebe.com/uniextract/data/UniExtract.exe"
Const $sMainNighlyUpdateURL = "https://update.bioruebe.com/uniextract/nightly/UniExtract.exe"
Const $sGetLinkURL = "https://update.bioruebe.com/uniextract/geturl.php?q="
Const $sUniExtract = @ScriptDir & "\UniExtract.exe"

If Not FileExists($sUniExtract) Then
	If MsgBox(16+4, $sUpdaterTitle, "Universal Extractor main executable not found in current directory." & @CRLF & @CRLF & "Path is " & $sUniExtract & @CRLF & @CRLF & "Do you want to redownload Universal Extractor?") == 6 Then _
		_UpdateUniExtract()
	Exit
EndIf

If $cmdline[0] < 1 Then Exit ShellExecute($sUniExtract, "/update")

Sleep(50)

If $cmdline[1] == "/pluginst" Then
	; To install plugins we just start UniExtract elevated
	Exit ShellExecute($sUniExtract, "/plugins")
ElseIf $cmdline[1] == "/main" Then
	_UpdateUniExtract(_ArraySearch($cmdline, "/nightly") > -1)
ElseIf $cmdline[1] == "/helper" Then
	Exit ShellExecute($sUniExtract, "/updatehelper")
ElseIf $cmdline[1] == "/ffmpeg" Then
	_GetFFMPEG()
EndIf

Func _UpdateUniExtract($bNightly = False)
	If Not ProcessWaitClose($sUniExtract, 10) Then Exit MsgBox(16, $sUpdaterTitle, "Failed to close Universal Extractor. Please terminate the process manually and try again.")

	_Download($bNightly? $sMainNighlyUpdateURL: $sMainUpdateURL)
	$error = @error

	Sleep(100)
	Exit ShellExecute($sUniExtract, $error? "": "/afterupdate")
EndFunc

Func _GetFFMPEG()
	Const $cmd = (FileExists(@ComSpec)? @ComSpec: @WindowsDir & '\system32\cmd.exe') & ' /d /c '
	Const $sOSArchDir = @ScriptDir & "\bin\" & (@OSArch = 'X64'? 'x64\': 'x86\')
	Const $sOSArch = @OSArch = 'X64'? '64': '32'
	Const $sLicenseFile = @ScriptDir & "\docs\FFmpeg_license.html"
	Const $7z = '""' & $sOSArchDir & '7z.exe"'

	$FFmpegURL = _INetGetSource($sGetLinkURL & "ffmpeg" & (StringInStr(@OSVersion, "WIN_XP")? "xp": "") & $sOSArch & "&r=0")
	$return = _Download($FFmpegURL, @TempDir, False)
	If @error Then Exit 1

	; Extract files, move them to scriptdir and delete files from tempdir
	Local $ret = RunWait($cmd & $7z & ' e -ir!ffmpeg.exe -y -o"' & $sOSArchDir & '" "' & $return & '"', @TempDir)
	FileDelete(@TempDir & $return)
	If $ret <> 0 Then Exit MsgBox(48, $sUpdaterTitle, "Failed to extract update package " & $return & "." & @CRLF & @CRLF & "Make sure Universal Extractor is up to date and try again, or unpack the file manually to " & $sOSArchDir)

	; Download license information
	If Not FileExists($sLicenseFile) Then _Download("https://ffmpeg.org/legal.html", $sLicenseFile, False, True)

	Run($sUniExtract)
EndFunc

Func _Download($sURL, $sDir = @ScriptDir, $bCreateBackup = True, $bIsFilePath = False)
	; Create GUI with progressbar
	Local $hGUI = GUICreate("Downloading", 466, 109, -1, -1, $WS_POPUPWINDOW, -1)
	GUICtrlCreateLabel($sURL, 8, 16, 446, 17, $SS_CENTER)
	Local $idProgress = GUICtrlCreateProgress(8, 46, 446, 25)
	GUISetState(@SW_SHOW)

	; Get file size
	Local $iBytesReceived = 0
	Local $iBytesTotal = InetGetSize($sURL)
	Local $idSize = GUICtrlCreateLabel($iBytesReceived & "/" & $iBytesTotal & " kb", 8, 76, 446, 17, $SS_CENTER)

	; Download File
	Local $sFile = $bIsFilePath? $sDir: $sDir & "\" & StringTrimLeft($sURL, StringInStr($sURL, "/", 0, -1))
	Local $sBackupFile = $sFile & ".bak"

	If $bCreateBackup And FileExists($sFile) Then FileMove($sFile, $sBackupFile)

	Local $hDownload = InetGet($sURL, $sFile, 1, 1)

	; Update progress bar
	While Not InetGetInfo($hDownload, 2)
		Sleep(50)
		If InetGetInfo($hDownload, 4) <> 0 Then
			GUIDelete($hGUI)
			If $bCreateBackup Then FileMove($sBackupFile, $sFile, 1)
			_DownloadError($sURL)
			Return SetError(1, 0, 0)
		EndIf
		$iBytesReceived = InetGetInfo($hDownload, 0)
		GUICtrlSetData($idProgress, Int($iBytesReceived / $iBytesTotal * 100))
		GUICtrlSetData($idSize, $iBytesReceived & "/" & $iBytesTotal & " kb")
	WEnd

	; Close GUI
	GUIDelete($hGUI)
	If Not FileExists($sFile) Then
		If $bCreateBackup Then FileMove($sBackupFile, $sFile, 1)
		_DownloadError($sURL)
		Return SetError(1, 0, 0)
	EndIf

	If $bCreateBackup Then FileDelete($sBackupFile)
	Return $sFile
EndFunc

Func _DownloadError($sURL)
	MsgBox(48, $sUpdaterTitle, 'The file ' & $sURL & ' could not be downloaded. Please ensure that you are connected to the internet and try again.')
EndFunc
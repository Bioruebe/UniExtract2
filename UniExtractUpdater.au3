#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=support\Icons\uniextract_exe.ico
#AutoIt3Wrapper_Res_Description=Update utility for Universal Extractor
#AutoIt3Wrapper_Res_Fileversion=1.0.0
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.1
 Author:         Bioruebe

 Script Function:
	Auto-updater for Universal Extractor

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

Const $sTitle = "Universal Extractor Updater"

If $cmdline[0] < 1 Then Exit MsgBox(16, $sTitle, "Updater called with invalid number of arguments.")
If Not FileExists($cmdline[1]) Then Exit MsgBox(16, $sTitle, "Invalid update package passed to updater.")
$OSArch = @OSArch = 'X64'? 'x64': 'x86'

If $cmdline[0] == 2 Then
	_UpdateFFMPEG()
Else
	 _UpdateUniExtract()
EndIf

Func _UpdateUniExtract()
	If Not ProcessWaitClose("UniExtract.exe", 10) Then Exit MsgBox(16, $sTitle, "Failed to close Universal Extractor. Please terminate the process manually and try again.")

	$sCmd = @ScriptDir & '\bin\' & $OSArch & '\7z.exe x -y -xr!UniExtract.ini -o"' & @ScriptDir & '" "' & $cmdline[1] & '"'
	;~ MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, 'Selection:' & @CRLF & '$sCmd' & @CRLF & @CRLF & 'Return:' & @CRLF & $sCmd) ;### Debug MSGBOX
	RunWait($sCmd)
	Sleep(100)
	FileDelete($cmdline[1])
	Run(@ScriptDir & "\UniExtract.exe /afterupdate")
EndFunc

Func _UpdateFFMPEG()
	; Binaries
	FileMove($cmdline[1] & "\bin\ffmpeg.exe", @ScriptDir & "\bin\" & $OSArch & "\ffmpeg.exe", 1)
	FileMove($cmdline[1] & "\licenses\*", @ScriptDir & "\docs\FFmpeg\", 8+1)
	DirRemove($cmdline[1], 1)

	; License files
	If $cmdline[2] <> 0 Then FileMove($cmdline[2], @ScriptDir & "\docs\FFmpeg\FFmpeg_license.html", 8 + 1)
EndFunc


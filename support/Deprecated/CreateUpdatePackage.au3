#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.1
 Author:         Bioruebe

 Script Function:
	Creates an update package with all changed files since last run

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <Array.au3>
#include <Crypt.au3>
#include <File.au3>
#include <String.au3>

; Settings - Default values work if this script is run from the \support subdirectory
$sDir = "..\"
$sSnapshotFile = "Snapshot.csv" ; path, hash
$sMainFile = "..\UniExtract.au3"
$s7z = "..\bin\x64\7z.exe"
$sFilter = "*|*.au3;*.xcf;*.bak;*.csv;standard.ini;UniExtract.ini;English_old.ini;passwords.txt;list.txt;.gitignore;ffmpeg.exe;arc_conv.exe;bootimg.exe;ci-extractor.exe;dcp_unpacker.exe;dgcac.exe;EnigmaVBUnpacker.exe;iscab.exe;i5comp.exe;mpq.wcx*;RPGDecrypter.exe;sim_unpacker.exe;Extractor.exe;extract.exe;ZD50149.DLL;ZD51145.DLL;gea.dll;gentee.dll;" & $sSnapshotFile & ";" & $sSnapshotFile & ".bak" & "|.git;backup;devdata;homepage;log;test;userlogs;Update;crass-0.4.14.0;IS_Languages;FFmpeg"

$aVersion = _StringBetween(FileRead($sMainFile), 'version = "', '"')
If @error Then Dim $aVersion = ["Update"]
$sOutdir = ".\" & $aVersion[0] & "\"
;~ _ArrayDisplay($aVersion)

Local $aSnapshot[0][2], $aChanged[0]
_FileReadToArray($sSnapshotFile, $aSnapshot, $FRTA_NOCOUNT, "|")
If $aSnapshot == 0 Then Local $aSnapshot[0][2]

$aFiles = _FileListToArrayRec($sDir, $sFilter, $FLTAR_FILES, $FLTAR_RECUR)
_ArrayDelete($aFiles, 0)
;~ _ArrayDisplay($aFiles)
;~ _ArrayDisplay($aSnapshot)

_Crypt_Startup()
For $sFile In $aFiles
	$dHash = _Crypt_HashFile($sDir & $sFile, $CALG_MD5)
	$iIndex = _ArraySearch($aSnapshot, $sFile)
	If $iIndex > -1 Then
		If $aSnapshot[$iIndex][1] = $dHash Then
			Cout("Unchanged file: " & $sFile)
			ContinueLoop
		EndIf

		Cout("Changed file: " & $sFile & "(" & $aSnapshot[$iIndex][1] & " <=> " & $dHash & ")")
		$aSnapshot[$iIndex][1] = $dHash
	Else
		Cout("New file: " & $sFile & " - " & $dHash)
		_ArrayAdd($aSnapshot, $sFile & "|" & $dHash)
	EndIf

	_ArrayAdd($aChanged, $sFile)
Next
_Crypt_Shutdown()

;~ _ArrayDisplay($aSnapshot)
_ArrayDisplay($aChanged)

; Create update package
If FileExists($sOutdir) Then DirRemove($sOutdir, 1)
If UBound($aChanged) < 1 Then Exit 0
$sArchive = $aVersion[0] & ".zip"
If FileExists($sArchive) Then Exit MsgBox(48, "Error", "File already exists: " & $sArchive)

Cout("Copying files")
For $sFile In $aChanged
	; 7zip cannot replace itself, so to extract new versions of 7zip, it has to be renamed
	FileCopy($sDir & $sFile, $sOutdir & $sFile & (StringInStr($sFile, "\7z.")? ".new": ""), $FC_CREATEPATH)
Next

Cout("Compressing")
RunWait($s7z & ' a -mx=9 "' & $aVersion[0] & '.zip" "' & $sOutdir & '*"')
;~ DirRemove($sOutdir, 1)

Cout("Saving snapshot")
FileMove($sSnapshotFile, $sSnapshotFile & ".bak", $FC_OVERWRITE)
$hFile = FileOpen($sSnapshotFile, $FO_OVERWRITE)
FileWrite($hFile, _ArrayToString($aSnapshot, "|"))
FileClose($hFile)
FileCopy($sSnapshotFile, $aVersion[0] & ".csv", $FC_OVERWRITE)

; Write data to stdout stream
Func Cout($Data)
	Local $Output = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & ":" & @MSEC & @TAB & $Data & @CRLF; & @CRLF
	ConsoleWrite($Output)
EndFunc
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.12.0
 Author:         Bioruebe

 Script Function:
	Update language files, replace %s with %1, %2, ..., %n
	Used to convert legacy language files to new UniExtract 2 format

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <Array.au3>
#include <String.au3>

; Run for each file like this:
;~ _SUpdate("..\lang\Chinese (Simplified).ini")

Func _SUpdate($sFile)
;~ 	ConsoleWrite("[SUpdate] " & $sFile & @CRLF)
	$hFile = FileOpen($sFile, 16384)
	$aFileContent = FileReadToArray($sFile)
	FileClose($hFile)

	;~ _ArrayDisplay($aFileContent)

	For $i = 0 To UBound($aFileContent)-1
		$sFirstChar = StringLeft($aFileContent[$i], 1)
		If $sFirstChar = ";" Or $sFirstChar = "[" Then ContinueLoop
		For $j = 1 To 9
			$aFileContent[$i] = StringReplace($aFileContent[$i], "%s", "%" & $j, 1)
		Next
	Next
	;~ _ArrayDisplay($aFileContent)

	$hFile = FileOpen($sFile, 32+2)
	FileWrite($hFile, _ArrayToString($aFileContent, @CRLF))
	FileClose($hFile)
EndFunc
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.12.0
 Author:         Bioruebe

 Script Function:
	Update language files, replace %s with %1, %2, ..., %n

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <Array.au3>
#include <String.au3>

;~ $sFile = "..\English.ini"
$sFile = "..\lang\German.ini"

$aFileContent = FileReadToArray($sFile)

;~ _ArrayDisplay($aFileContent)

For $i = 0 To UBound($aFileContent)-1
	$sFirstChar = StringLeft($aFileContent[$i], 1)
	If $sFirstChar = ";" Or $sFirstChar = "[" Then ContinueLoop
	For $j = 1 To 9
		$aFileContent[$i] = StringReplace($aFileContent[$i], "%s", "%" & $j, 1)
	Next
Next
;~ _ArrayDisplay($aFileContent)

$hFile = FileOpen($sFile, 2)
FileWrite($hFile, _ArrayToString($aFileContent, @CRLF))
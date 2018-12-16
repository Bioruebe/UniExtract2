#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.12.0
 Author:         Bioruebe

 Script Function:
	Update language files, replace %s with %name
	Used to convert legacy language files to new UniExtract 2 format

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <Array.au3>
#include <String.au3>

$sUpdatedFile = "..\English.ini"
Global $aTerms[0][2]

$aFileContent = FileReadToArray($sUpdatedFile)

;~ _ArrayDisplay($aFileContent)

For $i = 0 To UBound($aFileContent)-1
	$sFirst = StringLeft($aFileContent[$i], 1)
	$iPos = StringInStr($aFileContent[$i], "%name")
	If $sFirst = ";" Or $sFirst = "[" Or Not $iPos Then ContinueLoop
	$sTerm = _StringBetween($aFileContent[$i], "", " =")[0]
	$sPrev = StringLeft($aFileContent[$i], $iPos)
	$aReturn = StringRegExp($sPrev, "%s", 1)
	_ArrayAdd($aTerms, $sTerm & "|" & UBound($aReturn) + 1, 0, "|")
	ConsoleWrite("[" & $i & " -> " & UBound($aReturn) + 1 & "] " & $aFileContent[$i] & @CRLF)
Next
;~ _ArrayDisplay($aTerms)

;~ _NameUpdate("..\lang\Chinese (Simplified).ini")


Func _NameUpdate($sOldFile)
;~ 	ConsoleWrite("[NameUpdate] " & $sOldFile & @CRLF)

	$hOldFile = FileOpen($sOldFile, 16384)
	$aOldContent = FileReadToArray($hOldFile)
	FileClose($hOldFile)
	;~ _ArrayDisplay($aOldContent)

	For $i = 0 To UBound($aTerms)-1
		$iIndex = _ArraySearch($aOldContent, $aTerms[$i][0], 0, 0, 0, 1)
		If @error Then ContinueLoop
		$aOldContent[$iIndex] = StringReplace($aOldContent[$iIndex], "%s", "%name", $aTerms[$i][1], 0)
	Next

	;~ _ArrayDisplay($aOldContent)

	$hOldFile = FileOpen($sOldFile, 32+2)
	FileWrite($hOldFile, _ArrayToString($aOldContent, @CRLF))
	FileClose($hOldFile)
EndFunc
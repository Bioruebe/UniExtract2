#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         Bioruebe

 Script Function:
	Update Universal Extractor's language files (add new & delete old terms)

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <Array.au3>
#include <File.au3>

Global $aEnglish, $sVersion, $sContent, $sHeader = ""

Global $aExcludes[1] = ["German.ini"]
Const $sEnglishLanguageFile = "..\English.ini"
Const $sLanguageDir = "..\lang\"

Global Const $sTimestamp = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & "-" & @MIN & "-" & @SEC

_ReadTemplate()
_ProcessFiles()


; Read template file, build array with key names only to be inserted into new files
Func _ReadTemplate()
	_FileReadToArray($sEnglishLanguageFile, $aEnglish)
	If @error Then Exit ConsoleWrite("Error opening file: " & @error & @CRLF)

	; Copy header (translation instructions) from English.ini
	Local $i = 8
	While StringLeft($aEnglish[$i], 1) == ";"
		$sHeader &= $aEnglish[$i] & @CRLF
		$i += 1
	WEnd

	_Strip($aEnglish)
	;~ _ArrayDisplay($aEnglish)
EndFunc

; Read all language files and update them
Func _ProcessFiles()
	$aLanguageFiles = _FileListToArray($sLanguageDir, "*.ini", 1)
	;~ _ArrayDisplay($aLanguageFiles)

	For $i = 1 To $aLanguageFiles[0]
		Local $sFile = $aLanguageFiles[$i]

		If _ArraySearch($aExcludes, $sFile) < 0 Then
			_ProcessFile($sLanguageDir & $sFile)
		Else
			ConsoleWrite("Skipping file " & $sFile)
		EndIf

		ConsoleWrite(" (" & $i & "/" & $aLanguageFiles[0] & ")" & @CRLF)
	Next
EndFunc

; Read and preprocess a language file
Func _ReadFile($sFile)
	Local $hFile = FileOpen($sFile, 16384)
	Local $aReturn = FileReadToArray($hFile)
	If @error Then ConsoleWrite(@CRLF & "Error opening file: " & @error & @CRLF)
	FileClose($hFile)

	Local $iSize = UBound($aReturn)
	Local $aFile[$iSize][2], $aSplit

	For $i = 0 To $iSize - 1
		If StringLeft($aReturn[$i], 1) == ";" Then
			$aFile[$i][0] = StringStripWS($aReturn[$i], $STR_STRIPLEADING + $STR_STRIPTRAILING)
			ContinueLoop
		EndIf

		$aSplit = StringSplit($aReturn[$i], "=")
		$aFile[$i][0] = StringStripWS($aSplit[1], $STR_STRIPLEADING + $STR_STRIPTRAILING)
		If $aSplit[0] > 2 Then $aSplit[2] = _ArrayToString($aSplit, "=", 2, -1, "")
		If $aSplit[0] > 1 Then $aFile[$i][1] = StringStripWS($aSplit[2], $STR_STRIPLEADING + $STR_STRIPTRAILING)
	Next

;~ 	_ArrayDisplay($aFile)
	Return $aFile
EndFunc

; Process a single language file
Func _ProcessFile($sFile)
	ConsoleWrite("Processing file " & $sFile)

	Local $aTranslation = _ReadFile($sFile)

	_Update($aTranslation)

	Local $iReturn = FileMove($sFile, "..\backup\lang\" & $sVersion & "_" & $sTimestamp & "\", 8+1)
	If $iReturn <> 1 Then ConsoleWrite(@CRLF & "Error moving file" & @CRLF)

	$hFile = FileOpen($sFile, 32+2)
	FileWrite($hFile, $sContent)
	FileClose($hFile)
EndFunc

; Update a single language file
Func _Update(ByRef $aTranslation)
	Global $sContent = ""
	Local $iSize = UBound($aTranslation) - 1

	; Copy header from old file
	For $i = 0 To $iSize
		$sLine = $aTranslation[$i][0]
		If StringInStr($sLine, "Written for Universal Extractor") Then
			_AddLine("; Written for Universal Extractor " & $sVersion)
		ElseIf $i > 3 And ($sLine == ";" Or StringIsSpace($sLine)) Then
			_AddLine($sLine)
			ExitLoop
		ElseIf StringLeft($sLine, 1) = ";" Then
			_AddLine($sLine)
		EndIf
	Next

	_AddLine($sHeader)
	_ArraySort($aTranslation)
	Local $bIsHeader = True, $aSplit, $sLine
	For $i = 1 To $aEnglish[0]
		$iPos = -1

		; Skip header as it can have different lengths
		If $bIsHeader And (StringLeft($aEnglish[$i], 1) = ";" Or StringIsSpace($aEnglish[$i])) Then
			ContinueLoop
		Else
			$bIsHeader = False
			$iPos = _ArrayBinarySearch($aTranslation, $aEnglish[$i])
;~ 			ConsoleWrite(@CRLF & "Searching '" & $aEnglish[$i] & "' -> position " & $iPos)

			If $iPos = -1 Then
				If StringLeft($aEnglish[$i], 1) == ";" Or StringIsSpace($aEnglish[$i]) Then
					; Comments etc.
					_AddLine($aEnglish[$i])
				Else
					; New content, no translation in old file
					_AddLine($aEnglish[$i] & ' = ""')
				EndIf
			Else
				; Write old content to file
				$sLine = $aTranslation[$iPos][0]
				If $aTranslation[$iPos][1] Then $sLine &= " = " & $aTranslation[$iPos][1]
				_AddLine($sLine)
			EndIf
		EndIf
	Next
EndFunc

; Add a line to the file content
Func _AddLine($sLine)
	$sContent &= $sLine & @CRLF
EndFunc

; Strip values
Func _Strip(ByRef $aArray)
	For $i = 1 To $aArray[0]
		Local $iPos = StringInStr($aArray[$i], "=")
		If $iPos Then $aArray[$i] = StringStripWS(StringLeft($aArray[$i], $iPos - 2), $STR_STRIPLEADING + $STR_STRIPTRAILING)
		If StringInStr($aArray[$i], "Written for") Then $sVersion = StringReplace($aArray[$i], "; Written for Universal Extractor ", "")
	Next
EndFunc
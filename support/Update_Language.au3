#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         Bioruebe

 Script Function:
	Update Universal Extractor's language files (add new & delete old terms)

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <Array.au3>
#include <File.au3>

Dim $arr_new, $file_old, $version

Dim $exclude[1] = ["German.ini"]
$ini_new = "..\English.ini"
$sLanguageDir = "..\lang\"

$timestamp = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & "-" & @MIN & "-" & @SEC

$ini_old = _FileListToArray($sLanguageDir, "*.ini", 1)
;_ArrayDisplay($ini_old)

; Read template file, build array with key names only to be inserted into new files
_FileReadToArray($ini_new, $arr_new)
If @error Then Exit ConsoleWrite("Error opening file: " & @error & @CRLF)

; Copy header template from English.ini
$sHeader = ""
Local $i = 8
While StringLeft($arr_new[$i], 1) == ";"
	$sHeader &= $arr_new[$i] & @CRLF
	$i += 1
WEnd
$sHeader &= @CRLF

_Strip($arr_new)
;~ _ArrayDisplay($arr_new)


For $i=1 To $ini_old[0]
	Local $sFile = $ini_old[$i]

	If _ArraySearch($exclude, $sFile) < 0 Then
		_Start($sLanguageDir & $sFile)
	Else
		ConsoleWrite("Skipping file " & $sFile)
	EndIf

	ConsoleWrite(" (" & $i & "/" & $ini_old[0] & ")" & @CRLF)
Next

Func _Start($file)
	ConsoleWrite("Processing file " & $file)

	$hFile = FileOpen($file, 16384)
	$file_old = FileReadToArray($hFile)
	If @error Then ConsoleWrite(@CRLF & "Error opening file: " & @error & @CRLF)
	FileClose($hFile)
	_ArrayInsert($file_old, 0, UBound($file_old))

;~ 	_ArrayDisplay($file_old)

	Global $handle = FileOpen("new.ini", 32+2)

	_Update()

	FileClose($handle)

	$ret = FileMove($file, "..\backup\lang\" & $version & "_" & $timestamp & "\", 8+1)
	If $ret <> 1 Then ConsoleWrite(@CRLF & "Error moving file (1)" & @CRLF)
	$ret = FileMove("new.ini", $file, 1)
	If $ret <> 1 Then ConsoleWrite(@CRLF & "Error moving file (2)" & @CRLF)
EndFunc

Func _Update()
	; Copy header from old file
	For $i = 1 To $file_old[0]
		If StringInStr($file_old[$i], "Written for Universal Extractor") Then
			FileWriteLine($handle, "; Written for Universal Extractor " & $version)
		ElseIf $i > 3 And (StringStripWS($file_old[$i], 2) == ";" Or StringIsSpace($file_old[$i])) Then
			FileWriteLine($handle, $file_old[$i])
			ExitLoop
		ElseIf StringLeft($file_old[$i], 1) = ";" Then
			FileWriteLine($handle, $file_old[$i])
		EndIf
	Next

	FileWriteLine($handle, $sHeader)
	Local $bIsHeader = True
	For $i = 1 To $arr_new[0]
		$pos = -1
;~ 		ConsoleWrite("Searching " & $arr_new[$i] & " | Position " & $pos & " (Error: " & @error & ")" & @CRLF)

		; Skip header as it can have different lengths
		If $bIsHeader And (StringLeft($arr_new[$i], 1) = ";" Or StringIsSpace($arr_new[$i])) Then
			ContinueLoop
		Else
			$bIsHeader = False
			$pos = _ArraySearch($file_old, $arr_new[$i], 1, 0, 0, 1)
			If $pos = -1 Then
				If StringInStr($arr_new[$i], "=") Then
					FileWriteLine($handle, $arr_new[$i] & ' ""')
				Else
					FileWriteLine($handle, $arr_new[$i])
				EndIf
			Else
				FileWriteLine($handle, $file_old[$pos])
			EndIf
		EndIf
	Next
EndFunc

; Strip values
Func _Strip(ByRef $arr)
	For $i=1 To $arr[0]
		Local $pos = StringInStr($arr[$i], "=")
		If $pos Then $arr[$i] = StringLeft($arr[$i], $pos)
		If StringInStr($arr[$i], "Written for") Then $version = StringReplace($arr[$i], "; Written for Universal Extractor ", "")
	Next
EndFunc
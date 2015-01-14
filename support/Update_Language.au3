#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         Bioruebe

 Script Function:
	Update Universal Extractor's language files (add new & delete old terms)

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <Array.au3>
#include <File.au3>
#include "_GetIntersection.au3"

Dim $arr_new, $file_old, $version, $changes_arr, $changes_new_arr, $changes_old_arr, $changes_ini

Dim $exclude[2] = ["German.ini"]

; Template
$ini_new = "..\English.ini"

$dir = "..\lang\"

$changes_ini = "..\English_old.ini"

$timestamp = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & "-" & @MIN & "-" & @SEC

$ini_old = _FileListToArray($dir, "*.ini", 1)
;_ArrayDisplay($ini_old)

; Read template file
_FileReadToArray($ini_new, $arr_new)
If @error Then ConsoleWrite("Error opening file: " & @error & @CRLF)
_Strip($arr_new)


For $i=1 To $ini_old[0]
	CheckExcludes($ini_old[$i])
	ConsoleWrite(" (" & $i & "/" & $ini_old[0]+1 & ")" & @CRLF)
Next

ConsoleWrite("Creating changes.txt")
GetChanges()
ConsoleWrite(" (" & $i & "/" & $ini_old[0]+1 & ")" & @CRLF)

Func GetChanges()
	_FileReadToArray($changes_ini, $changes_old_arr)
	_FileReadToArray($ini_new, $changes_new_arr)

	$changes_arr = _GetIntersection($changes_old_arr, $changes_new_arr)

	;_ArrayDisplay($changes_arr)


	; Write to file
	FileMove($dir & "changes.txt", "..\backup\lang\" & $version & "_" & $timestamp & "\", 8+1)
	$handle = FileOpen($dir & "changes.txt", 2)
	FileWrite($handle, "Language file changes:" & @CRLF & @CRLF & "New/changed entries:" & @CRLF & _
			  "-----------------------------------------------------------" & @CRLF & @CRLF)

	WriteChanges($handle, 2)

	FileWrite($handle, @CRLF & @CRLF & "Deleted entries:" & @CRLF & _
			  "-----------------------------------------------------------" & @CRLF & @CRLF)

	WriteChanges($handle, 1)

	FileClose($handle)
EndFunc

Func WriteChanges($handle, $col)
	For $i=0 To UBound($changes_arr)-1
		If StringInStr($changes_arr[$i][$col], "=") Then FileWriteLine($handle, $changes_arr[$i][$col])
	Next
EndFunc

Func CheckExcludes($ini)
	For $j=0 To UBound($exclude)-1
		If $ini = $exclude[$j] Then
			ConsoleWrite("Skipping file " & $dir & $ini)
			Return
		EndIf
	Next
	_Start($dir & $ini)
EndFunc

Func _Start($file)
	ConsoleWrite("Processing file " & $file)

	_FileReadToArray($file, $file_old)
	If @error Then ConsoleWrite("Error opening file: " & @error & @CRLF)

	Global $handle = FileOpen("new.ini", 128+2) ;128 unicode

	_Update()

	FileClose($handle)

	$ret = FileMove($file, "..\backup\lang\" & $version & "_" & $timestamp & "\", 8+1)
	If $ret <> 1 Then ConsoleWrite("Error moving file (1)" & @CRLF)
	$ret = FileMove("new.ini", $file, 1)
	If $ret <> 1 Then ConsoleWrite("Error moving file (2)" & @CRLF)
EndFunc

Func _Update()
	For $i = 1 To $arr_new[0]
		$pos = 0
		$pos = _ArraySearch($file_old, $arr_new[$i], 1, 0, 0, 1)
		;ConsoleWrite("Searching " & $arr_new[$i] & " | Position " & $pos & " (Error: " & @error & ")" & @CRLF)
		If $i > 31 Then
			If $pos = -1 Then
				If StringInStr($arr_new[$i], "=") Then
					FileWriteLine($handle, $arr_new[$i] & '""')
				Else
					FileWriteLine($handle, $arr_new[$i])
				EndIf
			Else
				FileWriteLine($handle, $file_old[$pos])
			EndIf
		Else
			If StringInStr($file_old[$i], "Written for Universal Extractor") Then
				FileWriteLine($handle, $arr_new[$i])
				$version = StringReplace($arr_new[$i], "; Written for ", "")
			Else
				FileWriteLine($handle, $file_old[$i])
			EndIf
		EndIf
	Next
EndFunc

Func _Strip(ByRef $arr)
	For $i=1 To $arr[0]
		$pos = 0
		$pos = StringInStr($arr[$i], "=")
		If $pos Then $arr[$i] = StringLeft($arr[$i], $pos+1)
	Next
EndFunc
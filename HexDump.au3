#include-once

; Example
;~ ConsoleWrite(_HexDump(@WindowsDir & "\Notepad.exe", 256))

; #FUNCTION# ====================================================================================================================
; Name ..........: _HexDump
; Description ...: Open a specified file and create a hex dump of the first $iBytes bytes.
; Syntax ........: _HexDump($sFile, $iBytes)
; Parameters ....: $sFile               - Filename of the file to dump.
;                  $iBytes              - The length in bytes to dump.
; Return values .: A formatted hex dump string or an empty string and sets the @error flag to non-zero.
; 				   @error: 		1   - Error reading file
;								2|3 - DllCall 'CryptBinaryToString' failed
; Author ........: Bioruebe
; Modified ......:
; Remarks .......:
; Related .......: http://msdn.microsoft.com/en-us/library/windows/desktop/aa379887%28v=vs.85%29.aspx
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _HexDump($sFile, $iBytes)
	Local $hFile, $bData, $tInput, $aDllCall, $tOut

	$hFile = FileOpen($sFile, 16)
	If @error Then Return SetError(1, 0, "")
	$bData = FileRead($hFile, $iBytes)
	If @error Then Return SetError(1, 0, "")
	FileClose($hFile)

    $tInput = DllStructCreate("byte[" & BinaryLen($bData) & "]")
    DllStructSetData($tInput, 1, $bData)

	$aDllCall = DllCall("crypt32.dll", "int", "CryptBinaryToString", "ptr", DllStructGetPtr($tInput), "dword", DllStructGetSize($tInput), "dword", _
						0x0000000B, "ptr", 0, "dword*", 0)

    If @error Or Not $aDllCall[0] Then Return SetError(2, 0, "")

    $tOut = DllStructCreate("char[" & $aDllCall[5] & "]")
    $aDllCall = DllCall("crypt32.dll", "int", "CryptBinaryToString", "ptr", DllStructGetPtr($tInput), "dword", DllStructGetSize($tInput), "dword", _
						0x0000000B, "ptr", DllStructGetPtr($tOut), "dword*", $aDllCall[5])

    If @error Or Not $aDllCall[0] Then Return SetError(3, 0, "")

	Return DllStructGetData($tOut, 1)
EndFunc
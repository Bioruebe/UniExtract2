#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.12.0
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <String.au3>


MediaFileScan("C:\Users\Bioruebe\Documents\example.ogg")

Func MediaFileScan($f)
	Local $filetype_curr = ""

	$hDll = DllOpen("MediaInfo.dll")
	$hMI = DllCall($hDll, "ptr", "MediaInfo_New")

	$Open_Result = DllCall($hDll, "int", "MediaInfo_Open", "ptr", $hMI[0], "wstr", $f)
	$return = DllCall($hDll, "wstr", "MediaInfo_Inform", "ptr", $hMI[0], "int", 0)

	$hMI = DllCall($hDll, "none", "MediaInfo_Delete", "ptr", $hMI[0])
	DllClose($hDll)

;~ 	ConsoleWrite($return[0])
;~ 	MsgBox(262144, 'Debug line ~' & @ScriptLineNumber, $return[0]) ;### Debug MSGBOX

	; Format returned string to align in message box
	For $i in StringSplit($return[0], @CRLF, 2)
		$return = StringSplit($i, " : ", 2+1)
		If @error Then
			If Not StringIsSpace($i) Then $filetype_curr &= @CRLF & "[" & $i & "]" & @CRLF
			ContinueLoop
		EndIf
		$sType = StringStripWS($return[0], 4+2+1)
		$iLen = StringLen($sType)
		$filetype_curr &= $sType & _StringRepeat(@TAB, 3 - Floor($iLen / 10)) & (($iLen > 20 And $iLen < 25)? @TAB: "") & StringStripWS($return[1], 4+2+1) & @CRLF
	Next

;~ 	ConsoleWrite($filetype_curr)
;~ 	MsgBox(0, "", $filetype_curr)
EndFunc
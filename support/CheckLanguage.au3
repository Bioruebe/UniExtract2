#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         Bioruebe

 Script Function:
	Checks language file terms and returns terms which are not used in source code.
	Some terms are not used in the main script, but in the definition files, the installer or the updater!
	Don't delete them even if reported by this script!

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

Dim $error, $line, $handle, $script

$handle = FileOpen("..\UniExtract.au3")
$script = FileRead($handle)
FileClose($handle)

$handle = FileOpen("..\English.ini")

Do
	$line = FileReadLine($handle)
	If @error = -1 Then $error = 1
	$return = StringInStr($line, "=")
	If $return <> 0 Then
		If NOT StringInStr($script, StringLeft($line, $return - 2)) Then ConsoleWrite(StringLeft($line, $return - 2) & @CRLF)
	EndIf
Until $error = 1

FileClose($handle)
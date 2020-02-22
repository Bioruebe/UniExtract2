#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         Bioruebe

 Script Function:
	Returns files in /bin directory, which are not listed in helper binaries info file

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

$handle = FileOpen("..\helper_binaries_info.txt")
$infile = FileRead($handle)
FileClose($handle)

$search = FileFindFirstFile("..\bin\*")

If $search = -1 Then
    Exit
EndIf

While 1
    $file = FileFindNextFile($search)
    If @error Then ExitLoop
    If NOT StringInStr($infile, $file) Then ConsoleWrite($file & @CRLF)
WEnd

FileClose($search)
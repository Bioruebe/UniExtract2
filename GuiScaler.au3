#cs ----------------------------------------------------------------------------

 GUIScaler UDF by alpines
 https://autoit.de/thread/86505-guiscaler-guis-automatisch-zur-dpi-skalieren-lassen-windows-7-und-windows-10-per/

 Provides functions to apply correct scaling on high DPI screens or when using Windows' scaling modes.

#ce ----------------------------------------------------------------------------

#include-once

#include <GUIConstants.au3>
#include <GuiListView.au3>
#include <SendMessage.au3>
#include <WinAPIGdi.au3>
#include <WinAPISys.au3>
#include <WindowsConstants.au3>

;Scripting.Dictionary nimmt als Keys keine Handles an. Deshalb nutzen wir String($hWnd)
;	$__o_GUIControlInfos (von außen nicht zugänglich)
;		[n][0] = Control Id
;		[n][1] = Font die es bei der Erzeugung bekam
;		[n][2] = Created DPI (für ListViews wirds angepasst damit die Columns resized werden können.
;				 Die Fonts von TVs und LVs werden automatisch angepasst, deshalb können wir den Wert verändern,
;				 da wir beim Rescalen Spezialfälle für die Controls haben.
;		[n][3] = Resizing Method (näheres dazu in _GUI_SetResizing)
;		[n][4] = Control Position[4] (ControlGetPos) - Bei TreeView/ListView-Items kein Array
;
;	$__o_GUIInfos (von außen zugänglich mit _GUI_GetInfos(hWnd))
;		[0] = Width (bei 100%)
;		[1] = Height (bei 100%)
;		[2] = Aktuelle DPI
;		[3] = bCurrentlyResizing
Global $__o_GUIControlInfos = ObjCreate("Scripting.Dictionary")
Global $__o_GUIInfos = ObjCreate("Scripting.Dictionary")

Global Const $DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
If @OSBuild >= 9600 Then _WinAPI_SetProcessDpiAwarenessContext($DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)

Global Const $WM_DPICHANGED = 0x02E0
GUIRegisterMsg($WM_DPICHANGED, _GUI_WM_DPICHANGED)

Global $__t_ZeroPoint = DllStructCreate($tagPOINT)
DllStructSetData($__t_ZeroPoint, 1, 0)
DllStructSetData($__t_ZeroPoint, 2, 0)

Global Const $H_PRIMARYMONITOR = _WinAPI_MonitorFromPoint($__t_ZeroPoint, $MONITOR_DEFAULTTOPRIMARY)
Global Const $I_PRIMARYMONITOR_DPI = _WinAPI_GetDpiForMonitor($H_PRIMARYMONITOR)

Func _GUI_WM_DPICHANGED($hWnd, $iMsg, $wParam, $lParam)
	GUIRegisterMsg($WM_DPICHANGED, "")

	Local $tRect = DllStructCreate($tagRECT, $lParam)
	_GUI_Resize($hWnd, DllStructGetData($tRect, "Left"), DllStructGetData($tRect, "Top"), _WinAPI_HiWord($wParam))

	GUIRegisterMsg($WM_DPICHANGED, _GUI_WM_DPICHANGED)

	Return $GUI_RUNDEFMSG
EndFunc

Func _GUI_GetInfos($hWnd)
	Return $__o_GUIInfos.Item(String($hWnd))
EndFunc

Func _GUI_Resize($hGUI, $iX, $iY, $iDPI)
	;Wir müssen den Monitor speichern auf welchem das Fenster nach dem $WM_DPICHANGED ist.
	;Wenn wir nach dem Resizen plötzlich auf einem anderen Monitor landen verschieben wir da Fenster,
	;einfach nach (0, Default) wenn das Fenster zu weit links ist respektive nach (Default, 0) wenn es zu weit unten ist
	;Analog für (Width-fensterbreite, Default) (wenn es zu weit rechts vom anderen Monitor ist).
	Local $hMonitor = _WinAPI_MonitorFromWindow($hGUI)

	Local $aGUIInfos = $__o_GUIInfos.Item(String($hGUI))
	$aGUIInfos[2] = $iDPI
	$aGUIInfos[3] = True
	$__o_GUIInfos.Item(String($hGUI)) = $aGUIInfos

	GUISetState(@SW_LOCK, $hGUI)

	Local $aControls = $__o_GUIControlInfos.Item(String($hGUI))

	;GUI auf 100% bringen
	;Warum moven wir das Fenster 2x?
	;Beim ersten Move wird das Fenster zwar auf die richtige Größe verkleinert, allerdings bleibt die Titelleiste noch groß.
	;Senden wir noch ein WM_MOVE ab, so wird die Titelleiste verkleinert und die GUI hat nun die richtige Größe.
	;Kommentiert man das GUI_Move aus, kann man sehen wie das autoitinterne GUI Resizing die Controls anpasst,
	;obwohl das Fenster die Größe beibehalten hat (der ClientSize-Bereich vergrößert sich aber)
	_GUI_Move($hGUI, $iX, $iY, True)
	_GUI_Move($hGUI, $iX, $iY, True)

	For $i = 0 To UBound($aControls) - 1
		;ControlResizing verbieten, denn wir setzen die angepasste Größe
		GUICtrlSetResizing($aControls[$i][0], $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

		;Controlgrößen auf 100% bringen
		;SysTreeView- und ListView-Items returnen keine ControlPos.

		Local $aControlPos = $aControls[$i][4]

		If UBound($aControlPos) Then
			Local $fDPIScalefactor = $iDPI / 96

			;Achtung, wenn die Controls nur 1px groß (bzw. <2px)sind kann es sein, dass sie nicht dargestellt werden.
			;Beispiel Win10 150%

			GUICtrlSetPos($aControls[$i][0], _
				$aControlPos[0] * $fDPIScalefactor, _
				$aControlPos[1] * $fDPIScalefactor, _
				$aControlPos[2] * $fDPIScalefactor, _
				$aControlPos[3] * $fDPIScalefactor _
			)
		EndIf

		;Spezielle Fälle abfangen bei denen die Schrift automatisch mitskaliert wird von Windows und wir das nicht machen dürfen.
		Switch _WinAPI_GetClassName($aControls[$i][0])
			Case "SysTreeView32"
				ContinueLoop

			Case "SysListView32"
				;Header skalieren weils schöner aussieht.
				;Achtung, hier wird der Rundungsfehler in Kauf genommen weil der User ja die Header während
				;des Benutzens resizen (oder neue hinzufügen) kann. Deshalb speichern wir die Größen nicht ein.

				Local $LVM_FIRST = 0x1000
				Local $HDM_FIRST = 0x1200
				Local $LVM_GETHEADER = $LVM_FIRST + 31
				Local $HDM_GETITEMCOUNT = $HDM_FIRST + 0

				Local $hWndHeader = GUICtrlSendMsg($aControls[$i][0], $LVM_GETHEADER, 0, 0)
				Local $iCount = _SendMessage($hWndHeader, $HDM_GETITEMCOUNT, 0, 0)

				For $j = 0 To $iCount - 1
					Local $iColumnWidth = GUICtrlSendMsg($aControls[$i][0], $LVM_GETCOLUMNWIDTH, $j, 0)
					Local $fNormalizedWidth = $iColumnWidth / ($aControls[$i][2] / 96)
					Local $iNewWidth = $fNormalizedWidth * ($iDPI / 96)
					GUICtrlSendMsg($aControls[$i][0], $LVM_SETCOLUMNWIDTH, $j, $iNewWidth)
				Next

				$aControls[$i][2] = $iDPI
				ContinueLoop

		EndSwitch

		;Schriftart anpassen
		Local $tLogFontA = DllStructCreate("long lfHeight;long lfWidth;long lfEscapement;long lfOrientation;" & _
										   "long lfWeight;byte lfItalic;byte lfUnderline;byte lfStrikeOut;" & _
										   "byte lfCharSet;byte lfOutPrecision;byte lfClipPrecision;byte lfQuality;" & _
										   "byte lfPitchAndFamily;char lfFaceName[64];")

		_WinAPI_GetObject($aControls[$i][1], DllStructGetSize($tLogFontA), DllStructGetPtr($tLogFontA))

		Local $iInitialHeight = DllStructGetData($tLogFontA, "lfHeight")
		Local $fNormalizedHeight = $iInitialHeight / ($aControls[$i][2] / 96)
		Local $iNewHeight = -Floor(Abs($fNormalizedHeight * ($iDPI / 96)))

		DllStructSetData($tLogFontA, "lfHeight", $iNewHeight)

		;Neue Font erstellen
		Local $hFont = _WinAPI_CreateFontIndirect($tLogFontA)
		DllStructSetData($tLogFontA, "lfHeight", $iInitialHeight)

		;Alte Font disposen damit kein Memoryleak entsteht
		_WinAPI_DeleteObject($aControls[$i][5])

		;Neue Font abspeichern
		$aControls[$i][5] = $hFont

		GUICtrlSendMsg($aControls[$i][0], $WM_SETFONT, $aControls[$i][5], False)
	Next

	;Falls was aktualisiert wurde wieder abspeichern
	$__o_GUIControlInfos.Item(String($hGUI)) = $aControls

	;GUI auf die skalierte Größe bringen (GUICtrlSetResizing skaliert alles mit)
	;Gleiches Problem wie oben, deshalb 2x
	_GUI_Move($hGUI, $iX, $iY, False, $iDPI / 96)
	_GUI_Move($hGUI, $iX, $iY, False, $iDPI / 96)

	For $i = 0 To UBound($aControls) - 1
		;Alte Skalierung wiederherstellen
		GUICtrlSetResizing($aControls[$i][0], $aControls[$i][3])
	Next

	;Ist das Fenster nach dem Resizen auf dem alten Monitor? (hMonitor = neuer Monitor)
	;Beispiel:
	;	Es wird resized von 150% auf 100%.
	;	WM_DPICHANGED wurde gefeuert
	;	GUI skalier auf 100% (allerdings nach oben links hin weil X und Y Position ja beim Move erhalten bleiben)
	;	GUI ist vollständig auf 150% GUI und feuert wenn wir nichts unternehmen nochmal WM_DPICHANGED
	Local $hAfterSizeMonitor = _WinAPI_MonitorFromWindow($hGUI)

	If $hMonitor <> $hAfterSizeMonitor Then
		Local $aDesktopInfo = _GetDesktopInfoFromMonitor($hMonitor)

		If $iX <> Default and $iY <> Default Then
			If $iX < $aDesktopInfo[0] Then $iX = $aDesktopInfo[0]
			If $iY < $aDesktopInfo[1] Then $iY = $aDesktopInfo[1]

			_GUI_Move($hGUI, $iX, $iY, False, $iDPI / 96)
		EndIf
	EndIf

	GUISetState(@SW_UNLOCK, $hGUI)

	$aGUIInfos[3] = False
	$__o_GUIInfos.Item(String($hGUI)) = $aGUIInfos
EndFunc

#cs _GUI_SetResizing($hGUI, $iWidth, $iHeight, $aControls, $iCreatedDPI = _WinAPI_GetDpiForMonitor($H_PRIMARYMONITOR))

	Hier indexieren wir die Controls für das Resizing später.

	$hGUI = GUICreate-Handle
	$iWidth = Clientbreite bei 100%
	$iHeight = Clienthöhe bei 100%
	$aControls = Controls mit Resizeinformationen, entweder Format 1 oder 2 verwenden

		Format 1 (custom resizing):
			$aControls[n][0] = GUICtrlCreate-Handle
			$aControls[n][1] = GUICtrlSetResizing-Option ( $GUI_DOCKLEFT + $GUI_DOCKRIGHT + ... wie man halt will)

		Format 2 (autoscale resizing):
			$aControls[0] = GUICtrlCreate-Handle Anfang
			$aControls[1] = GUICtrlCreate-Handle Ende

			Da AutoIt intern die Ids fortlaufend hochzählt kann man einen kleinen Trick anwenden.
			Erzeugt man auf der GUI beispielsweise Label1, Label2, Button1, Button2, Button3

			So sind diese durchnummeriert und möchte man alle autoscalen kann man einfach
			Local $aControls[2] = [ $Label1, $Button3 ] übergeben.

			Das impliziert allerdings eine konsequent monoton steigende Controlreihenfolge.

	$iCreatedDPI = Die DPI mit der das Control erzeugt wurde. Dieser Wert wird für ListViews angepasst,
				   sobald das Fenster auf einen anderen Monitor mit andere DPI geschoben wurde.

				   Ansonsten nutzen es die anderen Controls um die Schrifthöhe zu normalisieren.
				   Es gelten Spezialfälle für TreeViews und ListViews, siehe näheres dazu in der _GUI_Resize.
#ce
Func _GUI_SetResizing($hGUI, $iWidth, $iHeight, $aControls, $iCreatedDPI = _WinAPI_GetDpiForMonitor($H_PRIMARYMONITOR))
	Local $aGUIInfos[4] = [ $iWidth, $iHeight, $iCreatedDPI, False ]
	Local $aControlInfo[0][6]

	If UBound($aControls, 2) Then
		ReDim $aControlInfo[UBound($aControls)][6]

		For $i = 0 To UBound($aControlInfo) - 1
			$aControlInfo[$i][0] = $aControls[$i][0]
			$aControlInfo[$i][1] = GUICtrlSendMsg($aControls[$i][0], $WM_GETFONT, 0, 0)
			$aControlInfo[$i][2] = $iCreatedDPI
			$aControlInfo[$i][3] = $aControls[$i][1]
			$aControlInfo[$i][4] = ControlGetPos($hGUI, "", $aControls[$i][0])
			$aControlInfo[$i][5] = Null
		Next
	Else
		ReDim $aControlInfo[$aControls[1] - $aControls[0] + 1][6]

		For $i = 0 To UBound($aControlInfo) - 1
			$aControlInfo[$i][0] = $i + $aControls[0]
			$aControlInfo[$i][1] = GUICtrlSendMsg($i + $aControls[0], $WM_GETFONT, 0, 0)
			$aControlInfo[$i][2] = $iCreatedDPI
			$aControlInfo[$i][3] = $GUI_DOCKAUTO
			$aControlInfo[$i][4] = ControlGetPos($hGUI, "", $i + $aControls[0])
			$aControlInfo[$i][5] = Null
		Next
	EndIf

	For $i = 0 To UBound($aControlInfo) - 1
		Switch _WinAPI_GetClassName($aControlInfo[$i][0])
			Case "SysListView32"
				Local $hWndHeader = GUICtrlSendMsg($aControlInfo[$i][0], $LVM_GETHEADER, 0, 0)
				Local $iCount = _SendMessage($hWndHeader, $HDM_GETITEMCOUNT, 0, 0)

				For $j = 0 To $iCount - 1
					Local $iColumnWidth = GUICtrlSendMsg($aControlInfo[$i][0], $LVM_GETCOLUMNWIDTH, $j, 0)
					Local $iNewWidth = $iColumnWidth * ($iCreatedDPI / 96)
					GUICtrlSendMsg($aControlInfo[$i][0], $LVM_SETCOLUMNWIDTH, $j, $iNewWidth)
				Next

				ContinueLoop
		EndSwitch
	Next

	$__o_GUIInfos.Item(String($hGUI)) = $aGUIInfos
	$__o_GUIControlInfos.Item(String($hGUI)) = $aControlInfo
EndFunc

#cs _GUI_GetBorderSize($hGUI)

	Wieso nutzen wir nicht SystemMetrics statt WinGetPos - WinGetClientSize?
		Da verschiedene Fenster verschiedene Styles haben, und dementsprechend auch andere Ränder
		ist es viel zu kompliziert das alles abzufragen, stattdessen subtrahieren wir hier die Werte einfach.

	Achtung: Beim WinMove nach dem WM_DPICHANGED ist Titelleiste immer noch groß.
			 Erst beim 2. WinMove wird die Titelleiste kleiner.
#ce
Func _GUI_GetBorderSize($hGUI)
	Local $aClientSize = WinGetClientSize($hGUI)
	Local $aSize = WinGetPos($hGUI)

	Local $aReturn[2] = [ $aSize[2] - $aClientSize[0], $aSize[3] - $aClientSize[1] ]
	Return $aReturn
EndFunc

#cs _GUI_Move($hGUI, $iX, $iY, $bIgnoreDPI = False, $fDPIScalefactor = _WinAPI_GetDpiForMonitor(_WinAPI_MonitorFromWindow($hGUI)) / 96)
	Diese Funktion dient dazu, die Fenster zu moven, dabei wird WinMove nochmal gewrappt, aber so das es nutzbarer ist.
	WinMove platziert die Fenster nämlich nach der Fenstergröße und nicht Clientgröße.

	Bei GUICreate gibt man aber die Clientgröße an und das irritiert sehr.
	Die Funktion movet das Fenster passend zur DPI (kann man auch ausschalten), so dass man nur die ursprüngliche Größe übergeben muss.
	Außerdem berechnet sie die Ränder mit, d.h. wenn ich folgendes Beispiel habe.

	$iX und $iY geben dabei die Koordinaten an, an welchem der Clientbereich anfangen soll. Einige Werte sind dabei speziell:

	Default = behalte X- respektive Y-Position bei
	-1 = Zentriere X- respektive Y-Achse
	-2 = schiebe es an die linke (X) respektive obere (Y) Seite des Monitors
	-3 = schiebe es an die rechte (X) respektive untere (Y) Seite des Monitors

	bIgnoreDPI wird verwendet um das Fenster (bevor es irgendwann hochskaliert wird) in die richtige Position zu bringen,
	damit man nicht WinMove nutzen und die Ränder selber berechnen muss.

	fDPIScalefactor wurde als Parameter aufgenommen, da die Fenster unterschiedlich skaliert werden können ab Win 8.1
	Das muss berücksichtigt werden, und deshalb gibts keinen globalen fDPIScalefactor mehr

	fDPIScalefactor = Default: Nehme die DPI des Monitors auf dem das Fenster ist
						 Wert: Skaliere das Fenster um den Faktor

					  Achtung: Der Skalierungsfaktor ist notwendig, da beim WM_DPICHANGED die DPI-Werte komisch sind.
							   Deswegen geben wir ab und zu den Skalierungsfaktor manuell an.
							   Wäre das nicht so hätten wir ruhig den Skalierungsfaktor komplett weglassen können.

	hMonitor = Default: Nehme den Monitor auf dem das Fenster platziert ist
					-1:	Nehme den Primärmonitor
				  hWnd: Nehme den angegeben Monitor

			   Achtung: Überschreibt fDPIScalefactor in jedem Fall mit der Skalierungsfaktor des Monitors,
						da ansonsten nicht richtig zentriert werden kann.

	hMonitor gibt an auf welchem Fenster die speziellen X- und Y-Parameter angewandt werden sollen.
	Die X- und Y-Werte spannen alle Desktop zusammen auf, deshalb gilt $hMonitor nur bei -1, -2 und -3.
#ce
Func _GUI_Move($hGUI, $iX, $iY, $bIgnoreDPI = False, $fDPIScalefactor = Default, $hMonitor = Default)
	Local $aGUIInfos = $__o_GUIInfos.Item(String($hGUI))
	Local $aBorders = _GUI_GetBorderSize($hGUI)
	Local $iWidth = $aGUIInfos[0]
	Local $iHeight = $aGUIInfos[1]

	Local $aDesktopInfo

	;Die Desktopgrößen holen auf dem das Fenster platziert werden soll
	If $hMonitor = Default Then
		If $fDPIScalefactor = Default Then $fDPIScalefactor = _WinAPI_GetDpiForMonitor(_WinAPI_MonitorFromWindow($hGUI)) / 96
		$aDesktopInfo = _GetDesktopInfoFromGUI($hGUI)
	Else
		$aDesktopInfo = _GetDesktopInfoFromMonitor(($hMonitor = -1) ? $H_PRIMARYMONITOR : $hMonitor)
		$fDPIScalefactor = ($hMonitor = -1) ? ($I_PRIMARYMONITOR_DPI / 96) : (_WinAPI_GetDpiForMonitor($hMonitor) / 96)
	EndIf

	; [0] = X-Offset des Desktops
	; [1] = Y-Offset des Desktops
	; [2] = Breite des Desktops
	; [3] = Höhe

	$w = $iWidth
	$h = $iHeight

	Switch $iX
		Case -1
			$iX = $aDesktopInfo[0] + ($aDesktopInfo[2] - $iWidth * ($bIgnoreDPI ? 1 : $fDPIScalefactor) - $aBorders[0]) / 2

		Case -2
			$iX = $aDesktopInfo[0]

		Case -3
			$iX = Ceiling($aDesktopInfo[0] + $aDesktopInfo[2] - $iWidth * ($bIgnoreDPI ? 1 : $fDPIScalefactor) - $aBorders[0])

	EndSwitch

	Switch $iY
		Case -1
			$iY = $aDesktopInfo[1] + ($aDesktopInfo[3] - $iHeight * ($bIgnoreDPI ? 1 : $fDPIScalefactor) - $aBorders[1]) / 2

		Case -2
			$iY = $aDesktopInfo[1]

		Case -3
			$iY = Ceiling($aDesktopInfo[1] + $aDesktopInfo[3] - $iHeight * ($bIgnoreDPI ? 1 : $fDPIScalefactor) - $aBorders[1])

	EndSwitch

	If Not $bIgnoreDPI Then
		$iWidth = $iWidth * $fDPIScalefactor
		$iHeight = $iHeight * $fDPIScalefactor
	EndIf

	$iWidth += $aBorders[0]
	$iHeight += $aBorders[1]

	Return WinMove($hGUI, "", $iX, $iY, $iWidth, $iHeight)
EndFunc

Func _GetDesktopInfoFromMonitor($hMonitor)
	Local $aMonitorInfo = _WinAPI_GetMonitorInfo($hMonitor)

	Local $aReturn[4] = [ _
		DllStructGetData($aMonitorInfo[1], "Left"), _
		DllStructGetData($aMonitorInfo[1], "Top"), _
		DllStructGetData($aMonitorInfo[1], "Right") - DllStructGetData($aMonitorInfo[1], "Left"), _
		DllStructGetData($aMonitorInfo[1], "Bottom") - DllStructGetData($aMonitorInfo[1], "Top") _
	]

	Return $aReturn
EndFunc

Func _GetDesktopInfoFromGUI($hGUI)
	Local $aWinPos = WinGetPos($hGUI)

	Local $tRect = DllStructCreate($tagRECT)
	DllStructSetData($tRect, "Left", $aWinPos[0])
	DllStructSetData($tRect, "Top", $aWinPos[1])
	DllStructSetData($tRect, "Right", $aWinPos[0] + $aWinPos[2])
	DllStructSetData($tRect, "Bottom", $aWinPos[1] + $aWinPos[3])

	Local $hMonitor = _WinAPI_MonitorFromRect($tRect)
	Local $aMonitorInfo = _WinAPI_GetMonitorInfo($hMonitor)

	Local $iDesktopXOffset = DllStructGetData($aMonitorInfo[1], "Left")
	Local $iDesktopYOffset = DllStructGetData($aMonitorInfo[1], "Top")

	Local $iDesktopWidth = DllStructGetData($aMonitorInfo[1], "Right") - DllStructGetData($aMonitorInfo[1], "Left")
	Local $iDesktopHeight = DllStructGetData($aMonitorInfo[1], "Bottom") - DllStructGetData($aMonitorInfo[1], "Top")

	Local $aReturn = [ $iDesktopXOffset, $iDesktopYOffset, $iDesktopWidth, $iDesktopHeight ]
	Return $aReturn
EndFunc

Func _WinAPI_SetProcessDpiAwarenessContext($iDpiAwareness)
    Return DllCall("user32.dll", "bool", "SetProcessDpiAwarenessContext", "int", $iDpiAwareness)
EndFunc

Func _WinAPI_GetDpiForMonitor($hMonitor = $H_PRIMARYMONITOR)
	If @OSBuild < 9600 Then Return RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "AppliedDPI")

	Local Const $MDT_DEFAULT = 0
    Local $aMonitors = _WinAPI_EnumDisplayMonitors()
    Local $tAxis = DllStructCreate("uint;uint;")

    Local $aReturn = DllCall("shcore.dll", _
		"long", "GetDpiForMonitor", _
			"long", $hMonitor, _
			"int", $MDT_DEFAULT, _
			"uint*", DllStructGetPtr($tAxis, 1), _
			"uint*", DllStructGetPtr($tAxis, 2) _
	)

	Return $aReturn[3]
EndFunc
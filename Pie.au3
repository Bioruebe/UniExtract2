#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.12.0
 Author:         WideBoyDixon, Bioruebe

 Script Function:
	Draw and animate 3D pie charts

	Based on the example script from WideBoyDixon:
	http://www.autoitscript.com/forum/index.php?showtopic=97241

#ce ----------------------------------------------------------------------------;

#include-once
#include <Array.au3>
#include <GDIPlus.au3>
#include <GUIConstants.au3>

Global Const $__PI = ATan(1) * 4

#Region Example
#cs
; Controls the size of the pie and also the depth
Const $PIE_DIAMETER = 400
Const $PIE_AREA = $PIE_DIAMETER + 2 * $PIE_DIAMETER * 0.025
Const $LERP_BY = 0.15

; Create random values
Const $NUM_VALUES = 8
Local $aChartValues[$NUM_VALUES][2]
For $i = 0 To $NUM_VALUES - 1
    $aChartValues[$i][0] = Random(5, 25, 1)
	$aChartValues[$i][1] = "Caption " & $i + 1
Next

; Create the GUI
$hWnd = GUICreate("Pie Chart", $PIE_AREA, $PIE_AREA + 100, Default, Default)
$idPic = GUICtrlCreatePic("", 0, 0, $PIE_AREA, $PIE_AREA)
$idLegend = GUICtrlCreatePic("", 0, $PIE_AREA, $PIE_AREA, 100)
GUISetState()

; Prepare values for the pie chart and setup GDI+ for both picture controls
$aHandles = _Pie_PrepareValues($aChartValues, $idPic)
$aLegendHandles = _Pie_CreateContext($idLegend, 0)

_ArraySort($aChartValues)
;~ _ArrayDisplay($aChartValues)

; Draw the initial pie chart and legend
_Pie_Draw($aHandles, $aChartValues, 1, 0)
_Pie_Draw_Legend($aLegendHandles, $aChartValues)

; Rotate pie if mouse is over control
Local $rot = 0, $asp = 1
While GUIGetMsg() <> $GUI_EVENT_CLOSE
    Sleep(10)

	_MouseOverRotation($rot, $asp, $aHandles, $aChartValues, $hWnd, $aHandles[4])
WEnd

; Cleanup
_Pie_Shutdown($aHandles, $aChartValues, False)
_Pie_Shutdown($aLegendHandles)
#ce
#EndRegion

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetDarkerColour
; Description ...: Get a darker version of a colour by extracting the RGB components
; Syntax ........: _GetDarkerColour($Colour)
; Parameters ....: $Colour              - the base color.
; Return values .: The new color
; Author ........: WideBoyDixon
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/97241-3d-pie-chart/
; Example .......: No
; ===============================================================================================================================
Func _GetDarkerColour($Colour)
    Local $Red, $Green, $Blue
    $Red = (BitAND($Colour, 0xff0000) / 0x10000) - 40
    $Green = (BitAND($Colour, 0x00ff00) / 0x100) - 40
    $Blue = (BitAND($Colour, 0x0000ff)) - 40
    If $Red < 0 Then $Red = 0
    If $Green < 0 Then $Green = 0
    If $Blue < 0 Then $Blue = 0
    Return ($Red * 0x10000) + ($Green * 0x100) + $Blue
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Pie_Draw
; Description ...: Draw the pie chart
; Syntax ........: _Pie_Draw($aHandles, $aValues, $nAspect, $nRotation[, $bPersistent = True])
; Parameters ....: $aHandles            - array of handles as returned by _Pie_CreateContext()
;                  $aValues             - array of values to draw as returned by _Pie_PrepareValues
;                  $nAspect             - aspect.
;                  $nRotation           - rotation.
;                  $bPersistent         - [optional] execute _Pie_Make_Persistent after drawing? Default is True.
; Return values .: None
; Author ........: WideBoyDixon
; Modified ......: Bioruebe
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/97241-3d-pie-chart/
; Example .......: Yes
; ===============================================================================================================================
Func _Pie_Draw($aHandles, $aValues, $nAspect, $nRotation, $bPersistent = True)
    Local $nCount, $nTotal = 0, $angleStart, $angleSweep, $X, $Y, $iSize = UBound($aValues)
    Local $pieArea = $aHandles[7] < $aHandles[8]? $aHandles[7]: $aHandles[8], $pieDiameter = (20 * $pieArea) / 21
	Local $pieLeft = $pieDiameter * 0.025, $pieTop = $pieArea / 2 - ($pieDiameter / 2) * $nAspect
	Local $pieHeight = $pieDiameter * $nAspect, $hPath, $pieDepth = $pieDiameter * 0.2

	If $nRotation > 360 Then $nRotation = Mod($nRotation, 360)

    _GDIPlus_GraphicsClear($aHandles[2], 0xFFFFFFFF)

	; Set the initial angles based on the fractional values
    Local $Angles[$iSize + 1]
    For $nCount = 0 To $iSize
        If $nCount = 0 Then
            $Angles[$nCount] = $nRotation
        Else
            $Angles[$nCount] = $Angles[$nCount - 1] + ($aValues[$nCount - 1][3] * 360)
        EndIf
    Next

    ; Adjust the angles based on the aspect
	For $nCount = 0 To $iSize
		$X = $pieDiameter * Cos($Angles[$nCount] * $__PI / 180)
		$Y = $pieDiameter * Sin($Angles[$nCount] * $__PI / 180)
		$Y -= ($pieDiameter - $pieHeight) * Sin($Angles[$nCount] * $__PI / 180)
		If $X = 0 Then
			$Angles[$nCount] = 90 + ($Y < 0) * 180
		Else
			$Angles[$nCount] = ATan($Y / $X) * 180 / $__PI
		EndIf
		If $X < 0 Then $Angles[$nCount] += 180
		If $X >= 0 And $Y < 0 Then $Angles[$nCount] += 360
		$X = $pieDiameter * Cos($Angles[$nCount] * $__PI / 180)
		$Y = $pieHeight * Sin($Angles[$nCount] * $__PI / 180)
	Next

    ; Decide which pieces to draw first and last
	Local $nStart = -1, $nEnd = -1
	For $nCount = 0 To $iSize - 1
		$angleStart = Mod($Angles[$nCount], 360)
		$angleSweep = Mod($Angles[$nCount + 1] - $Angles[$nCount] + 360, 360)
		If $angleStart <= 270 And ($angleStart + $angleSweep) >= 270 Then $nStart = $nCount
		If ($angleStart <= 90 And ($angleStart + $angleSweep) >= 90) Or ($angleStart <= 450 And ($angleStart + $angleSweep) >= 450) Then $nEnd = $nCount
		If $nEnd >= 0 And $nStart >= 0 Then ExitLoop
	Next

    ; Draw the first piece
	_Pie_DrawPiece($aHandles[2], $aValues, $pieLeft, $pieTop, $pieDiameter, $pieHeight, $pieDepth * (1 - $nAspect), $nStart, $Angles)

    ; Draw pieces "to the right"
	$nCount = Mod($nStart + 1, $iSize)
	While $nCount <> $nEnd
		_Pie_DrawPiece($aHandles[2], $aValues, $pieLeft, $pieTop, $pieDiameter, $pieHeight, $pieDepth * (1 - $nAspect), $nCount, $Angles)
		$nCount = Mod($nCount + 1, $iSize)
	WEnd

    ; Draw pieces "to the left"
	$nCount = Mod($nStart + $iSize - 1, $iSize)
	While $nCount <> $nEnd
		_Pie_DrawPiece($aHandles[2], $aValues, $pieLeft, $pieTop, $pieDiameter, $pieHeight, $pieDepth * (1 - $nAspect), $nCount, $Angles)
		$nCount = Mod($nCount + $iSize - 1, $iSize)
	WEnd

    ; Draw the last piece
	_Pie_DrawPiece($aHandles[2], $aValues, $pieLeft, $pieTop, $pieDiameter, $pieHeight, $pieDepth * (1 - $nAspect), $nEnd, $Angles)

	; Now draw the bitmap on to the device context of the window
    _GDIPlus_GraphicsDrawImage($aHandles[0], $aHandles[1], 0, 0)

	If $bPersistent Then _Pie_Make_Persistent($aHandles)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Pie_DrawPiece
; Description ...: Draw a single piece of the pie chart
; Syntax ........: _Pie_DrawPiece($hGraphics, $aValues, $iX, $iY, $iWidth, $iHeight, $iDepth, $nCount, $Angles)
; Parameters ....: $hGraphics           - handle to the graphics object.
;                  $aValues             - array of values to draw as returned by _Pie_PrepareValues
;                  $iX                  - x position.
;                  $iY                  - y position.
;                  $iWidth              - width.
;                  $iHeight             - height.
;                  $iDepth              - depth.
;                  $nCount              - piece count.
;                  $Angles              - angles.
; Return values .: None
; Author ........: WideBoyDixon
; Modified ......: Bioruebe
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/97241-3d-pie-chart/
; Example .......: No
; ===============================================================================================================================
Func _Pie_DrawPiece($hGraphics, $aValues, $iX, $iY, $iWidth, $iHeight, $iDepth, $nCount, $Angles)
    Local $hPath, $cX = $iX + ($iWidth / 2), $cY = $iY + ($iHeight / 2), $fDrawn = False
    Local $iStart = Mod($Angles[$nCount], 360), $iSweep = Mod($Angles[$nCount + 1] - $Angles[$nCount] + 360, 360)

	; Draw side
    $hPath = _GDIPlus_PathCreate()
    If $iStart < 180 And ($iStart + $iSweep > 180) Then
        _GDIPlus_PathAddArc($hPath, $iX, $iY, $iWidth, $iHeight, $iStart, 180 - $iStart)
        _GDIPlus_PathAddArc($hPath, $iX, $iY + $iDepth, $iWidth, $iHeight, 180, $iStart - 180)
        _GDIPlus_PathCloseFigure($hPath)
        _GDIPlus_GraphicsFillPath($hGraphics, $hPath, $aValues[$nCount][5])
        _GDIPlus_GraphicsDrawPath($hGraphics, $hPath, $aValues[$nCount][6])
        $fDrawn = True
    EndIf
    If $iStart + $iSweep > 360 Then
        _GDIPlus_PathAddArc($hPath, $iX, $iY, $iWidth, $iHeight, 0, $iStart + $iSweep - 360)
        _GDIPlus_PathAddArc($hPath, $iX, $iY + $iDepth, $iWidth, $iHeight, $iStart + $iSweep - 360, 360 - $iStart - $iSweep)
        _GDIPlus_PathCloseFigure($hPath)
        _GDIPlus_GraphicsFillPath($hGraphics, $hPath, $aValues[$nCount][5])
        _GDIPlus_GraphicsDrawPath($hGraphics, $hPath, $aValues[$nCount][6])
        $fDrawn = True
    EndIf
    If $iStart < 180 And (Not $fDrawn) Then
        _GDIPlus_PathAddArc($hPath, $iX, $iY, $iWidth, $iHeight, $iStart, $iSweep)
        _GDIPlus_PathAddArc($hPath, $iX, $iY + $iDepth, $iWidth, $iHeight, $iStart + $iSweep, -$iSweep)
        _GDIPlus_PathCloseFigure($hPath)
        _GDIPlus_GraphicsFillPath($hGraphics, $hPath, $aValues[$nCount][5])
        _GDIPlus_GraphicsDrawPath($hGraphics, $hPath, $aValues[$nCount][6])
    EndIf
    _GDIPlus_PathDispose($hPath)

	; Draw top
    _GDIPlus_GraphicsFillPie($hGraphics, $iX, $iY, $iWidth, $iHeight, $iStart, $iSweep, $aValues[$nCount][4])
    _GDIPlus_GraphicsDrawPie($hGraphics, $iX, $iY, $iWidth, $iHeight, $iStart, $iSweep, $aValues[$nCount][6])
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Pie_PrepareValues
; Description ...: Prepare the values to draw
; Syntax ........: _Pie_PrepareValues(Byref $aValues[, $hControl = 0])
; Parameters ....: $aValues             - [in/out] the values to draw as a 2-dimensional array:
;											[$i][0] - absolute values
;											[$i][1] - captions
;											[$i][2] - [optional] a third dimension might specify the colors to draw the pie pieces
;                  $hControl            - [optional] the handle to the control. This is a convenience argument to call
;													_Pie_CreateContext and return it's result
; Return values .: Either the return of _Pie_CreateContext (if $hControl is specified) or 0
; Author ........: Bioruebe
; Modified ......:
; Remarks .......: Based on the example script by WideBoyDixon (see Link)
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/97241-3d-pie-chart/
; Example .......: No
; ===============================================================================================================================
Func _Pie_PrepareValues(ByRef $aValues, $hControl = 0)
	Local $iSize = UBound($aValues), $nTotal = 0, $bSetColors = False
	If UBound($aValues, 2) < 3 Then $bSetColors = True
	ReDim $aValues[$iSize][7]

	; Calculate percentage values
	For $i = 0 To $iSize - 1
		$nTotal += $aValues[$i][0]
	Next

	_GDIPlus_Startup()

	For $i = 0 To $iSize - 1
		; Set the fractional values
		$aValues[$i][3] = $aValues[$i][0] / $nTotal

		; Create the brushes and pens
		If $bSetColors Then $aValues[$i][2] = (Random(0, 255, 1) * 0x10000) + (Random(0, 255, 1) * 0x100) + Random(0, 255, 1)

		$aValues[$i][4] = _GDIPlus_BrushCreateSolid(BitOR(0xff000000, $aValues[$i][2]))
		$aValues[$i][5] = _GDIPlus_BrushCreateSolid(BitOR(0xff000000, _GetDarkerColour($aValues[$i][2])))
		$aValues[$i][6] = _GDIPlus_PenCreate(BitOR(0xff000000, _GetDarkerColour(_GetDarkerColour($aValues[$i][2]))))
	Next

	If $hControl == 0 Then Return 0
	Return _Pie_CreateContext($hControl)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Pie_CreateContext
; Description ...: Set up GDI+
; Syntax ........: _Pie_CreateContext($idControl[, $iSmoothing = 2])
; Parameters ....: $idControl           - ControlID as returned by GuiCtrlCreatePic().
;                  $iSmoothing          - [optional] smoothing mode. Default is 2.
; Return values .: An array of handles: [0] Handle as returned from _GDIPlus_GraphicsCreateFromHWND
;										[1] Handle as returned from _GDIPlus_BitmapCreateFromGraphics
;										[2] Handle as returned from _GDIPlus_ImageGetGraphicsContext
;										[3] The control ID passed to this function
;										[4] HWND for the control
;										[5] X position of the control, absolute value
;										[6] Y position of the control, absolute value
;										[7] Width of the control, absolute value
;										[8] Height of the control, absolute value
; Author ........: Bioruebe
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Pie_CreateContext($idControl, $iSmoothing = 2)
	Local $aHandles[5], $hControl

	_GDIPlus_Startup()

	$hControl = GUICtrlGetHandle($idControl)
	$aControlPos = WinGetPos($hControl)

	$aHandles[0] = _GDIPlus_GraphicsCreateFromHWND($hControl)
	$aHandles[1] = _GDIPlus_BitmapCreateFromGraphics($aControlPos[2], $aControlPos[3], $aHandles[0])
	$aHandles[2] = _GDIPlus_ImageGetGraphicsContext($aHandles[1])
	$aHandles[3] = $idControl
	$aHandles[4] = $hControl
	_ArrayAdd($aHandles, WinGetPos($hControl))

	_GDIPlus_GraphicsSetSmoothingMode($aHandles[2], $iSmoothing)

	Return $aHandles
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Pie_Draw_Legend
; Description ...: Draw the legend for the pie chart
; Syntax ........: _Pie_Draw_Legend($aHandles, $aValues[, $iBulletWidth = 20[, $iColumns = 3[, $iFontSize = 10]]])
; Parameters ....: $aHandles            - array of handles as returned by _Pie_CreateContext()
;                  $aValues             - array of values to draw as returned by _Pie_PrepareValues
;                  $iBulletWidth        - [optional] the width and height of each bullet. Default is 20.
;                  $iColumns            - [optional] the amount of columns to create. Default is 3.
;                  $iFontSize           - [optional] the font size to use. Default is 10.
; Return values .: None
; Author ........: Bioruebe
; Modified ......:
; Remarks .......: This function does not include an overflow check. Too many values in $aValues may result in invisible bullet points.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _Pie_Draw_Legend($aHandles, $aValues, $iBulletWidth = 20, $iColumns = 3, $iFontSize = 10)
	_GDIPlus_GraphicsClear($aHandles[2], 0xFFFFFFFF)

	Local $iSize = UBound($aValues), $iOffset = $aHandles[7] * 0.1, $iBetween = $iBulletWidth / 2, _
		  $iBetweenCols = ($aHandles[7] - $iOffset) / $iColumns, $iRows = Int(($iSize - 1) / $iColumns), _
		  $iRowSpace = $aHandles[8] / ($iRows + 1), $iLabelBulletOffset = ($iBulletWidth - $iFontSize) / 3, $iX, $iY
	Local $hFontLegend = _GDIPlus_FontCreate(_GDIPlus_FontFamilyCreate('Arial'), $iFontSize, '')
	Local $hFormat = _GDIPlus_StringFormatCreate(0x0020)
    Local $hTextBrush = _GDIPlus_BrushCreateSolid(0xFF000000)

	$i = 0
	For $k = 0 To $iRows
		$iY = $iRowSpace * $k
		For $j = 0 To $iColumns - 1
			$iX = $iOffset + $iBetweenCols * $j
			_GDIPlus_GraphicsFillRect($aHandles[2], $iX, $iY, $iBulletWidth, $iBulletWidth, $aValues[$i][4])
			$iLabelOffsetY = Int(StringLen($aValues[$i][1]) / 12) * $iFontSize * 0.8
			_GDIPlus_GraphicsDrawStringEx($aHandles[2], $aValues[$i][1], $hFontLegend, _GDIPlus_RectFCreate($iX + $iBetween + $iBulletWidth, $iY + $iLabelBulletOffset - $iLabelOffsetY, $iBetweenCols - $iBulletWidth - $iBetween, $iRowSpace), $hFormat, $hTextBrush)
			$i += 1
			If $i >= $iSize Then ExitLoop 2
		Next
	Next

	_GDIPlus_GraphicsDrawImage($aHandles[0], $aHandles[1], 0, 0)

	_Pie_Make_Persistent($aHandles)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Pie_Make_Persistent
; Description ...: Create HBITMAP and send to picture control to automatically redraw after GUI is minimized and restored
; Syntax ........: _Pie_Make_Persistent($aHandles)
; Parameters ....: $aHandles            - array of handles as returned by _Pie_CreateContext().
; Return values .: None
; Author ........: Bioruebe
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Pie_Make_Persistent($aHandles)
	Local $hHBITMAP = _GDIPlus_BitmapCreateHBITMAPFromBitmap($aHandles[1])
    _WinAPI_DeleteObject(GUICtrlSendMsg($aHandles[3], 0x0172, 0, $hHBITMAP))
    _WinAPI_DeleteObject($hHBITMAP)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Pie_Shutdown
; Description ...: Dispose GDI+ objects
; Syntax ........: _Pie_Shutdown($aHandles[, $aValues = 0[, $bShutdownGDIPlus = True]])
; Parameters ....: $aHandles            - array of handles as returned by _Pie_CreateContext()
;                  $aValues             - [optional] array of values as returned by _Pie_Prepare to dispose used brushes.
;                  $bShutdownGDIPlus    - [optional] call _GDIPlus_Shutdown after disposing objects? Default is True.
; Return values .: None
; Author ........: Bioruebe
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Pie_Shutdown($aHandles, $aValues = 0, $bShutdownGDIPlus = True)
	If IsArray($aValues) Then
		For $i = 0 To UBound($aValues) - 1
			_GDIPlus_BrushDispose($aValues[$i][4])
			_GDIPlus_BrushDispose($aValues[$i][5])
			_GDIPlus_PenDispose($aValues[$i][6])
		Next
	EndIf

	_GDIPlus_GraphicsDispose($aHandles[2])
	_GDIPlus_BitmapDispose($aHandles[1])
	_GDIPlus_GraphicsDispose($aHandles[0])

	If $bShutdownGDIPlus Then _GDIPlus_Shutdown()
EndFunc


; #FUNCTION# ====================================================================================================================
; Name ..........: _MouseIsOverControl
; Description ...: Returns true if the mouse is over a given control and the GUI is currently active
; Syntax ........: _MouseIsOverControl($hWnd, $hControlOrPosArray[, $iOffset = 0])
; Parameters ....: $hWnd                - the GUI handle.
;                  $hControlOrPosArray  - either a control handle or an array as returned by WinGetPos.
;                  $iOffset             - [optional] an offset to use if $hControlOrPosArray is an array. Default is 0.
; 										   e.g. _Pie_CreateContext() returns the WinGetPos() values beginning at index 5, so the index would be 5
; Return values .: True or False
; Author ........: Bioruebe
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _MouseIsOverControl($hWnd, $hControlOrPosArray, $iOffset = 0)
    If Not WinActive($hWnd) Then Return False
    Local $aMousePos = MouseGetPos()
    Local $aWinPos = IsArray($hControlOrPosArray)? $hControlOrPosArray: WinGetPos($hControlOrPosArray) ; Yes, WinGetPos returns valid control positions; ControlGetPos uses absolute positions
    If ($aMousePos[0] < $aWinPos[$iOffset] Or $aMousePos[0] > $aWinPos[$iOffset] + $aWinPos[$iOffset + 2]) Or _
	   ($aMousePos[1] < $aWinPos[$iOffset + 1] Or $aMousePos[1] > $aWinPos[$iOffset + 1] + $aWinPos[$iOffset + 3]) Then Return False
	Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _MouseOverRotation
; Description ...: Convenience function to create a ratation/aspect change effect if the mouse is over the control
; Syntax ........: _MouseOverRotation(Byref $rot, Byref $asp, $aHandles, $aValues, $hWnd, $hControlOrPosArray[, $iOffset = 0[,
;                  $iRotationSpeed = 0.4[, $iLerpBy = 0.15]]])
; Parameters ....: $rot                 - [in/out] rotation variable.
;                  $asp                 - [in/out] aspect variable.
;                  $aHandles            - array of handles as returned by _Pie_CreateContext()
;                  $aValues             - array of values to draw as returned by _Pie_PrepareValues
;                  $hWnd                - handle to the GUI.
;                  $hControlOrPosArray  - either a control handle or an array as returned by WinGetPos.
;                  $iOffset             - [optional] an offset to use if $hControlOrPosArray is an array. Default is 0.
;                  $iRotationSpeed      - [optional] speed of rotation. Default is 0.4.
;                  $iLerpBy             - [optional] interpolation speed. Default is 0.15.
; Return values .: None
; Author ........: Bioruebe
; Modified ......:
; Remarks .......: This should be called in the GUIGetMsg() loop.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _MouseOverRotation(ByRef $rot, ByRef $asp, $aHandles, $aValues, $hWnd, $hControlOrPosArray, $iOffset = 0, $iRotationSpeed = 0.4, $iLerpBy = 0.15)
	If _MouseIsOverControl($hWnd, $hControlOrPosArray, $iOffset) Then
		$rot += $iRotationSpeed
		If $rot > 360 Then $rot = 0

		$asp = _Cerp($asp, 0.30, $iLerpBy)
	Else
		If $asp > 0.99 Then
			$asp = 1
			_Pie_Draw($aHandles, $aValues, $asp, $rot, True)
			Return
		EndIf
		$asp = _Cerp($asp, 1, $iLerpBy)
	EndIf

;~ 	ConsoleWrite("DRAW " & $hControlOrPosArray & " - " & $asp & @CRLF)
	_Pie_Draw($aHandles, $aValues, $asp, $rot, False)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Lerp
; Description ...: Linear interpolation between $a and $b by $t
; Syntax ........: _Lerp($a, $b, $t)
; Parameters ....: $a                   - a.
;                  $b                   - b.
;                  $t                   - t.
; Return values .: Interpolated value
; Author ........: Bioruebe
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Lerp($a, $b, $t)
	Return $a * (1 - $t) + $b * $t
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Cerp
; Description ...: Cosine interpolation between $a and $b by $t
; Syntax ........: _Cerp($a, $b, $t)
; Parameters ....: $a                   - a.
;                  $b                   - b.
;                  $t                   - t.
; Return values .: Interpolated value
; Author ........: Bioruebe
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _Cerp($a, $b, $t)
	Local $ft = $t * $__PI
	Local $f = (1 - Cos($ft)) / 2

	Return $a * (1 - $f) + $b * $f
EndFunc
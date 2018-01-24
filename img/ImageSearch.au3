#include-once
; ------------------------------------------------------------------------------
;
; AutoIt Version: 3.0
; Language:       English
; Description:    Functions that assist with Image Search
;                 Require that the ImageSearchDLL.dll be loadable
;
; ------------------------------------------------------------------------------

;===============================================================================
;
; Description:      Find the position of an image on the desktop
; Syntax:           _ImageSearchArea, _ImageSearch
; Parameter(s):
;                   $findImage - the image to locate on the desktop
;                   $tolerance - 0 for no tolerance (0-255). Needed when colors of
;                                image differ from desktop. e.g GIF
;                   $resultPosition - Set where the returned x,y location of the image is.
;                                     1 for centre of image, 0 for top left of image
;                   $x $y - Return the x and y location of the image
;
; Return Value(s):  On Success - Returns 1
;                   On Failure - Returns 0
;
; Note: Use _ImageSearch to search the entire desktop, _ImageSearchArea to specify
;       a desktop region to search
;
;===============================================================================
Global $aClientSize = WinGetClientSize("[CLASS:TibiaClient]")

Func _ImageSearch($findImage,$resultPosition,ByRef $x, ByRef $y,$tolerance)
   return _ImageSearchArea($findImage,$resultPosition,0,0,@DesktopWidth,@DesktopHeight,$x,$y,$tolerance)
;~ 	  return _ImageSearchArea($findImage,$resultPosition,0,0,$aClientSize[0],$aClientSize[1],$x,$y,$tolerance)
EndFunc

;~ teste()
Func teste()
	  Local $aClientSize = WinGetClientSize("[CLASS:TibiaClient]")
	  Run("notepad.exe")

	  Local $wi = String($aClientSize[0])
	  Local $he = String($aClientSize[1])
	  WinWaitActive("[CLASS:Notepad]")
	  Send("wid: "&$wi&" "&"he: "& $he &" ")
	  WinActivate("[CLASS:TibiaClient]")
	  AutoItSetOption("MouseCoordMode",2)
;~ 	  movemouse($wi,$he)
	  movemouse($wi -170,$he/2)

EndFunc



Func _ImageSearchLOOT($findImage,$resultPosition,ByRef $x, ByRef $y,$tolerance)

;~    Local $hWnd = WinWait("[CLASS:TibiaClient]")]
;~    Run("notepad.exe")
;~    Sleep(1000)
;~    Send("X "&$aClientSize[0])
;~    Send("Y "&$aClientSize[1])

   Local $x1 = $aClientSize[0]-170
   Local $y1 = 380

   return _ImageSearchArea($findImage,$resultPosition,$x1,$y1,$aClientSize[0],$aClientSize[1],$x,$y,$tolerance)
EndFunc

Func _ImageSearchArea($findImage,$resultPosition,$x1,$y1,$right,$bottom,ByRef $x, ByRef $y, $tolerance)
;~ 	MsgBox(0,"asd","" & $x1 & " " & $y1 & " " & $right & " " & $bottom)
	if $tolerance>0 then $findImage = "*" & $tolerance & " " & $findImage
	Local $result = DllCall("img\ImageSearchDLL.dll","str","ImageSearch","int",$x1,"int",$y1,"int",$right,"int",$bottom,"str",$findImage)
;~ If ($result <> Null) Then
;~    MsgBox(0,"","retor")
;~ EndIf
	; If error exit
    if $result[0]="0" then return 0

	; Otherwise get the x,y location of the match and the size of the image to
	; compute the centre of search
	$array = StringSplit($result[0],"|")

;~    _ArrayDisplay($array,"array")

   $x=Int(Number($array[2]))
   $y=Int(Number($array[3]))

   if $resultPosition=1 then
      $x=$x + Int(Number($array[4])/2)
      $y=$y + Int(Number($array[5])/2)
   endif
   return 1
EndFunc

;===============================================================================
;
; Description:      Wait for a specified number of seconds for an image to appear
;
; Syntax:           _WaitForImageSearch, _WaitForImagesSearch
; Parameter(s):
;					$waitSecs  - seconds to try and find the image
;                   $findImage - the image to locate on the desktop
;                   $tolerance - 0 for no tolerance (0-255). Needed when colors of
;                                image differ from desktop. e.g GIF
;                   $resultPosition - Set where the returned x,y location of the image is.
;                                     1 for centre of image, 0 for top left of image
;                   $x $y - Return the x and y location of the image
;
; Return Value(s):  On Success - Returns 1
;                   On Failure - Returns 0
;
;
;===============================================================================
Func _WaitForImageSearch($findImage,$waitSecs,$resultPosition,ByRef $x, ByRef $y,$tolerance)
	$waitSecs = $waitSecs * 1000
	$startTime=TimerInit()
	While TimerDiff($startTime) < $waitSecs
		sleep(100)
		$result=_ImageSearch($findImage,$resultPosition,$x, $y,$tolerance)
		if $result > 0 Then
			return 1
		EndIf
	WEnd
	return 0
EndFunc

;===============================================================================
;
; Description:      Wait for a specified number of seconds for any of a set of
;                   images to appear
;
; Syntax:           _WaitForImagesSearch
; Parameter(s):
;					$waitSecs  - seconds to try and find the image
;                   $findImage - the ARRAY of images to locate on the desktop
;                              - ARRAY[0] is set to the number of images to loop through
;								 ARRAY[1] is the first image
;                   $tolerance - 0 for no tolerance (0-255). Needed when colors of
;                                image differ from desktop. e.g GIF
;                   $resultPosition - Set where the returned x,y location of the image is.
;                                     1 for centre of image, 0 for top left of image
;                   $x $y - Return the x and y location of the image
;
; Return Value(s):  On Success - Returns the index of the successful find
;                   On Failure - Returns 0
;
;
;===============================================================================
Func _WaitForImagesSearch($findImage,$waitSecs,$resultPosition,ByRef $x, ByRef $y,$tolerance)
	$waitSecs = $waitSecs * 1000
	$startTime=TimerInit()
	While TimerDiff($startTime) < $waitSecs
		for $i = 1 to $findImage[0]
		    sleep(100)
		    $result=_ImageSearch($findImage[$i],$resultPosition,$x, $y,$tolerance)
		    if $result > 0 Then
			    return $i
		    EndIf
		Next
	WEnd
	return 0
EndFunc


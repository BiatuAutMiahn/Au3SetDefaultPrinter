#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Res\ohiohealth.ico
#AutoIt3Wrapper_Outfile_x64=..\SetDefaultPrinter (GUI).exe
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Array.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiComboBox.au3>
#include "Includes\PrintMgr.au3"

Local $aPrinterList = _PrintMgr_EnumPrinter()
If $aPrinterList[0]==0 Then
    MsgBox(16,"Error","No printers found")
    Exit 1
EndIf
#Region ### START Koda GUI section ### Form=
DllCall("User32.dll", "bool", "SetProcessDpiAwarenessContext" , "HWND", "DPI_AWARENESS_CONTEXT" -1)
$Form1 = GUICreate("SetDefaultPrinter", 256, 96+8)
GUISetFont(10, 400, 0, "Consolas")
$idComboBox = GUICtrlCreateCombo("", 8, 32, 256-16, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$sPtr=""
For $i = 1 to $aPrinterList[0]
    $sPtr&=$aPrinterList[$i]
    If $i<$aPrinterList[0] Then $sPtr&="|"
Next
GUICtrlSetData($idComboBox,$sPtr,$aPrinterList[1])
GUICtrlCreateLabel("Select a Printer:", 8, 8, 128, 17)
$idConfirm = GUICtrlCreateButton("Make Default", (256/2)+8, 64, 96, 32)
$idCancel1 = GUICtrlCreateButton("Cancel", (256/2)-96-8, 64, 96, 32)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit 2
        Case $idCancel1
            Exit 2
        Case $idConfirm
            GUICtrlRead($idComboBox)
            $iIdx=_GUICtrlComboBox_GetCurSel($idComboBox)+1
            _PrintMgr_SetDefaultPrinter($aPrinterList[$iIdx])
            If @error Then
                If @extended==1 Then
                    MsgBox(16,"Error","Failed to connect to WMI Namespace.")
                ElseIf @extended==2 Then
                    MsgBox(16,"Error","Error Querying Printer")
                ElseIf @extended==3 Then
                    MsgBox(16,"Error","Error setting default printer."&@LF&"Error: "&@error)
                EndIf
                Exit 1
            EndIf
            MsgBox(64,"Success",'"'&$aPrinterList[$iIdx]&'" is now the default printer.')
	EndSwitch
WEnd

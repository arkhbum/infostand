#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=kiosk.ico
#AutoIt3Wrapper_Outfile_x64=kiosk.exe
#AutoIt3Wrapper_Compression=3
#AutoIt3Wrapper_Res_Comment=kiosk.exe “Путь к файлу.pptx” 2 5 300, 2 – Номер слайда меню (по умолчанию – 2)
#AutoIt3Wrapper_Res_Description=Программа запуска презентации на инфостенде
#AutoIt3Wrapper_Res_Fileversion=0.2.0.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=Infostand presentation launcher
#AutoIt3Wrapper_Res_ProductVersion=1
#AutoIt3Wrapper_Res_CompanyName=АО АРХБУМ
#AutoIt3Wrapper_Res_LegalCopyright=АО АРХБУМ, 2023
#AutoIt3Wrapper_Res_LegalTradeMarks=http://rudotcom.github.io
#AutoIt3Wrapper_Res_Language=1049
#AutoIt3Wrapper_Tidy_Stop_OnError=n
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo /sf /sv
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <FileOperations.au3>
#include <FileConstants.au3>
$oMyError = ObjEvent("AutoIt.Error","MyErrFunc")    ; Initialize a COM error handler

Global $slideTimer
Global $slideNumber
Global $updateTimer
Global $idleShow
Global $menuSlide
Global $slideCount
Global $fileTimeStamp
Global $remoteFileName
Global $localFileName
Global $Presentation
Global $sourceDir
Global $PPT
Global $localFileTimeStamp

HotKeySet("^+!{-}", "Terminate")
HotKeySet("{Esc}", "RunSlideShow")

If $CmdLine[0] = 0 Or $CmdLine[0] > 4 Then
  MsgBox( 0, "Предупреждение", "Запуск программы с параметрами: " & @CRLF & @CRLF & "kiosk.exe ""Путь к файлу.pptx"" 2 5 300" & _
  @CRLF & @CRLF & "где:" & _
  @CRLF & "2 - номер слайда меню" & _
  @CRLF & "5 - частота смены слайдов - 5 сек" & _
  @CRLF & "300 - время до перехода в режим бездействия (автоматического переключения слайдов) - 300 сек")
  Exit
Endif

$remoteFileName = $CmdLine[1]
$menuSlide = $CmdLine[0] > 1 ? $CmdLine[2] : 2
$slideDelay = $CmdLine[0] > 2 ? $CmdLine[3] * 1000 : 5000
$idleTime = $CmdLine[0] > 3 ? $CmdLine[4] * 1000 : 300000

$updateDelay = 10000 ; частота проверки обновления презентации

$localFileName = @ScriptDir & '\~presentation.pptx'
$localFileTimeStamp = FileGetTime($localFileName, 0, 1)

CheckForUpdate()
StartChromeIfNotExists()
OpenPresentation()

$stillTimer = TimerInit()
$slideTimer = TimerInit()
$idleShow = True

While 1
	CheckForUpdate()

	$moved = False

	$pos1 = MouseGetPos()
	Sleep(50)
	$pos2 = MouseGetPos()

	If Abs($pos1[0] - $pos2[0]) Or Abs($pos1[1] - $pos2[1]) > 20 Then $moved = True

	If $idleShow Then
		If $moved Then
			$stillTimer = TimerInit()
			GoToMenu()
		Else
			RunSlideShow()
		EndIf
	EndIf

	If Not $idleShow And TimerDiff($stillTimer) > $idleTime And Not $moved Then
		RunSlideShow()
		CloseChrome()
	EndIf

	If WinActive( "[CLASS:Chrome_WidgetWin_1]" ) Then
		ToolTip( "Esc - вернуться к презентации" )
	EndIf

	StartChromeIfNotExists()

WEnd



; функции
Func OpenPresentation()
	ToolTip("Открываю презентацию.")
	$PPT = ObjCreate("PowerPoint.Application")
	If IsObj($PPT) Then
		$Presentation = $PPT.Presentations.Open($localFileName, False, False, False)
		If IsObj($Presentation) Then
			$slideCount = $Presentation.Slides.Count
		Else
			Return False
		EndIf
	EndIf
EndFunc   ;==>OpenPresentation

Func RunSlideShow()
	ToolTip("Слайд №" & $slideNumber & ' из ' & $slideCount)
	If TimerDiff($slideTimer) > $slideDelay Then
		$slideTimer = TimerInit()
		$slideNumber += 1
		If $slideNumber > $slideCount Then
			$slideNumber = 1
		EndIf
		If $slideNumber = $menuSlide Then
			$slideNumber += 1
		EndIf

		If IsObj($Presentation) And IsObj($Presentation.SlideShowSettings) Then
			With $Presentation.SlideShowSettings.Run.View
				.GotoSlide($slideNumber, False)
			EndWith
			WinActivate ( "[CLASS:screenClass]", "" )
		Else
			Local $pOpen = OpenPresentation()
			If IsObj($Presentation) Then
				RunSlideShow()
			EndIf
			ToolTip("Ожидание презентации")
			Return False
		EndIf
	EndIf
;~ 	$Presentation.Slides($menuSlide).SlideShowTransition.Hidden = True
	$idleShow = True
EndFunc   ;==>RunSlideShow

Func GoToMenu()
	If IsObj($Presentation) And IsObj($Presentation.SlideShowSettings) Then
		With $Presentation.SlideShowSettings.Run.View
			.GotoSlide($menuSlide, False)
		EndWith
		$idleShow = False
		ToolTip("Esc - вернуться к автоматической презентации")
	Else
			Local $pOpen = OpenPresentation()
			If $pOpen Then
				GoToMenu()
			EndIf
		ToolTip("Ожидание презентации")
	EndIf
EndFunc   ;==>GoToMenu

Func StartChromeIfNotExists()
	If Not ProcessExists("chrome.exe") Then
		Run("C:\Program Files\Google\Chrome\Application\chrome.exe -kiosk --disable-extensions")
	EndIf
EndFunc   ;==>StartChrome

Func CloseChrome()
	If WinExists('[CLASS:Chrome_WidgetWin_1]') Then
		$List = WinList('[CLASS:Chrome_WidgetWin_1]')
		For $i = 1 To $List[0][0]
			If BitAND(WinGetState($List[$i][1]), 2) Then
				WinClose($List[$i][1])
				WinWaitClose($List[$i][1])
			EndIf
		Next
	EndIf
EndFunc   ;==>CloseChrome

Func Terminate() ; по нажатию клавиши Escape
	If IsObj($PPT) Then
		ToolTip("Закрываю презентацию.")
		$PPT.Quit
	EndIf
	CloseChrome()
	Exit 0 ; завершение работы
EndFunc   ;==>Terminate

Func CheckForUpdate()
	$localFileTimeStamp = FileGetTime($localFileName, 0, 1)
	If TimerDiff($updateTimer) > $updateDelay Then
		$updateTimer = TimerInit()
		$remoteFileTimeStamp = FileGetTime($remoteFileName, 0, 1)
		If $localFileTimeStamp <> $remoteFileTimeStamp Then
			ToolTip("Презентация в режиме обновления у исполнителя. Ожидайте!")
			Sleep( 2000 )
			Local $hFileOpen = FileOpen($remoteFileName, 0)
			If $hFileOpen = -1 Then
				ToolTip("Пока не получилось получить обновление, повторяю попытку!")
				Sleep( 5000 )
				Return False
			EndIf
			FileClose ( $hFileOpen )
			ToolTip("Обновление презентации, подождите...")
			If IsObj($PPT) Then
				ToolTip("Закрываю презентацию.")
				$PPT.Quit
			EndIf
			FileCopy($remoteFileName, $localFileName, 1)
			OpenPresentation()
			RunSlideShow()
			$localFileTimeStamp = $remoteFileTimeStamp
		EndIf
	EndIf


EndFunc   ;==>CheckForUpdate

Func MyErrFunc()
	RunSlideShow()
	Return False
  Msgbox(48,"AutoItCOM Test","Произошла ошибка !"    & @CRLF  & @CRLF & _
             "err.description is: " & @TAB & $oMyError.description  & @CRLF & _
             "err.windescription:"   & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "       & @TAB & hex($oMyError.number,8)  & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @CRLF & _
             "err.source is: "       & @TAB & $oMyError.source       & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile     & @CRLF & _
             "err.helpcontext is: " & @TAB & $oMyError.helpcontext, 5)
Endfunc
